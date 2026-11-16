import { SaxesParser } from "saxes";
import type { DocumentReadLimits, DeepReadonly } from "../runtime.js";
import { DocumentError } from "../errors.js";
import { invalid, checkLimit } from "../document/validation.js";
import type { WordNativeContent } from "../contracts/word.js";

const COLLECTION = "urn:doffi:word:native-fragments",
  RELATIONSHIPS = new Set([
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
  ]);
type Fragment = DeepReadonly<WordNativeContent["fragments"][number]>;
interface References {
  relationships: Set<string>;
  identities: Set<string>;
}
interface Entry extends References {
  from: number;
  to: number;
}
interface Resource extends References {
  source: string;
  rootIdentities: Set<string>;
  entries?: Map<string, Entry>;
}
const identity = (value: string): string | undefined =>
  /^\{?[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}\}?$/i.test(value)
    ? value.replace(/[{}]/g, "").toUpperCase()
    : undefined;

/** Generic bounded XML/OPC envelope reader. It knows the doffi collection and
 * relationship attribute namespaces, never Word element/property semantics. */
export class NativeXmlResources {
  private nodes = 0;
  private readonly cache = new Map<string, Resource>();
  constructor(
    private readonly files: ReadonlyMap<string, Uint8Array>,
    private readonly limits: DocumentReadLimits,
  ) {}
  read(path: string): Resource {
    const cached = this.cache.get(path);
    if (cached) return cached;
    const bytes = this.files.get(path);
    if (!bytes) invalid("A native XML resource is missing.");
    checkLimit(bytes.length, this.limits.maxFileBytes);
    let source: string;
    try {
      const encoding =
        bytes[0] === 0xff && bytes[1] === 0xfe
          ? "utf-16le"
          : bytes[0] === 0xfe && bytes[1] === 0xff
            ? "utf-16be"
            : "utf-8";
      source = new (
        globalThis as unknown as {
          TextDecoder: new (
            encoding: string,
            options: { fatal: boolean },
          ) => { decode(bytes: Uint8Array): string };
        }
      ).TextDecoder(encoding, { fatal: true }).decode(bytes);
    } catch {
      return invalid("A native XML resource has invalid text encoding.");
    }
    const result: Resource = {
        source,
        relationships: new Set(),
        identities: new Set(),
        rootIdentities: new Set(),
      },
      parser = new SaxesParser({ xmlns: true }),
      stack: { entry: Entry | undefined; elements: number }[] = [];
    const count = (): void => {
      checkLimit(++this.nodes, this.limits.maxNodes);
    };
    parser.on("doctype", () => invalid("Native XML cannot contain a DTD."));
    parser.on("processinginstruction", () =>
      invalid("Native XML cannot contain processing instructions."),
    );
    parser.on("error", () => invalid("A native XML resource is malformed."));
    parser.on("comment", count);
    const text = (text: string): void => {
      if (!text) return;
      count();
      if (result.entries && stack.length <= 2 && text.trim())
        invalid("A native XML collection has unexpected text.");
    };
    parser.on("text", text);
    parser.on("cdata", text);
    parser.on("opentag", (tag) => {
      count();
      checkLimit(stack.length + 1, this.limits.maxDepth);
      let entry = stack.at(-1)?.entry;
      if (!stack.length && tag.uri === COLLECTION && tag.local === "fragments")
        result.entries = new Map();
      if (stack.length) stack.at(-1)!.elements++;
      if (result.entries && stack.length === 1) {
        const key = Object.values(tag.attributes).find((a) => !a.uri && a.local === "id")?.value;
        if (tag.uri !== COLLECTION || tag.local !== "item" || !key || result.entries.has(key))
          invalid("A native XML collection has an invalid or repeated entry.");
        entry = {
          from: source.lastIndexOf("<", parser.position - 1),
          to: -1,
          relationships: new Set(),
          identities: new Set(),
        };
        result.entries.set(key, entry);
      }
      for (const attribute of Object.values(tag.attributes)) {
        if (RELATIONSHIPS.has(attribute.uri)) {
          result.relationships.add(attribute.value);
          entry?.relationships.add(attribute.value);
        }
        const id = identity(attribute.value);
        if (id) {
          result.identities.add(id);
          entry?.identities.add(id);
          if (!stack.length) result.rootIdentities.add(id);
        }
      }
      stack.push({ entry, elements: 0 });
    });
    parser.on("closetag", () => {
      const current = stack.pop()!;
      if (result.entries && stack.length === 1) {
        if (current.elements !== 1)
          invalid("A native XML collection entry requires exactly one XML element.");
        current.entry!.to = parser.position;
      }
    });
    try {
      parser.write(source).close();
    } catch (error) {
      if (error instanceof DocumentError) throw error;
      return invalid("A native XML resource is malformed.");
    }
    this.cache.set(path, result);
    return result;
  }
  fragment(fragment: Fragment): References | undefined {
    const resource = this.read(fragment.assetPath);
    if (!fragment.fragmentKey) return resource;
    if (!resource.entries) return;
    const entry = resource.entries.get(fragment.fragmentKey);
    if (!entry) invalid("A native XML collection entry is missing.");
    return entry;
  }
  slice(path: string, keys: ReadonlySet<string>): string | undefined {
    const resource = this.read(path);
    if (!resource.entries) return;
    for (const key of keys)
      if (!resource.entries.has(key)) invalid("A native XML collection entry is missing.");
    if (keys.size === resource.entries.size) return;
    let at = 0,
      result = "";
    for (const [key, entry] of resource.entries) {
      if (keys.has(key)) continue;
      result += resource.source.slice(at, entry.from);
      at = entry.to;
    }
    return result + resource.source.slice(at);
  }
}

/** Reconstruct old generic OPC references and opaque UUID identity edges.
 * A referenced identity on a part root belongs to that part and its declared
 * incoming owner. Unknown envelope formats keep their conservative closure. */
export function indexLegacyNativeResources(
  native: WordNativeContent,
  previous: DeepReadonly<WordNativeContent> | undefined,
  xml: NativeXmlResources,
): { relationships: Set<string>; parts: Set<string> } {
  const former = { relationships: new Set<string>(), parts: new Set<string>() },
    all = [...(previous?.fragments ?? []), ...native.fragments],
    references = new Map<Fragment, References>(),
    wanted = new Set<string>(),
    currentFragments = new Map<Fragment, WordNativeContent["fragments"][number]>(
      native.fragments.map((f) => [f, f]),
    );
  for (const fragment of all) {
    if (
      fragment.fragmentFormat !== "json" &&
      (fragment.relationshipIds === undefined || fragment.dependencyParts === undefined)
    ) {
      const index = xml.fragment(fragment);
      if (index) {
        references.set(fragment, index);
        for (const id of index.identities) wanted.add(id);
      }
    }
  }
  const dependencies = new Map<string, Set<string>>(),
    declared = new Set(native.parts.map((p) => p.sourcePart));
  if (wanted.size)
    for (const part of native.parts) {
      if (!part.sourcePart.endsWith(".xml") && !/(?:\/|\+)xml(?:;|$)/i.test(part.contentType))
        continue;
      for (const id of xml.read(part.assetPath).rootIdentities) {
        if (!wanted.has(id)) continue;
        const roots = dependencies.get(id) ?? new Set<string>();
        const owners = native.relationships.filter(
          (r) => !r.external && r.target === part.sourcePart && declared.has(r.sourcePart),
        );
        for (const owner of owners) roots.add(owner.sourcePart);
        if (!owners.length) roots.add(part.sourcePart);
        dependencies.set(id, roots);
      }
    }
  for (const fragment of all) {
    const index = references.get(fragment),
      relationships = fragment.relationshipIds ?? (index ? [...index.relationships] : undefined),
      parts =
        fragment.dependencyParts ??
        (index
          ? [...new Set([...index.identities].flatMap((id) => [...(dependencies.get(id) ?? [])]))]
          : undefined);
    if (relationships)
      for (const id of relationships) former.relationships.add(fragment.sourcePart + "\n" + id);
    if (parts) for (const path of parts) former.parts.add(path);
    const current = currentFragments.get(fragment);
    if (current) {
      if (relationships) current.relationshipIds ??= [...relationships];
      if (parts) current.dependencyParts ??= [...parts];
    }
  }
  return former;
}

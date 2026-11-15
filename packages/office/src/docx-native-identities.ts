import { DocumentError } from "@doffi/core";
import type { DeepReadonly, WordMetadata } from "@doffi/core";
import { WP, A, PIC } from "./ooxml/xml.js";
import type { XmlNode } from "./ooxml/xml.js";
import { NativeWordIdentities } from "./docx-native-word-identities.js";
const VML = "urn:schemas-microsoft-com:vml",
  O = "urn:schemas-microsoft-com:office:office";
const MARKER = "urn:doffi:transient:native",
  WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
type Native = DeepReadonly<NonNullable<WordMetadata["nativeContent"]>>;
interface Origin {
  part: string;
  scope: string;
}
const kind = (node: XmlNode): "frame" | "object" | undefined =>
  node.uri === WP && node.name === "docPr"
    ? "frame"
    : node.name === "cNvPr" && [A, PIC, WPS, WPG].includes(node.uri)
      ? "object"
      : undefined;
/** Resolve native numeric drawing identities against the current compiler's IDs.
 * Shapes inside an opaque group use that group's object scope; Word frames and
 * cross-object endpoints use the original part's frame identity. */
export class NativeDrawingIdentities {
  private readonly word: NativeWordIdentities;
  private readonly origins = new WeakMap<XmlNode, Origin>();
  private readonly generated = new Map<string, string>();
  private readonly frames = new Map<string, string>();
  private readonly objects = new Map<string, string>();
  private readonly used = new Set<string>();
  private next = 1;
  private readonly vml = new Map<string, string>();
  private readonly usedVml = new Set<string>();
  private nextVml = 1;
  constructor(native: Native, roots: Iterable<XmlNode>) {
    const documents = [...roots];
    this.word = new NativeWordIdentities(documents, (node) => this.origins.get(node));
    const scan = (node: XmlNode): void => {
      const type = kind(node),
        id = node.attrs.get(" id"),
        owner = node.attrs.get(MARKER + " identity");
      if (node.uri === VML && node.attrs.has(" id")) this.usedVml.add(node.attrs.get(" id")!);
      if (type && id !== undefined) {
        this.used.add(id);
        if (owner) this.generated.set(owner + "\n" + type, id);
      }
      for (const child of node.content) if (typeof child !== "string") scan(child);
    };
    for (const root of documents) scan(root);
    for (const identity of native.identities ?? []) {
      const value = this.generated.get(identity.nodeId + "\n" + identity.kind);
      if (value !== undefined)
        (identity.kind === "frame" ? this.frames : this.objects).set(
          identity.sourcePart + "\n" + identity.sourceId,
          value,
        );
    }
  }
  mark(root: XmlNode, part: string, scope: string): void {
    const origin = { part, scope },
      scan = (node: XmlNode): void => {
        this.origins.set(node, origin);
        for (const child of node.content) if (typeof child !== "string") scan(child);
      };
    scan(root);
  }
  private allocate(): string {
    while (this.used.has(String(this.next))) this.next++;
    if (this.next > 4294967295)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "Native drawing identities exceed the UInt32 budget.",
      );
    const id = String(this.next++);
    this.used.add(id);
    return id;
  }
  rewrite(roots: Iterable<XmlNode>): void {
    const documents = [...roots];
    const definitions = (node: XmlNode): void => {
      const origin = this.origins.get(node),
        type = kind(node),
        id = node.attrs.get(" id");
      if (origin && node.uri === VML && id) {
        const key = origin.part + "\n" + id;
        let value = this.vml.get(key);
        if (!value) {
          const prefix = node.name === "shapetype" ? "_x0000_t" : "_x0000_s";
          do {
            value = prefix + this.nextVml++;
          } while (this.usedVml.has(value));
          this.usedVml.add(value);
          this.vml.set(key, value);
        }
        const alias = node.attrs.get(O + " spid");
        if (alias) this.vml.set(origin.part + "\n" + alias, value);
        (node.attrs as Map<string, string>).set(" id", value);
        if (alias) (node.attrs as Map<string, string>).set(O + " spid", value);
      }
      if (origin && type && id !== undefined && (type === "frame" || id !== "0")) {
        const key = origin.part + "\n" + (type === "object" ? origin.scope + "\n" : "") + id,
          map = type === "frame" ? this.frames : this.objects;
        let value = map.get(key);
        if (value === undefined) {
          value = this.allocate();
          map.set(key, value);
        }
        (node.attrs as Map<string, string>).set(" id", value);
      }
      for (const child of node.content) if (typeof child !== "string") definitions(child);
    };
    for (const root of documents) definitions(root);
    const references = (node: XmlNode, group = false): void => {
      group ||= node.uri === WPG && ["wgp", "grpSp"].includes(node.name);
      const origin = this.origins.get(node),
        id = node.attrs.get(" id");
      if (origin) {
        for (const [key, value] of node.attrs) {
          if (node.uri === VML && key === " type" && value.startsWith("#")) {
            const target = this.vml.get(origin.part + "\n" + value.slice(1));
            if (target) (node.attrs as Map<string, string>).set(key, "#" + target);
          }
          if (node.uri === O && node.name === "OLEObject" && key === " ShapeID") {
            const target = this.vml.get(origin.part + "\n" + value);
            if (target) (node.attrs as Map<string, string>).set(key, target);
          }
        }
      }
      if (origin && node.uri === A && ["stCxn", "endCxn"].includes(node.name) && id !== undefined) {
        const local = this.objects.get(origin.part + "\n" + origin.scope + "\n" + id),
          frame = this.frames.get(origin.part + "\n" + id),
          object = this.objects.get(origin.part + "\n" + id),
          target = group ? (local ?? object ?? frame) : (frame ?? local ?? object);
        if (target !== undefined) (node.attrs as Map<string, string>).set(" id", target);
      }
      for (const child of node.content) if (typeof child !== "string") references(child, group);
    };
    for (const root of documents) references(root);
    this.word.rewrite(documents);
  }
}

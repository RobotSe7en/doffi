import { createAssetDescriptor, DocumentError } from "@doffi/core";
import type { Annotation, ContentDocumentManifest, PresentationMetadata } from "@doffi/core";
import { openOpc } from "../ooxml/package.js";
import type { OpcPackage } from "../ooxml/package.js";
import { fail, R, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { declarations } from "../ooxml/styles.js";
import type { Css } from "../ooxml/styles.js";
import type { Html } from "../drawingml/html.js";
import { rootName } from "../drawingml/values.js";
import { serializeXml } from "../ooxml/serialize.js";
import { emfPreview } from "./emf-preview.js";
export * from "../drawingml/html.js";
export * from "../drawingml/values.js";

export interface Surface { id: string; root: Html; source: XmlNode; themeId: string; masterId?: string; layoutId?: string }
export interface NativeParagraph { nodeId: string; start: number; text: string; nativeText: string }
export class Context {
  serial = 0;
  work = 0;
  outputBytes = 0;
  readonly css: string[] = [];
  readonly unavailableThemeStyles: Set<string> = new Set();
  readonly assets: ContentDocumentManifest["assets"] = [];
  readonly files: Map<string, Uint8Array> = new Map();
  readonly surfaces: Map<string, Surface> = new Map();
  readonly nativeNodes: Map<string, Map<string, string>> = new Map();
  readonly slidePaths: Map<string, string> = new Map();
  readonly annotations: Annotation[] = [];
  readonly paragraphs: Map<string, NativeParagraph[]> = new Map();
  readonly workbooks: Map<string, OpcPackage> = new Map();
  readonly dataSetsByRange: Map<string, string> = new Map();
  nestedBytes: number;
  nestedFiles: number;
  readonly metadata: PresentationMetadata = { size: { unit: "pt", width: 960, height: 540 }, commentFormat: "modern",
    slides: [], themes: [], masters: [], layouts: [], notesMasters: [], stories: [], objects: [], charts: [], dataSets: [], comments: [], hyperlinks: [] };
  constructor(readonly opc: OpcPackage, readonly compatiblePreview = false) { this.nestedBytes = [...opc.files.values()].reduce((n, bytes) => n + bytes.length, 0); this.nestedFiles = opc.files.size; }
  retainPreview(feature: Exclude<NonNullable<PresentationMetadata["nativePreview"]>[number]["feature"], "emf-image">, ownerId: string, source: XmlNode, message: string): void {
    if (!this.compatiblePreview) unsupported(source, message);
    const visit = (node: XmlNode): void => {
      this.count();
      if ([...node.attrs.keys()].some(key => key.startsWith(R + " ")))
        unsupported(node, "A preserved PPTX preview fragment cannot have package relationships.");
      for (const item of node.content) if (typeof item !== "string") visit(item);
    };
    visit(source);
    const xml = serializeXml(source);
    this.bytes(xml.length * 6);
    (this.metadata.nativePreview ??= []).push({ feature, ownerId, part: source.part, xml });
  }
  workbook(path: string): OpcPackage {
    const cached = this.workbooks.get(path);
    if (cached) return cached;
    const limits = { ...this.opc.limits, maxPackageBytes: this.opc.limits.maxPackageBytes - this.nestedBytes, maxFiles: this.opc.limits.maxFiles - this.nestedFiles };
    if (limits.maxPackageBytes <= 0 || limits.maxFiles <= 0) throw new DocumentError("ResourceLimitExceeded", "Nested chart packages exceed the shared budget.");
    const book = openOpc(this.opc.files.get(path)!, { limits }, this.opc.xmlBudget);
    this.nestedBytes += [...book.files.values()].reduce((n, data) => n + data.length, 0); this.nestedFiles += book.files.size;
    this.workbooks.set(path, book); return book;
  }
  count(amount = 1): void {
    this.work += amount;
    if (this.work > this.opc.limits.maxNodes) throw new DocumentError("ResourceLimitExceeded", "PPTX mapping exceeds its work budget.");
  }
  bytes(amount: number): void {
    this.outputBytes += amount;
    if (this.outputBytes > this.opc.limits.maxPackageBytes) throw new DocumentError("ResourceLimitExceeded", "PPTX mapped output exceeds its byte budget.");
  }
  make(tag: string, attrs: Record<string, string> = {}, children: (Html | string)[] = []): Html {
    this.count();
    this.bytes(tag.length * 2 + 64 + Object.entries(attrs).reduce((n, [k, v]) => n + k.length + v.length * 6, 0));
    return { tag, attrs: { "data-node-id": this.id("node"), ...attrs }, children };
  }
  id(prefix: string): string { return prefix + "-" + ++this.serial; }
  content(node: Html, value: string): void { this.bytes(value.length * 6); node.children.push(value); }
  style(node: Html, css: Css): void {
    const source = declarations(css);
    this.bytes(source.length * 6);
    if (source) node.attrs.style = source;
  }
  rule(name: string, css: Css): void { const text = "." + name + "{" + declarations(css) + "}"; this.bytes(text.length * 4); this.css.push(text); }
  themeStyle(name: string, reference: XmlNode): string {
    if (this.unavailableThemeStyles.has(name)) unsupported(reference, "A referenced theme style cannot be rendered.");
    if (!this.css.some(rule => rule.startsWith("." + name + "{"))) fail(reference, "A theme style reference is missing.");
    return name;
  }
  relation(source: string, id: string | undefined, kind: string, external = false): { target: string; external: boolean } {
    const relation = id ? this.opc.relationships(source).get(id) : undefined;
    if (!relation || relation.type !== R + "/" + kind) fail({ part: source, name: "Relationship" }, "A required relationship is missing or has the wrong type.");
    if (relation.external && !external) unsupported({ part: source, name: "Relationship" } as XmlNode, "External PPTX resources are not supported.");
    return relation;
  }
  related(source: string, kind: string, required = true): string | undefined {
    const list = [...this.opc.relationships(source).values()].filter(r => r.type === R + "/" + kind);
    if (list.length > 1 || required && !list.length) fail({ part: source, name: kind }, "A singleton PPTX relationship is missing or ambiguous.");
    return list[0] ? this.relation(source, list[0].id, kind).target : undefined;
  }
  part(path: string, kind: string, name: string, uri: string): XmlNode {
    if (this.opc.contentType(path) !== "application/vnd.openxmlformats-officedocument." + kind + "+xml") fail({ part: path, name }, "A PPTX part content type is invalid.");
    const root = this.opc.read(path); rootName(root, name, uri); return root;
  }
  private storeAsset(path: string, mime: string, data: Uint8Array): string {
    if (!this.files.has(path)) {
      this.bytes(data.byteLength);
      this.assets.push(createAssetDescriptor(path, mime, data));
      this.files.set(path, data);
    }
    return path;
  }
  image(source: string, id: string | undefined, ownerId?: string): string {
    const path = this.relation(source, id, "image").target, data = this.opc.files.get(path)!;
    const mime = this.opc.contentType(path);
    const nativePath = "assets/" + path;
    if (mime && ["image/png", "image/jpeg", "image/webp"].includes(mime)) return this.storeAsset(nativePath, mime, data);
    if (ownerId && this.compatiblePreview && (mime === "image/emf" || mime === "image/x-emf")) {
      const previewPath = "preview-assets/" + path + ".png";
      if (!this.files.has(previewPath)) {
        this.storeAsset(nativePath, mime, data);
        this.storeAsset(previewPath, "image/png", emfPreview(data, path, this.opc.limits.maxFileBytes, () => this.count()));
      }
      (this.metadata.nativePreview ??= []).push({ feature: "emf-image", ownerId, part: path, assetPath: nativePath });
      return previewPath;
    }
    return unsupported({ part: path, name: "image" } as XmlNode, "This image container is not supported.");
  }
}

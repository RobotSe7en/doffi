import type { WorkbookMetadata } from "../contracts/workbook.js";
import type { WorkbookObjectTextOperation } from "../contracts/workbook-object-text-operations.js";
import type { WorkbookModel } from "../runtime.js";
import { invalid } from "../document/validation.js";
import type { ParsedHtml } from "../document/html.js";
import { textFormattingProperties } from "../document/text-decoration.js";
import { WorkbookContentEdit } from "./workbook-content.js";

export function validateWorkbookObjectTextTarget(base: WorkbookModel, operation: WorkbookObjectTextOperation): void {
  const object = base.snapshot().metadata.workbook.objects?.find(object => object.nodeId === operation.nodeId);
  if (!object || object.kind !== "text" && object.kind !== "geometry") invalid("Paragraph editing requires a base text box or shape.");
  const paragraphId = "paragraphId" in operation ? operation.paragraphId : operation.beforeParagraphId;
  if (paragraphId !== undefined) {
    const paragraph = base.getNode(paragraphId);
    if (paragraph?.tagName !== "p" || paragraph.parentNodeId !== object.nodeId) invalid("The paragraph must belong to its drawing in the base snapshot.");
  }
}
export function sourceDescendants(source: ParsedHtml, nodeId: string): Set<string> {
  const result = new Set<string>(), pending = [nodeId];
  while (pending.length) { const id = pending.pop()!; result.add(id); pending.push(...source.peekNode(id)!.childNodeIds); }
  return result;
}
/** Explicit replacements remove only their own dangling native hyperlinks. */
export function pruneWorkbookObjectLinks(workbook: WorkbookMetadata, content: WorkbookContentEdit, previousIds: ReadonlySet<string>): void {
  if (!workbook.objectHyperlinks?.some(link => previousIds.has(link.nodeId))) return;
  const source = content.parse(); workbook.objectHyperlinks = workbook.objectHyperlinks.filter(link => !previousIds.has(link.nodeId) || !!source.peekNode(link.nodeId));
}
export function editWorkbookObjectText(workbook: WorkbookMetadata, content: WorkbookContentEdit, operation: WorkbookObjectTextOperation): void {
  const object = workbook.objects?.find(object => object.nodeId === operation.nodeId);
  if (!object || object.kind !== "text" && object.kind !== "geometry") invalid("An earlier operation removed this text drawing.");
  const source = content.parse(), root = source.peekNode(object.nodeId)!;
  const paragraph = (id: string): string => {
    const node = source.peekNode(id); if (node?.tagName !== "p" || node.parentNodeId !== object.nodeId) invalid("An earlier operation removed this drawing paragraph."); return id;
  };
  if (operation.op === "insertWorkbookObjectParagraph") {
    const before = operation.beforeParagraphId === undefined ? undefined : paragraph(operation.beforeParagraphId), added = content.edit().insertSource(object.nodeId, operation.html, before);
    if (added.rootNodeIds.length !== 1 || content.parse().peekNode(added.rootNodeIds[0]!)?.tagName !== "p") invalid("Paragraph insertion requires exactly one source paragraph."); return;
  }
  const id = paragraph(operation.paragraphId);
  if (operation.op === "replaceWorkbookObjectText") { content.edit().replaceSourceText(id, operation); return; }
  if (operation.op === "formatWorkbookObjectText") { content.edit().formatSourceText(id, operation, operation.properties); return; }
  if (operation.op === "setWorkbookObjectParagraphStyle") {
    for (const [property, value] of Object.entries(textFormattingProperties(operation.properties))) {
      if (!/^(?:color|font-(?:family|size|weight|style)|text-(?:align|indent|decoration(?:-(?:line|style|color))?)|line-height|letter-spacing|margin-(?:top|bottom)|direction|--office-(?:font|text)-[a-z0-9-]+)$/.test(property)) invalid("The property is outside the drawing paragraph subset.");
      content.edit().sourceDrawingStyle(id, property, value);
    }
    return;
  }
  const previous = sourceDescendants(source, id);
  if (root.childNodeIds.length === 1) content.edit().replaceSourceContent(id, ""); else content.edit().removeSource(id);
  if (workbook.objectHyperlinks) workbook.objectHyperlinks = workbook.objectHyperlinks.filter(link => !previous.has(link.nodeId));
}

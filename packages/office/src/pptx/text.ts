import { A, attr, child, fail, R, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { PresentationObject } from "@doffi/core";
import { Context, value } from "./context.js";
import type { Html } from "./context.js";
import { importText as drawingText, importTextStyle as drawingTextStyle } from "../drawingml/text.js";
export { runProperties, paragraphProperties } from "../drawingml/text.js";

export function importHyperlink(context: Context, hyperlink: XmlNode | undefined, targetNode: Html, part: string): void {
  if (!hyperlink) return;
  if (value(hyperlink, "action") && value(hyperlink, "action") !== "ppaction://hlinksldjump") unsupported(hyperlink, "Active presentation actions are not supported.");
  if (child(hyperlink, "snd", A)) unsupported(hyperlink, "Hyperlink sound actions are not supported.");
  const id = attr(hyperlink, "id", R), relation = id && context.opc.relationships(part).get(id);
  if (!relation) fail(hyperlink, "A hyperlink relationship is missing.");
  if (relation.type === R + "/slide" && !relation.external) {
    const slideId = context.slidePaths.get(relation.target);
    if (!slideId) fail(hyperlink, "A hyperlink targets an unknown slide.");
    context.metadata.hyperlinks.push({ nodeId: targetNode.attrs["data-node-id"]!, slideId });
  } else {
    const target = context.relation(part, id, "hyperlink", true);
    if (!target.external || !/^(https?:\/\/|mailto:)/i.test(target.target) || /[\u0000-\u0020\u007f]/.test(target.target)) unsupported(hyperlink, "This hyperlink target is not supported.");
    context.metadata.hyperlinks.push({ nodeId: targetNode.attrs["data-node-id"]!, target: target.target });
  }
}

export function importTextStyle(context: Context, source: XmlNode | undefined, className: string): void {
  drawingTextStyle(context, source, className);
}
export function importText(context: Context, body: XmlNode, target: Html, part: string, eastAsianVertical = false): Extract<PresentationObject, { kind: "text" }>["text"] {
  return drawingText({
    eastAsianVertical,
    id: prefix => context.id(prefix), make: (tag, attrs) => context.make(tag, attrs),
    style: (node, css) => context.style(node, css), rule: (name, css) => context.rule(name, css),
    bytes: amount => context.bytes(amount), content: (node, text) => context.content(node, text),
    hyperlink: (source, node, part) => importHyperlink(context, source, node, part),
    projection: (nodeId, paragraphs) => { context.paragraphs.set(nodeId, paragraphs); },
    field(source, node) {
      const type = value(source, "type");
      if (!type || !/^(slidenum|datetime(?:FigureOut|[1-9]|1[0-3])?)$/.test(type)) unsupported(source, "This presentation field type is not supported.");
      node.attrs["data-office-field"] = type;
      if (value(source, "id")) node.attrs["data-office-field-id"] = value(source, "id")!;
    },
  }, body, target, part);
}

import { W } from "./xml.js";
import type { XmlNode } from "./xml.js";
const order = ["pos", "numFmt", "numStart", "numRestart", "footnote", "endnote"];
/** CT_FtnProps/CT_EdnProps order also applies inside native property histories. */
export function sortWordNoteProperties(root: XmlNode): void {
  if (root.uri === W && ["footnotePr", "endnotePr"].includes(root.name)) {
    const known = root.content
      .filter(
        (node): node is XmlNode =>
          typeof node !== "string" && node.uri === W && order.includes(node.name),
      )
      .sort((a, b) => order.indexOf(a.name) - order.indexOf(b.name));
    let index = 0;
    for (let slot = 0; slot < root.content.length; slot++) {
      const node = root.content[slot]!;
      if (typeof node !== "string" && node.uri === W && order.includes(node.name))
        root.content[slot] = known[index++]!;
    }
  }
  for (const child of root.content) if (typeof child !== "string") sortWordNoteProperties(child);
}

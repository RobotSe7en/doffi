import { defaultTreeAdapter, html, parseFragment } from "parse5";
import type { DefaultTreeAdapterMap } from "parse5";
import type { SourceStyleMap, SourceStyleNode } from "@doffi/core";
import { font } from "./style-values.js";
import type { WorkbookStylesWriter } from "./styles.js";
import { element, stringText } from "./xml.js";

type Child = DefaultTreeAdapterMap["childNode"];
const row = defaultTreeAdapter.createElement("tr", html.NS.HTML, []);
const adapter = defaultTreeAdapter;
export function richCellText(source: string, plain: string, styles: WorkbookStylesWriter): string {
  if (!/<span(?:[\t\n\f\r />])/i.test(source)) return stringText(plain);
  const fragment = parseFragment(row, source, {}), cell = fragment.childNodes[0];
  if (!cell || !adapter.isElementNode(cell)) return stringText(plain);
  const runs: { properties: string; text: string }[] = [];
  const add = (text: string, properties: string): void => {
    if (!text) return;
    const prior = runs.at(-1);
    if (prior?.properties === properties) prior.text += text;
    else runs.push({ properties, text });
  };
  const visit = (node: Child, parent: SourceStyleNode, inherited: string, previous: SourceStyleMap): void => {
    if (adapter.isTextNode(node)) add(node.value, inherited);
    else if (adapter.isElementNode(node)) {
      if (node.tagName === "br") { add("\n", inherited); return; }
      const projection = { tag: node.tagName, attrs: Object.fromEntries(node.attrs.map(attr => [attr.name, attr.value])), parent };
      const raw = { ...previous, ...styles.resolver.specified(projection) }, computed = styles.resolver.get(projection);
      const css = Object.fromEntries(Object.keys(raw).map(key => [key, computed[key]]).filter(entry => entry[1] !== undefined));
      const properties = font(raw, css, "run") || inherited;
      for (const child of node.childNodes) visit(child, projection, properties, raw);
    }
  };
  for (const child of cell.childNodes) visit(child, styles.themeScope, "", {});
  return runs.some(run => run.properties) ? runs.map(run => element("r", run.properties + stringText(run.text))).join("") : stringText(plain);
}

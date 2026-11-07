import { expect, test } from "vitest";
import { parseHtmlTree, serializeHtmlTree } from "../src/document/html-parser.js";
import { readLimits } from "../src/document/validation.js";
const tags = new Set(["html", "head", "body", "div"]);
const source = (text: string) => '<!doctype html><html><head></head><body><div>' + text + '</div></body></html>';
test.each(["&#13;&#10;", "&#013;&#xA;", "&#xD;&#13;", "&#x00d;&#x1F600;"])("a CR immediately followed by a numeric entity remains exact: %s", text => {
  const tree = parseHtmlTree(source(text), readLimits(), tags), html = serializeHtmlTree(tree);
  expect(html).toContain("&#13;"); expect(serializeHtmlTree(parseHtmlTree(html, readLimits(), tags))).toBe(html);
  if (text === "&#13;&#10;") expect(html).toContain('<div>&#13;\n</div>');
});
test.each(["&#13;&#0;", "&#13;&#x1;", "&#13;&#x7f;", "&#13;&#xD800;", "&#13;&#999999999;", "&#13;&#10"])("a CR cannot hide a following invalid numeric entity: %s", text => {
  expect(() => parseHtmlTree(source(text), readLimits(), tags)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
});

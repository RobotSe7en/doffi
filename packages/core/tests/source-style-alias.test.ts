import { expect, test } from "vitest";
import { DEFAULT_DOCUMENT_READ_LIMITS as limits, SourceStyleResolver } from "../src/index.js";
import { SourceStyleEdit } from "../src/session/source-styles.js";

const encode = (text: string) => new TextEncoder().encode(text), decode = (bytes: Uint8Array) => new TextDecoder().decode(bytes);
test("repeated isolated drawing style edits keep linear source CSS instead of copying shadowed history", () => {
  const files = new Map([["styles.css", encode('.shared{color:blue;--office-line-width:2pt}')]]), edit = new SourceStyleEdit(files, ["styles.css"], limits);
  let previous = "shared", halfway = 0;
  for (let index = 0; index < 100; index++) {
    const name = "local-" + index, properties = edit.drawingProperties({ color: index % 2 ? "#112233" : "#334455" });
    edit.aliasClass(previous, name, new Set(Object.keys(properties))); edit.overrideLocalClass(name, properties, {}, "span"); previous = name;
    if (index === 49) halfway = files.get("styles.css")!.length;
  }
  const css = decode(files.get("styles.css")!); expect(css.length).toBeLessThan(halfway * 2.5); expect(css.length).toBeLessThan(100_000);
  const resolver = new SourceStyleResolver([css], limits), own = resolver.specified({ tag: "span", attrs: { class: previous }, parent: undefined });
  expect(own).toMatchObject({ color: "#112233", "--office-line-width": "2pt", "--office-text-source": "direct" });
  expect(resolver.classStyle("shared").color).toBe("blue");
});

test("discarding overwritten alias declarations preserves descendant rules, other consumers and contextual cascade order", () => {
  const files = new Map([["styles.css", encode('div.parent span.shared{color:red!important;--office-line-width:2pt}.shared .child{color:green}.shared{font-size:11pt}.parent .shared{font-size:13pt}')]]), edit = new SourceStyleEdit(files, ["styles.css"], limits);
  edit.aliasClass("shared", "local", new Set(["color"])); edit.overrideLocalClass("local", { color: "blue" }, {}, "span");
  const resolver = new SourceStyleResolver([decode(files.get("styles.css")!)], limits), outer = { tag: "div", attrs: { class: "parent" }, parent: undefined };
  const local = { tag: "span", attrs: { class: "local" }, parent: outer }, shared = { tag: "span", attrs: { class: "shared" }, parent: outer };
  expect(resolver.get(local)).toMatchObject({ color: "blue", "font-size": "13pt", "--office-line-width": "2pt" }); expect(resolver.get(shared).color).toBe("red");
  expect(resolver.get({ tag: "span", attrs: { class: "child" }, parent: local }).color).toBe("green");
});

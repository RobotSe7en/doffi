import { expect, test } from "vitest";
import { DEFAULT_DOCUMENT_READ_LIMITS as limits, SourceStyleResolver, resolveSourceCssVariables } from "../src/index.js";
import type { SourceStyleNode } from "../src/index.js";

const node = (tag: string, attrs: Record<string, string> = {}, parent?: SourceStyleNode): SourceStyleNode => ({ tag, attrs, parent });

test("portable source cascade resolves selectors, priorities, shorthand and inherited computed font roles", () => {
  const root = node("html"), body = node("body", {}, root), outer = node("div", { class: "outer", style: '--face:"宋体";font-size:10pt;--accent:#123456' }, body);
  const paragraph = node("p", { class: "selected", "data-role": "value", style: "margin-left:4pt!important;color:purple" }, outer);
  const child = node("strong", {}, paragraph);
  const styles = new SourceStyleResolver(['.outer p[data-role=value]{color:var(--accent)!important;margin:1pt 2pt;font-size:1.5em;font-family:var(--face)}.selected{color:red;margin-left:9pt!important}'], limits);
  expect(styles.get(paragraph)).toMatchObject({ color: "#123456", "margin-top": "1pt", "margin-left": "4pt", "font-size": "15pt", "font-family": '"宋体"' });
  expect(styles.get(child)).toMatchObject({ color: "#123456", "font-size": "15pt", "font-weight": "bold" });
  expect(styles.specified(paragraph)).toMatchObject({ color: "var(--accent)", "font-size": "1.5em", "font-family": "var(--face)" });
  expect(Object.isFrozen(styles.get(child))).toBe(true); expect(Object.isFrozen(styles.specified(paragraph))).toBe(true);
});

test("own decoration queries distinguish ancestor paint from local none without changing native combined export", () => {
  const paragraph = node("p", { style: "color:red;text-decoration:underline double currentColor" });
  const child = node("span", { style: "color:blue;text-decoration:none" }, paragraph);
  const combined = new SourceStyleResolver([], limits), own = new SourceStyleResolver([], limits, new Set(), { combineDecorations: false });
  expect(combined.get(child)).toMatchObject({ "text-decoration-line": "underline", "text-decoration-style": "double", "text-decoration-color": "red" });
  expect(own.get(paragraph)).toMatchObject({ "text-decoration-line": "underline", "text-decoration-style": "double" });
  expect(own.get(child)).toMatchObject({ "text-decoration-line": "none", "text-decoration-style": "solid", color: "blue" });
  const secondPaint = node("s", { style: "text-decoration-color:blue" }, paragraph);
  expect(() => combined.get(secondPaint)).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
  expect(own.get(secondPaint)).toMatchObject({ "text-decoration-line": "line-through", "text-decoration-color": "blue" });
});

test("source resolver captures options and revision IDs while retaining source snapshot defaults", () => {
  const options = { initial: { color: "red", "font-size": "11pt" }, combineDecorations: false }, revisions = new Set(["revision"]);
  const styles = new SourceStyleResolver([], limits, revisions, options);
  options.initial.color = "blue"; options.combineDecorations = true; revisions.clear();
  expect(styles.get(node("span", { style: "color:currentColor" }))).toMatchObject({ color: "red", "font-size": "11pt" });
  expect(styles.get(node("ins", { "data-node-id": "revision" }))["text-decoration-line"]).toBe("none");
  expect(styles.get(node("span", {}, node("u")))["text-decoration-line"]).toBe("none");
});

test("source resolver bounds aggregate CSS, inline source, ancestry and cyclic variable expansion", () => {
  expect(() => new SourceStyleResolver([".a{color:red}".repeat(50), ".b{color:blue}".repeat(50)], { ...limits, maxCssNodes: 800 })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  expect(() => new SourceStyleResolver([".a{color:red}".repeat(20)], { ...limits, maxFileBytes: 200 })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  expect(() => new SourceStyleResolver([], limits).get(node("p", { style: "color:url(https://invalid.example/)" }))).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  const cyclic = { tag: "span", attrs: {}, parent: undefined as SourceStyleNode | undefined }; cyclic.parent = cyclic;
  expect(() => new SourceStyleResolver([], limits).get(cyclic)).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
  let deep = node("p"); for (let i = 0; i < 8; i++) deep = node("span", {}, deep);
  // The smaller depth remains sufficient to parse the built-in base rules.
  expect(() => new SourceStyleResolver([], { ...limits, maxDepth: 8 }).get(deep)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  const variables = node("p", { style: "--a:var(--b,var(--a));--b:red;color:var(--a,purple)" });
  expect(new SourceStyleResolver([], limits).get(variables).color).toBe("purple");
  expect(() => resolveSourceCssVariables("var(--a) var(--a)", () => "x".repeat(20), 0, 35)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
});

test("fallbacks preserve authored declarations and independent class queries stay frozen", () => {
  const parent = node("div", { style: "color:red" }), p = node("p", { class: "local" }, parent);
  const styles = new SourceStyleResolver(["p{color:green}.local{font-weight:bold}"], limits);
  expect(styles.withFallback(p, { color: "blue" })).toMatchObject({ color: "blue", "font-weight": "bold" });
  expect(styles.specified(p)).toEqual({ "font-weight": "bold", "margin-top": "0", "margin-right": "0", "margin-bottom": "0", "margin-left": "0", "line-height": "normal" });
  expect(styles.classStyle("local")["font-weight"]).toBe("bold");
  expect(Object.isFrozen(styles.specifiedClass("local"))).toBe(true);
});


test("authored declaration queries do not require native length conversion or computed variable values", () => {
  const styles = new SourceStyleResolver([".s{font-size:1vw;text-decoration:underline double var(--ink)}"], limits);
  const target = node("span", { class: "s" });
  expect(styles.specified(target)).toMatchObject({ "font-size": "1vw", "text-decoration-line": "underline", "text-decoration-style": "double", "text-decoration-color": "var(--ink)" });
  expect(() => styles.get(target)).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
});

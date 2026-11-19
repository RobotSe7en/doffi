import { expect, test } from "vitest";
import { strToU8 } from "fflate";
import { SaxesParser } from "saxes";
import { DEFAULT_DOCUMENT_READ_LIMITS } from "@doffi/core";
import { xml } from "../src/ooxml/xml.js";
import { serializeXml } from "../src/ooxml/serialize.js";
test("native namespace reuse remains bounded and preserves shadowed and default QName bindings", () => {
  let source =
      '<r:root xmlns:r="urn:root" xmlns:p="urn:first" xmlns:x="urn:types" xmlns="urn:default" x:type="Thing"><p:outer x:type="p:Type"/><p:inner xmlns:p="urn:second" x:type="p:Other"><p:leaf/></p:inner></r:root>',
    namespaceCount: number | undefined;
  for (let pass = 0; pass < 8; pass++) {
    source = serializeXml(
      xml(strToU8(source), "native.xml", DEFAULT_DOCUMENT_READ_LIMITS, { nodes: 0 }),
    );
    const count = source.match(/xmlns(?::[\w.-]+)?=/g)?.length ?? 0;
    if (namespaceCount !== undefined) expect(count).toBe(namespaceCount);
    namespaceCount = count;
    const parser = new SaxesParser({ xmlns: true }),
      values: string[] = [],
      scopes: Record<string, string>[] = [];
    parser.on("opentag", (tag) => {
      const scope = { ...scopes.at(-1), ...tag.ns };
      scopes.push(scope);
      const value = Object.values(tag.attributes).find(
        (attribute) => attribute.uri === "urn:types" && attribute.local === "type",
      )?.value;
      if (!value) return;
      const [prefix, name] = value.includes(":") ? value.split(":") : ["", value];
      values.push(scope[prefix!] + " " + name);
    });
    parser.on("closetag", () => {
      scopes.pop();
    });
    parser.write(source).close();
    expect(values).toEqual(["urn:default Thing", "urn:first Type", "urn:second Other"]);
  }
});

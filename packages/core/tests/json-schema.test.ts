import { readFileSync } from "node:fs";
import { join } from "node:path";
import { Ajv2020 } from "ajv/dist/2020.js";
import { expect, test } from "vitest";
import { CONTRACT_NAMES, DocumentError, getContractJsonSchema, parseContract } from "@doffi/core";
import { invalidCases, manifest, repository, validCases } from "./fixtures.js";

const ajv = new Ajv2020({ strict: true, allErrors: true });
const validators = new Map(CONTRACT_NAMES.map((name) => [name, ajv.compile(getContractJsonSchema(name))]));

test("axis edits have strict generated and runtime contracts", () => {
  const valid = { op: "editWorkbookAxis", sheetId: "sheet-1", axis: "row", action: "insert", index: 255, count: 2 };
  expect(validators.get("document-operation")!(valid)).toBe(true);
  expect(parseContract("document-operation", valid)).toEqual(valid);
  for (const invalid of [{ ...valid, count: 0 }, { ...valid, index: -1 }, { ...valid, axis: "sheet" }, { ...valid, action: "guess" }, { ...valid, partial: true }]) {
    expect(validators.get("document-operation")!(invalid)).toBe(false); expect(() => parseContract("document-operation", invalid)).toThrow(DocumentError);
  }
});

test("merge operations and explicit clearing share strict runtime and generated JSON contracts", () => {
  const target = { sheetId: "sheet-1", range: { from: { row: 0, column: 0 }, to: { row: 1, column: 1 } } };
  const valid = [{ op: "mergeWorkbookCells", ...target }, { op: "mergeWorkbookCells", ...target, coveredContent: "clear" }, { op: "unmergeWorkbookCells", ...target }];
  const invalid = [{ ...valid[0], coveredContent: "guess" }, { ...valid[2], coveredContent: "clear" }, { ...valid[0], partial: true }, { ...valid[0], range: undefined }];
  for (const operation of valid) { expect(validators.get("document-operation")!(operation)).toBe(true); expect(parseContract("document-operation", operation)).toEqual(operation); }
  for (const operation of invalid) { expect(validators.get("document-operation")!(operation)).toBe(false); expect(() => parseContract("document-operation", operation)).toThrow(DocumentError); }
});

test("worksheet filter commands require either a strict definition or explicit null removal", () => {
  const command = { op: "setWorkbookAutoFilter", sheetId: "sheet-1", filter: { range: { from: { row: 0, column: 0 }, to: { row: 3, column: 1 } }, columns: [] } };
  for (const operation of [command, { ...command, filter: null }]) {
    expect(validators.get("document-operation")!(operation)).toBe(true); expect(parseContract("document-operation", operation)).toEqual(operation);
  }
  for (const operation of [{ ...command, filter: undefined }, { ...command, filter: { ...command.filter, hiddenRows: [] } }, { ...command, partial: true }]) {
    expect(validators.get("document-operation")!(operation)).toBe(false); expect(() => parseContract("document-operation", operation)).toThrow(DocumentError);
  }
});

test("unlocated workbook comments are allowed only as explicit recovery annotations in both contracts", () => {
  const anchor = { kind: "cells", sheetId: "sheet-1", ranges: [] };
  expect(validators.get("anchor")!(anchor)).toBe(false);
  expect(() => parseContract("anchor", anchor)).toThrow(DocumentError);
  for (const body of [{ instruction: "定位待确认" }, { bodyNodeId: "comment-body" }]) for (const status of ["open", "resolved", "closed", "needs_reanchor"]) {
    const annotations = [{ annotationId: "comment", documentId: "book", baseRevision: 0, anchor, status, ...body }];
    expect(validators.get("annotations")!(annotations)).toBe(status === "needs_reanchor");
    if (status === "needs_reanchor") expect(parseContract("annotations", annotations)).toEqual(annotations);
    else expect(() => parseContract("annotations", annotations)).toThrow(DocumentError);
  }
});

test.each(validCases)("independently validates %s with JSON Schema", (name, input) => {
  const validate = validators.get(name)!;
  expect(validate(input), JSON.stringify(validate.errors)).toBe(true);
});

test.each(invalidCases.map(([name, input], index) => [index, name, input] as const))(
  "JSON Schema rejects invalid case %i (%s)", (_index, name, input) => {
    expect(validators.get(name)!(input)).toBe(false);
  },
);

test.each(["\u2028", "\u2029"])("both validators reject dot segments after %j", (separator) => {
  const validate = validators.get("document-manifest")!;
  for (const suffix of ["../../outside.html", "./outside.html", ".", ".."]) {
    const input = { ...manifest, entrypoint: "a" + separator + "/" + suffix };
    expect(validate(input)).toBe(false);
    expect(() => parseContract("document-manifest", input)).toThrow(DocumentError);
  }
  const canonical = { ...manifest, entrypoint: "a" + separator + "/inside.html" };
  expect(validate(canonical)).toBe(true);
  expect(parseContract("document-manifest", canonical)).toEqual(canonical);
});

test.each([
  { field: "message", minimum: 1, maximum: 256 },
  { field: "issue.code", minimum: 1, maximum: 64 },
  { field: "issue.path", minimum: 0, maximum: 128 },
  { field: "issue.message", minimum: 1, maximum: 256 },
])("both validators count Unicode code points in $field", ({ field, minimum, maximum }) => {
  for (const character of ["a", "中", "😀", "\n", "\u2028", "\u2029"]) {
    for (const length of [0, Math.floor(maximum / 2) + 1, maximum - 1, maximum, maximum + 1]) {
      const value = character.repeat(length);
      const error = {
        code: "InvalidDocument",
        message: field === "message" ? value : "Invalid document.",
        details: [{
          code: field === "issue.code" ? value : "invalid_type",
          path: [field === "issue.path" ? value : "message"],
          message: field === "issue.message" ? value : "Invalid value.",
        }],
      };
      const event = { protocolVersion: 1, documentId: "d", revision: 0, type: "error", error };
      const valid = length >= minimum && length <= maximum;
      for (const [name, input] of [["document-error", error], ["document-event", event]] as const) {
        expect(validators.get(name)!(input), name + ": " + field + ", " + JSON.stringify(value)).toBe(valid);
        if (valid) expect(parseContract(name, input)).toEqual(input);
        else expect(() => parseContract(name, input)).toThrow(DocumentError);
      }
    }
  }
});

test.each(CONTRACT_NAMES)("build artifact %s matches the current source contract", (name) => {
  const path = join(repository, "packages/core/dist/schemas", name + ".json");
  expect(JSON.parse(readFileSync(path, "utf8"))).toEqual(getContractJsonSchema(name));
});

test("schemas have unique versioned identifiers and generation does not share mutable output", () => {
  const ids = CONTRACT_NAMES.map((name) => getContractJsonSchema(name).$id);
  expect(new Set(ids).size).toBe(CONTRACT_NAMES.length);
  const schema = getContractJsonSchema("document-manifest");
  const branches = schema.anyOf as { properties: Record<string, Record<string, unknown>> }[];
  const properties = branches[0]!.properties;
  properties.format!.const = "changed";
  delete (schema as Record<string, unknown>).anyOf;
  expect(getContractJsonSchema("document-manifest")).toHaveProperty("anyOf.0.properties.format.const", "doffi");
});

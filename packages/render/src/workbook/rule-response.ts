import { DocumentError, WorkbookValueSchema } from "@doffi/core";
import type { WorkbookRuleIssue, WorkbookRuleReadRequest, WorkbookRuleReadResponse, WorkbookSnapshot } from "@doffi/core";
import { isRecord, keys } from "./read-transport.js";
import { checkWindowIdentity } from "./window-source.js";

const reasons = new Set<WorkbookRuleIssue["reason"]>(["unsupported-function", "circular-reference", "circular-name", "ambiguous-name-scope", "calculation-required",
  "formula-error", "non-numeric-threshold", "empty-scale", "invalid-threshold-order", "invalid-list-reference", "type-mismatch"]);
const key = (point: { readonly row: number; readonly column: number }): number => point.row * 16_384 + point.column;
function invalid(): never { throw new DocumentError("InvalidDocument", "The workbook rule provider returned an invalid or incomplete result."); }

/** Budget and capture a host response before it can affect a candidate view.
 * Required coordinates, rule order and stopping are checked against the same
 * immutable source, including responses received through a Worker port. */
export function checkedWorkbookRules(value: unknown, request: WorkbookRuleReadRequest, snapshot: WorkbookSnapshot): WorkbookRuleReadResponse {
  if (!isRecord(value) || !keys(value, ["documentId", "revision", "sheetId", "cells", "steps", "readCells", "evaluatedCells"])
    || typeof value.documentId !== "string" || !Number.isSafeInteger(value.revision) || value.sheetId !== request.sheetId || !Array.isArray(value.cells)) invalid();
  checkWindowIdentity(value as unknown as { documentId: string; revision: number }, snapshot);
  for (const [name, maximum] of [["steps", request.maxSteps], ["readCells", request.maxValues], ["evaluatedCells", request.maxValues]] as const)
    if (!Number.isSafeInteger(value[name]) || (value[name] as number) < 0 || (value[name] as number) > maximum) invalid();
  if (value.cells.length > request.maxCells) throw new DocumentError("ResourceLimitExceeded", "Workbook rule results exceed their coordinate budget.");
  const sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === request.sheetId)!;
  const conditional = [...sheet.conditionalRules ?? []].sort((a, b) => a.priority - b.priority), validations = sheet.dataValidations ?? [];
  let inspected = 0;
  const contains = (rule: { readonly ranges: readonly { readonly from: { readonly row: number; readonly column: number }; readonly to: { readonly row: number; readonly column: number } }[] }, row: number, column: number): boolean =>
    rule.ranges.some(range => {
      if (++inspected > request.maxSteps) throw new DocumentError("ResourceLimitExceeded", "Workbook rule response inspection exceeds its work budget.");
      return row >= range.from.row && row <= range.to.row && column >= range.from.column && column <= range.to.column;
    });
  const required = new Map<number, { readonly row: number; readonly column: number }>(request.includeCurrent ? request.coordinates.map(point => [key(point), point]) : []);
  const definitions = [...conditional, ...validations], selected = new Set<string>();
  for (const point of request.coordinates) for (const rule of definitions) if (contains(rule, point.row, point.column)) {
    selected.add(rule.ruleId); required.set(key(point), point);
    if (selected.size > request.maxRules) throw new DocumentError("ResourceLimitExceeded", "Workbook rule results exceed their rule budget.");
  }
  if (value.cells.length !== required.size) invalid();
  let units = 0, evaluations = 0;
  const account = (amount: number): void => { units += amount; if (units > request.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Workbook rule results exceed their text budget."); };
  function issue(value: unknown): WorkbookRuleIssue {
    if (!isRecord(value) || !keys(value, ["reason", "functionName"]) || !reasons.has(value.reason as WorkbookRuleIssue["reason"]) || value.functionName !== undefined && (typeof value.functionName !== "string" || value.functionName.length > 8192)) invalid();
    if (typeof value.functionName === "string") account(value.functionName.length);
    return Object.freeze({ reason: value.reason as WorkbookRuleIssue["reason"], ...(typeof value.functionName === "string" ? { functionName: value.functionName } : {}) });
  }
  const cells = value.cells.map(raw => {
    if (!isRecord(raw) || !keys(raw, ["row", "column", "current", "conditional", "validation"]) || !Number.isSafeInteger(raw.row) || !Number.isSafeInteger(raw.column)
      || (raw.row as number) < 0 || (raw.row as number) >= 1_048_576 || (raw.column as number) < 0 || (raw.column as number) >= 16_384
      || !required.delete(key(raw as unknown as { row: number; column: number })) || !Array.isArray(raw.conditional) || !isRecord(raw.current)) invalid();
    const row = raw.row as number, column = raw.column as number, sourceRules = conditional.filter(rule => contains(rule, row, column));
    let current: WorkbookRuleReadResponse["cells"][number]["current"];
    if (raw.current.kind === "value" && keys(raw.current, ["kind", "value"])) {
      const parsed = WorkbookValueSchema.safeParse(raw.current.value); if (!parsed.success) invalid();
      if (parsed.data.type === "text") account(parsed.data.value.length);
      current = Object.freeze({ kind: "value", value: Object.freeze(parsed.data) });
    } else if (raw.current.kind === "unavailable" && keys(raw.current, ["kind", "reason", "functionName"])) {
      const unavailable = issue({ reason: raw.current.reason, ...(raw.current.functionName !== undefined ? { functionName: raw.current.functionName } : {}) });
      if (!["unsupported-function", "circular-reference", "circular-name", "ambiguous-name-scope", "calculation-required"].includes(unavailable.reason)) invalid();
      current = Object.freeze({ kind: "unavailable", ...unavailable }) as typeof current;
    } else invalid();
    let stopped = false;
    if (raw.conditional.length > sourceRules.length) invalid();
    const results = raw.conditional.map((result, index) => {
      if (++evaluations > request.maxSteps) throw new DocumentError("ResourceLimitExceeded", "Too many workbook rule evaluations.");
      const definition = sourceRules[index]!;
      if (stopped || !isRecord(result) || !keys(result, ["ruleId", "state", "issue", "scale"]) || result.ruleId !== definition.ruleId || !["match", "miss", "unavailable"].includes(result.state as string)) invalid();
      account(definition.ruleId.length);
      const failed = result.issue === undefined ? undefined : issue(result.issue);
      if (result.state === "unavailable" && !failed) invalid();
      let scale: { fromStop: number; toStop: number; ratio: number } | undefined;
      if (definition.kind === "colorScale" && result.state === "match") {
        const raw = result.scale;
        if (!isRecord(raw) || !keys(raw, ["fromStop", "toStop", "ratio"]) || !Number.isSafeInteger(raw.fromStop) || !Number.isSafeInteger(raw.toStop) || typeof raw.ratio !== "number" || !Number.isFinite(raw.ratio)
          || (raw.fromStop as number) < 0 || (raw.toStop as number) >= definition.stops.length || (raw.toStop as number) < (raw.fromStop as number)
          || (raw.toStop as number) - (raw.fromStop as number) > 1 || raw.ratio < 0 || raw.ratio > 1 || raw.fromStop === raw.toStop && raw.ratio !== 0) invalid();
        scale = Object.freeze({ fromStop: raw.fromStop as number, toStop: raw.toStop as number, ratio: raw.ratio });
      } else if (result.scale !== undefined) invalid();
      stopped = definition.stopIfTrue && result.state !== "miss";
      return Object.freeze({ ruleId: definition.ruleId, state: result.state as "match" | "miss" | "unavailable", ...(failed ? { issue: failed } : {}), ...(scale ? { scale } : {}) });
    });
    if (!stopped && results.length !== sourceRules.length) invalid();
    const validation = validations.find(rule => contains(rule, row, column));
    let validationResult: WorkbookRuleReadResponse["cells"][number]["validation"];
    if (validation) {
      const result = raw.validation;
      if (!isRecord(result) || !keys(result, ["ruleId", "state", "issue"]) || result.ruleId !== validation.ruleId || !["valid", "invalid", "unavailable", "none"].includes(result.state as string)
        || (validation.kind === "none") !== (result.state === "none")) invalid();
      const failed = result.issue === undefined ? undefined : issue(result.issue); if (result.state === "unavailable" && !failed) invalid();
      account(validation.ruleId.length);
      validationResult = Object.freeze({ ruleId: validation.ruleId, state: result.state as "valid" | "invalid" | "unavailable" | "none", ...(failed ? { issue: failed } : {}) });
    } else if (raw.validation !== undefined) invalid();
    return Object.freeze({ row, column, current, conditional: Object.freeze(results), ...(validationResult ? { validation: validationResult } : {}) });
  });
  return Object.freeze({ documentId: request.documentId, revision: request.revision, sheetId: request.sheetId, cells: Object.freeze(cells), steps: value.steps as number, readCells: value.readCells as number, evaluatedCells: value.evaluatedCells as number });
}

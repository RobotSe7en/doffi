import { captureWorkbookModel, readWorkbookRules } from "@doffi/core";
import type { WorkbookCalculationResult, WorkbookModel, WorkbookRuleIssue, WorkbookRuleReadRequest, WorkbookRuleReadResponse, WorkbookSnapshot } from "@doffi/core";

export interface WorkbookRuleDiagnostic {
  readonly code: "rule-current-value-unavailable" | "conditional-rule-unavailable" | "conditional-formula-error" | "validation-unavailable" | "validation-invalid";
  readonly sheetId: string; readonly row: number; readonly column: number; readonly ruleId?: string; readonly issue?: WorkbookRuleIssue;
}
export function workbookRuleDiagnostics(response: WorkbookRuleReadResponse): readonly WorkbookRuleDiagnostic[] {
  const result: WorkbookRuleDiagnostic[] = [];
  for (const cell of response.cells) {
    const location = { sheetId: response.sheetId, row: cell.row, column: cell.column };
    if (cell.current.kind === "unavailable") result.push(Object.freeze({ ...location, code: "rule-current-value-unavailable", issue: Object.freeze({ reason: cell.current.reason,
      ...(cell.current.functionName ? { functionName: cell.current.functionName } : {}) }) }));
    for (const rule of cell.conditional) if (rule.issue) result.push(Object.freeze({ ...location, ruleId: rule.ruleId, code: rule.state === "unavailable" ? "conditional-rule-unavailable" : "conditional-formula-error", issue: rule.issue }));
    if (cell.validation?.state === "unavailable") result.push(Object.freeze({ ...location, code: "validation-unavailable", ruleId: cell.validation.ruleId, issue: cell.validation.issue! }));
    if (cell.validation?.state === "invalid") result.push(Object.freeze({ ...location, code: "validation-invalid", ruleId: cell.validation.ruleId, ...(cell.validation.issue ? { issue: cell.validation.issue } : {}) }));
  }
  return Object.freeze(result);
}

export interface WorkbookRuleSource {
  readonly snapshot: WorkbookSnapshot;
  /** Request current values for every projected coordinate, including cells
   * without a rule. Source cells and imported caches remain immutable. */
  readonly currentValues?: boolean;
  read(request: WorkbookRuleReadRequest, signal: AbortSignal): Promise<WorkbookRuleReadResponse>;
}

/** The host adapter supplies event-loop checkpoints. Core owns neither timers
 * nor the host's Worker lifetime, and reads remain bound to one revision. */
export function modelRuleSource(model: WorkbookModel, calculation?: WorkbookCalculationResult): WorkbookRuleSource {
  model = captureWorkbookModel(model);
  return Object.freeze({ snapshot: model.snapshot(), ...(calculation === undefined ? {} : { currentValues: true }), read(request: WorkbookRuleReadRequest, signal: AbortSignal) {
    return readWorkbookRules(model, request, { signal, yield: () => new Promise<void>(resolve => setTimeout(resolve, 0)), ...(calculation === undefined ? {} : { calculation }) });
  } });
}

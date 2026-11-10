import { DocumentError } from "../errors.js";
import type { WorkbookModel } from "../runtime.js";
import type { WorkbookCalculationResult } from "./calculation.js";
import type { ScalarResolved, WorkbookScalarContext } from "./scalar.js";
import { currentWorkbookModel } from "./source.js";

interface CalculationSource {
  readonly model: WorkbookModel;
  readonly context: WorkbookScalarContext;
  readonly values: ReadonlyMap<string, ScalarResolved>;
}
const sources = new WeakMap<WorkbookCalculationResult, CalculationSource>();
export function registerWorkbookCalculation(result: WorkbookCalculationResult, source: CalculationSource): void { sources.set(result, source); }
/** An identity pair cannot substitute for an actual completed first-party
 * batch. Captured old revisions remain readable through their original model. */
export function workbookCalculationSource(model: WorkbookModel, result: WorkbookCalculationResult): CalculationSource {
  const source = sources.get(result);
  if (!source) throw new DocumentError("InvalidOperation", "Calculated reading requires a completed first-party calculation batch.");
  if (source.model !== currentWorkbookModel(model)) throw new DocumentError("RevisionConflict", "The calculation batch belongs to another workbook source.");
  return source;
}

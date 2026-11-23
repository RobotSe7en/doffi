import type { WorkbookModel } from "../runtime.js";
import type { WorkbookCalculationResult } from "./calculation.js";
import { workbookCalculationSource } from "./calculation-source.js";
import { createWorkbookScalarKernel } from "./scalar.js";
import type { WorkbookScalarContext, WorkbookScalarEvaluator, WorkbookScalarLimits } from "./scalar.js";

/** Read rules, charts and expressions against one completed batch, including
 * long dependency chains and the exact time/locale used for that batch. */
export function createWorkbookCalculatedEvaluator(model: WorkbookModel, calculation: WorkbookCalculationResult, limits: Partial<WorkbookScalarLimits> = {}): WorkbookScalarEvaluator {
  let source: ReturnType<typeof workbookCalculationSource> | undefined = workbookCalculationSource(model, calculation);
  const kernel = createWorkbookScalarKernel(source.model, limits, source.context, {
    cachedCell: location => source?.values.get(JSON.stringify([location.sheetId, location.row, location.column])),
  });
  return Object.freeze({ evaluate: kernel.evaluate, cell: kernel.cell, dispose() { source = undefined; kernel.dispose(); } });
}

/** Evaluate expressions against the saved formula values of one revision.
 * Keep this explicit snapshot separate from the current calculation cache. */
export function createWorkbookCachedEvaluator(model: WorkbookModel, limits: Partial<WorkbookScalarLimits> = {}, context: WorkbookScalarContext = {}): WorkbookScalarEvaluator {
  const kernel = createWorkbookScalarKernel(model, limits, context, {
    cachedCell: (_location, cell) => cell.formula?.cacheState === "imported" || cell.formula?.cacheState === "valid"
      ? cell.value : { kind: "unavailable", reason: "calculation-required" },
  });
  return Object.freeze({ evaluate: kernel.evaluate, cell: kernel.cell, dispose: kernel.dispose });
}

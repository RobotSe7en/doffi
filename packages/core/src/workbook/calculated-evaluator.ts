import type { WorkbookModel } from "../runtime.js";
import type { WorkbookCalculationResult } from "./calculation.js";
import { workbookCalculationSource } from "./calculation-source.js";
import { createWorkbookScalarKernel } from "./scalar.js";
import type { WorkbookScalarEvaluator, WorkbookScalarLimits } from "./scalar.js";

/** Read rules, charts and expressions against one completed batch, including
 * long dependency chains and the exact time/locale used for that batch. */
export function createWorkbookCalculatedEvaluator(model: WorkbookModel, calculation: WorkbookCalculationResult, limits: Partial<WorkbookScalarLimits> = {}): WorkbookScalarEvaluator {
  let source: ReturnType<typeof workbookCalculationSource> | undefined = workbookCalculationSource(model, calculation);
  const kernel = createWorkbookScalarKernel(source.model, limits, source.context, {
    cachedCell: location => source?.values.get(JSON.stringify([location.sheetId, location.row, location.column])),
  });
  return Object.freeze({ evaluate: kernel.evaluate, cell: kernel.cell, dispose() { source = undefined; kernel.dispose(); } });
}

import type { TransactionPreparationProgress } from "../runtime.js";
import type { WorkbookCalculationResult, WorkbookCalculator } from "../workbook/calculation.js";
import type { WorkbookScalarContext } from "../workbook/scalar.js";
import type { SessionPreparationControl } from "./lifecycle.js";
import { DocumentError } from "../errors.js";

export type WorkbookPreparationStep =
  | { readonly kind: "progress"; readonly value: TransactionPreparationProgress }
  | { readonly kind: "calculation"; readonly calculator: WorkbookCalculator; readonly context?: WorkbookScalarContext };
export type WorkbookPreparation<Value> = Generator<WorkbookPreparationStep, Value | undefined, WorkbookCalculationResult | undefined>;

/** Both entry points drive the same isolated candidate kernel. Only its
 * checkpoints and existing calculator execution mode differ. */
export function finishWorkbookPreparation<Value>(steps: WorkbookPreparation<Value>): Value {
  let calculation: WorkbookCalculationResult | undefined;
  try {
    for (;;) {
      const next = steps.next(calculation); calculation = undefined;
      if (next.done) { if (next.value === undefined) throw new DocumentError("InvalidOperation", "Workbook preparation did not complete."); return next.value; }
      if (next.value.kind === "calculation") calculation = next.value.calculator.calculateSync(next.value.context ? { context: next.value.context } : {});
    }
  } finally { steps.return(undefined); }
}
export async function finishWorkbookPreparationAsync<Value>(steps: WorkbookPreparation<Value>, control: SessionPreparationControl): Promise<Value> {
  let calculation: WorkbookCalculationResult | undefined;
  try {
    for (;;) {
      control.guard(); const next = steps.next(calculation); calculation = undefined;
      if (next.done) { if (next.value === undefined) throw new DocumentError("InvalidOperation", "Workbook preparation did not complete."); return next.value; }
      if (next.value.kind === "progress") await control.checkpoint(next.value.value);
      else {
        const request = next.value; let progress: TransactionPreparationProgress = { phase: "calculation", completed: 0, total: 0 };
        await control.checkpoint(progress);
        calculation = await request.calculator.calculate({ ...(request.context ? { context: request.context } : {}), ...(control.signal ? { signal: control.signal } : {}),
          yield: () => control.checkpoint(progress), onProgress: value => { progress = { phase: "calculation", completed: value.completed, total: value.total }; } });
        control.guard();
      }
    }
  } finally { steps.return(undefined); }
}

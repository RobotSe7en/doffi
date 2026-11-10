import type { TransactionPreparationProgress, WorkbookLayoutMeasurer, WorkbookLayoutRequest, WorkbookLayoutMeasurement } from "../runtime.js";
import type { WorkbookCalculationResult, WorkbookCalculator } from "../workbook/calculation.js";
import type { WorkbookScalarContext } from "../workbook/scalar.js";
import type { SessionPreparationControl } from "./lifecycle.js";
import { DocumentError } from "../errors.js";

export type WorkbookPreparationStep =
  | { readonly kind: "progress"; readonly value: TransactionPreparationProgress }
  | { readonly kind: "layout"; readonly measure: WorkbookLayoutMeasurer; readonly request: WorkbookLayoutRequest; readonly accept: (value: WorkbookLayoutMeasurement) => void }
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
      else if (next.value.kind === "layout") {
        let cancelled = false;
        const signal = Object.freeze({ get aborted() { return cancelled; } });
        try {
          const value = next.value.measure(Object.freeze({ ...next.value.request, signal }));
          if (value && typeof (value as Promise<unknown>).then === "function") {
            void Promise.resolve(value).catch(() => {});
            throw new DocumentError("InvalidOperation", "An asynchronous workbook layout measurer requires asynchronous preparation.");
          }
          next.value.accept(value as WorkbookLayoutMeasurement);
        } finally { cancelled = true; }
      }
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
      else if (next.value.kind === "layout") {
        await control.checkpoint({ phase: "validation", completed: 0, total: 0 });
        const request = control.signal ? Object.freeze({ ...next.value.request, signal: control.signal }) : next.value.request;
        const value = await next.value.measure(request);
        control.guard(); next.value.accept(value); control.guard();
      }
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

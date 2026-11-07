import { captureWorkbookModel, DocumentError, readWorkbookChangeImpact } from "@doffi/core";
import type { WorkbookCalculationResult, WorkbookChangeImpact, WorkbookModel } from "@doffi/core";
import { modelWindowSource } from "./window-source.js";
import { modelAssetSource } from "./assets.js";
import { modelChartSource } from "./chart-source.js";
import { modelRuleSource } from "./rule-source.js";
import type { WorkbookPreviewInput, WorkbookPreviewSource } from "./preview-contracts.js";
const sources = new WeakMap<WorkbookPreviewSource, { readonly model: WorkbookModel; readonly calculation?: WorkbookCalculationResult }>();
const noRules = Object.freeze({});
const updates = new WeakMap<object, WeakMap<object, { readonly cells: WeakRef<object>; readonly rules?: WeakRef<object>; readonly impact: WorkbookChangeImpact }>>();
/** Trusted adapters attest a complete cell/calculation delta against the exact
 * previous providers. IDs alone cannot select a cache, and weak links cannot
 * retain an unbounded chain of old revisions after the owner releases them. */
export function linkWorkbookPreviewSource<Source extends WorkbookPreviewSource>(source: Source, previous: WorkbookPreviewSource, impact: WorkbookChangeImpact): Source {
  if (typeof WeakRef === "undefined") return source;
  let values = updates.get(source.cells); if (!values) { values = new WeakMap(); updates.set(source.cells, values); }
  values.set(source.rules ?? noRules, { cells: new WeakRef(previous.cells), ...(previous.rules ? { rules: new WeakRef(previous.rules) } : {}), impact });
  return source;
}
export function workbookPreviewChangeImpact(previous: WorkbookPreviewSource, source: WorkbookPreviewSource): WorkbookChangeImpact | undefined {
  const link = updates.get(source.cells)?.get(source.rules ?? noRules);
  if (!link || link.cells.deref() !== previous.cells || link.rules?.deref() !== previous.rules) return;
  return link.impact;
}

/** Direct adapter for a host-owned immutable source and optional completed
 * calculation. Worker hosts create these providers inside their Worker. */
export function createWorkbookPreviewSource(model: WorkbookModel, calculation?: WorkbookCalculationResult, previous?: WorkbookPreviewSource): WorkbookPreviewSource {
  model = captureWorkbookModel(model);
  const before = previous === undefined ? undefined : sources.get(previous);
  let impact: WorkbookChangeImpact | undefined;
  if (before?.model.snapshot().manifest.documentId === model.snapshot().manifest.documentId) try {
    impact = readWorkbookChangeImpact(before.model, model, before.calculation, calculation);
  } catch (error) {
    // An authorized custom immutable provider has no private core proof. Its
    // ordinary read path remains available without the optional optimization.
    if (calculation !== undefined || before.calculation !== undefined || !(error instanceof DocumentError) || error.code !== "InvalidOperation") throw error;
  }
  const result = Object.freeze({ kind: "workbook" as const, cells: modelWindowSource(model), assets: modelAssetSource(model), charts: modelChartSource(model, calculation), rules: modelRuleSource(model, calculation), ...(impact ? { impact } : {}) });
  sources.set(result, { model, ...(calculation === undefined ? {} : { calculation }) });
  return impact && previous ? linkWorkbookPreviewSource(result, previous, impact) : result;
}
export function workbookPreviewSource(input: WorkbookPreviewInput): WorkbookPreviewSource {
  if (input.kind !== "workbook") throw new DocumentError("UnsupportedFormat", "Workbook preview requires a workbook model or authorized read provider.");
  return Object.freeze("cells" in input ? { kind: "workbook", cells: input.cells, assets: input.assets, charts: input.charts, ...(input.rules ? { rules: input.rules } : {}), ...(input.impact ? { impact: input.impact } : {}) }
    : createWorkbookPreviewSource(input));
}

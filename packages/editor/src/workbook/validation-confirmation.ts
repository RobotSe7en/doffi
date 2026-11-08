import { DocumentError } from "@doffi/core";
import type { TransactionPrepareOptions } from "@doffi/core";
import type { WorkbookEditorOptions, WorkbookValidationConfirmation, WorkbookValidationWarning } from "./editor-contracts.js";
import type { WorkbookInputState } from "./input-controller.js";

interface ConfirmationContext {
  readonly documentId: string;
  readonly baseRevision: number;
  readonly confirm: WorkbookEditorOptions["confirmValidationWarnings"];
  readonly assertCurrent: () => void;
}
/** Keep time, locale, assets and preexisting acknowledgments fixed across every
 * private retry. A host callback cannot mutate the options being retried. */
function copyOptions(value: TransactionPrepareOptions): TransactionPrepareOptions {
  return { ...value,
    ...(value.assets ? { assets: value.assets.map(asset => ({ ...asset, descriptor: { ...asset.descriptor }, bytes: asset.bytes instanceof Uint8Array ? new Uint8Array(asset.bytes) : asset.bytes })) } : {}),
    ...(value.calculationContext ? { calculationContext: { ...value.calculationContext,
      ...(value.calculationContext.now ? { now: { ...value.calculationContext.now } } : {}) } } : {}),
    ...(value.validation ? { validation: { ...value.validation, acceptedRuleIds: [...value.validation.acceptedRuleIds] } } : {}) };
}
function retryOptions(error: unknown, options: TransactionPrepareOptions, context: ConfirmationContext): TransactionPrepareOptions {
  if (!context.confirm || !(error instanceof DocumentError) || error.code !== "InvalidOperation" || !error.details.length) throw error;
  const warnings: WorkbookValidationWarning[] = [];
  for (const issue of error.details) {
    const [sheetId, row, column, ruleId] = issue.path;
    if (issue.code !== "validation-confirmation-required" || issue.path.length !== 4 || typeof sheetId !== "string" || !sheetId ||
      typeof ruleId !== "string" || !ruleId || typeof row !== "number" || !Number.isSafeInteger(row) || row < 0 || row > 1_048_575 ||
      typeof column !== "number" || !Number.isSafeInteger(column) || column < 0 || column > 16_383) throw error;
    warnings.push(Object.freeze({ sheetId, row, column, ruleId, message: issue.message }));
  }
  const accepted = new Set(options.validation?.acceptedRuleIds);
  if (warnings.every(warning => accepted.has(warning.ruleId))) throw error;
  context.assertCurrent();
  const event: WorkbookValidationConfirmation = Object.freeze({ documentId: context.documentId, baseRevision: context.baseRevision, warnings: Object.freeze(warnings) });
  if (context.confirm(event) !== true) throw error;
  context.assertCurrent();
  for (const warning of warnings) accepted.add(warning.ruleId);
  return { ...options, validation: { ...options.validation, acceptedRuleIds: [...accepted] } };
}
export function withWorkbookValidationConfirmation<T>(apply: (options: TransactionPrepareOptions) => T, options: TransactionPrepareOptions, context: ConfirmationContext): T {
  if (!context.confirm) return apply(options);
  let current = copyOptions(options);
  for (;;) { try { return apply(current); } catch (error) { current = retryOptions(error, current, context); } }
}
export async function withWorkbookAsyncValidationConfirmation<T>(apply: (options: TransactionPrepareOptions) => Promise<T>, options: TransactionPrepareOptions, context: ConfirmationContext): Promise<T> {
  if (!context.confirm) return apply(options);
  let current = copyOptions(options);
  for (;;) { try { return await apply(current); } catch (error) { current = retryOptions(error, current, context); } }
}
/** Excludes transport activity flags, which change while a draft is prepared. */
export function workbookValidationInputKey(state: WorkbookInputState): string {
  return JSON.stringify({ selection: state.selection, text: state.text, mode: state.mode, dirty: state.dirty, composing: state.composing });
}

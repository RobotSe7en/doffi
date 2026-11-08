import { parseWorkbookFormula, rewriteWorkbookFormula } from "@doffi/core";
import type { DeepReadonly, WorkbookCoordinate, WorkbookDefinedName, WorkbookFormulaExpression, WorkbookFormulaRewriteBindings, WorkbookMetadata, WorkbookSheet } from "@doffi/core";

interface CopyNameBudget {
  guard(): void; step(): void; account(units: number): void; yield?: () => Promise<void>;
}
interface WorkbookNameCopyPlan {
  readonly names: readonly WorkbookDefinedName[];
  copy(expression: string, from: WorkbookCoordinate, to: WorkbookCoordinate): Promise<string>;
}
/** Commands plan new local definitions against their immutable source. The
 * first-party core rewriter owns every formula transformation; this helper
 * owns only scope lookup, name identities and cooperative graph discovery. */
export async function createWorkbookNameCopyPlan(workbook: DeepReadonly<WorkbookMetadata>, source: DeepReadonly<WorkbookSheet>, target: DeepReadonly<WorkbookSheet>, budget: CopyNameBudget): Promise<WorkbookNameCopyPlan> {
  const scopes = new Map<string | undefined, Map<string, DeepReadonly<WorkbookDefinedName>>>(), sheets = new Map<string, string>(), used = new Set<string>(), tables = new Set<string>();
  let work = 0, serial = 0, completed = 0;
  async function pause(): Promise<void> { await budget.yield!(); budget.guard(); }
  function tick(): Promise<void> | undefined {
    budget.guard(); budget.step(); return ++work % 256 === 0 && budget.yield ? pause() : undefined;
  }
  for (const sheet of workbook.sheets) {
    const waiting = tick(); if (waiting) await waiting; sheets.set(sheet.name.toUpperCase(), sheet.sheetId);
    for (const table of sheet.tables ?? []) { const waiting = tick(); if (waiting) await waiting; tables.add(table.name.toUpperCase()); }
  }
  for (const definition of workbook.names ?? []) {
    const waiting = tick(); if (waiting) await waiting;
    const scope = scopes.get(definition.sheetId) ?? new Map<string, DeepReadonly<WorkbookDefinedName>>(); scope.set(definition.name.toUpperCase(), definition); scopes.set(definition.sheetId, scope); used.add(definition.nameId);
  }
  const sourceNames = scopes.get(source.sheetId), targetNames = scopes.get(target.sheetId) ?? new Map<string, DeepReadonly<WorkbookDefinedName>>(); scopes.set(target.sheetId, targetNames);
  const planned: WorkbookDefinedName[] = [], originals: DeepReadonly<WorkbookDefinedName>[] = [];
  const bindings: WorkbookFormulaRewriteBindings = {
    resolveName: (name, owner, explicit, workbook) => {
      const token = name.toUpperCase(), local = scopes.get(sheets.get(owner.toUpperCase()) ?? "")?.get(token);
      return (workbook === 0 ? scopes.get(undefined)?.get(token) : local ?? (explicit ? undefined : scopes.get(undefined)?.get(token)))?.nameId;
    },
    isTableName: (name, owner, workbook) => !scopes.get(undefined)?.has(name.toUpperCase()) && (workbook === 0 || !scopes.get(sheets.get(owner.toUpperCase()) ?? "")?.has(name.toUpperCase())) && tables.has(name.toUpperCase()),
  };
  const copy = (expression: string, from: WorkbookCoordinate, to: WorkbookCoordinate): string => rewriteWorkbookFormula(expression, { ...from, sheet: source.name }, { ...to, sheet: target.name }, { kind: "copy" }, bindings);
  async function discover(expression: string, origin: WorkbookCoordinate): Promise<void> {
    const pending = [{ expression, origin, includeExplicit: false }];
    while (pending.length) {
      const input = pending.pop()!, stack: WorkbookFormulaExpression[] = [parseWorkbookFormula(input.expression, input.origin)];
      while (stack.length) {
        const waiting = tick(); if (waiting) await waiting; const node = stack.pop()!;
        if (node.kind === "name" && node.workbook === undefined && (node.sheet === undefined || input.includeExplicit && node.sheet.toUpperCase() === source.name.toUpperCase())) {
          const token = node.name.toUpperCase(), original = sourceNames?.get(token);
          if (original && !targetNames.has(token)) {
            let nameId: string; do { const waiting = tick(); if (waiting) await waiting; nameId = "copied-name-" + ++serial; } while (used.has(nameId)); used.add(nameId);
            const value: WorkbookDefinedName = { ...original, origin: { ...original.origin }, sheetId: target.sheetId, nameId };
            targetNames.set(token, value); planned.push(value); originals.push(original);
            pending.push({ expression: original.expression, origin: original.origin, includeExplicit: true });
          }
        } else if (node.kind === "binary") stack.push(node.right, node.left);
        else if (node.kind === "unary") stack.push(node.operand);
        else if (node.kind === "call") for (const argument of node.arguments) if (argument) stack.push(argument);
      }
    }
  }
  return {
    names: planned as readonly WorkbookDefinedName[],
    async copy(expression: string, from: WorkbookCoordinate, to: WorkbookCoordinate): Promise<string> {
      await discover(expression, from);
      while (completed < planned.length) {
        const original = originals[completed]!, clone = planned[completed]!, waiting = tick(); if (waiting) await waiting;
        clone.expression = rewriteWorkbookFormula(copy(original.expression, original.origin, original.origin), { ...original.origin, sheet: source.name }, { ...original.origin, sheet: target.name }, { kind: "renameSheet", from: source.name, to: target.name }, bindings);
        budget.account(JSON.stringify(clone).length); completed++;
      }
      budget.guard(); return copy(expression, from, to);
    },
  };
}

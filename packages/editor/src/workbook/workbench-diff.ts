import type { DeepReadonly, WorkbookEditProposal, WorkbookInteractionRequest, WorkbookInteractionCell, WorkbookProposalReview, WorkbookSnapshot } from "@doffi/core";
import type { WorkbookPreviewSource } from "@doffi/render";
import type { WorkbookWorkbenchCandidate } from "./workbench-contracts.js";
import { sessionMapping } from "./session-transport.js";

interface WorkbookProposalDiffInput {
  readonly candidate: WorkbookWorkbenchCandidate;
  readonly before: WorkbookSnapshot;
  readonly source: WorkbookPreviewSource;
  readonly request: WorkbookInteractionRequest;
  readonly proposal: WorkbookEditProposal;
  readonly review: DeepReadonly<WorkbookProposalReview>;
  readonly candidateHost: HTMLElement;
  readonly signal: AbortSignal;
  readonly onError: (error: unknown) => void;
}
const cellKey = (at: { readonly sheetId: string; readonly row: number; readonly column: number }) => JSON.stringify([at.sheetId, at.row, at.column]);
function address(at: { readonly row: number; readonly column: number }): string {
  let name = "", column = at.column + 1;
  while (column) { column--; name = String.fromCharCode(65 + column % 26) + name; column = Math.floor(column / 26); }
  return name + (at.row + 1);
}
type CellSource = Pick<WorkbookInteractionCell, "value" | "formula" | "cssClass" | "formatId" | "html">;
const sourceValue = (cell: CellSource | undefined) => cell ? { value: cell.value, formula: cell.formula, cssClass: cell.cssClass, formatId: cell.formatId, html: cell.html } : { value: { type: "blank" } };

/** Each candidate owns one aborted-on-release view. Source rows and calculated
 * differences are paged independently; only the current source page is read. */
export async function renderWorkbookProposalDiff(container: HTMLElement, input: WorkbookProposalDiffInput): Promise<void> {
  const { candidate, before, source, request, proposal, review, signal } = input, document = container.ownerDocument;
  const afterSnapshot = source.cells.snapshot, sheets = new Map([...before.metadata.workbook.sheets, ...afterSnapshot.metadata.workbook.sheets].map(sheet => [sheet.sheetId, sheet.name]));
  container.replaceChildren(); container.hidden = false;
  const title = document.createElement("h2"), summary = document.createElement("p"); title.textContent = "提案差异"; summary.textContent = proposal.summary; container.append(title, summary);
  function action(title: string, work: () => void | Promise<void>, parent: HTMLElement): HTMLButtonElement {
    const button = document.createElement("button"); button.type = "button"; button.textContent = title; parent.append(button);
    button.addEventListener("click", () => {
      try { void Promise.resolve(work()).catch(error => { if (!signal.aborted) input.onError(error); }); }
      catch (error) { if (!signal.aborted) input.onError(error); }
    }, { signal }); return button;
  }
  const sources = document.createElement("details"), sourceTitle = document.createElement("summary"), sourcePage = document.createElement("div"), sourceActions = document.createElement("div");
  sourceTitle.textContent = "直接修改的值、公式与格式"; sourcePage.setAttribute("aria-label", "工作簿来源差异分页"); sources.append(sourceTitle, sourceActions, sourcePage); container.append(sources);
  const ranges = candidate.changedRanges ?? [], total = ranges.reduce((count, entry) => count + (entry.range.to.row - entry.range.from.row + 1) * (entry.range.to.column - entry.range.from.column + 1), 0);
  const beforeCells = new Map(request.context.cells.map(cell => [cellKey(cell), cell])), sheetIds = new Set(afterSnapshot.metadata.workbook.sheets.map(sheet => sheet.sheetId));
  let page = 0, pageSequence = 0;
  async function sourceRows(): Promise<void> {
    const ticket = ++pageSequence, locations: { sheetId: string; row: number; column: number }[] = []; let skip = page * 50;
    for (const entry of ranges) {
      const width = entry.range.to.column - entry.range.from.column + 1, count = (entry.range.to.row - entry.range.from.row + 1) * width;
      if (skip >= count) { skip -= count; continue; }
      for (let at = skip; at < count && locations.length < 50; at++) locations.push({ sheetId: entry.sheetId, row: entry.range.from.row + Math.floor(at / width), column: entry.range.from.column + at % width });
      skip = 0; if (locations.length >= 50) break;
    }
    const windows = locations.filter(at => sheetIds.has(at.sheetId)).map(at => ({ ...at, rows: 1, columns: 1 }));
    const values = windows.length ? await source.cells.read({ documentId: candidate.documentId, revision: candidate.baseRevision, maxCells: 50, maxCodeUnits: request.context.limits.maxBytes, windows }, signal)
      .catch(reason => { if (signal.aborted) return undefined; throw reason; }) : { cells: [] };
    if (!values || signal.aborted || ticket !== pageSequence) return;
    const afterCells = new Map(values.cells.map(cell => [cellKey(cell), cell])); sourcePage.replaceChildren();
    const label = document.createElement("p"); label.textContent = total ? `${page * 50 + 1}–${Math.min(total, (page + 1) * 50)} / ${total} 个单元格` : "没有单元格源修改。"; sourcePage.append(label);
    for (const at of locations) {
      const row = document.createElement("details"), caption = document.createElement("summary"), previous = document.createElement("pre"), next = document.createElement("pre");
      caption.textContent = (sheets.get(at.sheetId) ?? at.sheetId) + "!" + address(at);
      previous.setAttribute("aria-label", "修改前单元格来源"); next.setAttribute("aria-label", "修改后单元格来源");
      previous.textContent = JSON.stringify(sourceValue(beforeCells.get(cellKey(at))), null, 2); next.textContent = JSON.stringify(sourceValue(afterCells.get(cellKey(at))), null, 2);
      row.append(caption, previous, next); sourcePage.append(row);
    }
    previous.disabled = page === 0; next.disabled = (page + 1) * 50 >= total;
  }
  const previous = action("上一页源差异", () => { if (page > 0) { page--; return sourceRows(); } return undefined; }, sourceActions);
  const next = action("下一页源差异", () => { if ((page + 1) * 50 < total) { page++; return sourceRows(); } return undefined; }, sourceActions);
  const derived = document.createElement("details"), derivedTitle = document.createElement("summary"), derivedRows = document.createElement("div"), derivedActions = document.createElement("div");
  derivedTitle.textContent = "当前结果与只读消费者重算（" + review.calculatedChanges.length + "）"; derived.append(derivedTitle, derivedActions, derivedRows); container.append(derived);
  let derivedPage = 0;
  function drawDerived(): void {
    derivedRows.replaceChildren();
    for (const value of review.calculatedChanges.slice(derivedPage * 50, (derivedPage + 1) * 50)) {
      const row = document.createElement("p"); row.textContent = (sheets.get(value.sheetId) ?? value.sheetId) + "!" + address(value) + (value.access === "read-only" ? " · 只读源重算：" : " · 当前结果：") + JSON.stringify(value.before) + " → " + JSON.stringify(value.after); derivedRows.append(row);
    }
    derivedPrevious.disabled = derivedPage === 0; derivedNext.disabled = (derivedPage + 1) * 50 >= review.calculatedChanges.length;
  }
  const derivedPrevious = action("上一页计算差异", () => { derivedPage--; drawDerived(); }, derivedActions), derivedNext = action("下一页计算差异", () => { derivedPage++; drawDerived(); }, derivedActions); drawDerived();
  const structural = document.createElement("details"), structuralTitle = document.createElement("summary"), structuralText = document.createElement("pre");
  structuralTitle.textContent = "结构、规则、绘图、批注锚点与样式";
  const contentChanged = before.content?.path !== afterSnapshot.content?.path || before.content?.html !== afterSnapshot.content?.html;
  structuralText.textContent = JSON.stringify({ before: { metadata: before.metadata, annotations: before.annotations, styles: before.styles, ...(contentChanged ? { content: before.content ?? null } : {}) }, after: { metadata: afterSnapshot.metadata, annotations: afterSnapshot.annotations, styles: afterSnapshot.styles, ...(contentChanged ? { content: afterSnapshot.content ?? null } : {}) }, nodeChanges: candidate.changes, anchorMapping: "model" in candidate ? sessionMapping(candidate.mapping) : candidate.mapping }, null, 2);
  structural.append(structuralTitle, structuralText); container.append(structural, input.candidateHost);
  await sourceRows();
}

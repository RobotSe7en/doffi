import { expect } from "@playwright/test";
import type { Page } from "@playwright/test";
import { createDocumentSession } from "@doffi/core";
import type { Annotation, DeepReadonly, DocumentOperation, WorkbookComment, WorkbookModel } from "@doffi/core";

export function workbookCopyCommentsSource(model: WorkbookModel) {
  const s = createDocumentSession(model), root = s.getAnnotations()[0]!, note = s.snapshot().metadata.workbook.comments!.find(binding => binding.annotationId === root.annotationId)!;
  if (note.format !== "classic" || !note.box || !note.commentPrBox) throw Error("Expected independent dual-note source");
  const apply = (operations: DocumentOperation[]) => { const m = s.snapshot().manifest; s.commit(s.prepare({ protocolVersion: 1, documentId: m.documentId, baseRevision: m.revision, operations })); };
  const at = (row: number, column: number) => ({ kind: "cells" as const, sheetId: "sheet-1", ranges: [{ from: { row, column }, to: { row, column } }] });
  apply([
    { op: "updateWorkbookComment", annotationId: root.annotationId, changes: { anchor: at(1, 16) } },
    { op: "setWorkbookNoteBox", annotationId: root.annotationId, box: { ...note.box, visible: true, anchor: { kind: "two-cell", from: { row: 1, column: 17, offsetXPt: 2, offsetYPt: 3 }, to: { row: 3, column: 19, offsetXPt: 2, offsetYPt: 3 }, moveWithCells: true, sizeWithCells: false } } },
    { op: "updateWorkbookSheet", sheetId: "sheet-1", changes: { rowDimensions: [{ from: 5, to: 8, heightPt: 30 }], columnDimensions: [{ from: 4, to: 6, width: 25, unit: "characters" }] } },
    { op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 1, column: 16, source: { kind: "value", value: { type: "number", value: 7 } } }, { row: 5, column: 3, source: { kind: "value", value: { type: "number", value: 9 } } }] },
    { op: "formatWorkbookRanges", sheetId: "sheet-1", ranges: [{ from: { row: 1, column: 16 }, to: { row: 2, column: 16 } }], properties: { "font-weight": "bold" } },
    { op: "addWorkbookComment", anchor: at(2, 16), html: "<p>blank-thread禰󠄀😀</p>", author: { name: "现代作者" } },
    { op: "addWorkbookComment", anchor: at(5, 3), html: "<p>old-target</p>" },
    { op: "addWorkbookComment", anchor: at(6, 3), html: "<p>old-blank-target</p>" },
  ]);
  const modern = s.getAnnotations().find(item => item.anchor.kind === "cells" && item.anchor.ranges[0]!.from.column === 16 && item.annotationId !== root.annotationId)!;
  apply([{ op: "addWorkbookCommentReply", annotationId: modern.annotationId, html: "<p>copied-reply😀</p>" }]);
  const original = s.snapshot(), source = s.serialize(); s.dispose(); return { source, original, classicId: root.annotationId, modernId: modern.annotationId };
}
interface State { annotations: DeepReadonly<Annotation>[]; comments: DeepReadonly<WorkbookComment>[]; content: string; history: { undoSteps: number }; errors: string[] }
export async function exerciseWorkbookCopyComments(page: Page, remote: boolean, original: DeepReadonly<Annotation>[], classicId: string, modernId: string): Promise<void> {
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as State;
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const bounds = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.displayedDrawingBounds() : window.workbookEditor.displayedDrawingBounds(), remote);
  const select = async (value: string) => { const address = page.getByLabel("选择单元格或区域", { exact: true }); await address.fill(value); await address.press("Enter"); await settled(); };
  const at = (state: State, row: number) => state.annotations.find(item => item.anchor.kind === "cells" && item.anchor.ranges.some(range => range.from.row === row && range.from.column === 3))!;
  const sourceClassic = original.find(item => item.annotationId === classicId)!, sourceModern = original.find(item => item.annotationId === modernId)!;
  await select("Q2:Q3"); const box = (await bounds()).objects.find(item => item.nodeId === sourceClassic.bodyNodeId)!; expect(box).toBeDefined();
  const before = await inspect();
  const paste = async (mode = "contents", skipBlanks = false) => {
    await select("Q2:Q3"); await page.getByRole("button", { name: "复制选区", exact: true }).click();
    await page.getByLabel("区域粘贴方式", { exact: true }).selectOption(mode); await page.getByLabel("区域粘贴时跳过空白", { exact: true }).setChecked(skipBlanks);
    await select("D6"); await page.getByRole("button", { name: "粘贴选区", exact: true }).click(); await settled();
  };
  const undo = async () => { await page.getByRole("button", { name: "撤销", exact: true }).click(); await settled(); const state = await inspect(); expect(state).toMatchObject({ errors: [], history: { undoSteps: 0 }, comments: before.comments, content: before.content }); };
  await paste(); let state = await inspect(); expect(state).toMatchObject({ errors: [], history: { undoSteps: 1 } });
  const classic = at(state, 5), modern = at(state, 6); expect(classic.bodyNodeId).not.toBe(sourceClassic.bodyNodeId); expect(modern.bodyNodeId).not.toBe(sourceModern.bodyNodeId);
  expect(modern.replies).toHaveLength(sourceModern.replies!.length); expect(modern.replies![0]!.replyId).not.toBe(sourceModern.replies![0]!.replyId);
  const note = state.comments.find(item => item.annotationId === classic.annotationId)!; expect(note).toMatchObject({ format: "classic", boxSource: "vml" }); expect(note).toHaveProperty("commentPrBox");
  const copiedBox = (await bounds()).objects.find(item => item.nodeId === classic.bodyNodeId)!; expect(copiedBox).toBeDefined();
  expect(Math.abs(copiedBox.widthPt - box.widthPt)).toBeLessThan(.2); expect(Math.abs(copiedBox.heightPt - box.heightPt)).toBeLessThan(.2); await undo();
  for (const mode of ["values", "formulas", "formats"]) { await paste(mode); state = await inspect(); expect(state.comments).toEqual(before.comments); expect(state.annotations.map(item => item.annotationId)).toEqual(before.annotations.map(item => item.annotationId)); await undo(); }
  await paste("contents", true); state = await inspect(); expect(at(state, 6).annotationId).toBe(at(before, 6).annotationId); expect(at(state, 5).annotationId).not.toBe(at(before, 5).annotationId); await undo();
}

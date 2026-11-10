import { expect } from '@playwright/test';
import type { Page } from '@playwright/test';
import type { WorkbookSnapshot } from '@doffi/core';

export async function exerciseWorkbookFilterHeaders(page: Page, remote: boolean): Promise<void> {
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = async () => await page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote) as { sheets: WorkbookSnapshot['metadata']['workbook']['sheets'] };
  const choose = async (value: string) => { const input = page.getByRole('textbox', { name: '选择单元格或区域', exact: true }); await input.fill(value); await input.press('Enter'); await settled(); };
  const tab = async (name: string) => { await page.getByRole('tab', { name, exact: true }).click(); await settled(); };
  const shown = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.displayedCell(0, 0) : window.workbookEditor.displayedCell(0, 0), remote);
  const before = (await inspect()).sheets;
  for (const [selection, from, to, expected] of [
    ['B2:D2', 1, 4, '80'], ['A1:D3', 0, 2, '90'], ['B2:D6', undefined, undefined, '#REF!'], ['B2:D5', 1, 1, '40'],
  ] as const) {
    await tab('Data'); await choose(selection); await page.getByRole('button', { name: '删除选区行', exact: true }).click(); await settled();
    const filter = (await inspect()).sheets[0]!.autoFilter;
    if (from === undefined) expect(filter).toBeUndefined();
    else expect(filter).toMatchObject({ range: { from: { row: from, column: 1 }, to: { row: to, column: 3 } }, columns: before[0]!.autoFilter!.columns });
    await tab('Observer'); await choose('A1'); expect(await shown()).toBe(expected);
    expect(await inspect()).toMatchObject({ history: { undoSteps: 1 }, errors: [] });
    await page.getByRole('button', { name: '撤销', exact: true }).click(); await settled(); expect(await shown()).toBe('70');
    expect((await inspect()).sheets).toEqual(before);
    expect(await inspect()).toMatchObject({ history: { undoSteps: 0, redoSteps: 1 }, errors: [] });
  }
  if (remote) expect(await page.evaluate(() => window.workbookRemoteEditor.finish())).toMatchObject({ sessionsCreated: 1, endpointDisposed: true, reopenedEqual: true });
}

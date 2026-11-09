import { expect } from '@playwright/test';
import type { Page } from '@playwright/test';
export async function exerciseWorkbookFilterDisplay(page: Page, remote: boolean, sheets: { name: string; sheetId: string }[], formatId: string): Promise<void> {
  const settled = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.settled() : window.workbookEditor.settled(), remote);
  const inspect = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.inspect() : window.workbookEditor.inspect(), remote);
  const subtotal = () => page.evaluate(remote => remote ? window.workbookRemoteEditor.displayedCell(0, 3) : window.workbookEditor.displayedCell(0, 3), remote);
  for (const sheet of sheets) {
    await page.getByRole('tab', { name: sheet.name, exact: true }).click(); await settled();
    const address = page.getByLabel('选择单元格或区域', { exact: true }); await address.fill('D1'); await address.press('Enter'); await settled();
    expect(await subtotal()).toBe('10');
    await page.evaluate(async ({ remote, sheetId, formatId }) => {
      const view = remote ? window.workbookRemoteEditor : window.workbookEditor;
      await view.perform([{ op: 'setWorkbookCellFormats', sheetId, cells: [{ row: 1, column: 0, formatId }] }]);
    }, { remote, sheetId: sheet.sheetId, formatId });
    expect(await subtotal()).toBe('0'); expect(await inspect()).toMatchObject({ errors: [], history: { undoSteps: 1 } });
    await page.getByRole('button', { name: '撤销', exact: true }).click(); await settled(); expect(await subtotal()).toBe('10');
    expect(await inspect()).toMatchObject({ errors: [], history: { undoSteps: 0 } });
  }
}

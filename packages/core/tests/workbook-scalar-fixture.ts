import { createHash } from 'node:crypto';
import { createWorkbook, openDocument } from '../src/index.js';
import type { WorkbookDocumentMetadata, WorkbookManifest, WorkbookValue } from '../src/index.js';

const encoder = new TextEncoder(), decoder = new TextDecoder();
const escape = (value: string) => value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('"', '&quot;');
type Cell = { row: number; column: number; value: WorkbookValue; formula?: string; cacheState?: 'imported' | 'valid' | 'stale' | 'unavailable' };
export function scalarWorkbook(cells: readonly Cell[], edit: (metadata: WorkbookDocumentMetadata) => void = () => {}) {
  const files = new Map(createWorkbook({ documentId: 'scalar-kernel' })), manifest = JSON.parse(decoder.decode(files.get('manifest.json')!)) as WorkbookManifest;
  const metadata = JSON.parse(decoder.decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata; edit(metadata);
  files.set('metadata.json', encoder.encode(JSON.stringify(metadata)));
  const chunks = new Map<number, Cell[]>();
  for (const cell of cells) { const key = Math.floor(cell.row / 256) * 512 + Math.floor(cell.column / 32); chunks.set(key, [...chunks.get(key) ?? [], cell]); }
  for (const [key, values] of [...chunks].sort((a, b) => a[0] - b[0])) {
    const row = Math.floor(key / 512) * 256, column = key % 512 * 32, path = `sheets/sheet-1/${row}-${column}.html`, rows = new Map<number, Cell[]>();
    for (const cell of values) rows.set(cell.row, [...rows.get(cell.row) ?? [], cell]);
    const html = '<!doctype html><html><head></head><body><table data-sheet-id="sheet-1" data-row="' + row + '" data-column="' + column + '"><tbody>'
      + [...rows].sort((a, b) => a[0] - b[0]).map(([row, values]) => '<tr data-row="' + row + '">' + values.sort((a, b) => a.column - b.column).map(cell => {
        const value = cell.value, scalar = value.type === 'number' || value.type === 'boolean' || value.type === 'error' ? ` data-value="${escape(String(value.value))}"` : '';
        return `<td data-column="${cell.column}" data-value-type="${value.type}"${scalar}${cell.formula ? ` data-formula="${escape(cell.formula)}" data-cache-state="${cell.cacheState ?? 'imported'}"` : ''}>${value.type === 'blank' ? '' : escape(String(value.value))}</td>`;
      }).join('') + '</tr>').join('') + '</tbody></table></body></html>';
    const bytes = encoder.encode(html); files.set(path, bytes); manifest.chunks.push({ path, sheetId: 'sheet-1', row, column, sha256: createHash('sha256').update(bytes).digest('hex') });
  }
  files.set('manifest.json', encoder.encode(JSON.stringify(manifest))); return openDocument(files, { kind: 'workbook' });
}

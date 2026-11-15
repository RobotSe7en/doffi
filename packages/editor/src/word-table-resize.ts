import { DocumentError } from '@doffi/core';
import type { DocumentEditingSession } from '@doffi/core';
interface TableCell {
    row: number;
    column: number;
    rows: number;
    columns: number;
}
/** Guides live outside the input root. One completed gesture commits one typed
 * table edit, with the same tracking, conflicts and undo as the property panel. */
export function bindWordTableResize(root: HTMLElement, session: DocumentEditingSession, onError: (error: unknown) => void, signal: AbortSignal): void {
    const document = root.ownerDocument, view = document.defaultView!, guides: HTMLButtonElement[] = [];
    let active: HTMLElement | undefined, dragging = false;
    const clear = (): void => { for (const guide of guides.splice(0))
        guide.remove(); active = undefined; };
    signal.addEventListener('abort', clear, { once: true });
    function show(cell: HTMLElement): void {
        clear();
        active = cell;
        const table = cell.closest<HTMLTableElement>('table'), info = session.snapshot().metadata.word?.tables.find(item => item.nodeId === table?.dataset.nodeId);
        if (!table || !info)
            return;
        const source = session.getNode(info.nodeId)!, rows = source.childNodeIds.flatMap(id => { const node = session.getNode(id)!; return node.tagName === 'tr' ? [id] : node.childNodeIds.filter(id => session.getNode(id)!.tagName === 'tr'); }), occupied: string[][] = Array.from({ length: rows.length }, () => []), cells = new Map<string, TableCell>();
        for (const [row, rowId] of rows.entries())
            for (const id of session.getNode(rowId)!.childNodeIds) {
                const node = session.getNode(id)!;
                if (!['td', 'th'].includes(node.tagName))
                    continue;
                let column = 0;
                while (occupied[row]![column])
                    column++;
                const height = Number(node.attributes.rowspan ?? 1) || rows.length - row, width = Number(node.attributes.colspan ?? 1);
                cells.set(id, { row, column, rows: height, columns: width });
                for (let y = row; y < row + height; y++)
                    for (let x = column; x < column + width; x++)
                        occupied[y]![x] = id;
            }
        const selection = cells.get(cell.dataset.nodeId!);
        if (!selection)
            return;
        const rect = cell.getBoundingClientRect(), grid = [...info.grid], column = selection.column + selection.columns - 1, row = selection.row + selection.rows - 1, ratio = grid.reduce((a, b) => a + b, 0) / table.getBoundingClientRect().width,bidi=table.dataset.sdkWordTableBidi==='true';
        const heights = rows.map(id => { const element = root.querySelector<HTMLElement>('[data-node-id="' + view.CSS.escape(id) + '"]'); return Math.max(1, (element?.getBoundingClientRect().height ?? 0) * .75 || 20); });
        const commit = (axis: 'column' | 'row', delta: number, baseRevision: number): void => { if (Math.abs(delta) < .01)
            return; try {
            const snapshot = session.snapshot();
            if (snapshot.manifest.revision !== baseRevision)
                throw new DocumentError('RevisionConflict', '拖动期间文档已更改，请重试。');
            if(axis==='column'&&bidi)delta=-delta;
            const change = axis === 'column' ? { kind: 'resize' as const, grid: [...grid] } : { kind: 'resize' as const, rowIndex:row, rowHeight:heights[row]! };
            if ('grid' in change) {
                const minimum=Math.min(6,grid[column]!,grid[column+1]??6),maximum=column+1<grid.length?grid[column+1]!-minimum:Infinity,amount=Math.max(minimum-grid[column]!,Math.min(maximum,delta));
                change.grid[column]! += amount;
                if (column + 1 < grid.length)
                    change.grid[column + 1]! -= amount;
            }
            else
                change.rowHeight = Math.max(1, heights[row]! + delta);
            session.commit(session.prepare({ protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision, operations: [{ op: 'editWordTable', nodeId: info!.nodeId, change }] }));
        }
        catch (error) {
            onError(error);
        } };
        for (const axis of ['column', 'row'] as const) {
            const guide = document.createElement('button');
            guide.type = 'button';
            guide.setAttribute('aria-label', axis === 'column' ? '调整第 ' + (column + 1) + ' 列宽度' : '调整第 ' + (row + 1) + ' 行高度');
            guide.dataset.sdkTableResize = axis;
            guide.style.cssText = `position:fixed;z-index:1000;box-sizing:border-box;border:1px solid #2563eb;background:#2563eb66;padding:0;border-radius:2px;cursor:${axis === 'column' ? 'col' : 'row'}-resize;left:${axis === 'column' ? (bidi?rect.left:rect.right) - 4 : rect.left + Math.max(0, rect.width / 2 - 12)}px;top:${axis === 'column' ? rect.top + Math.max(0, rect.height / 2 - 12) : rect.bottom - 4}px;width:${axis === 'column' ? 8 : 24}px;height:${axis === 'column' ? 24 : 8}px`;
            document.body.append(guide);
            guides.push(guide);
            guide.addEventListener('keydown', event => { const negative = axis === 'column' ? 'ArrowLeft' : 'ArrowUp', positive = axis === 'column' ? 'ArrowRight' : 'ArrowDown'; if (event.key !== negative && event.key !== positive)
                return; event.preventDefault(); commit(axis, (event.key === negative ? -1 : 1) * (event.shiftKey ? .25 : 1), session.snapshot().manifest.revision); }, { signal });
            guide.addEventListener('pointerdown', event => {
                if (event.button !== 0)
                    return;
                event.preventDefault();
                event.stopPropagation();
                guide.focus({preventScroll:true});
                dragging = true;
                guide.setPointerCapture(event.pointerId);
                const start = axis === 'column' ? event.clientX : event.clientY, baseRevision = session.snapshot().manifest.revision;
                let delta = 0;
                const property = axis === 'column' ? 'left' : 'top', initial = parseFloat(guide.style[property]);
                const move = (next: PointerEvent): void => { const distance = (axis === 'column' ? next.clientX : next.clientY) - start; delta = Math.round(distance * (axis === 'column' ? ratio : .75) * 4) / 4; guide.style[property] = initial + distance + 'px'; };
                const end = (): void => { guide.removeEventListener('pointermove', move); guide.removeEventListener('pointerup', end); guide.removeEventListener('pointercancel', cancel); guide.removeEventListener('keydown', escape); dragging = false; guide.style[property] = initial + 'px'; commit(axis, delta, baseRevision); };
                const cancel = (): void => { delta = 0; end(); };
                const escape = (event: KeyboardEvent): void => { if (event.key === 'Escape')
                    cancel(); };
                guide.addEventListener('pointermove', move);
                guide.addEventListener('pointerup', end, { once: true });
                guide.addEventListener('pointercancel', cancel, { once: true });
                guide.addEventListener('keydown', escape);
                signal.addEventListener('abort', cancel, { once: true });
            }, { signal });
        }
    }
    document.addEventListener('pointermove', event => { if (dragging || guides.some(guide => guide.contains(event.target as Node)))
        return; const element = event.target as Element, cell = element.closest?.<HTMLElement>('td,th'); if (cell && root.contains(cell)&&!cell.closest('[data-word-drawing]')) {
        if (cell !== active)
            show(cell);
    }
    else if (active)
        clear(); }, { signal });
}

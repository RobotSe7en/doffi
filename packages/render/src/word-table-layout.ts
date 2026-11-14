import { DocumentError } from '@doffi/core';
import type { RenderLimits } from './contracts.js';
import { markViewNode, ownViewStyle, viewNodeRole } from './view-nodes.js';
/** Exact row heights clip the view; all hidden cell content remains in source. */
export function composeWordTableLayout(document: Document, limits: Readonly<RenderLimits>): void {
    const view = document.defaultView!;
    let count = document.querySelectorAll('*').length;
    for (const table of document.querySelectorAll<HTMLTableElement>('table')) {
        const rows = [...table.rows], heights = rows.map(row => { const css = view.getComputedStyle(row), rule = css.getPropertyValue('--word-row-height-rule').trim(); if (rule === 'auto') {
            ownViewStyle(row, 'height');
            row.style.height = 'auto';
            return undefined;
        } if (rule !== 'exact')
            return undefined; const raw = row.style.height, match = /^([\d.]+)(pt|px|in|cm|mm|pc|em|rem)$/.exec(raw); if (!match)
            return undefined; const factors: Record<string, number> = { pt: 4 / 3, px: 1, in: 96, cm: 96 / 2.54, mm: 96 / 25.4, pc: 16, em: parseFloat(css.fontSize), rem: parseFloat(view.getComputedStyle(document.documentElement).fontSize) }; return Number(match[1]) * factors[match[2]!]!; });
        for (const [index, row] of rows.entries())
            for (const cell of row.cells) {
                const span = cell.rowSpan || rows.length - index, covered = heights.slice(index, index + span);
                if (!covered.some(height => height !== undefined))
                    continue;
                const target = covered.reduce<number>((sum, height, offset) => sum + (height ?? rows[index + offset]!.getBoundingClientRect().height), 0), css = view.getComputedStyle(cell), padding = (parseFloat(css.paddingTop) || 0) + (parseFloat(css.paddingBottom) || 0), border = (parseFloat(css.borderTopWidth) || 0) + (parseFloat(css.borderBottomWidth) || 0), box = markViewNode(document.createElement('div'), 'unwrap'), content = markViewNode(document.createElement('div'), 'unwrap');
                count += 2;
                if (count > limits.maxNodes)
                    throw new DocumentError('ResourceLimitExceeded', 'Fixed row clipping exceeds the rendering node budget.');
                box.className = 'sdk-word-row-clip';
                box.style.cssText = `height:${Math.max(0, target - padding - border)}px;min-height:0;overflow:hidden;display:flex;flex-direction:column;justify-content:${css.verticalAlign === 'middle' ? 'center' : css.verticalAlign === 'bottom' ? 'flex-end' : 'flex-start'}`;
                content.style.cssText = 'display:flow-root;flex:none;min-height:0';
                content.append(...cell.childNodes);
                box.append(content);
                cell.append(box);
            }
    }
}
/** Exclude clipped geometry from pagination and source hit projections. */
export function wordRowVisibleRect(owner: Element, rect: DOMRect): DOMRect | undefined {
    let left = rect.left, right = rect.right, top = rect.top, bottom = rect.bottom, clipped = false;
    for (let node: Element | null = owner; node; node = node.parentElement)
        if (node.classList.contains('sdk-word-row-clip') && viewNodeRole(node) === 'unwrap') {
            clipped = true;
            const clip = node.getBoundingClientRect();
            left = Math.max(left, clip.left);
            right = Math.min(right, clip.right);
            top = Math.max(top, clip.top);
            bottom = Math.min(bottom, clip.bottom);
        }
    if (right - left < .01 || bottom - top < .1)
        return;
    return clipped ? new DOMRect(left, top, right - left, bottom - top) : rect;
}

/** An explicit row override wins over the table style's repeat-header flag. */
export function wordTableRepeatHeader(row:HTMLTableRowElement):boolean{
 const direct=row.dataset.wordRepeatHeader;if(direct!==undefined)return direct==='true';
 return row.parentElement?.localName==='thead'||row.ownerDocument.defaultView!.getComputedStyle(row).getPropertyValue('--word-row-repeat-header').trim()==='1';
}

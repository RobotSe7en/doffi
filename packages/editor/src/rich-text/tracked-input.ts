import { isDocumentTextBoundary } from '@doffi/core';
import type { DocumentEditingSession, DocumentOperation, ContentDocumentModel,ReconciledTextEdit } from '@doffi/core';
import { sourceText, offsetAt } from './dom.js';
import type { SelectionBookmark, CaretBookmark } from './dom.js';
/** Text tracking uses the same canonical paragraph positions as native input;
 * composition steps collapse to one exact replacement at commit. */
export function trackedWordInput(session: DocumentEditingSession, previous: HTMLElement, next: HTMLElement, properties: ReadonlyMap<string, string>,textEdits:ReconciledTextEdit[]): DocumentOperation[] | undefined {
    if (!session.snapshot().metadata.word?.trackChanges)
        return;
    const selector = 'p,h1,h2,h3,h4,h5,h6', before = [...previous.querySelectorAll<HTMLElement>(selector)], after = [...next.querySelectorAll<HTMLElement>(selector)];
    if (previous.matches(selector))
        before.unshift(previous);
    if (next.matches(selector))
        after.unshift(next);
    const oldIds = before.map(node => node.dataset.nodeId), newIds = after.map(node => node.dataset.nodeId);
    if (JSON.stringify(oldIds) !== JSON.stringify(newIds))
        return[{op:'reconcileWordTracked',nodeId:previous.dataset.nodeId!,html:next.innerHTML,textEdits,date:new Date().toISOString()}];
    const operations: DocumentOperation[] = [];
    for (let index = 0; index < before.length; index++) {
        const old = sourceText(before[index]!, () => false), current = sourceText(after[index]!, () => false);
        if (old === current)
            continue;
        let from = 0;
        while (from < old.length && from < current.length && old[from] === current[from])
            from++;
        while (from > 0 && (!isDocumentTextBoundary(old, from) || !isDocumentTextBoundary(current, from)))
            from--;
        let oldEnd = old.length, end = current.length;
        while (oldEnd > from && end > from && old[oldEnd - 1] === current[end - 1]) {
            oldEnd--;
            end--;
        }
        while (oldEnd < old.length && (!isDocumentTextBoundary(old, oldEnd) || !isDocumentTextBoundary(current, end))) {
            oldEnd++;
            end++;
        }
        operations.push({ op: 'trackWordText', nodeId: oldIds[index]!, from, to: oldEnd, expectedText: old.slice(from, oldEnd), text: current.slice(from, end), date: new Date().toISOString(), ...(properties.size ? { properties: Object.fromEntries(properties) } : {}) });
    }
    return operations.length ? operations : undefined;
}
export function trackedSelection(value: SelectionBookmark | undefined, session: DocumentEditingSession, model: ContentDocumentModel): SelectionBookmark | undefined {
    if (!value)
        return;
    const known = new Set(session.snapshot().metadata.word?.revisions.map(revision => revision.revisionId ?? revision.nodeId)), deleted = model.snapshot().metadata.word?.revisions.filter(revision => revision.kind === 'delete' && !known.has(revision.revisionId ?? revision.nodeId)) ?? [];
    if (!deleted.length)
        return value;
    const document = new DOMParser().parseFromString(model.snapshot().html, 'text/html');
    const point = (entry: CaretBookmark): CaretBookmark => { if (!entry.nodeId)
        return entry; const owner = document.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(entry.nodeId) + '"]'); if (!owner)
        return entry; const ranges = deleted.flatMap(revision => { const node = document.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(revision.nodeId) + '"]'); return node && owner.contains(node) ? [{ from: offsetAt(owner, node, 0, () => false), length: sourceText(node, () => false).length }] : []; }).sort((a, b) => a.from - b.from); let extra = 0; for (const range of ranges)
        if (entry.offset >= range.from - extra)
            extra += range.length; return { ...entry, offset: entry.offset + extra }; };
    return { anchor: point(value.anchor), focus: point(value.focus) };
}

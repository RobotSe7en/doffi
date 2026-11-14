import type { DocumentEditingSession, DocumentOperation } from '@doffi/core';
export type WordRevisionDisplay = 'all' | 'final' | 'original';
export function createWordReviewControls(parent: ShadowRoot, before: HTMLElement, session: DocumentEditingSession, onError: (error: unknown) => void, display: () => WordRevisionDisplay, setDisplay: (mode: WordRevisionDisplay) => void, focus: (nodeId: string) => void): {
    open(): void;
    dispose(): void;
} {
    const document = before.ownerDocument;
    let panel: HTMLElement | undefined, controller: AbortController | undefined;
    const close = (): void => { controller?.abort(); panel?.remove(); panel = undefined; };
    function open(): void {
        close();
        controller = new AbortController();
        const signal = controller.signal, snapshot = session.snapshot(), word = snapshot.metadata.word, root = new DOMParser().parseFromString(snapshot.html, 'text/html').querySelector<HTMLElement>('[data-office-kind="document"]')!.dataset.nodeId!;
        panel = document.createElement('section');
        panel.setAttribute('role', 'region');
        panel.setAttribute('aria-label', 'Word 修订审阅');
        panel.style.cssText = 'padding:12px;background:white;border:1px solid #dbe3ef;max-height:340px;overflow:auto';
        parent.insertBefore(panel, before);
        const apply = (operations: DocumentOperation[]): void => { session.commit(session.prepare({ protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision, operations })); open(); };
        const button = (name: string, run: () => void, container: HTMLElement = panel!): void => { const node = document.createElement('button'); node.type = 'button'; node.textContent = name; node.addEventListener('click', () => { try {
            run();
        }
        catch (error) {
            onError(error);
        } }, { signal }); container.append(node); };
        button('关闭修订审阅', close);
        button('刷新修订列表', open);
        const tracking = document.createElement('input'), author = document.createElement('input'), label = document.createElement('label');
        tracking.type = 'checkbox';
        tracking.checked = word?.trackChanges ?? false;
        tracking.setAttribute('aria-label', '记录修订');
        label.textContent = '记录修订';
        label.append(tracking);
        author.setAttribute('aria-label', '修订作者');
        author.value = word?.trackingAuthor ?? '';
        panel.append(label, author);
        button('应用修订设置', () => apply([{ op: 'setWordReview', nodeId: root, trackChanges: tracking.checked, author: author.value }]));
        const mode = document.createElement('select');
        mode.setAttribute('aria-label', '修订显示');
        for (const [value, name] of [['all', '显示所有标记'], ['final', '显示最终状态'], ['original', '显示原始状态']] as const) {
            const option = document.createElement('option');
            option.value = value;
            option.textContent = name;
            mode.append(option);
        }
        mode.value = display();
        mode.addEventListener('change', () => setDisplay(mode.value as WordRevisionDisplay), { signal });
        panel.append(mode);
        const authors = document.createElement('select');
        authors.setAttribute('aria-label', '筛选修订作者');
        for (const name of ['全部作者', ...new Set(word?.revisions.map(revision => revision.author ?? ''))]) {
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name || '未署名';
            authors.append(option);
        }
        panel.append(authors);
        const all = (decision: 'accept' | 'reject'): void => apply([{ op: 'reviewWordRevisions', nodeId: root, decision, ...(authors.selectedIndex ? { author: authors.value } : {}) }]);
        button('接受全部显示的修订', () => all('accept'));
        button('拒绝全部显示的修订', () => all('reject'));
        const list = document.createElement('div');
        panel.append(list);
        const labels = { insert: '插入', delete: '删除', moveFrom: '移动来源', moveTo: '移动目标', format: '格式更改', paragraphInsert: '插入段落标记', paragraphDelete: '删除段落标记', rowInsert: '插入行', rowDelete: '删除行', cellInsert: '插入单元格', cellDelete: '删除单元格' };
        const render = (): void => { list.replaceChildren(); for (const revision of word?.revisions ?? []) {
            if (authors.selectedIndex && authors.value !== (revision.author ?? ''))
                continue;
            const row = document.createElement('div'), text = document.createElement('p');
            text.textContent = labels[revision.kind] + ' · ' + (revision.author || '未署名') + ' · ' + (session.getNode(revision.nodeId)?.text.slice(0, 80) || '标记');
            row.append(text);
            button('定位修订', () => focus(revision.nodeId), row);
            button('接受此修订', () => apply([{ op: 'reviewWordRevision', nodeId: revision.nodeId, revisionId: revision.revisionId ?? revision.nodeId, decision: 'accept' }]), row);
            button('拒绝此修订', () => apply([{ op: 'reviewWordRevision', nodeId: revision.nodeId, revisionId: revision.revisionId ?? revision.nodeId, decision: 'reject' }]), row);
            list.append(row);
        } };
        authors.addEventListener('change', render, { signal });
        render();
    }
    return { open, dispose: close };
}

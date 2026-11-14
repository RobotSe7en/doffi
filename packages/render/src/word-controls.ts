import type { ContentDocumentModel } from '@doffi/core';
import { ownViewAttribute, ownViewStyle } from './view-nodes.js';
export function composeWordControls(document: Document, model: ContentDocumentModel): void {
    for (const control of model.snapshot().metadata.word?.controls ?? []) {
        const node = document.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(control.nodeId) + '"]');
        if (!node)
            continue;
        for (const name of ['data-word-control-lock', 'aria-label', 'role'])
            ownViewAttribute(node, name);
        node.setAttribute('data-word-control-lock', control.lock ?? 'unlocked');
        node.setAttribute('aria-label', control.alias ?? control.tag ?? '内容控件');
        node.setAttribute('role', control.kind === 'checkBox' ? 'checkbox' : 'group');
        if (control.kind === 'checkBox') {
            ownViewAttribute(node, 'aria-checked');
            ownViewAttribute(node, 'tabindex');
            node.tabIndex = 0;
            node.setAttribute('aria-checked', String(control.checked ?? false));
        }
        if (['checkBox', 'dropDown', 'date'].includes(control.kind) || ['contentLocked', 'sdtContentLocked'].includes(control.lock ?? '')) {
            ownViewAttribute(node, 'contenteditable');
            node.contentEditable = 'false';
        }
        ownViewStyle(node, 'outline', 'border-radius');
        node.style.outline = '1px solid #b7c5d5';
        node.style.borderRadius = '2px';
    }
}

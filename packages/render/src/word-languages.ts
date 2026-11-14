import { DocumentError, wordScriptRole } from '@doffi/core';
import type { ContentDocumentModel, WordScriptRole } from '@doffi/core';
import type { RenderLimits } from './contracts.js';
import { ownViewAttribute, markViewNode, viewNodeRole } from './view-nodes.js';
/** Font/language role wrappers belong to the view and retain exact source text. */
export function composeWordLanguages(document: Document, model: ContentDocumentModel, limits: Readonly<RenderLimits>): void {
    if (!model.snapshot().metadata.word)
        return;
    for (const node of document.querySelectorAll<HTMLElement>('[data-sdk-word-script-view]'))
        if (viewNodeRole(node) === 'unwrap')
            node.replaceWith(...node.childNodes);
    const walker = document.createTreeWalker(document.body, 4), texts: Text[] = [];
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const parent = node.parentElement;
        if (parent?.closest('[data-office-kind="document"]') && !parent.closest('[data-word-drawing],[data-word-native-object],[data-word-equation],[data-word-note-reference],[data-word-content-control="checkBox"]'))
            texts.push(node as Text);
    }
    let count = document.querySelectorAll('*').length;
    for (const text of texts) {
        const parent = text.parentElement!, css = document.defaultView!.getComputedStyle(parent);
        if (!['ascii', 'hAnsi', 'eastAsia', 'cs'].some(role => css.getPropertyValue('--word-font-' + role).trim()) && !['val', 'eastAsia', 'bidi'].some(role => css.getPropertyValue('--word-language-' + role).trim()))
            continue;
        let role: WordScriptRole | undefined, part = '';
        const fragment = document.createDocumentFragment();
        const append = (): void => { if (!part || !role)
            return; if (++count > limits.maxNodes)
            throw new DocumentError('ResourceLimitExceeded', 'Script roles exceed the render node budget.'); const span = markViewNode(document.createElement('span'), 'unwrap'); span.dataset.sdkWordScriptView = ''; span.dataset.wordScript = role; span.style.fontFamily = 'var(--word-font-' + role + ',var(--word-font-ascii,inherit))'; if (role === 'cs') {
            const size = css.getPropertyValue('--word-font-size-cs').trim(), bold = css.getPropertyValue('--word-bold-cs').trim(), italic = css.getPropertyValue('--word-italic-cs').trim();
            if (size)
                span.style.fontSize = size;
            if (bold)
                span.style.fontWeight = bold === '1' ? '700' : '400';
            if (italic)
                span.style.fontStyle = italic === '1' ? 'italic' : 'normal';
        } span.textContent = part; fragment.append(span); part = ''; };
        for (const character of text.data) {
            const next = wordScriptRole(character, role);
            if (next !== role) {
                append();
                role = next;
            }
            part += character;
        }
        append();
        text.replaceWith(fragment);
    }
    for (const node of document.querySelectorAll<HTMLElement>('[data-node-id],[data-sdk-word-script-view]')) {
        const css = document.defaultView!.getComputedStyle(node), role = node.dataset.wordScript === 'eastAsia' ? 'eastAsia' : node.dataset.wordScript === 'cs' ? 'bidi' : 'val', raw = css.getPropertyValue('--word-language-' + role).trim() || css.getPropertyValue('--word-language-val').trim();
        if (!raw)
            continue;
        let language = raw;
        if (/^"/.test(raw))
            try {
                language = JSON.parse(raw) as string;
            }
            catch {
                continue;
            }
        else if (/^'[^']*'$/.test(raw))
            language = raw.slice(1, -1);
        if (!language || /[\u0000-\u0020\u007f]/.test(language))
            continue;
        ownViewAttribute(node, 'lang');
        node.lang = language;
    }
}

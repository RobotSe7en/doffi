import { DocumentError } from '@doffi/core';
import type { ContentDocumentModel, SourceStyleNode } from '@doffi/core';
import type { RenderLimits } from './contracts.js';
import { createSourceStyleResolver } from './source-style-probe.js';
import { markViewNode, ownViewStyle } from './view-nodes.js';

/** Native runs have one font box. HTML ancestors would otherwise retain their
 * old large strut, highlight, decoration or superscript after an inner reset.
 * Project only affected paragraphs; source wrappers and identities stay put. */
export function composeWordTextResets(document: Document, model: ContentDocumentModel, limits: Readonly<RenderLimits>): void {
    if (!model.snapshot().metadata.word) return;
    const paragraphs = new Set([...document.querySelectorAll<HTMLElement>('[data-word-reset-text="true"]')].flatMap(node => {
        const paragraph = node.closest<HTMLElement>('p,h1,h2,h3,h4,h5,h6'); return paragraph ? [paragraph] : [];
    }));
    if (!paragraphs.size) return;
    const resolver = createSourceStyleResolver(document), projections = new WeakMap<HTMLElement, SourceStyleNode>(), view = document.defaultView!;
    const project = (node: HTMLElement): SourceStyleNode => {
        let value = projections.get(node); if (!value) { value = { tag: node.localName, attrs: node.dataset.nodeId ? model.getNode(node.dataset.nodeId)?.attributes ?? {} : {}, parent: node.parentElement ? project(node.parentElement) : undefined }; projections.set(node, value); } return value;
    };
    const properties = ['font-family', 'font-size', 'font-weight', 'font-style', 'font-variant', 'text-transform', 'color', 'background-color', 'vertical-align', 'letter-spacing', 'line-height', 'white-space', 'text-decoration-line', 'text-decoration-style', 'text-decoration-color'];
    const paint: { text: Text; values: Record<string, string> }[] = [], containers = new Set<HTMLElement>();
    let count = document.querySelectorAll('*').length;
    for (const paragraph of paragraphs) {
        const walker = document.createTreeWalker(paragraph, 4);
        for (let node = walker.nextNode(); node; node = walker.nextNode()) {
            const parent = node.parentElement!;
            if (!node.textContent || parent.closest('math,[data-word-drawing],[data-word-native-object],[data-word-note-reference],[data-word-note-number],[data-word-note-separator],[data-word-equation]')) continue;
            const css = resolver.get(project(parent)), values = Object.fromEntries([...properties, ...Object.keys(css).filter(key => key.startsWith('--word-'))].flatMap(key => css[key] === undefined ? [] : [[key, css[key]!]]));
            // Review paint remains owned by ins/del; authored run resets do not
            // remove the separate insert/delete indication.
            if (parent.closest('ins,del')) values.color = view.getComputedStyle(parent).color;
            paint.push({ text: node as Text, values });
            for (let ancestor: HTMLElement | null = parent; ancestor && ancestor !== paragraph; ancestor = ancestor.parentElement)
                if (!['ins', 'del'].includes(ancestor.localName)) containers.add(ancestor);
        }
    }
    // Resolve every source style before adding any view styles or wrappers.
    for (const container of containers) {
        ownViewStyle(container, 'display', 'background-color', 'text-decoration-line', 'vertical-align');
        container.style.display = 'contents'; container.style.backgroundColor = 'transparent'; container.style.textDecorationLine = 'none'; container.style.verticalAlign = 'baseline';
    }
    for (const { text, values } of paint) {
        if (++count > limits.maxNodes) throw new DocumentError('ResourceLimitExceeded', 'Word text reset projection exceeds the render node budget.');
        const wrapper = markViewNode(document.createElement('span'), 'unwrap'); wrapper.className = 'sdk-word-text-run';
        for (const [name, value] of Object.entries(values)) wrapper.style.setProperty(name, value);
        text.before(wrapper); wrapper.append(text);
    }
}

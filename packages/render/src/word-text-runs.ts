import { DocumentError } from '@doffi/core';
import type { ContentDocumentModel, SourceStyleNode } from '@doffi/core';
import type { RenderLimits } from './contracts.js';
import { createSourceStyleResolver } from './source-style-probe.js';
import { markViewNode, ownViewStyle } from './view-nodes.js';

/** Native runs have one font box. HTML ancestors would otherwise retain their
 * old large strut, highlight, decoration or superscript after an inner edit.
 * Project only affected paragraphs; source wrappers and identities stay put. */
export function composeWordTextRuns(document: Document, model: ContentDocumentModel, limits: Readonly<RenderLimits>): void {
    if (!model.snapshot().metadata.word) return;
    const revisions=new Map(model.snapshot().metadata.word!.revisions.filter(revision=>['insert','delete','moveFrom','moveTo'].includes(revision.kind)).map(revision=>[revision.nodeId,['insert','moveTo'].includes(revision.kind)?'insert':'delete']));
    const paragraphs = new Set([...document.querySelectorAll<HTMLElement>('[data-word-reset-text="true"],[data-word-current-format="true"],[data-word-run-overlay="true"],ins,del')].flatMap(node => {
        if(['ins','del'].includes(node.localName)&&!revisions.has(node.dataset.nodeId??''))return[];
        const paragraph = node.closest<HTMLElement>('p,h1,h2,h3,h4,h5,h6'); return paragraph ? [paragraph] : [];
    }));
    if (!paragraphs.size) return;
    const resolver = createSourceStyleResolver(document,new Set(revisions.keys())), projections = new WeakMap<HTMLElement, SourceStyleNode>();
    const project = (node: HTMLElement): SourceStyleNode => {
        let value = projections.get(node); if (!value) { value = { tag: node.localName, attrs: node.dataset.nodeId ? model.getNode(node.dataset.nodeId)?.attributes ?? {} : {}, parent: node.parentElement ? project(node.parentElement) : undefined }; projections.set(node, value); } return value;
    };
    const properties = ['font-family', 'font-size', 'font-weight', 'font-style', 'font-variant', 'text-transform', 'color', 'background-color', 'vertical-align', 'letter-spacing', 'line-height', 'white-space', 'text-decoration-line', 'text-decoration-style', 'text-decoration-color'];
    const paint: { text: Text; values: Record<string, string>;review?:string|undefined }[] = [], containers = new Set<HTMLElement>();
    let count = document.querySelectorAll('*').length;
    for (const paragraph of paragraphs) {
        const walker = document.createTreeWalker(paragraph, 4);
        for (let node = walker.nextNode(); node; node = walker.nextNode()) {
            const parent = node.parentElement!;
            if (!node.textContent || parent.closest('math,[data-word-drawing],[data-word-native-object],[data-word-note-reference],[data-word-note-number],[data-word-note-separator],[data-word-equation]')) continue;
            const css = resolver.get(project(parent)), values = Object.fromEntries([...properties, ...Object.keys(css).filter(key => key.startsWith('--word-'))].flatMap(key => css[key] === undefined ? [] : [[key, css[key]!]]));
            const owner=parent.closest<HTMLElement>('ins,del'),review=owner?revisions.get(owner.dataset.nodeId??''):undefined;
            paint.push({ text: node as Text, values,review });
            for (let ancestor: HTMLElement | null = parent; ancestor && ancestor !== paragraph; ancestor = ancestor.parentElement)
                containers.add(ancestor);
        }
    }
    // Resolve every source style before adding any view styles or wrappers.
    for(const paragraph of paragraphs){ownViewStyle(paragraph,'text-decoration-line');paragraph.style.textDecorationLine='none';}
    for (const container of containers) {
        ownViewStyle(container, 'display', 'background-color', 'text-decoration-line', 'vertical-align');
        container.style.display = 'contents'; container.style.backgroundColor = 'transparent'; container.style.textDecorationLine = 'none'; container.style.verticalAlign = 'baseline';
    }
    for (const { text, values,review } of paint) {
        if (++count > limits.maxNodes) throw new DocumentError('ResourceLimitExceeded', 'Word text run projection exceeds the render node budget.');
        const wrapper = markViewNode(document.createElement('span'), 'unwrap'); wrapper.className = 'sdk-word-text-run';
        for (const [name, value] of Object.entries(values)) wrapper.style.setProperty(name, value);
        if(review){wrapper.dataset.sdkWordTextReview=review;const lines=(values['text-decoration-line']??'none').split(/\s+/).filter(line=>line!=='none');for(const[key,value]of Object.entries({'color':values.color??'#000000','background':values['background-color']??'transparent','decoration':lines.join(' ')||'none','insert-decoration':[...new Set([...lines,'underline'])].join(' '),'delete-decoration':[...new Set([...lines,'line-through'])].join(' ')}))wrapper.style.setProperty('--sdk-word-authored-'+key,value);}
        text.before(wrapper); wrapper.append(text);
    }
}

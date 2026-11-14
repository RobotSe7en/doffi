import { openDocument } from '@doffi/core';
import type { ContentDocumentModel, WordDocumentMetadata } from '@doffi/core';
/** A disposable format-only projection. Text, identities and the editing
 * session remain unchanged; SourceProjection restores authored attributes. */
export function wordReviewPreview(model: ContentDocumentModel): ContentDocumentModel {
    const snapshot = model.snapshot(), word = snapshot.metadata.word;
    if (!word?.revisions.some(revision => revision.kind === 'format'))
        return model;
    const metadata = JSON.parse(JSON.stringify(snapshot.metadata)) as WordDocumentMetadata, document = new DOMParser().parseFromString(snapshot.html, 'text/html');
    for (const revision of word.revisions) {
        if (revision.kind !== 'format' || !revision.before)
            continue;
        const before = revision.before, node = document.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(revision.nodeId) + '"]');
        if (!node)
            continue;
        for (const [key, value] of Object.entries(before.attributes ?? {}))
            if (value === null)
                node.removeAttribute(key);
            else
                node.setAttribute(key, value);
        for (const [key, value] of Object.entries(before.properties ?? {}))
            node.style.setProperty(key, value);
        if (before.grid) {
            const table = metadata.word!.tables.find(table => table.nodeId === revision.nodeId);
            if (table)
                table.grid = [...before.grid];
        }
        if (before.tabs) {
            const index = metadata.word!.tabs.findIndex(tab => tab.paragraphNodeId === revision.nodeId), tabs = { paragraphNodeId: revision.nodeId, defaultInterval: before.tabs.defaultInterval, stops: before.tabs.stops.map(stop => ({ ...stop })) };
            if (index < 0)
                metadata.word!.tabs.push(tabs);
            else
                metadata.word!.tabs[index] = tabs;
        }
        if (before.section) {
            const index = metadata.sections.findIndex(section => section.rootNodeId === revision.nodeId), current = metadata.sections[index];
            if (current)
                metadata.sections[index] = { sectionId: current.sectionId, rootNodeId: current.rootNodeId, ...JSON.parse(JSON.stringify(before.section)) as Omit<WordDocumentMetadata['sections'][number], 'sectionId'|'rootNodeId'> };
        }
    }
    const files = new Map(model.serialize()), encoder = new TextEncoder();
    files.set(snapshot.manifest.entrypoint, encoder.encode('<!doctype html>' + document.documentElement.outerHTML));
    files.set(snapshot.manifest.metadata, encoder.encode(JSON.stringify(metadata)));
    return openDocument(files, { kind: 'document' });
}

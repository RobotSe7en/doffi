import type { ContentDocumentModel } from '@doffi/core';
/** Explain atomic object edits whose current data lives in Word metadata. */
export function wordDiffDetails(model: ContentDocumentModel, nodeId: string): string {
    const metadata = model.snapshot().metadata, word = metadata.word;
    if (!word)
        return '';
    const values: Record<string, unknown> = {};
    for (const name of ['drawings', 'equations', 'controls', 'fields', 'tables', 'revisions', 'noteReferences'] as const) {
        const value = word[name]?.find(item => item.nodeId === nodeId);
        if (value) {
            const copy = JSON.parse(JSON.stringify(value)) as Record<string, unknown>;
            if (name === 'drawings' && copy.kind === 'chart')
                delete (copy.chart as Record<string, unknown>).native;
            values[name] = copy;
        }
    }
    const section = metadata.sections?.find(section => section.rootNodeId === nodeId);
    if (section)
        values.section = section;
    const story = word.stories.find(story => story.rootNodeId === nodeId);
    if (story)
        values.story = story;
    if (model.getNode(nodeId)?.parentNodeId === null) {
        values.styles = word.styles;
        values.numbering = word.numbering;
        values.pageSetup = metadata.pageSetup;
    }
    return Object.keys(values).length ? JSON.stringify(values, null, 2) : '';
}

import { DocumentError } from '../errors.js';
import type { WordDocumentMetadata } from '../contracts/document.js';
import type { WordOperation } from '../contracts/word-operations.js';
import type { DocumentEdit } from '../document/edit.js';
import type { ParsedHtml } from '../document/html.js';
export function applyWordWatermark(metadata: WordDocumentMetadata, edit: DocumentEdit, parse: () => ParsedHtml, operation: Extract<WordOperation, {
    op: 'setWordWatermark';
}>, apply: (operation: WordOperation) => void): void {
    const word = metadata.word!, section = metadata.sections.find(section => section.rootNodeId === operation.nodeId);
    if (!section)
        throw new DocumentError('InvalidOperation', 'A watermark requires its section.');
    const sectionIndex = metadata.sections.indexOf(section), variants = ['default', ...(section.differentFirstPage ? ['first'] : []), ...(word.evenAndOddHeaders ? ['even'] : [])] as ('default' | 'first' | 'even')[];
    const effective = (index: number, variant: typeof variants[number]): string | undefined => { for (let n = index; n >= 0; n--) {
        const id = metadata.sections[n]!.headers?.[variant];
        if (id)
            return id;
    } return; };
    const belongs = (nodeId: string, root: string): boolean => { for (let node = parse().peekNode(nodeId); node; node = node.parentNodeId ? parse().peekNode(node.parentNodeId) : undefined)
        if (node.nodeId === root)
            return true; return false; };
    for (const variant of variants) {
        let storyId = effective(sectionIndex, variant), story = word.stories.find(story => story.storyId === storyId);
        const existing = story ? word.drawings?.filter(drawing => drawing.watermark && belongs(drawing.nodeId, story!.rootNodeId)) ?? [] : [];
        if (!operation.watermark && !existing.length)
            continue;
        const shared = storyId && metadata.sections.some((_section, index) => index !== sectionIndex && effective(index, variant) === storyId);
        if (story && (!section.headers?.[variant] || shared)) {
            apply({ op: 'cloneWordStory', nodeId: section.rootNodeId, storyId: story.storyId, variant });
            storyId = section.headers![variant];
            story = word.stories.find(story => story.storyId === storyId);
        }
        else if (!story) {
            apply({ op: 'insertWordStory', nodeId: section.rootNodeId, kind: 'header', variant, html: '<p></p>' });
            storyId = section.headers![variant];
            story = word.stories.find(story => story.storyId === storyId);
        }
        for (const drawing of [...word.drawings ?? []])
            if (drawing.watermark && belongs(drawing.nodeId, story!.rootNodeId))
                apply({ op: 'removeWordDrawing', nodeId: drawing.nodeId });
        const value = operation.watermark;
        if (!value)
            continue;
        let paragraph = parse().peekNode(story!.rootNodeId)!.childNodeIds.find(id => /^p$|^h[1-6]$/.test(parse().peekNode(id)!.tagName));
        if (!paragraph)
            paragraph = edit.insertSource(story!.rootNodeId, '<p></p>').rootNodeIds[0]!;
        const setup = section.pageSetup ?? metadata.pageSetup, width = value.kind === 'image' ? value.width : setup.width * .8, height = value.kind === 'image' ? value.height : Math.max(value.fontSize * 2, setup.height * .15), before = new Set(word.drawings?.map(drawing => drawing.nodeId));
        apply({ op: 'insertWordDrawing', nodeId: paragraph, offset: 0, geometry: { x: 0, y: 0, width, height, rotation: value.rotation ?? (value.kind === 'text' ? -45 : 0), flipH: false, flipV: false }, layout: { mode: 'floating', horizontal: 'page', vertical: 'page', alignH: 'center', alignV: 'center', wrap: 'none', behindText: true, allowOverlap: true }, object: value.kind === 'text' ? { kind: 'textBox', text: value.text } : { kind: 'image', assetPath: value.assetPath, alt: value.alt } });
        const drawing = word.drawings!.find(drawing => !before.has(drawing.nodeId))!;
        drawing.watermark = true;
        drawing.name = 'DoffiWatermark-' + drawing.nodeId;
        if (value.kind === 'text' && drawing.kind === 'textBox') {
            const root = word.stories.find(story => story.storyId === drawing.textStoryId)!.rootNodeId, p = parse().peekNode(root)!.childNodeIds[0]!;
            apply({ op: 'setWordDrawingText', nodeId: drawing.nodeId, text: { vertical: 'middle', wrap: false, insets: { top: 0, left: 0, right: 0, bottom: 0 } } });
            apply({ op: 'setWordParagraph', nodeId: p, properties: { 'text-align': 'center', 'margin-top': '0pt', 'margin-bottom': '0pt','line-height':value.fontSize*1.2+'pt' } });
            apply({ op: 'formatWordText', nodeId: p, from: 0, to: value.text.length, expectedText: value.text, properties: { 'font-family': JSON.stringify(value.fontFamily), ...Object.fromEntries(['ascii', 'hAnsi', 'eastAsia', 'cs'].map(role => ['--word-font-' + role, JSON.stringify(value.fontFamily)])), 'font-size': value.fontSize + 'pt', color: value.color } });
        }
    }
}

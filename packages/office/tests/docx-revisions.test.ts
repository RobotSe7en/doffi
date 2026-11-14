import { expect, test } from 'vitest';
import { strToU8, strFromU8, zipSync, unzipSync } from 'fflate';
import { createDocumentSession, openDocument } from '@doffi/core';
import { exportDocx, importDocx } from '@doffi/office';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function fixture(body: string): Uint8Array { return zipSync({ '[Content_Types].xml': strToU8('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'), '_rels/.rels': strToU8('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'), 'word/document.xml': strToU8('<w:document xmlns:w="' + W + '" xmlns:x="urn:future"><w:body>' + body + '</w:body></w:document>') }); }
function formatting() { return fixture('<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="8" w:author="甲"><w:pPr><w:jc w:val="right"/></w:pPr></w:pPrChange></w:pPr><w:r><w:rPr><w:b/><x:after/><w:rPrChange w:id="9" w:author="甲"><w:rPr><w:i/><x:before/></w:rPr></w:rPrChange></w:rPr><w:t>文字漢\u{e0100}</w:t></w:r></w:p>'); }
test('native format revisions preserve previous properties and unrelated residuals across native reopens', () => { let bytes = formatting(); for (let pass = 0; pass < 2; pass++) {
    const model = openDocument(importDocx(bytes, { documentId: 'format' }), { kind: 'document' });
    expect(model.snapshot().metadata.word!.revisions.map(item => item.property)).toEqual(['run', 'paragraph']);
    const next = unzipSync(exportDocx(model).bytes);
    expect(strFromU8(next['word/document.xml']!)).toContain('rPrChange');
    expect(strFromU8(next['word/document.xml']!)).toContain('pPrChange');
    expect(strFromU8(next['word/document.xml']!)).toContain('before');
    bytes = exportDocx(model).bytes;
} });
test.each(['accept', 'reject'] as const)('reviewing %s restores or removes the correct known and unknown formatting, then reopens', decision => { const session = createDocumentSession(openDocument(importDocx(formatting(), { documentId: 'format' }), { kind: 'document' })), root = session.snapshot().metadata.word!.stories.find(story => story.kind === 'main')!; let article = root.rootNodeId; while (session.getNode(article)!.parentNodeId)
    article = session.getNode(article)!.parentNodeId!; session.commit(session.prepare({ protocolVersion: 1, documentId: 'format', baseRevision: 0, operations: [{ op: 'reviewWordRevisions', nodeId: article, decision }] })); const xml = strFromU8(unzipSync(exportDocx(session).bytes)['word/document.xml']!); expect(xml).not.toContain('PrChange'); expect(/<(?:[\w.-]+:)?before(?:\s|\/|>)/.test(xml)).toBe(decision === 'reject'); expect(/<(?:[\w.-]+:)?after(?:\s|\/|>)/.test(xml)).toBe(decision === 'accept'); expect(xml).toContain('w:val="' + (decision === 'reject' ? 'right' : 'center') + '"'); const model = openDocument(importDocx(exportDocx(session).bytes, { documentId: 'again' }), { kind: 'document' }); expect(model.snapshot().metadata.word!.revisions).toEqual([]); session.undo(); expect(session.snapshot().metadata.word!.revisions).toHaveLength(2); });

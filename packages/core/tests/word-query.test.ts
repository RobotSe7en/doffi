import {expect,test}from 'vitest';
import {createWordDocument,createDocumentSession,openDocument,findWordText,getWordOutline,getWordTextStatistics}from '@doffi/core';
function document(){const session=createDocumentSession(openDocument(createWordDocument({documentId:'query'}),{kind:'document'}));session.commit(session.prepare({protocolVersion:1,documentId:'query',baseRevision:0,operations:[{op:'replaceFragment',nodeId:'section-root-1',html:'<h1>标题😀</h1><p>Office office办公 OfficeX</p><p>office 😀 office</p>'}]}));return session;}
test('Word search preserves UTF-16 offsets and supports case and word boundaries',()=>{
 const model=document();expect(findWordText(model,'office')).toHaveLength(5);expect(findWordText(model,'office',{matchCase:true})).toHaveLength(3);expect(findWordText(model,'Office',{wholeWord:true})).toHaveLength(3);
 const emoji=findWordText(model,'😀');expect(emoji).toHaveLength(2);expect(emoji[0]!.to-emoji[0]!.from).toBe(2);expect(findWordText(model,'')).toEqual([]);
});
test('whole replacement shares one history step and title navigation is derived from source',()=>{
 const session=document(),before=session.snapshot(),matches=findWordText(session,'office');session.commit(session.prepare({protocolVersion:1,documentId:'query',baseRevision:1,operations:[...matches].reverse().map(match=>({op:'replaceText' as const,...match,text:'Word'}))}));expect(findWordText(session,'office')).toEqual([]);session.undo();expect(session.snapshot().html).toBe(before.html);expect(getWordOutline(session).map(h=>[h.level,h.text])).toEqual([[1,'标题😀']]);expect(getWordTextStatistics(session).paragraphs).toBe(3);
});
test('literal regex punctuation is searched without interpretation or paragraph crossing',()=>{
 const session=document();expect(findWordText(session,'Office.*')).toEqual([]);expect(findWordText(session,'OfficeXoffice')).toEqual([]);
});

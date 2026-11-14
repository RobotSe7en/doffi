import { DocumentError } from '../errors.js';
import type { WordMetadata } from '../contracts/word.js';
import type { WordReferenceOperation } from '../contracts/word-reference.js';
import type { DocumentEdit } from '../document/edit.js';
import type { ParsedHtml } from '../document/html.js';
function fail(message:string):never{throw new DocumentError('InvalidOperation',message);}
const escape=(text:string):string=>text.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replace(/\r\n?/g,'\n').replaceAll('\n','<br>');
export function isWordReferenceOperation(operation:{readonly op:string}):operation is WordReferenceOperation{return['insertWordNote','removeWordNote','convertWordNote','setWordNoteOptions','insertWordBookmark','removeWordBookmark','insertWordLink','setWordLink','removeWordLink'].includes(operation.op);}
export function applyWordReference(word:WordMetadata,edit:DocumentEdit,parse:()=>ParsedHtml,operation:WordReferenceOperation):void{
 const tree=parse(),node=tree.peekNode(operation.nodeId);if(!node)fail('The Word reference target is missing.');
 if(operation.op==='setWordNoteOptions'){if(node.parentNodeId!==null)fail('Shared note settings require the document root.');word[operation.kind==='footnote'?'footnoteOptions':'endnoteOptions']=operation.options;return;}
 if(operation.op==='insertWordBookmark'||operation.op==='insertWordLink'){
  if(!/^p$|^h[1-6]$/.test(node.tagName))fail('Insert a bookmark or link in a source paragraph.');const text=tree.getNode(operation.nodeId)!.text;if(operation.from>operation.to||text.slice(operation.from,operation.to)!==operation.expectedText)fail('The selected reference text no longer matches.');if(operation.op==='insertWordBookmark'&&!/^[\p{L}_][\p{L}\p{N}_]{0,39}$/u.test(operation.name))fail('A bookmark name must start with a letter or underscore and contain at most 40 letters, digits or underscores.');
  if(operation.op==='insertWordLink'&&!operation.expectedText.includes('\ufffc'))edit.formatSourceText(operation.nodeId,operation,{color:'#0563c1','text-decoration':'underline'});
  const attribute=operation.op==='insertWordBookmark'?'data-word-bookmark':'data-word-link';const end=edit.insertSourceAt(operation.nodeId,operation.to,'<span '+attribute+'="end"></span>').rootNodeIds[0]!,start=(operation.from===operation.to?edit.insertSource(parse().peekNode(end)!.parentNodeId!,'<span '+attribute+'="start"></span>',end):edit.insertSourceAt(operation.nodeId,operation.from,'<span '+attribute+'="start"></span>')).rootNodeIds[0]!;
  if(operation.op==='insertWordBookmark')word.bookmarks.push({bookmarkId:start,name:operation.name,startNodeId:start,endNodeId:end});else word.hyperlinks.push({nodeId:start,endNodeId:end,...(operation.target?{target:operation.target}:{}),...(operation.bookmark?{bookmark:operation.bookmark}:{})});return;
 }
 if(operation.op==='removeWordBookmark'){
  if(node.parentNodeId!==null)fail('Removing a shared bookmark requires the document root.');const bookmark=word.bookmarks.find(bookmark=>bookmark.bookmarkId===operation.bookmarkId);if(!bookmark)fail('The bookmark is missing.');for(const link of word.hyperlinks.filter(link=>link.bookmark===bookmark.name&&!link.target)){if(link.endNodeId){edit.removeSource(link.nodeId);edit.removeSource(link.endNodeId);}else{edit.attributes(link.nodeId,{'data-word-link':undefined});}}word.hyperlinks=word.hyperlinks.filter(link=>link.bookmark!==bookmark.name||!!link.target);edit.removeSource(bookmark.startNodeId);edit.removeSource(bookmark.endNodeId);word.bookmarks=word.bookmarks.filter(item=>item!==bookmark);return;
 }
 if(operation.op==='setWordLink'||operation.op==='removeWordLink'){
  const link=word.hyperlinks.find(link=>link.nodeId===operation.nodeId);if(!link)fail('The hyperlink is missing.');if(operation.op==='setWordLink'){delete link.target;delete link.bookmark;if(operation.target)link.target=operation.target;if(operation.bookmark)link.bookmark=operation.bookmark;}else{if(link.endNodeId){edit.removeSource(link.nodeId);edit.removeSource(link.endNodeId);}else edit.unwrapSource(link.nodeId);word.hyperlinks=word.hyperlinks.filter(item=>item!==link);}return;
 }
 if(operation.op==='insertWordNote'){
  if(!/^p$|^h[1-6]$/.test(node.tagName))fail('Insert a footnote or endnote in a source paragraph.');
  const roots=new Set(word.stories.filter(story=>story.kind==='main').map(story=>story.rootNodeId));let owner:string|null=operation.nodeId;while(owner&&!roots.has(owner))owner=tree.peekNode(owner)!.parentNodeId;if(!owner)fail('Footnote and endnote references require the main story.');
  const article=tree.nodeIds().find(id=>tree.peekNode(id)!.parentNodeId===null)!,storyId=edit.insertSource(article,'<div data-office-story="'+operation.kind+'"><p><span data-word-note-number="true">\ufffc</span> '+escape(operation.text)+'</p></div>').rootNodeIds[0]!,noteId=storyId;
  word.stories.push({storyId,rootNodeId:storyId,kind:operation.kind});(word.notes??=[]).push({noteId,storyId,kind:operation.kind});
  const marker=edit.insertSourceAt(operation.nodeId,operation.offset,'<span data-word-note-reference="'+operation.kind+'">\ufffc</span>').rootNodeIds[0]!;(word.noteReferences??=[]).push({nodeId:marker,noteId,...(operation.customMark?{customMark:operation.customMark}:{})});return;
 }
 const reference=word.noteReferences?.find(item=>item.nodeId===operation.nodeId),note=word.notes?.find(item=>item.noteId===reference?.noteId);if(!reference||!note)fail('The note reference is missing.');
 if(operation.op==='removeWordNote'){edit.removeSource(operation.nodeId);return;}
 note.kind=operation.kind;const story=word.stories.find(story=>story.storyId===note.storyId)!;story.kind=operation.kind;edit.attributes(story.rootNodeId,{'data-office-story':operation.kind});for(const marker of word.noteReferences??[])if(marker.noteId===note.noteId)edit.attributes(marker.nodeId,{'data-word-note-reference':operation.kind});
}

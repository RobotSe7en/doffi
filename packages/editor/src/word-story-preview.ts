import { DocumentError, openDocument } from '@doffi/core';
import type { ContentDocumentModel, WordDocumentMetadata } from '@doffi/core';
/** A temporary rendering projection for one auxiliary story. The editing
 * session always retains the complete original document as its authority. */
export function wordStoryPreview(model:ContentDocumentModel,storyId:string|undefined):ContentDocumentModel{
 if(!storyId)return model;
 const snapshot=model.snapshot(),word=snapshot.metadata.word,story=word?.stories.find(story=>story.storyId===storyId);
 if(!story)throw new DocumentError('InvalidOperation','The requested Word story does not exist.');if(story.kind==='main')return model;
 const document=new DOMParser().parseFromString(snapshot.html,'text/html'),article=document.querySelector<HTMLElement>('[data-office-kind="document"]')!,root=document.querySelector<HTMLElement>('[data-node-id="'+CSS.escape(story.rootNodeId)+'"]')!;
 const roots=[root],stories=new Set([story.storyId]);for(let i=0;i<roots.length;i++)for(const drawing of word!.drawings??[]){if(!('textStoryId'in drawing)||!drawing.textStoryId||stories.has(drawing.textStoryId)||!roots[i]!.querySelector('[data-node-id="'+CSS.escape(drawing.nodeId)+'"]'))continue;const definition=word!.stories.find(story=>story.storyId===drawing.textStoryId)!,text=document.querySelector<HTMLElement>('[data-node-id="'+CSS.escape(definition.rootNodeId)+'"]')!;stories.add(definition.storyId);roots.push(text);}
 article.replaceChildren(...roots);root.setAttribute('data-office-story','main');
 const ids=new Set([article.dataset.nodeId!,...roots.flatMap(root=>[root.dataset.nodeId!,...[...root.querySelectorAll<HTMLElement>('[data-node-id]')].map(node=>node.dataset.nodeId!)])]);
 const metadata=JSON.parse(JSON.stringify(snapshot.metadata)) as WordDocumentMetadata,definitions=metadata.word!;delete definitions.comments;
 if(definitions.noteReferences)definitions.noteReferences=definitions.noteReferences.filter(reference=>ids.has(reference.nodeId));if(definitions.notes)definitions.notes=definitions.notes.filter(note=>definitions.noteReferences?.some(reference=>reference.noteId===note.noteId));
 if(definitions.equations)definitions.equations=definitions.equations.filter(equation=>ids.has(equation.nodeId));if(definitions.controls)definitions.controls=definitions.controls.filter(control=>ids.has(control.nodeId));
 definitions.stories=[{storyId:story.storyId,kind:'main',rootNodeId:story.rootNodeId},...definitions.stories.filter(item=>item.storyId!==story.storyId&&stories.has(item.storyId))];if(definitions.drawings)definitions.drawings=definitions.drawings.filter(item=>ids.has(item.nodeId));definitions.fields=definitions.fields.filter(field=>ids.has(field.nodeId));definitions.tables=definitions.tables.filter(table=>ids.has(table.nodeId));definitions.revisions=definitions.revisions.filter(revision=>ids.has(revision.nodeId));definitions.tabs=definitions.tabs.filter(tab=>ids.has(tab.paragraphNodeId));definitions.bookmarks=definitions.bookmarks.filter(bookmark=>ids.has(bookmark.startNodeId)&&ids.has(bookmark.endNodeId));
 const bookmarks=new Set(definitions.bookmarks.map(bookmark=>bookmark.name));definitions.hyperlinks=definitions.hyperlinks.filter(link=>{if(!ids.has(link.nodeId))return false;if(!link.target&&link.bookmark&&!bookmarks.has(link.bookmark)){root.querySelector('[data-node-id="'+CSS.escape(link.nodeId)+'"]')?.removeAttribute('data-word-link');return false;}return true;});
 if(definitions.nativeContent)definitions.nativeContent.fragments=definitions.nativeContent.fragments.filter(fragment=>!fragment.nodeId||ids.has(fragment.nodeId));
 const section=metadata.sections.find(section=>Object.values(section.headers??{}).includes(storyId)||Object.values(section.footers??{}).includes(storyId));
 metadata.pageSetup=section?.pageSetup??metadata.pageSetup;metadata.sections=[{sectionId:'story-editor-section',rootNodeId:story.rootNodeId,pageSetup:metadata.pageSetup,breakType:'nextPage'}];
 const files=new Map(model.serialize()),encoder=new TextEncoder();files.set(snapshot.manifest.entrypoint,encoder.encode('<!doctype html>'+document.documentElement.outerHTML));files.set(snapshot.manifest.metadata,encoder.encode(JSON.stringify(metadata)));files.set(snapshot.manifest.annotations,encoder.encode('[]'));
 return openDocument(files,{kind:'document'});
}

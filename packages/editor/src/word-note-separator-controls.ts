import type{DocumentEditingSession,DocumentOperation}from'@doffi/core';
export function wordNoteSeparatorControls(panel:HTMLElement,session:DocumentEditingSession,button:(name:string,run:()=>void)=>void,apply:(operations:DocumentOperation[])=>void,editStory:(id:string)=>void,close:()=>void):void{
 const document=panel.ownerDocument,kind=document.createElement('select'),type=document.createElement('select'),text=document.createElement('textarea');
 const options=(select:HTMLSelectElement,name:string,items:readonly(readonly[string,string])[]):void=>{select.setAttribute('aria-label',name);for(const[value,label]of items){const option=document.createElement('option');option.value=value;option.textContent=label;select.append(option);}panel.append(select);};
 options(kind,'分隔内容所属注释',[['footnote','脚注'],['endnote','尾注']]);options(type,'注释分隔内容类型',[['separator','分隔符'],['continuationSeparator','续页分隔符'],['continuationNotice','续页提示']]);
 text.setAttribute('aria-label','注释分隔内容');text.rows=2;panel.append(text);
 const selected=()=>session.snapshot().metadata.word?.notes?.find(note=>note.kind===kind.value&&note.type===type.value),load=():void=>{const note=selected(),story=note&&session.snapshot().metadata.word?.stories.find(story=>story.storyId===note.storyId);text.value=story?session.getNode(story.rootNodeId)!.text:'';};kind.addEventListener('change',load);type.addEventListener('change',load);load();
 const set=(content?:string):void=>{const root=new DOMParser().parseFromString(session.snapshot().html,'text/html').querySelector<HTMLElement>('[data-office-kind="document"]')!.dataset.nodeId!;apply([{op:'setWordNoteSeparator',nodeId:root,kind:kind.value as 'footnote'|'endnote',type:type.value as 'separator'|'continuationSeparator'|'continuationNotice',...(content!==undefined?{text:content}:{})}]);};
 button('应用注释分隔内容',()=>set(text.value));button('恢复默认注释分隔内容',()=>{set();load();});button('编辑分隔内容格式',()=>{if(!selected())set();editStory(selected()!.storyId);close();});
}

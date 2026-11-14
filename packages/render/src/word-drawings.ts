import { DocumentError,evaluatePresentationPreset } from '@doffi/core';
import type { ContentDocumentModel,DeepReadonly,WordDrawing } from '@doffi/core';
import type { RenderLimits,RenderedPage } from './contracts.js';
import { createDrawingGeometryView } from './drawing-geometry.js';
import { markViewNode,ownViewAttribute,ownViewStyle } from './view-nodes.js';
import { numeric,property,resolvedColor } from './presentation/appearance.js';

const style=(node:HTMLElement,values:Readonly<Record<string,string>>):void=>{ownViewStyle(node,...Object.keys(values));for(const[key,value]of Object.entries(values))node.style.setProperty(key,value);};
/** Shared Word composition for preview and managed PDF/PNG. Source object
 * characters and text stories stay intact; paint and instances are derived. */
export function composeWordDrawings(document:Document,model:ContentDocumentModel,limits:Readonly<RenderLimits>):{measure(pages:readonly RenderedPage[],root:HTMLElement):void}{
 const word=model.snapshot().metadata.word,view=document.defaultView!,definitions=new Map((word?.drawings??[]).map(item=>[item.nodeId,item])),byId=new Map([...document.querySelectorAll<HTMLElement>('[data-node-id]')].map(node=>[node.dataset.nodeId!,node])),floating:{definition:DeepReadonly<WordDrawing>;frame:HTMLElement;paragraph:HTMLElement}[]=[];
 let serial=0,count=document.querySelectorAll('*').length;const composed=new Set<string>();
 const account=(node:Element):void=>{count+=1+node.querySelectorAll('*').length;if(count>limits.maxNodes)throw new DocumentError('ResourceLimitExceeded','Word drawings exceed the render node budget.');};
 const compose=(item:DeepReadonly<WordDrawing>,group=false):void=>{
  const node=byId.get(item.nodeId)!;if(!node)return;const g=item.geometry,computed=view.getComputedStyle(node),originalColor=computed.color,originalFont=computed.fontSize;
  if(composed.has(item.nodeId))return;composed.add(item.nodeId);
  if('textStoryId'in item&&item.textStoryId){const story=word!.stories.find(story=>story.storyId===item.textStoryId)!,root=byId.get(story.rootNodeId)!;for(const child of definitions.values())if(root.contains(byId.get(child.nodeId)??null))compose(child);}
  let frame=node;if(item.kind==='image'){frame=markViewNode(document.createElement('span'),'unwrap');node.before(frame);frame.append(node);}
  ownViewAttribute(node,'contenteditable');node.contentEditable='false';ownViewAttribute(frame,'data-sdk-word-object');frame.dataset.sdkWordObject=item.nodeId;
  style(frame,{display:'inline-block',position:group?'absolute':'relative',width:g.width+'pt',height:g.height+'pt','box-sizing':'border-box',padding:'0',border:'0',margin:'0','vertical-align':'middle','text-indent':'0','transform-origin':'50% 50%',transform:`rotate(${g.rotation}deg) scale(${g.flipH?-1:1},${g.flipV?-1:1})`,'font-size':'0','line-height':'0',color:'transparent','background-color':'transparent',...(group?{left:g.x+'pt',top:g.y+'pt'}:{})});
  if(property(computed,'--office-shadow-kind')==='outer'){const angle=numeric(computed,'--office-shadow-direction',0)*Math.PI/180,distance=numeric(computed,'--office-shadow-distance',0);style(frame,{filter:`drop-shadow(${distance*Math.cos(angle)}pt ${distance*Math.sin(angle)}pt ${numeric(computed,'--office-shadow-blur',0)/2}pt ${resolvedColor(document,computed,'--office-shadow')})`});}
  if(item.kind==='group'){
   const wrapper=markViewNode(document.createElement('span'),'unwrap'),c=item.coordinateSpace;wrapper.style.cssText=`position:absolute;left:0;top:0;width:${c.width}pt;height:${c.height}pt;transform-origin:0 0;transform:scale(${g.width/c.width},${g.height/c.height}) translate(${-c.x}pt,${-c.y}pt)`;wrapper.append(...node.childNodes);node.append(wrapper);
   for(const child of [...wrapper.children]as HTMLElement[]){const definition=definitions.get(child.dataset.nodeId!);if(definition)compose(definition,true);}
  }else if(item.kind==='image'){
   const c=item.crop??{top:0,right:0,bottom:0,left:0};frame.style.overflow='hidden';style(node,{display:'block',position:'absolute',left:-c.left/(1-c.left-c.right)*100+'%',top:-c.top/(1-c.top-c.bottom)*100+'%',width:100/(1-c.left-c.right)+'%',height:100/(1-c.top-c.bottom)+'%','max-width':'none',margin:'0',padding:'0',border:'0',transform:'none'});
   const border=createDrawingGeometryView(document,computed,{preset:'rect',geometry:evaluatePresentationPreset('rect',g.width,g.height),widthPt:g.width,heightPt:g.height,id:'word-drawing-'+ ++serial,paint:'stroke',maxNodes:limits.maxNodes-count});frame.append(border);account(border);
  }else{
   const preset=item.kind==='connector'?(item.connector.routing==='elbow'?'bentConnector3':'straightConnector1'):item.kind==='shape'?item.preset:'rect',geometry=evaluatePresentationPreset(preset,g.width,g.height,item.kind==='shape'?item.adjustments:{});
   const svg=createDrawingGeometryView(document,computed,{preset,geometry,widthPt:g.width,heightPt:g.height,id:'word-drawing-'+ ++serial,fillFallback:item.kind==='shape'&&preset!=='line'?'#4472c4':'none',strokeFallback:item.kind==='connector'||preset==='line'?'#224466':'none',orientation:{rotationDegrees:g.rotation,flipH:g.flipH,flipV:g.flipV},maxNodes:limits.maxNodes-count});node.append(svg);account(svg);
   if('textStoryId'in item&&item.textStoryId){const story=word!.stories.find(story=>story.storyId===item.textStoryId)!,original=byId.get(story.rootNodeId)!;const text=markViewNode(document.createElement('span')),body=item.text??{},insets=body.insets??{left:7.2,right:7.2,top:3.6,bottom:3.6},rect=geometry.textRect??{x:0,y:0,width:g.width,height:g.height};
    text.className='sdk-word-drawing-text';text.style.cssText=`position:absolute;box-sizing:border-box;display:flex;flex-direction:column;left:${rect.x}pt;top:${rect.y}pt;width:${rect.width}pt;height:${rect.height}pt;padding:${insets.top}pt ${insets.right}pt ${insets.bottom}pt ${insets.left}pt;overflow:hidden;white-space:${body.wrap===false?'pre':'pre-wrap'};overflow-wrap:anywhere;color:${originalColor};font-size:${originalFont};line-height:normal;justify-content:${body.vertical==='middle'?'center':body.vertical==='bottom'?'flex-end':'flex-start'};writing-mode:${body.direction&&body.direction!=='horizontal'?'vertical-rl':'horizontal-tb'};transform:rotate(${(body.rotation??0)+(body.direction==='vertical270'?180:0)}deg)`;
    const sourceStyle=view.getComputedStyle(original);text.style.fontSize=sourceStyle.fontSize;text.style.color=sourceStyle.color;
    const clone=original.cloneNode(true)as HTMLElement;clone.classList.remove('sdk-auxiliary');clone.style.cssText+=';display:block;border:0;padding:0;margin:0';for(const child of [clone,...clone.querySelectorAll<HTMLElement>('*')]){child.removeAttribute('data-node-id');child.removeAttribute('data-sdk-word-object');child.removeAttribute('data-sdk-word-selected');if(child.classList.contains('sdk-story-label'))child.remove();}text.append(clone);node.append(text);account(text);
    if(body.autoFit==='normal'){const scale=body.fontScale??Math.min(1,Math.max(.01,(rect.height-insets.top-insets.bottom)*4/3/Math.max(1,clone.scrollHeight)));text.style.fontSize=parseFloat(originalFont)*scale+'px';}
   }
  }
  if(!group&&item.layout.mode==='floating'){
   const paragraph=node.closest<HTMLElement>('p,h1,h2,h3,h4,h5,h6');if(paragraph){const layout=item.layout,wrap=layout.wrap??'square',d=layout.distances??{left:0,right:0,top:0,bottom:0},relative=wrap!=='none'&&!layout.behindText;
    if(relative){const side=layout.wrapSide==='left'||layout.alignH==='right'?'right':'left';style(frame,{float:side,'margin-left':(d.left+(side==='left'?Math.max(0,g.x):0))+'pt','margin-right':d.right+'pt','margin-top':(d.top+Math.max(0,g.y))+'pt','margin-bottom':d.bottom+'pt',...(wrap==='topAndBottom'?{clear:'both','shape-outside':'inset(0 -20000pt 0 -20000pt)'}:{'shape-outside':(wrap==='tight'||wrap==='through')&&layout.wrapPolygon?'polygon('+layout.wrapPolygon.map(p=>`${p.x/216}% ${p.y/216}%`).join(',')+')':'margin-box'})});
    }else{style(paragraph,{position:'relative',isolation:'isolate'});style(frame,{position:'absolute',left:g.x+'pt',top:g.y+'pt','z-index':String(layout.behindText?-1:Math.min(2147483647,layout.zOrder??1))});floating.push({definition:item,frame,paragraph});}
   }
  }
 };
 for(const item of definitions.values()){let parent=byId.get(item.nodeId)?.parentElement,child=false;while(parent){if(definitions.has(parent.dataset.nodeId??'')){child=true;break;}parent=parent.parentElement;}if(!child)compose(item);}
 return{measure(pages,root){const origin=root.getBoundingClientRect();for(const{definition:item,frame,paragraph}of floating){const p=paragraph.getBoundingClientRect(),g=item.geometry,l=item.layout,page=pages.find(page=>(p.top-origin.top)*.75>=page.rect.y&&(p.top-origin.top)*.75<page.rect.y+page.rect.height);if(!page)continue;const horizontal=l.horizontal??'column',vertical=l.vertical??'paragraph',left=origin.left+page.rect.x/.75,top=origin.top+page.rect.y/.75,setup=page.setup,m=setup.margins;
  const hStart=horizontal==='page'?left:horizontal==='margin'?left+m.left/.75:p.left,hWidth=horizontal==='page'?setup.width:horizontal==='margin'?setup.width-m.left-m.right:p.width*.75;
  const vStart=vertical==='page'?top:vertical==='margin'?top+m.top/.75:p.top,vHeight=vertical==='page'?setup.height:vertical==='margin'?setup.height-m.top-m.bottom:p.height*.75;
  const aligned=(kind:string|undefined,size:number,object:number,offset:number):number=>kind==='center'?(size-object)/2:kind==='right'||kind==='bottom'||kind==='outside'?size-object:kind==='left'||kind==='top'||kind==='inside'?0:offset;
  frame.style.left=(hStart-p.left+aligned(l.alignH,hWidth,g.width,g.x)/.75)+'px';frame.style.top=(vStart-p.top+aligned(l.alignV,vHeight,g.height,g.y)/.75)+'px';
 }}};
}

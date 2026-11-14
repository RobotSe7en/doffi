import{DocumentError,isDocumentTextBoundary,wordScriptRole}from'@doffi/core';
import type{WordScriptRole}from'@doffi/core';
import type{RenderLimits}from'./contracts.js';
import{markViewNode,viewNodeRole,isMeasuredViewContent}from'./view-nodes.js';
type Role='cjk'|'latin'|'cs'|'off';
interface Piece{node:Text;role:Role;boundary:number;offset:number}
/** Fixed character units use Normal's pitch. Latin and complex-script text
 * retain shaping across source runs; only their measured line ends get padding. */
export function wordCharacterGrid(document:Document,root:HTMLElement,pitch:number,mode:'snapToChars'|'linesAndChars',tracking:number,limits:Readonly<RenderLimits>,diagnostic:(message:string)=>void):{measure():void;dispose():void}{
 const view=document.defaultView!,canvas=document.createElement('canvas'),context=canvas.getContext('2d')!,widths=new Map<string,number>(),segmenter=new Intl.Segmenter('und',{granularity:'grapheme'});let reported=false,wrappers:HTMLElement[]=[];
 const clear=():void=>{for(const wrapper of wrappers.splice(0))if(wrapper.parentNode)wrapper.replaceWith(...wrapper.childNodes);};
 return{dispose:clear,measure(){
  clear();let nodes=document.getElementsByTagName('*').length,work=0;
  const reserve=():void=>{if(++nodes>limits.maxNodes)throw new DocumentError('ResourceLimitExceeded','Character grid projection exceeds the render node budget.');};
  const wrap=(node:Text):HTMLElement=>{reserve();const span=markViewNode(document.createElement('span'),'unwrap');span.className='sdk-grid-run';node.before(span);span.append(node);wrappers.push(span);return span;};
  for(const paragraph of root.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6,li')){
   if(paragraph.closest('[data-sdk-word-endnotes]')||paragraph.querySelector('p,h1,h2,h3,h4,h5,h6,li'))continue;
   let boundary=0,offset=0,pieces:Piece[]=[];const pads:HTMLElement[]=[],writing=view.getComputedStyle(paragraph).writingMode,vertical=writing.startsWith('vertical'),cell=paragraph.closest<HTMLElement>('td,th'),flipped=(cell?.dataset.wordTextDirection??(cell?view.getComputedStyle(cell).getPropertyValue('--word-cell-text-direction').trim():''))==='btLr',progress=(rect:DOMRect):number=>vertical?(writing==='vertical-rl'?-1:1)*(flipped?-1:1)*rect.left:rect.top;
   const visit=(node:Node):void=>{
    if(++work>limits.maxNodes*32)throw new DocumentError('ResourceLimitExceeded','Character grid measurement exceeds its work budget.');
    if(node.nodeType===1){const element=node as HTMLElement;if(viewNodeRole(node)==='remove'&&!isMeasuredViewContent(node)||element.matches('br,img,[data-word-tab],[data-word-drawing],[data-word-native-object],[data-word-note-reference],[data-word-equation]')){boundary++;return;}const css=view.getComputedStyle(element);if(css.display==='none'||css.visibility==='hidden'){boundary++;return;}}
    if(node.nodeType!==3){for(const child of [...node.childNodes])visit(child);return;}
    const text=node as Text,css=view.getComputedStyle(text.parentElement!);if(!text.length)return;if(parseFloat(css.fontSize)<.01){boundary++;return;}
    const off=css.getPropertyValue('--word-run-snap-to-grid').trim()==='0';
    if(mode==='linesAndChars'){if(!off&&tracking){const span=wrap(text);span.style.letterSpacing=((parseFloat(css.letterSpacing)||0)+tracking)+'px';}return;}
    const fragment=document.createDocumentFragment(),segments:{text:string;role:Role}[]=[];let previousRole:WordScriptRole|undefined;
    for(const character of text.data){previousRole=wordScriptRole(character,previousRole);const role:Role=off||character==='\n'||character==='\t'?'off':previousRole==='eastAsia'?'cjk':previousRole==='cs'?'cs':'latin',last=segments.at(-1);if(last?.role===role)last.text+=character;else segments.push({text:character,role});}
    for(const segment of segments){
     const role=segment.role;
     if(role!=='cjk'){const part=document.createTextNode(segment.text);fragment.append(part);pieces.push({node:part,role,boundary,offset});offset+=part.length;continue;}
     let group='',advance=NaN;const append=():void=>{if(!group)return;reserve();const span=markViewNode(document.createElement('span'),'unwrap'),part=document.createTextNode(group),delta=pitch-advance;span.className='sdk-grid-run';span.style.cssText=`letter-spacing:${delta}px;position:relative;${vertical?'top':'left'}:${(css.direction==='rtl'?-1:1)*delta/2}px`;span.append(part);fragment.append(span);wrappers.push(span);pieces.push({node:part,role,boundary,offset});offset+=part.length;group='';};
     const font=`${css.fontStyle} ${css.fontWeight} ${css.fontSize} ${css.fontFamily}`;context.font=font;context.fontKerning=css.fontKerning as CanvasFontKerning;
     for(const cluster of segmenter.segment(segment.text)){const key=font+'\n'+cluster.segment;let width=widths.get(key);if(width===undefined){width=context.measureText(cluster.segment).width;widths.set(key,width);}if(group&&Math.abs(width-advance)>.01)append();advance=width;group+=cluster.segment;}append();
    }
    text.replaceWith(fragment);
   };visit(paragraph);
   if(mode!=='snapToChars')continue;
   let previous='';
   for(let pass=0;pass<6;pass++){
    const split:Piece[]=[],rectangles=new Map<Text,DOMRect>();
    for(const piece of pieces){
     if(piece.role!=='latin'&&piece.role!=='cs'){split.push(piece);continue;}
     const range=document.createRange();range.selectNodeContents(piece.node);const lines=[...range.getClientRects()].filter(rect=>rect.width>.01&&rect.height>.1),starts=new Set([0]);
     if(lines.length>limits.maxNodes)throw new DocumentError('ResourceLimitExceeded','Character grid lines exceed the render budget.');
     for(const line of lines.slice(1)){let low=0,high=piece.node.length;while(low<high){if(++work>limits.maxNodes*32)throw new DocumentError('ResourceLimitExceeded','Character grid measurement exceeds its work budget.');const middle=Math.floor((low+high)/2),probe=document.createRange();probe.setStart(piece.node,middle);probe.setEnd(piece.node,Math.min(piece.node.length,middle+1));if(progress(probe.getBoundingClientRect())<progress(line)-.5)low=middle+1;else high=middle;}while(low<piece.node.length&&!isDocumentTextBoundary(piece.node.data,low))low++;if(low<piece.node.length)starts.add(low);}
     const offsets=[...starts].sort((a,b)=>a-b),parts:Piece[]=[];let current=piece.node;for(let index=offsets.length-1;index>0;index--){const at=offsets[index]!,next=current.splitText(at);parts.unshift({...piece,node:next,offset:piece.offset+at});}parts.unshift({...piece,node:current});
     for(const part of parts){const probe=document.createRange();probe.selectNodeContents(part.node);rectangles.set(part.node,probe.getBoundingClientRect());}split.push(...parts);
    }
    pieces=split;
    const groups:{pieces:Piece[];rect:DOMRect;width:number}[]=[];let group:typeof groups[number]|undefined;
    for(const piece of pieces){const rect=rectangles.get(piece.node);if(!rect||rect.width<.01||rect.height<.1){group=undefined;continue;}const first=group?.pieces[0],same=group&&first?.role===piece.role&&first.boundary===piece.boundary&&(vertical?Math.min(group.rect.right,rect.right)-Math.max(group.rect.left,rect.left)>Math.min(group.rect.width,rect.width)*.4:Math.min(group.rect.bottom,rect.bottom)-Math.max(group.rect.top,rect.top)>Math.min(group.rect.height,rect.height)*.4);if(!same){group={pieces:[],rect,width:0};groups.push(group);}group!.pieces.push(piece);group!.width+=vertical?rect.height:rect.width;}
    const signature=groups.map(group=>group.pieces[0]!.offset+':'+(group.pieces.at(-1)!.offset+group.pieces.at(-1)!.node.length)+':'+Math.round(group.width*64)).join('|');
    if(signature===previous)break;previous=signature;
    const oldPads=new Set(pads.splice(0));for(const pad of oldPads)if(pad.parentNode){pad.replaceWith(...pad.childNodes);nodes--;}wrappers=wrappers.filter(wrapper=>!oldPads.has(wrapper));
    for(const group of groups){const space=Math.max(0,Math.ceil((group.width-.02)/pitch)*pitch-group.width);if(space<.01)continue;const first=group.pieces[0]!,last=group.pieces.at(-1)!,start=wrap(first.node);pads.push(start);if(first.role==='latin')start.style[vertical?'paddingTop':'paddingLeft']=space/2+'px';if(last.node===first.node)start.style[vertical?'paddingBottom':'paddingRight']=(first.role==='latin'?space/2:space)+'px';else{const end=wrap(last.node);pads.push(end);end.style[vertical?'paddingBottom':'paddingRight']=(first.role==='latin'?space/2:space)+'px';}}
    if(pass===5&&!reported){reported=true;diagnostic('Character-grid line blocks reached the bounded adjustment limit.');}
   }
  }
 }};
}

import type { DeepReadonly, DocumentSection, PageSetup } from '@doffi/core';
import { DocumentError } from '@doffi/core';
import type { RenderLimits } from './contracts.js';
import { markViewNode, ownViewStyle, viewNodeRole } from './view-nodes.js';

/** Native margins exclude the binding gutter. Keep the effective box derived. */
export function wordPageSetup(section:DeepReadonly<DocumentSection>,fallback:DeepReadonly<PageSetup>):DeepReadonly<PageSetup>{
 const setup=section.pageSetup??fallback;
 if(!section.gutter)return setup;
 const side=section.gutterPosition??'left';
 return {...setup,margins:{...setup.margins,[side]:setup.margins[side]+section.gutter}};
}
/** Facing-page parity is physical and independent of printed page labels. */
export function facingWordPageSetup(setup:DeepReadonly<PageSetup>,index:number,mirror:boolean):DeepReadonly<PageSetup>{return mirror&&index%2===1?{...setup,margins:{...setup.margins,left:setup.margins.right,right:setup.margins.left}}:setup;}
/** A grid changes measured text, never authored paragraph/run formatting. */
export function applyWordGrid(document:Document,root:HTMLElement,section:DeepReadonly<DocumentSection>,limits:Readonly<RenderLimits>):void{
 const grid=section.grid;if(!grid||grid.type==='default')return;
 const view=document.defaultView!;let nodes=document.getElementsByTagName('*').length;
 for(const paragraph of root.querySelectorAll<HTMLElement>('p,h1,h2,h3,h4,h5,h6,li')){
  const computed=view.getComputedStyle(paragraph);
  if(computed.getPropertyValue('--word-paragraph-snap-to-grid').trim()==='0')continue;
  if(grid.linePitch&&grid.type!=='snapToChars'&&paragraph.dataset.wordLineRule!=='exact'){
   const pitch=grid.linePitch*4/3,line=parseFloat(computed.lineHeight)||parseFloat(computed.fontSize)*1.2;
   ownViewStyle(paragraph,'line-height');paragraph.style.lineHeight=(Math.max(1,Math.ceil((line-.01)/pitch))*pitch)+'px';
  }
  if(!['linesAndChars','snapToChars'].includes(grid.type)||grid.characterSpacing===undefined)continue;
  // charSpace is the signed increment over Normal's character pitch in 1/4096 pt.
  const spacing=grid.characterSpacing/4096;
  const walker=document.createTreeWalker(paragraph,4),texts:Text[]=[];for(let node=walker.nextNode();node;node=walker.nextNode())texts.push(node as Text);
  for(const text of texts){
   let excluded=false;for(let parent:Node|null=text.parentElement;parent&&parent!==paragraph;parent=parent.parentNode)if(viewNodeRole(parent)==='remove'){excluded=true;break;}if(excluded)continue;
   const style=view.getComputedStyle(text.parentElement!);if(style.getPropertyValue('--word-run-snap-to-grid').trim()==='0')continue;
   const tracking=(parseFloat(style.letterSpacing)||0)+spacing*4/3,parts:{text:string;snap:boolean}[]=[];let previous=false;
   for(const character of text.data){const snap:boolean=grid.type==='linesAndChars'||(/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}\u3000-\u303f\uff01-\uff60]/u.test(character)||/[\p{Mark}\u200d\ufe00-\ufe0f\u{e0100}-\u{e01ef}]/u.test(character)&&previous);const last=parts.at(-1);if(last&&last.snap===snap)last.text+=character;else parts.push({text:character,snap});previous=snap;}
   const fragment=document.createDocumentFragment();for(const part of parts){if(!part.snap){fragment.append(document.createTextNode(part.text));continue;}if(++nodes>limits.maxNodes)throw new DocumentError('ResourceLimitExceeded','Document grid projection exceeds the rendering node budget.');const span=markViewNode(document.createElement('span'),'unwrap');span.style.letterSpacing=tracking+'px';span.textContent=part.text;fragment.append(span);}text.replaceWith(fragment);
  }
 }
}

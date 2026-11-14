import { editSourceTable } from "./table-edit.js";
import { copyWordStorySemantics } from './word-story-copy.js';
import { wordTextDefaults,clearCoveredWordRuns } from '../document/word-format.js';
import { applyWordDrawing,isWordDrawingOperation } from './word-drawing-edit.js';
import { applyWordReference,isWordReferenceOperation } from './word-reference-edit.js';
import { checkLimit } from "../document/validation.js";
import type { DocumentReadLimits } from "../runtime.js";
import { DocumentError } from '../errors.js';
import type { WordOperation } from '../contracts/word-operations.js';
import type { WordDocumentMetadata } from '../contracts/document.js';
import type { DocumentEdit } from '../document/edit.js';
import type { ParsedHtml } from '../document/html.js';
import type { SourceStyleEdit } from './source-styles.js';
function fail(message:string):never{throw new DocumentError('InvalidOperation',message);}
export function isWordOperation(operation:{readonly op:string}):operation is WordOperation{return isWordDrawingOperation(operation)||isWordReferenceOperation(operation)||['resetWordText','cloneWordStory','setWordOptions','setWordTabs','insertWordField','setWordSection','splitWordSection','insertWordStory','setWordStoryContent','insertWordTable','editWordTable','formatWordTable','formatWordCell','formatWordText','setWordParagraph','setWordStyle','setWordNumbering','setWordList'].includes(operation.op);}
export function applyWordOperation(metadata:WordDocumentMetadata,edit:DocumentEdit,parse:()=>ParsedHtml,styles:SourceStyleEdit,operation:WordOperation,limits:DocumentReadLimits):void{
  if (!metadata.word) {
    const source=parse(),rootId=source.nodeIds().find(id=>source.peekNode(id)!.parentNodeId===null)!;
    metadata.word={stories:[{storyId:'main',kind:'main',rootNodeId:rootId}],styles:[],numbering:[],hyperlinks:[],bookmarks:[],fields:[],tabs:[],revisions:[],tables:[]};
    edit.attributes(rootId,{'data-office-story':'main'});
  }
  const word=metadata.word;
  const node=parse().peekNode(operation.nodeId);if(!node)fail('The Word source target is missing.');
  if(isWordDrawingOperation(operation)){applyWordDrawing(metadata,edit,parse,operation,limits);return;}
  if(isWordReferenceOperation(operation)){applyWordReference(word,edit,parse,operation);return;}
  if(operation.op==='resetWordText'){
    if(!/^p$|^h[1-6]$/.test(node.tagName))fail('Resetting text formatting requires a paragraph range.');
    if(operation.resetParagraph){if(operation.from!==0||operation.to!==parse().getNode(operation.nodeId)!.text.length)fail('Paragraph formatting can only be reset for a complete paragraph.');const normal=word.styles.find(style=>style.kind==='paragraph'&&style.styleId==='Normal');applyWordOperation(metadata,edit,parse,styles,{op:'setWordParagraph',nodeId:operation.nodeId,clearDirectFormatting:true,...(normal?{styleId:normal.styleId}:{})},limits);}
    const source=parse(),defaults=wordTextDefaults(word,source,operation.nodeId,styles.resolver());edit.formatSourceText(operation.nodeId,operation,defaults);clearCoveredWordRuns(word,parse(),edit,operation.nodeId,operation.from,operation.to);return;
  }
  if(operation.op==='setWordOptions'){
    if(node.parentNodeId!==null)fail('Shared Word settings require the document root.');
    if(operation.evenAndOddHeaders!==undefined)word.evenAndOddHeaders=operation.evenAndOddHeaders;
    if(operation.automaticHyphenation!==undefined)word.automaticHyphenation=operation.automaticHyphenation;
    return;
  }
  if(operation.op==='setWordTabs'){
    if(!/^p$|^h[1-6]$/.test(node.tagName))fail('Tab stops require a paragraph.');
    const definition={paragraphNodeId:operation.nodeId,defaultInterval:operation.defaultInterval,stops:operation.stops},index=word.tabs.findIndex(item=>item.paragraphNodeId===operation.nodeId);
    if(index<0)word.tabs.push(definition);else word.tabs[index]=definition;return;
  }
  if(operation.op==='insertWordField'){
    if(!/^p$|^h[1-6]$/.test(node.tagName))fail('A field insertion requires a paragraph.');
    const cached=operation.cachedText??'1',escape=(value:string):string=>value.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;');
    const source='<span data-word-field="simple">'+escape(cached)+'</span>',id=(operation.offset===undefined?edit.insertSource(operation.nodeId,source):edit.insertSourceAt(operation.nodeId,operation.offset,source)).rootNodeIds[0]!;
    word.fields.push({nodeId:id,instruction:operation.instruction,resultState:'cached',resultNodeIds:[id]});return;
  }
  if(operation.op==='setWordSection') {
    const section=metadata.sections.find(section=>section.rootNodeId===operation.nodeId);if(!section)fail('Page changes require a section root.');
    Object.assign(section,operation.changes);return;
  }
  if(operation.op==='splitWordSection') {
    if(!/^p$|^h[1-6]$/.test(node.tagName))fail('A section break requires a paragraph.');
    const index=metadata.sections.findIndex(section=>section.rootNodeId===node.parentNodeId),section=metadata.sections[index];if(!section)fail('A section break requires a paragraph directly in its section.');
    const main=word.stories.find(story=>story.kind==='main')!,html=parse(),siblings=html.peekNode(main.rootNodeId)!.childNodeIds,ordinal=siblings.indexOf(section.rootNodeId);
    const rootId=edit.insertSource(main.rootNodeId,'<section></section>',siblings[ordinal+1]).rootNodeIds[0]!;
    const children=html.peekNode(section.rootNodeId)!.childNodeIds,tail=children.slice(children.indexOf(operation.nodeId)+1);if(tail.length)edit.moveSource(tail,rootId);else edit.insertSource(rootId,'<p></p>');
    const next={...section,sectionId:rootId,rootNodeId:rootId,breakType:operation.breakType};delete next.pageNumberStart;metadata.sections.splice(index+1,0,next);return;
  }
  if(operation.op==='insertWordStory') {
    const section=metadata.sections.find(section=>section.rootNodeId===operation.nodeId);if(!section)fail('Header/footer creation requires a section.');
    const main=word.stories.find(story=>story.kind==='main')!,html=parse(),rootId=html.nodeIds().find(id=>html.peekNode(id)!.parentNodeId===null)!;
    if(main.rootNodeId===rootId){const children=[...html.peekNode(rootId)!.childNodeIds],wrapper=edit.insertSource(rootId,'<div data-office-story="main"></div>').rootNodeIds[0]!;edit.moveSource(children,wrapper);edit.attributes(rootId,{'data-office-story':undefined});main.rootNodeId=wrapper;for(const item of metadata.sections)if(item.rootNodeId===rootId)item.rootNodeId=wrapper;}
    const storyRoot=edit.insertSource(rootId,'<div data-office-story="'+operation.kind+'">'+operation.html+'</div>').rootNodeIds[0]!;
    word.stories.push({storyId:storyRoot,kind:operation.kind,rootNodeId:storyRoot});const normal=word.styles.find(style=>style.kind==='paragraph'&&style.styleId==='Normal');if(normal){const tree=parse(),visit=(id:string):void=>{const paragraph=tree.peekNode(id)!;if(paragraph.tagName==='p'&&!paragraph.attributes['data-word-style'])edit.attributes(id,{'data-word-style':normal.styleId,class:[paragraph.attributes.class,normal.cssClass].filter(Boolean).join(' ')});for(const child of paragraph.childNodeIds)visit(child);};visit(storyRoot);}const key=operation.kind==='header'?'headers':'footers';section[key]={...section[key],[operation.variant]:storyRoot};if(operation.variant==='first')section.differentFirstPage=true;if(operation.variant==='even')word.evenAndOddHeaders=true;return;
  }
  if(operation.op==='cloneWordStory'){
    const section=metadata.sections.find(section=>section.rootNodeId===operation.nodeId),story=word.stories.find(story=>story.storyId===operation.storyId);if(!section||!story||!['header','footer'].includes(story.kind))fail('Copying a header/footer requires its source story and target section.');
    const source=parse(),article=source.nodeIds().find(id=>source.peekNode(id)!.parentNodeId===null)!,clone=edit.cloneSource(story.rootNodeId,article),rootId=clone.rootNodeIds[0]!;
    copyWordStorySemantics(word,clone.identities,edit,source,rootId);word.stories.push({...story,storyId:rootId,rootNodeId:rootId});const key=story.kind==='header'?'headers':'footers';section[key]={...section[key],[operation.variant]:rootId};if(operation.variant==='first')section.differentFirstPage=true;if(operation.variant==='even')word.evenAndOddHeaders=true;return;
  }
  if(operation.op==='setWordStoryContent'){
    if(!word.stories.some(story=>story.rootNodeId===operation.nodeId&&story.kind!=='main'))fail('Auxiliary story editing requires an existing header, footer or comment story.');edit.replaceSourceContent(operation.nodeId,operation.html);return;
  }
  if(operation.op==='insertWordTable') {
    if (!/^p$|^h[1-6]$/.test(node.tagName)) fail('Insert a table beside an existing paragraph.');
    checkLimit(operation.rows * operation.columns * 3, limits.maxNodes);
    const width=operation.width??metadata.pageSetup.width-metadata.pageSetup.margins.left-metadata.pageSetup.margins.right;
    const source='<table style="border-collapse:collapse;table-layout:fixed;width:'+width+'pt">'+Array.from({length:operation.rows},()=>'<tr>'+Array.from({length:operation.columns},()=>'<td style="border:0.5pt solid #000000;padding:4pt;vertical-align:top"><p></p></td>').join('')+'</tr>').join('')+'</table>';
    const parent=edit.sourceParent(operation.nodeId),siblings=parse().peekNode(parent)!.childNodeIds,index=siblings.indexOf(operation.nodeId);
    const reference=operation.position==='before'?operation.nodeId:siblings[index+1];
    const tableId=edit.insertSource(parent,source,reference).rootNodeIds[0]!;
    word.tables.push({nodeId:tableId,grid:Array.from({length:operation.columns},()=>width/operation.columns)});return;
  }
  if(operation.op==='editWordTable'||operation.op==='formatWordTable') {
    if(node.tagName!=='table')fail('Word table commands require a table.');
    let info=word.tables.find(table=>table.nodeId===operation.nodeId);
    const source=parse(),rowIds=node.childNodeIds.flatMap(id=>{const row=source.peekNode(id)!;return row.tagName==='tr'?[id]:row.childNodeIds;});
    if(!info){
      const occupied:number[][]=Array.from({length:rowIds.length},()=>[]);let columns=0;
      for(const[row,id]of rowIds.entries())for(const cellId of source.peekNode(id)!.childNodeIds){const cell=source.peekNode(cellId)!;let column=0;while(occupied[row]![column])column++;const width=Number(cell.attributes.colspan??1),height=Number(cell.attributes.rowspan??1)||rowIds.length-row;checkLimit((column+width)*rowIds.length,limits.maxNodes);if(row+height>rowIds.length)fail('A merged cell exceeds the table.');for(let r=row;r<row+height;r++)for(let c=column;c<column+width;c++)occupied[r]![c]=1;columns=Math.max(columns,column+width);}
      if(!columns)fail('An editable table requires a nonempty grid.');const width=metadata.pageSetup.width-metadata.pageSetup.margins.left-metadata.pageSetup.margins.right;
      info={nodeId:operation.nodeId,grid:Array.from({length:columns},()=>width/columns)};word.tables.push(info);
    }
    if(operation.op==='editWordTable') {
      const firstCell=rowIds.flatMap(id=>source.peekNode(id)!.childNodeIds).map(id=>source.peekNode(id)!).find(cell=>['td','th'].includes(cell.tagName));
      const escape=(value:string):string=>value.replaceAll('&','&amp;').replaceAll('"','&quot;').replaceAll('<','&lt;');
      const cellAttributes=firstCell?['class','style'].filter(key=>firstCell.attributes[key]).map(key=>' '+key+'="'+escape(firstCell.attributes[key]!)+'"').join(''):'';
      const previousWidth=info.grid.reduce((a,b)=>a+b,0);
      const result=editSourceTable({edit,parse,limits,newCellHtml:'<td'+cellAttributes+'><p></p></td>'},operation.nodeId,info.grid,{width:previousWidth,height:Math.max(1,rowIds.length)*20},operation.change);
      const preserveWidth=operation.change.kind==='insertColumns'||operation.change.kind==='deleteColumns';
      info.grid=preserveWidth?result.grid.map(width=>width*previousWidth/result.width):result.grid;
      if(operation.change.kind==='resize'&&operation.change.grid)edit.sourceStyle(operation.nodeId,'width',result.width+'pt');
      if(preserveWidth||operation.change.kind==='resize'&&operation.change.grid||operation.change.kind==='splitCell')for(const cell of result.cells)edit.sourceStyle(cell.id,'width',info.grid.slice(cell.column,cell.column+cell.columns).reduce((a,b)=>a+b,0)+'pt');
      return;
    }
    if(operation.styleId!==undefined){const style=operation.styleId===null?undefined:word.styles.find(style=>style.styleId===operation.styleId&&style.kind==='table');if(operation.styleId!==null&&!style)fail('The native table style is missing.');const old=word.styles.find(style=>style.styleId===node.attributes['data-word-style']);const classes=(node.attributes.class??'').split(/\s+/).filter(value=>value&&value!==old?.cssClass);if(style)classes.push(style.cssClass);edit.attributes(operation.nodeId,{'data-word-style':operation.styleId??undefined,class:classes.join(' ')});}
    for(const[property,value]of Object.entries(operation.properties??{}))edit.sourceStyle(operation.nodeId,property,value);
    if(operation.repeatHeaderRows!==undefined){if(operation.repeatHeaderRows>rowIds.length)fail('Repeated headers exceed the table.');for(const[index,id]of rowIds.entries())edit.attributes(id,{'data-word-repeat-header':index<operation.repeatHeaderRows?'true':undefined});}
    if(operation.allowRowSplit!==undefined)for(const id of rowIds)edit.sourceStyle(id,'break-inside',operation.allowRowSplit?'auto':'avoid');return;
  }
  if(operation.op==='formatWordCell') {
    if(!['td','th'].includes(node.tagName))fail('Cell formatting requires a table cell.');
    for(const[property,value]of Object.entries(operation.properties))edit.sourceStyle(operation.nodeId,property,value);
    if(operation.textDirection!==undefined)edit.attributes(operation.nodeId,{'data-word-text-direction':operation.textDirection});return;
  }
  if(operation.op==='formatWordText'){edit.formatSourceText(operation.nodeId,operation,operation.properties);return;}
  if(operation.op==='setWordStyle'||operation.op==='setWordNumbering'){
    if(node.tagName!=='article'||node.parentNodeId!==null)fail('Shared Word definitions require the complete document root.');
    if(operation.op==='setWordNumbering'){
      if(operation.copyFromNumberingId){const previous=word.numbering.find(item=>item.numberingId===operation.copyFromNumberingId);if(!previous)fail('The numbering appearance source is missing.');for(const level of operation.levels){const source=previous.levels.find(item=>item.level===level.level)??previous.levels.at(-1)!;if(source.cssClass!==level.cssClass)styles.cloneClass(source.cssClass,level.cssClass);}}
      for(const item of operation.levelProperties??[]){const level=operation.levels.find(level=>level.level===item.level);if(!level)fail('Numbering appearance requires an existing level.');styles.setClass(level.cssClass,item.properties);}
      const index=word.numbering.findIndex(n=>n.numberingId===operation.numberingId),definition={numberingId:operation.numberingId,levels:operation.levels};if(index<0)word.numbering.push(definition);else word.numbering[index]=definition;return;
    }
    const index=word.styles.findIndex(s=>s.styleId===operation.style.styleId),previous=index<0?undefined:word.styles[index];
    if(previous&&previous.kind!==operation.style.kind)fail('A Word style cannot change its native kind.');
    const definition={...operation.style};
    if(operation.paragraphProperties){styles.setClass(definition.cssClass,operation.paragraphProperties);definition.paragraphProperties=[...new Set([...(definition.paragraphProperties??[]),...Object.keys(operation.paragraphProperties)])];}
    if(operation.runProperties){definition.runCssClass??=definition.cssClass+'-run';styles.setClass(definition.runCssClass,operation.runProperties);definition.runProperties=[...new Set([...(definition.runProperties??[]),...Object.keys(operation.runProperties)])];}
    if(index<0)word.styles.push(definition);else word.styles[index]=definition;
    for (const descendant of word.styles) {
      if (descendant.styleId === definition.styleId) continue;
      const chain: typeof word.styles = []; let current: typeof definition | undefined = descendant;
      const seen = new Set<string>();
      while (current && current.styleId !== definition.styleId && !seen.has(current.styleId)) { seen.add(current.styleId); chain.push(current); current = current.basedOn ? word.styles.find(s => s.styleId === current!.basedOn) : undefined; }
      if (!current || current.styleId !== definition.styleId) continue;
      const paragraph = Object.fromEntries(Object.entries(operation.paragraphProperties ?? {}).filter(([key]) => !chain.some(s => s.paragraphProperties?.includes(key))));
      const run = Object.fromEntries(Object.entries(operation.runProperties ?? {}).filter(([key]) => !chain.some(s => s.runProperties?.includes(key))));
      if (Object.keys(paragraph).length) styles.setClass(descendant.cssClass, paragraph);
      if (Object.keys(run).length) { styles.setClass(descendant.cssClass, run); if (descendant.runCssClass) styles.setClass(descendant.runCssClass, run); }
    }
    return;
  }
  if(!/^p$|^h[1-6]$|^li$/.test(node.tagName))fail('Paragraph commands require a paragraph or list item.');
  if(operation.op==='setWordList'){
    const definition=word.numbering.find(n=>n.numberingId===operation.numberingId);
    if(operation.numberingId!==null&&!definition?.levels.some(l=>l.level===(operation.level??0)))fail('A list requires an existing numbering level.');
    edit.attributes(operation.nodeId,{'data-word-numbering':operation.numberingId??undefined,'data-word-level':operation.numberingId===null?undefined:String(operation.level??0)});
    edit.sourceStyle(operation.nodeId,'margin-left',operation.numberingId===null?'0pt':((operation.level??0)+1)*24+'pt');edit.sourceStyle(operation.nodeId,'text-indent',operation.numberingId===null?'0pt':'-18pt');return;
  }
  if(operation.clearDirectFormatting)edit.attributes(operation.nodeId,{style:undefined});
  if(operation.styleId!==undefined){
    const definition=operation.styleId===null?undefined:word.styles.find(s=>s.styleId===operation.styleId&&s.kind==='paragraph');
    if(operation.styleId!==null&&!definition)fail('A paragraph style is missing or has the wrong kind.');
    const previous=word.styles.find(s=>s.styleId===node.attributes['data-word-style']);
    const ancestry = (style: typeof definition): typeof word.styles => { const result: typeof word.styles = [], seen = new Set<string>(); for(let current=style;current&&!seen.has(current.styleId);current=current.basedOn?word.styles.find(s=>s.styleId===current!.basedOn):undefined){result.unshift(current);seen.add(current.styleId);}return result; };
    const oldClasses = new Set(ancestry(previous).map(style=>style.cssClass));
    const classes=(node.attributes.class??'').split(/\s+/).filter(c=>c&&!oldClasses.has(c));
    classes.push(...ancestry(definition).map(style=>style.cssClass));
    edit.attributes(operation.nodeId,{'data-word-style':operation.styleId??undefined,class:classes.join(' ')});
    const previousRun=previous?.runCssClass,nextRun=definition?.runCssClass;
    const tree=parse(),visit=(id:string):void=>{const child=tree.peekNode(id)!;if(child.attributes['data-word-run']){const runClasses=(child.attributes.class??'').split(/\s+/).filter(c=>c&&c!==previousRun);if(nextRun)runClasses.push(nextRun);edit.attributes(id,{class:[...new Set(runClasses)].join(' ')});}for(const next of child.childNodeIds)visit(next);};visit(operation.nodeId);
    const level=definition?.styleId.match(/^Heading([1-6])$/)?.[1];edit.renameSource(operation.nodeId,level?'h'+level:'p');
  }
  for(const[property,value]of Object.entries(operation.properties??{})){
    if(!/^(?:margin-(?:top|right|bottom|left)|text-(?:align|indent)|line-height|break-(?:before|after|inside)|widows|orphans|background-color|border(?:-(?:top|right|bottom|left)(?:-(?:style|color|width))?)?|--word-[a-z-]+)$/.test(property))fail('A paragraph property is outside the Word format contract.');
    edit.sourceStyle(operation.nodeId,property,value);
  }
}

import{removeWordRows,wordTableRows}from'./word-row-revisions.js';
import{removeWordCells}from'./word-cell-revisions.js';
import type{DocumentReadLimits}from'../runtime.js';
import{expandWordRevisionGroups}from'../word-revision-groups.js';
import{reconcileTrackedWord}from'./word-tracked-structure.js';
import type { WordMetadata } from '../contracts/word.js';
import type{DocumentSection}from'../contracts/word-page.js';
import type { WordRevision, WordRevisionOperation } from '../contracts/word-revision.js';
import type { DocumentEdit } from '../document/edit.js';
import type { ParsedHtml } from '../document/html.js';
import { DocumentError } from '../errors.js';
import { isTextBoundary } from '../document/text.js';
import type {SourceStyleEdit} from './source-styles.js';
import {insertWordReplacement} from './word-replacement-context.js';
const fail = (message: string): never => { throw new DocumentError('InvalidOperation', message); };
export const wordRevisionId = (revision: Pick<WordRevision, 'nodeId' | 'revisionId'>): string => revision.revisionId ?? revision.nodeId;
export function isWordRevisionOperation(operation: {
    readonly op: string;
}): operation is WordRevisionOperation { return ['reconcileWordTracked','setWordReview', 'reviewWordRevision', 'reviewWordRevisions', 'trackWordText','trackWordInlineContent'].includes(operation.op); }

export function trackWordInlineContent(word:WordMetadata,edit:DocumentEdit,parse:()=>ParsedHtml,operation:Extract<WordRevisionOperation,{op:'trackWordInlineContent'}>,copy?:{inserted(identities:ReadonlyMap<string,string>):void;keepEmpty:boolean}):void{
    const source=parse(),node=source.getNode(operation.nodeId),{from,to,expectedText,html}=operation;
    if(!node||!/^p$|^h[1-6]$/.test(node.tagName)||from>to||to>node.text.length||node.text.slice(from,to)!==expectedText||!isTextBoundary(node.text,from)||!isTextBoundary(node.text,to))fail('Tracked inline content requires an exact paragraph range.');
    edit.assertInlineSource(operation.nodeId,html);
    const start=source.textOffset(operation.nodeId)!+from,end=source.textOffset(operation.nodeId)!+to,author=operation.author??word.trackingAuthor??'',date=operation.date??new Date().toISOString();
    const inside=(id:string,parent:string):boolean=>{for(let current:string|null=id;current;current=source.peekNode(current)!.parentNodeId)if(current===parent)return true;return false;};
    if(from!==to)for(const id of [...word.drawings??[],...word.controls??[],...word.equations??[],...word.noteReferences??[]].map(item=>item.nodeId)){if(!source.hasNode(id)||!inside(id,operation.nodeId))continue;const begin=source.textOffset(id)!,finish=begin+source.textLength(id)!;if(start<finish&&end>begin)throw new DocumentError('UnsupportedFeature','Review or edit this native object through its own operation before replacing its range.');}
    for(const revision of word.revisions){if(!source.hasNode(revision.nodeId)||!inside(revision.nodeId,operation.nodeId)&&!inside(operation.nodeId,revision.nodeId))continue;
        const begin=source.textOffset(revision.nodeId)!,finish=begin+source.textLength(revision.nodeId)!;
        if(['delete','moveFrom'].includes(revision.kind)&&start<finish&&end>begin)fail('Deleted revision content must be reviewed before editing it.');
        if(revision.kind==='insert'&&revision.author===author&&start>=begin&&end<=finish){
            for(const id of edit.wrapSourceText(revision.nodeId,start-begin,end-begin,'span'))edit.removeSource(id);
            const content=edit.insertSourceAt(revision.nodeId,start-begin,html,undefined,!!copy);copy?.inserted(content.identities);
            if(!parse().getNode(revision.nodeId)!.text&&!copy?.keepEmpty){edit.removeSource(revision.nodeId);word.revisions=word.revisions.filter(item=>item!==revision);}return;
        }
    }
    const deleted=edit.wrapSourceText(operation.nodeId,from,to,'del');for(const id of deleted)edit.attributes(id,{'data-word-run-overlay':'true'});
    const content=edit.insertSourceAt(operation.nodeId,to,'<ins data-word-run-overlay="true">'+html+'</ins>',undefined,!!copy),inserted=content.rootNodeIds[0]!;copy?.inserted(content.identities);const current=parse();
    const live=current.getNode(inserted)!.text.length>0||!!copy?.keepEmpty;if(!live)edit.removeSource(inserted);
    const groupId=deleted[0]??inserted;
    for(const[kind,ids]of [['delete',deleted],['insert',live?[inserted]:[]]]as const)for(const id of ids)word.revisions.push({nodeId:id,revisionId:id,sourceId:id,groupId,kind,author,date});
}
export function trackWordText(word: WordMetadata, edit: DocumentEdit, parse: () => ParsedHtml, operation: {
    nodeId: string;
    from: number;
    to: number;
    expectedText: string;
    text: string;
    insertionContextNodeId?:string | undefined;
    properties?: Record<string, string> | undefined;
    author?: string | undefined;
    date?: string | undefined;
},styles:SourceStyleEdit): void {
    const source = parse(), node = source.getNode(operation.nodeId) ?? fail('The tracked text source is missing.');
    const { from, to, expectedText, text } = operation;
    if (from > to || node.text.slice(from, to) !== expectedText || !isTextBoundary(node.text, from) || !isTextBoundary(node.text, to) || /[\u0000\r\ud800-\udfff]/u.test(text))
        fail('The tracked text range or Unicode content is invalid.');
    if (expectedText === text)
        return;
    const inside=(id:string,parent:string):boolean=>{for(let current:string|null=id;current;current=source.peekNode(current)!.parentNodeId)if(current===parent)return true;return false;};
    const atomicContext=(id:string):boolean=>{const node=source.peekNode(id)!;return ['br','img','del'].includes(node.tagName)||['data-word-paragraph-mark','data-word-tab','data-word-break','data-word-drawing','data-word-native-object','data-word-note-reference','data-word-equation'].some(key=>node.attributes[key]!==undefined)||['checkBox','dropDown','date','picture'].includes(node.attributes['data-word-content-control']??'');};
    let context=operation.insertionContextNodeId;
    if(context){const owner=source.getNode(context),offset=source.textOffset(operation.nodeId)!+from,begin=owner?source.textOffset(context)!:NaN;if(from!==to||!owner||!inside(context,operation.nodeId)||offset<begin||offset>begin+owner.text.length)fail('The insertion context must own this zero-length source boundary.');for(let current:string|null=context;current;current=source.peekNode(current)!.parentNodeId){if(atomicContext(current))fail('An atomic or deleted source cannot own inserted text.');if(current===operation.nodeId)break;}}
    else if(from===to&&from>0){const point=source.textOffset(operation.nodeId)!+from;let selectedContext=operation.nodeId;for(;;){const child:string|undefined=source.peekNode(selectedContext)!.childNodeIds.find(id=>{const start=source.textOffset(id)!,length=source.textLength(id)!;return length>0&&start<point&&start+length>=point;});if(!child||atomicContext(child))break;selectedContext=child;}context=selectedContext;}
    const properties={...operation.properties};if(properties['font-weight']!==undefined&&properties['--word-bold-cs']===undefined)properties['--word-bold-cs']=properties['font-weight']==='bold'||Number(properties['font-weight'])>=600?'1':'0';if(properties['font-style']!==undefined&&properties['--word-italic-cs']===undefined)properties['--word-italic-cs']=properties['font-style']==='italic'?'1':'0';
    const escape = (value: string): string => value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;');
    const formatInserted=(id:string,from:number):void=>{if(text&&Object.keys(properties).length){const changed:string[]=[];edit.formatSourceText(id,{from,to:from+text.length,expectedText:text},properties,change=>changed.push(change.nodeId),undefined,false,true);for(const nodeId of changed)edit.attributes(nodeId,{'data-word-current-format':'true'});}};
    const author = operation.author ?? word.trackingAuthor ?? '', date = operation.date ?? new Date().toISOString(), start = source.textOffset(operation.nodeId)! + from, end = source.textOffset(operation.nodeId)! + to;
    for (const revision of word.revisions) {
        const owner = source.getNode(revision.nodeId);
        if (!owner)
            continue;
        const begin = source.textOffset(owner.nodeId)!, finish = begin + owner.text.length;
        if(!inside(owner.nodeId,operation.nodeId)&&!inside(operation.nodeId,owner.nodeId))continue;
        if (revision.kind === 'insert' && revision.author === author && start >= begin && end <= finish && (!context||inside(context,owner.nodeId))) {
            if(context&&from===to&&text)edit.insertSourceAt(revision.nodeId,start-begin,escape(text).split('\n').join('<br>'),context);else edit.replaceSourceText(revision.nodeId, { ...operation, from: start - begin, to: end - begin });
            if(!parse().getNode(revision.nodeId)!.text){edit.removeSource(revision.nodeId);word.revisions=word.revisions.filter(item=>item!==revision);}
            else formatInserted(revision.nodeId,start-begin);
            return;
        }
        if (['delete', 'moveFrom'].includes(revision.kind) && start < finish && end > begin)
            fail('Deleted revision content must be reviewed before editing it.');
    }
    const deleted = edit.wrapSourceText(operation.nodeId, from, to, 'del');
    for(const id of deleted)edit.attributes(id,{'data-word-run-overlay':'true'});
    const content=escape(text).split('\n').join('<br>');
    const inserted = text ? from!==to?insertWordReplacement(word,edit,source,styles,operation.nodeId,from,to,content):edit.insertSourceAt(operation.nodeId, to, '<ins data-word-run-overlay="true">' + content + '</ins>',context).rootNodeIds : [], groupId = deleted[0] ?? inserted[0]!;
    for(const id of inserted)formatInserted(id,0);
    for (const [kind, ids] of [['delete', deleted], ['insert', inserted]] as const)
        for (const id of ids)
            word.revisions.push({ nodeId: id, revisionId: id, sourceId: id, groupId, kind, author, date });
}
function nativeReview(word: WordMetadata, revision: WordRevision, decision: 'accept' | 'reject'): void {
    if (!word.nativeContent)
        return;
    const property = ({ run: 'rPr', paragraph: 'pPr', table: 'tblPr', row: 'trPr', cell: 'tcPr', section: 'sectPr', tableGrid: 'tblGrid' } as const)[revision.property ?? 'run'], change = property + 'Change';
    word.nativeContent.fragments = word.nativeContent.fragments.flatMap(fragment => {
        if (fragment.nodeId !== revision.nodeId)
            return [fragment];
        if (fragment.revisionSourceId === revision.sourceId) {
            if (decision === 'reject' && fragment.revisionPrevious) {
                const path = [...fragment.propertyPath],indices=fragment.propertyIndices?[...fragment.propertyIndices]:undefined, index = path.findIndex(name => name.endsWith(' ' + change));
                if (index >= 0) {
                    {path.splice(index, 1);indices?.splice(index,1);}
                    if (path[index] === path[index - 1]||fragment.ownerName===property&&path[index]?.endsWith(' '+property))
                        {path.splice(index, 1);indices?.splice(index,1);}
                }
                const next = { ...fragment, propertyPath: path,...(indices?{propertyIndices:indices}:{}) };
                delete next.revisionSourceId;
                delete next.revisionPrevious;
                return [next];
            }
            return [];
        }
        if (decision === 'reject' && !revision.before?.inheritNativeProperties && revision.kind === 'format' && fragment.kind === 'property' && (fragment.ownerName===property||fragment.propertyPath[0]?.endsWith(' ' + property) || !fragment.propertyPath.length && fragment.propertyName?.endsWith(' ' + property)))
            return [];
        return [fragment];
    });
}
export function applyWordRevision(word: WordMetadata, edit: DocumentEdit, parse: () => ParsedHtml, operation: WordRevisionOperation,limits:DocumentReadLimits,sections:DocumentSection[],styles:SourceStyleEdit): void {
    if (operation.op === 'setWordReview') {
        if (parse().peekNode(operation.nodeId)!.parentNodeId !== null)
            fail('Review settings require the document root.');
        word.trackChanges = operation.trackChanges;
        if (operation.author !== undefined)
            word.trackingAuthor = operation.author;
        return;
    }
    if(operation.op==='reconcileWordTracked'){reconcileTrackedWord(word,edit,parse,operation);return;}
    if(operation.op==='trackWordInlineContent'){trackWordInlineContent(word,edit,parse,operation);return;}
    if (operation.op === 'trackWordText') {
        trackWordText(word, edit, parse, operation,styles);
        return;
    }
    let selected: WordRevision[];
    if (operation.op === 'reviewWordRevisions') {
        if (parse().peekNode(operation.nodeId)!.parentNodeId !== null)
            fail('Reviewing all revisions requires the document root.');
        selected = word.revisions.filter(revision => operation.author === undefined || revision.author === operation.author);
    }
    else {
        const revision = word.revisions.find(revision => wordRevisionId(revision) === operation.revisionId && revision.nodeId === operation.nodeId) ?? fail('The selected revision is missing.');
        selected = word.revisions.filter(item => item === revision || revision.groupId && item.groupId === revision.groupId || revision.moveId && item.moveId === revision.moveId);
    }
    selected=[...expandWordRevisionGroups(word.revisions,selected.map(wordRevisionId))];
    const beforeReview=parse(),linkedText=(source:ParsedHtml,link:WordMetadata['hyperlinks'][number]):number=>source.hasNode(link.nodeId)&&(!link.endNodeId||source.hasNode(link.endNodeId))?(link.endNodeId?source.textOffset(link.endNodeId)!-source.textOffset(link.nodeId)!:source.textLength(link.nodeId)!):0;
    const nonemptyLinks=new Set(word.hyperlinks.filter(link=>linkedText(beforeReview,link)>0).map(link=>link.nodeId));
    const rowRemovals=new Set(selected.filter(revision=>operation.decision==='accept'?revision.kind==='rowDelete':revision.kind==='rowInsert').map(revision=>revision.nodeId));if(rowRemovals.size)removeWordRows(word,edit,parse,rowRemovals,limits);
    const cellRemovals=new Set(selected.filter(revision=>operation.decision==='accept'?revision.kind==='cellDelete':revision.kind==='cellInsert').map(revision=>revision.nodeId));if(cellRemovals.size)removeWordCells(word,edit,parse,cellRemovals,limits);
    const removals = new Set<string>();
    for (const revision of selected) {
        nativeReview(word, revision, operation.decision);
        const node = parse().peekNode(revision.nodeId);
        if (!node)
            continue;
        if (revision.kind === 'format') {
            if (operation.decision === 'reject') {
                const before = revision.before ?? fail('The revision has no previous formatting.');
                if(before.conditionalFormat!==undefined){word.conditionalFormats=word.conditionalFormats?.filter(item=>item.nodeId!==revision.nodeId);if(before.conditionalFormat){word.conditionalFormats??=[];word.conditionalFormats.push({nodeId:revision.nodeId,format:{...before.conditionalFormat}});}}
                if(before.section){const section=sections.find(section=>section.rootNodeId===revision.nodeId)??fail('The reviewed section is missing.');for(const key of Object.keys(section))if(key!=='rootNodeId'&&key!=='sectionId')Reflect.deleteProperty(section,key);Object.assign(section,JSON.parse(JSON.stringify(before.section)));}
                if (before.attributes)
                    edit.attributes(revision.nodeId, Object.fromEntries(Object.entries(before.attributes).map(([name, value]) => [name, value ?? undefined])));
                if (before.tagName)
                    edit.renameSource(revision.nodeId, before.tagName);
                if (before.tabs) {
                    const old = word.tabs.findIndex(tab => tab.paragraphNodeId === revision.nodeId), tab = { paragraphNodeId: revision.nodeId, ...before.tabs };
                    if (old < 0)
                        word.tabs.push(tab);
                    else
                        word.tabs[old] = tab;
                }
                for (const [key, value] of Object.entries(before.properties ?? {}))
                    edit.sourceStyle(revision.nodeId, key, value);
                if(before.tableStyleOptions){const table=word.tables.find(table=>table.nodeId===revision.nodeId);if(table)for(const key of ['lookMask','rowBandSize','columnBandSize']as const){delete table[key];const value=before.tableStyleOptions[key];if(value!==undefined)table[key]=value;}}
                if (before.grid) {
                    const table = word.tables.find(table => table.nodeId === revision.nodeId);
                    if (table){
                        table.grid = [...before.grid];
                        const current=parse();for(const row of wordTableRows(current,table.nodeId))for(const id of current.peekNode(row)!.childNodeIds){const cell=current.peekNode(id)!,side=cell.attributes['data-word-grid-gap'];if(side!=='before'&&side!=='after')continue;const count=Number(cell.attributes.colspan??1),width=(side==='before'?table.grid.slice(0,count):table.grid.slice(-count)).reduce((a,b)=>a+b,0);edit.sourceStyle(id,'width',width+'pt');}
                    }
                }
            }
            continue;
        }
        const insertion = ['insert', 'moveTo', 'paragraphInsert','paragraphMoveTo', 'rowInsert', 'cellInsert'].includes(revision.kind), remove = operation.decision === 'accept' ? !insertion : insertion;
        if (['insert', 'delete', 'moveFrom', 'moveTo'].includes(revision.kind)) {
            if (remove) {
                edit.removeSource(revision.nodeId);
                removals.add(revision.nodeId);
            }
            else
                edit.unwrapSource(revision.nodeId);
        }
        else if (revision.kind.startsWith('paragraph')) {
            if (remove)
                edit.mergeSourceParagraphMark(revision.nodeId);
        }
        else if (remove) {
            edit.removeSource(revision.nodeId);
            removals.add(revision.nodeId);
        }
    }
    const moves=new Set(selected.flatMap(revision=>revision.moveId?[revision.moveId]:[]));for(const move of word.moves??[])if(moves.has(move.moveId))for(const id of [move.startNodeId,move.endNodeId]){const source=parse(),node=source.peekNode(id);if(!node)continue;const parent=node.parentNodeId?source.peekNode(node.parentNodeId):undefined,wrapper=parent&&parent.tagName==='div'&&parent.childNodeIds.length===1&&Object.keys(parent.attributes).every(name=>name==='data-node-id')&&source.getNode(parent.nodeId)!.text===''&&!word.stories.some(story=>story.rootNodeId===parent.nodeId);edit.removeSource(wrapper?parent!.nodeId:id);}if(word.moves)word.moves=word.moves.filter(move=>!moves.has(move.moveId));
    const reviewed = new Set(selected.map(wordRevisionId));
    word.revisions = word.revisions.filter(revision => !reviewed.has(wordRevisionId(revision)) && !removals.has(revision.nodeId));
    // Native hyperlinks can surround inserted runs. Rejecting their complete
    // content must remove the empty link while retaining any range markers
    // belonging to other semantics inside its wrapper.
    for(const link of word.hyperlinks)if(nonemptyLinks.has(link.nodeId)&&linkedText(parse(),link)===0&&parse().hasNode(link.nodeId)){
        if(link.endNodeId){edit.removeSource(link.nodeId);if(parse().hasNode(link.endNodeId))edit.removeSource(link.endNodeId);}
        else edit.unwrapSource(link.nodeId);
    }
}

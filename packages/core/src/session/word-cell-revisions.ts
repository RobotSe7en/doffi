import type {WordMetadata} from '../contracts/word.js';
import type {ParsedHtml} from '../document/html.js';
import type {DocumentEdit} from '../document/edit.js';
import type {DocumentReadLimits} from '../runtime.js';
import {DocumentError} from '../errors.js';
import {checkLimit} from '../document/validation.js';
import {editSourceTable} from './table-edit.js';
import {moveWordContinuationFormats} from './word-conditional-format.js';
import {wordTableRows} from './word-row-revisions.js';
import {recordWordFormat} from './word-format-revisions.js';

interface Cell {id:string;row:number;column:number;rows:number;columns:number;gap:boolean}
/** A bounded, transient projection of the authoritative HTML cells. */
export function wordTableCells(source:ParsedHtml,tableId:string,columns:number,limits:DocumentReadLimits):{rows:string[];cells:Cell[];occupied:Cell[][]}{
    const rows=wordTableRows(source,tableId),cells:Cell[]=[],occupied:Cell[][]=Array.from({length:rows.length},()=>[]);
    checkLimit(rows.length*columns,limits.maxNodes);
    for(const[row,id]of rows.entries())for(const child of source.peekNode(id)!.childNodeIds){
        const node=source.peekNode(child)!;let column=0;while(occupied[row]![column])column++;
        const cell={id:child,row,column,rows:Number(node.attributes.rowspan??1)||rows.length-row,columns:Number(node.attributes.colspan??1),gap:!!node.attributes['data-word-grid-gap']};
        if(row+cell.rows>rows.length||column+cell.columns>columns)throw new DocumentError('InvalidOperation','A revision cell exceeds its table grid.');
        for(let r=row;r<row+cell.rows;r++)for(let c=column;c<column+cell.columns;c++){if(occupied[r]![c])throw new DocumentError('InvalidOperation','Revision cells overlap.');occupied[r]![c]=cell;}
        cells.push(cell);
    }
    return{rows,cells,occupied};
}

/** Remove complete columns through the normal grid boundary. A single native
 * cell review shifts that row's remaining cells and retains a trailing omission. */
export function removeWordCells(word:WordMetadata,edit:DocumentEdit,parse:()=>ParsedHtml,nodeIds:ReadonlySet<string>,limits:DocumentReadLimits):void{
    const tables=new Set<string>();for(const id of nodeIds){let node=parse().peekNode(id);while(node&&node.tagName!=='table')node=node.parentNodeId?parse().peekNode(node.parentNodeId):undefined;if(node)tables.add(node.nodeId);}
    for(const tableId of tables){
        const info=word.tables.find(table=>table.nodeId===tableId);if(!info)throw new DocumentError('InvalidOperation','A cell review requires its native grid.');
        const old=wordTableCells(parse(),tableId,info.grid.length,limits),columns:number[]=[],candidates=new Set<number>();
        for(const cell of old.cells)if(nodeIds.has(cell.id))for(let column=cell.column;column<cell.column+cell.columns;column++)candidates.add(column);
        const spanHistory=new Map(word.revisions.filter(revision=>revision.kind==='format'&&revision.property==='cell'&&revision.before?.attributes&&'colspan'in revision.before.attributes).map(revision=>[revision.nodeId,Number(revision.before!.attributes!.colspan??1)]));
        const counts=[0];for(let column=0;column<info.grid.length;column++)counts.push(counts.at(-1)!+(candidates.has(column)?1:0));
        const expanded=new Set(old.cells.filter(cell=>spanHistory.has(cell.id)&&cell.columns - spanHistory.get(cell.id)! === counts[cell.column+cell.columns]!-counts[cell.column]!&&cell.columns>spanHistory.get(cell.id)!).map(cell=>cell.id));
        for(let column=0;column<info.grid.length;column++)if(candidates.has(column)&&old.occupied.every(row=>row[column]&&(row[column]!.gap||nodeIds.has(row[column]!.id)||expanded.has(row[column]!.id))))columns.push(column);
        if(columns.length===info.grid.length){edit.removeSource(tableId);continue;}
        const groups:{at:number;count:number}[]=[];for(const column of columns){const last=groups.at(-1);if(last&&last.at+last.count===column)last.count++;else groups.push({at:column,count:1});}
        const width=info.grid.reduce((a,b)=>a+b,0);
        for(const change of groups.reverse()){
            moveWordContinuationFormats(info,{kind:'deleteColumns',...change});
            const result=editSourceTable({edit,parse,limits,wordOmissions:true},tableId,info.grid,{width,height:old.rows.length*20},{kind:'deleteColumns',...change});
            info.grid=result.grid.map(value=>value*width/result.width);
            for(const cell of result.cells)edit.sourceStyle(cell.id,'width',info.grid.slice(cell.column,cell.column+cell.columns).reduce((a,b)=>a+b,0)+'pt');
        }
        const current=parse(),grid=wordTableCells(current,tableId,info.grid.length,limits);
        for(const[row,rowId]of grid.rows.entries()){
            const removed=grid.cells.filter(cell=>cell.row===row&&nodeIds.has(cell.id));if(!removed.length)continue;
            if(removed.some(cell=>cell.rows!==1)||grid.occupied[row]!.some(cell=>cell.rows!==1))throw new DocumentError('UnsupportedFeature','Review the complete column when a cell deletion intersects a vertical merge.');
            const count=removed.reduce((sum,cell)=>sum+cell.columns,0),after=grid.cells.find(cell=>cell.row===row&&current.peekNode(cell.id)!.attributes['data-word-grid-gap']==='after'),span=count+(after?.columns??0);
            for(const cell of removed)edit.removeSource(cell.id);
            const gapWidth=info.grid.slice(-span).reduce((a,b)=>a+b,0);
            if(after){edit.attributes(after.id,{colspan:span===1?undefined:String(span)});edit.sourceStyle(after.id,'width',gapWidth+'pt');}
            else edit.insertSource(rowId,'<td data-word-grid-gap="after"'+(span>1?' colspan="'+span+'"':'')+' style="width:'+gapWidth+'pt"></td>');
            edit.attributes(rowId,{'data-word-after-width-present':'true','data-word-after-width-type':'dxa','data-word-after-width-value':String(Math.round(gapWidth*20))});
        }
    }
}

export function trackWordColumnDeletion(word:WordMetadata,edit:DocumentEdit,parse:()=>ParsedHtml,tableId:string,at:number,count:number,limits:DocumentReadLimits):void{
    const info=word.tables.find(table=>table.nodeId===tableId)!;
    if(at+count>info.grid.length||count===info.grid.length)throw new DocumentError('InvalidOperation','Tracked column deletion must retain a nonempty table and stay within its grid.');
    const grid=wordTableCells(parse(),tableId,info.grid.length,limits),selected=grid.cells.filter(cell=>!cell.gap&&cell.column<at+count&&cell.column+cell.columns>at);
    if(selected.some(cell=>cell.column<at||cell.column+cell.columns>at+count||cell.rows!==1))throw new DocumentError('UnsupportedFeature','A tracked column deletion must contain complete cells without vertical merges.');
    const author=word.trackingAuthor??'',date=new Date().toISOString(),groupId=selected[0]?.id+'-column-delete',cancel=new Set<string>();
    if(selected.some(cell=>word.revisions.some(revision=>revision.nodeId===cell.id&&revision.kind==='cellInsert'&&revision.author!==author)))throw new DocumentError('UnsupportedFeature','Review another author\'s pending inserted cells before deleting their column.');
    for(const cell of selected){
        if(word.revisions.some(revision=>revision.nodeId===cell.id&&revision.kind==='cellInsert'&&revision.author===author)){cancel.add(cell.id);continue;}
        if(!word.revisions.some(revision=>revision.nodeId===cell.id&&revision.kind==='cellDelete'))word.revisions.push({nodeId:cell.id,sourceId:cell.id+'-cell-delete',revisionId:cell.id+'-cell-delete',groupId,kind:'cellDelete',author,date});
    }
    if(cancel.size)removeWordCells(word,edit,parse,cancel,limits);
}

/** Grid/property history and cell marks form one structural insertion. The
 * original grid is retained once, even after successive column insertions. */
export function prepareWordColumnInsertion(word:WordMetadata,source:ParsedHtml,tableId:string,limits:DocumentReadLimits):{cells:Cell[];groupId:string}{
    const info=word.tables.find(table=>table.nodeId===tableId)!,grid=wordTableCells(source,tableId,info.grid.length,limits);
    let revision=word.revisions.find(revision=>revision.nodeId===tableId&&revision.property==='tableGrid');
    const id=tableId+'-format-grid',groupId=revision?.groupId??id+'-insertion';
    if(!revision){revision={nodeId:tableId,sourceId:id,revisionId:id,kind:'format',property:'tableGrid',before:{grid:[...info.grid],inheritNativeProperties:true}};word.revisions.push(revision);}
    revision.groupId=groupId;revision.author=word.trackingAuthor??'';revision.date=new Date().toISOString();
    const formats=new Map(word.revisions.filter(revision=>revision.kind==='format').map(revision=>[revision.nodeId+'\n'+revision.property,revision]));
    for(const cell of grid.cells)if(!cell.gap){recordWordFormat(word,source,cell.id,'cell',formats);formats.get(cell.id+'\ncell')!.groupId=groupId;}
    for(const row of grid.rows)if(source.peekNode(row)!.childNodeIds.some(id=>source.peekNode(id)!.attributes['data-word-grid-gap'])){recordWordFormat(word,source,row,'row',formats);formats.get(row+'\nrow')!.groupId=groupId;}
    const ids=new Set(grid.cells.map(cell=>cell.id));for(const revision of word.revisions)if(revision.kind==='cellInsert'&&ids.has(revision.nodeId))revision.groupId=groupId;
    return{cells:grid.cells,groupId};
}

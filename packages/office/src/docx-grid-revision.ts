import {DocumentError} from '@doffi/core';
import type {WordMetadata,WordRevision} from '@doffi/core';
import {unsupported} from './ooxml/xml.js';
import type {XmlNode} from './ooxml/xml.js';
interface Node {tag:string;attrs:Record<string,string>;children:(Node|string)[]}
/** A column-count history is editable only when the native marks and old cell
 * spans describe the complete inverse. Ambiguous native structures stay native. */
export function bindWordGridInsertion(word:WordMetadata,table:Node,revision:WordRevision,source:XmlNode,maxCells:number):void{
    const tableId=table.attrs['data-node-id']!,info=word.tables.find(info=>info.nodeId===tableId)!,previous=revision.before!.grid!;
    const rows=table.children.filter((node):node is Node=>typeof node!=='string'&&node.tag==='tr'),cells:{node:Node;column:number;columns:number;rows:number;row:number}[]=[],occupied:typeof cells[]=Array.from({length:rows.length},()=>[]);
    if(rows.length*info.grid.length>maxCells)throw new DocumentError('ResourceLimitExceeded','Native grid revision mapping exceeds its cell budget.');
    const own=new Map(word.revisions.map(item=>[item.nodeId+'\n'+(item.property??item.kind),item])),inserted=new Set<number>();
    for(const[row,node]of rows.entries())for(const child of node.children){if(typeof child==='string')continue;let column=0;while(occupied[row]![column])column++;const cell={node:child,column,columns:Number(child.attrs.colspan??1),rows:Number(child.attrs.rowspan??1),row};cells.push(cell);for(let r=row;r<row+cell.rows;r++)for(let c=column;c<column+cell.columns;c++)occupied[r]![c]=cell;if(own.has(child.attrs['data-node-id']+'\ncellInsert'))for(let c=column;c<column+cell.columns;c++)inserted.add(c);}
    if(previous.length+inserted.size!==info.grid.length||!inserted.size)unsupported(source,'The changed native column count has no complete inserted-cell inverse.');
    const counts=[0];for(let column=0;column<info.grid.length;column++)counts.push(counts.at(-1)!+(inserted.has(column)?1:0));
    for(const cell of cells){const id=cell.node.attrs['data-node-id']!,gap=cell.node.attrs['data-word-grid-gap'],mark=own.get(id+'\ncellInsert'),format=own.get(id+'\ncell'),covered=counts[cell.column+cell.columns]!-counts[cell.column]!,old=format?.before?.attributes&&'colspan'in format.before.attributes?Number(format.before.attributes.colspan??1):cell.columns;if(!gap&&!mark&&cell.columns-old!==covered)unsupported(source,'A changed native grid does not match its original cell spans.');if(gap==='before'||gap==='after'){const row=own.get(rows[cell.row]!.attrs['data-node-id']+'\nrow'),previous=row?.before?.rowOmissions?.[gap]??cell.columns;if(cell.columns-previous!==covered)unsupported(source,'A changed native grid does not match its original row omissions.');}}
    const groupId=tableId+'-column-insertion',ids=new Set([tableId,...rows.map(row=>row.attrs['data-node-id']!),...cells.map(cell=>cell.node.attrs['data-node-id']!)]);revision.groupId=groupId;
    for(const item of word.revisions)if(ids.has(item.nodeId)&&(item.kind==='cellInsert'||item.kind==='format'&&['cell','row'].includes(item.property??''))){item.groupId=groupId;if(item.kind==='cellInsert'){if(item.author!==undefined)revision.author=item.author;if(item.date!==undefined)revision.date=item.date;}}
}

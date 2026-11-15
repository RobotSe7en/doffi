import type { ParsedHtml } from './html.js';
const cache=new WeakMap<ParsedHtml,Map<string,{before:number;after:number}>>();
/** Keep edge spans compressed, including regions crossing refined rows. */
export function wordRowOmissions(source:ParsedHtml,rowId:string):{before:number;after:number}{
    let values=cache.get(source);if(!values){values=new Map();cache.set(source,values);}const known=values.get(rowId);if(known)return known;
    let table=source.peekNode(rowId)!;while(table.tagName!=='table'&&table.parentNodeId)table=source.peekNode(table.parentNodeId)!;
    const rows=table.childNodeIds.flatMap(id=>{const row=source.peekNode(id)!;return row.tagName==='tr'?[id]:row.childNodeIds.filter(id=>source.peekNode(id)!.tagName==='tr');}),active:{before?:{until:number;count:number};after?:{until:number;count:number}}={};
    for(const[index,id]of rows.entries()){
        for(const side of ['before','after']as const)if(active[side]&&active[side]!.until<=index)delete active[side];
        for(const child of source.peekNode(id)!.childNodeIds){const node=source.peekNode(child)!,side=node.attributes['data-word-grid-gap'];if(side!=='before'&&side!=='after')continue;active[side]={until:index+(Number(node.attributes.rowspan??1)||rows.length-index),count:Number(node.attributes.colspan??1)};}
        values.set(id,{before:active.before?.count??0,after:active.after?.count??0});
    }
    return values.get(rowId)??{before:0,after:0};
}

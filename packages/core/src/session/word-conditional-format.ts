import type {WordMetadata} from '../contracts/word.js';
import type {TableChange} from '../contracts/table-operations.js';
/** Optimization coordinates follow ordinary grid edits; merge/split rebuild
 * their native continuation structure, so old continuation caches are retired. */
export function moveWordContinuationFormats(table:WordMetadata['tables'][number],change:TableChange):void {
 if(!table.continuationFormats)return;
 if(change.kind==='insertRows'){
  for(const entry of table.continuationFormats)if(entry.row>=change.at)entry.row+=change.count;
 }else if(change.kind==='insertColumns'){
  for(const entry of table.continuationFormats)if(entry.column>=change.at)entry.column+=change.count;
 }else if(change.kind==='deleteRows'){
  table.continuationFormats=table.continuationFormats.filter(entry=>entry.row<change.at||entry.row>=change.at+change.count).map(entry=>({...entry,row:entry.row>=change.at+change.count?entry.row-change.count:entry.row}));
 }else if(change.kind==='deleteColumns'){
  table.continuationFormats=table.continuationFormats.filter(entry=>entry.column<change.at||entry.column>=change.at+change.count).map(entry=>({...entry,column:entry.column>=change.at+change.count?entry.column-change.count:entry.column}));
 }else if(change.kind==='mergeCells'||change.kind==='splitCell')delete table.continuationFormats;
}

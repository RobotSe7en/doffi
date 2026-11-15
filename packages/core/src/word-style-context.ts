import type { WordMetadata } from './contracts/word.js';
import type { DeepReadonly } from './runtime.js';
import { DocumentError } from './errors.js';
type Css = Record<string, string>;
type Word = DeepReadonly<Pick<WordMetadata,'styles'>>;
type Style = Word['styles'][number];
type Read = (name: string) => Css;
const empty = () => ({ p: {} as Css, r: {} as Css, table: {} as Css, cell: {} as Css,row:{} as Css });
const toggles: Record<string, [
    string,
    string
]> = { 'font-weight': ['bold', 'normal'], 'font-style': ['italic', 'normal'], 'font-variant': ['small-caps', 'normal'], 'text-transform': ['uppercase', 'none'] };
export function applyWordStyleRun(base: Css, own: Css): Css { const result = { ...base, ...own }; for (const [key, [active, inactive]] of Object.entries(toggles))
    if (own[key] === active)
        result[key] = base[key] === active || key === 'font-weight' && Number(base[key]) >= 600 ? inactive : active; return result; }
function chain(word: Word, id: string | undefined): Style[] { const result: Style[] = [], seen = new Set<string>(); while (id) {
    if (seen.has(id))
        throw new DocumentError('InvalidDocument', 'Style inheritance is cyclic.');
    seen.add(id);
    const style = word.styles.find(style => style.styleId === id);
    if (!style)
        break;
    result.unshift(style);
    id = style.basedOn;
} return result; }
export function resolveWordStyleRun(word: Word, read: Read, id: string | undefined, inherited: Css): Css { let result = { ...inherited }; for (const style of chain(word, id)) {
    if (style.components?.r)
        result = applyWordStyleRun(result, read(style.components.r));
    else {
        const css = { ...read(style.cssClass), ...(style.runCssClass ? read(style.runCssClass) : {}) };
        const keys = style.runProperties ?? Object.keys(css).filter(key => /^(font-|color$|--word-|--office-|text-|letter-spacing$|word-spacing$|vertical-align$)/.test(key));
        result = applyWordStyleRun(result, Object.fromEntries(keys.flatMap(key => css[key] !== undefined ? [[key, css[key]!]] : [])));
    }
} return result; }
/** Native precedence: bands, columns, rows, then corner cells. */
export function wordTableRegions(row: number, column: number, rowCount: number, columns: number, span: number, mask: number, rowBand = 0, columnBand = 0): string[] { const result: string[] = []; const firstRow = !!(mask & 32), lastRow = !!(mask & 64), firstCol = !!(mask & 128), lastCol = !!(mask & 256); if (rowBand > 0 && !(mask & 512) && !(row === 0 && firstRow) && !(row === rowCount - 1 && lastRow))
    result.push(Math.floor((row - (firstRow ? 1 : 0)) / rowBand) % 2 === 0 ? 'band1Horz' : 'band2Horz'); if (columnBand > 0 && !(mask & 1024) && !(column === 0 && firstCol) && !(column + span === columns && lastCol))
    result.push(Math.floor((column - (firstCol ? 1 : 0)) / columnBand) % 2 === 0 ? 'band1Vert' : 'band2Vert'); if (column === 0 && firstCol)
    result.push('firstCol'); if (column + span === columns && lastCol)
    result.push('lastCol'); if (row === 0 && firstRow)
    result.push('firstRow'); if (row === rowCount - 1 && lastRow)
    result.push('lastRow'); if (result.includes('firstRow') && result.includes('firstCol'))
    result.push('nwCell'); if (result.includes('firstRow') && result.includes('lastCol'))
    result.push('neCell'); if (result.includes('lastRow') && result.includes('firstCol'))
    result.push('swCell'); if (result.includes('lastRow') && result.includes('lastCol'))
    result.push('seCell'); return result; }
export function resolveWordTableStyle(word: Word, read: Read, id: string | undefined, regions: readonly string[]) {
    const result = empty(), styles = chain(word, id);
    for (const style of styles) {
        if (style.components)
            for (const key of ['p', 'r', 'table', 'cell','row'] as const) {
                const name = style.components[key];
                if (name)
                    result[key] = {...result[key],...read(name)};
            }
        else {
            Object.assign(result.table, read(style.cssClass));
        }
    }
    for (const region of regions.filter(region=>region!=='wholeTable'))
        for (const style of styles) {
            const condition = style.conditions?.find(item => item.region === region);
            if (!condition)
                continue;
            if (condition.components)
                for (const key of ['p', 'r', 'table', 'cell','row'] as const) {
                    const name = condition.components[key];
                    if (name)
                        result[key] = {...result[key],...read(name)};
                }
            else
                Object.assign(result.cell, read(condition.cssClass));
        }
    return result;
}

const regionOrder=['band1Horz','band2Horz','band1Vert','band2Vert','firstCol','lastCol','firstRow','lastRow','nwCell','neCell','swCell','seCell'];
/** Only four column samples are needed to cover both bands and table edges. */
export function wordTableRowRegions(row:number,rows:number,columns:number,mask:number,rowBand=0,columnBand=0):string[]{
 const start=mask&128?1:0,selected=new Set([0,columns-1,start,start+columnBand].filter(column=>column>=0&&column<columns).flatMap(column=>wordTableRegions(row,column,rows,columns,1,mask,rowBand,columnBand)));
 return regionOrder.filter(region=>selected.has(region));
}
export function wordTableRowCellProperties(css:Record<string,string>):Record<string,string>{return Object.fromEntries(Object.entries(css).filter(([key])=>key.startsWith('padding-')||key==='white-space'||key==='vertical-align'));}

/** Physical continuation cells resolve their own row/column conditions even
 * though the editable HTML represents the merged content with one rowspan. */
export function wordTableCellProperties(word:Word,read:Read,styleId:string|undefined,info:DeepReadonly<WordMetadata['tables'][number]>,row:number,column:number,rows:number,span:number):Css{
 const styles=chain(word,styleId),setting=(key:'tableLookMask'|'rowBandSize'|'columnBandSize')=>[...styles].reverse().find(style=>style[key]!==undefined)?.[key]??0;
 const mask=info.lookMask??setting('tableLookMask'),rowBand=info.rowBandSize??setting('rowBandSize'),columnBand=info.columnBandSize??setting('columnBandSize'),base=resolveWordTableStyle(word,read,styleId,[]),value=resolveWordTableStyle(word,read,styleId,wordTableRegions(row,column,rows,info.grid.length,span,mask,rowBand,columnBand)),rowStyle=resolveWordTableStyle(word,read,styleId,wordTableRowRegions(row,rows,info.grid.length,mask,rowBand,columnBand)),inner:Css={};
 if(base.table['--word-border-inside-h']&&row>0)inner['border-top']='var(--word-border-inside-h)';if(base.table['--word-border-inside-v']&&column>0)inner['border-left']='var(--word-border-inside-v)';
 return{'padding-left':'5.4pt','padding-right':'5.4pt',...value.cell,...wordTableRowCellProperties(rowStyle.cell),...inner};
}

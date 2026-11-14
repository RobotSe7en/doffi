import { DocumentError,copyWorkbookFormula } from '@doffi/core';
import type { DocumentReadLimits } from '@doffi/core';
import { openOpc } from './ooxml/package.js';
import { R, S, attr, child, elements, text } from './ooxml/xml.js';
import type { XmlNode } from './ooxml/xml.js';
import { serializeXml } from './ooxml/serialize.js';
import { writeOpcZip } from './document-zip.js';
type Cell = string | number | null;
export interface ChartWorkbookBudget {
    bytes: number;
    nodes: number;
}
const error = (message: string): never => { throw new DocumentError('UnsupportedFeature', message); };
function address(raw: string): {
    row: number;
    column: number;
} {
    const match = /^\$?([A-Z]{1,3})\$?([1-9][0-9]*)$/i.exec(raw);
    if (!match)
        error('A chart data binding requires an A1 cell range.');
    let column = 0;
    for (const character of match![1]!.toUpperCase())
        column = column * 26 + character.charCodeAt(0) - 64;
    const row = Number(match![2]);
    if (column > 16384 || row > 1048576)
        error('A chart data range exceeds worksheet dimensions.');
    return { row, column };
}
function column(value: number): string {
    let result = '';
    for (; value; value = Math.floor((value - 1) / 26))
        result = String.fromCharCode(65 + (value - 1) % 26) + result;
    return result;
}
/** Edits only chart-bound cells in the original embedded package. Other cells,
 * formulas, styles, named ranges and workbook parts retain their source bytes. */
export class ChartWorkbookPatch {
    private readonly opc;
    private readonly workbook: XmlNode;
    private readonly workbookPart: string;
    private readonly changed = new Map<string, XmlNode>();
    constructor(bytes: Uint8Array, private readonly limits: DocumentReadLimits, budget: ChartWorkbookBudget) {
        const remaining = limits.maxPackageBytes - budget.bytes;
        if (remaining <= 0)
            throw new DocumentError('ResourceLimitExceeded', 'Embedded chart workbooks exceed their aggregate budget.');
        this.opc = openOpc(bytes, { limits: { ...limits, maxPackageBytes: remaining, maxNodes: Math.max(1, limits.maxNodes - budget.nodes) } });
        this.workbookPart = this.opc.related('', 'officeDocument') ?? error('A chart workbook main part is missing.');
        this.workbook = this.opc.read(this.workbookPart);
        if (this.workbook.uri !== S || this.workbook.name !== 'workbook')
            error('A chart workbook has an invalid main part.');
        budget.bytes += [...this.opc.files.values()].reduce((sum, bytes) => sum + bytes.length, 0);
        this.sharedBudget = budget;
    }
    private expandedFormulaUnits=0;
    private readonly sharedGroups=new WeakMap<XmlNode,Map<string,XmlNode[]>>();
    private readonly rowIndexes = new WeakMap<XmlNode, Map<number, XmlNode>>();
    private readonly cellIndexes = new WeakMap<XmlNode, Map<string, XmlNode>>();
    private indexSheet(root: XmlNode): Map<number, XmlNode> {
        const known = this.rowIndexes.get(root);
        if (known)
            return known;
        const rows = new Map<number, XmlNode>(),groups=new Map<string,XmlNode[]>();this.sharedGroups.set(root,groups);
        let previous = 0;
        for (const row of elements(child(root, 'sheetData', S), 'row', S)) {
            const r = attr(row, 'r', '') === undefined ? previous + 1 : Number(attr(row, 'r', ''));
            if (!Number.isSafeInteger(r) || r <= previous || r > 1048576)
                error('Native chart worksheet rows require ordered valid positions.');
            previous = r;
            (row.attrs as Map<string, string>).set(' r', String(r));
            rows.set(r, row);
            const cells = new Map<string, XmlNode>();
            let columnIndex = 0;
            for (const cell of elements(row, 'c', S)) {
                const value = attr(cell, 'r', ''), position = value === undefined ? { row: r, column: columnIndex + 1 } : address(value);
                if (position.row !== r || position.column <= columnIndex || position.column > 16384)
                    error('Native chart cells require ordered valid positions.');
                columnIndex = position.column;
                const ref = column(columnIndex) + r;
                (cell.attrs as Map<string, string>).set(' r', ref);
                cells.set(ref, cell);const formula=child(cell,'f',S);if(attr(formula,'t','')==='shared'){const id=attr(formula,'si','');if(id===undefined)error('A native shared formula identity is missing.');const members=groups.get(id!)??[];members.push(cell);groups.set(id!,members);}
            }
            this.cellIndexes.set(row, cells);
        }
        this.rowIndexes.set(root, rows);
        return rows;
    }
    private separateSharedFormula(root:XmlNode,cell:XmlNode):void{
     const formula=child(cell,'f',S);if(attr(formula,'t','')!=='shared')return;const id=attr(formula,'si','')!,members=this.sharedGroups.get(root)?.get(id),master=members?.find(cell=>text(child(cell,'f',S)!).trim());if(!members||!master)error('A shared chart formula has no source.');const source=text(child(master!,'f',S)!),from=address(attr(master!,'r','')!),expanded=members!.map(cell=>{const to=address(attr(cell,'r','')!),formula=copyWorkbookFormula('='+source,{row:from.row-1,column:from.column-1},{row:to.row-1,column:to.column-1}).slice(1);this.expandedFormulaUnits+=formula.length;if(this.expandedFormulaUnits>this.limits.maxFileBytes)throw new DocumentError('ResourceLimitExceeded','Separating native chart formulas exceeds the output text budget.');return{cell,formula};});
     for(const item of expanded){const node=child(item.cell,'f',S)!;for(const name of ['t','si','ref'])(node.attrs as Map<string,string>).delete(' '+name);node.content.splice(0,node.content.length,item.formula);}this.sharedGroups.get(root)!.delete(id);
    }
    private readonly sharedBudget: ChartWorkbookBudget;
    private range(formula: string, seen = new Set<string>()): {
        sheet: string;
        from: {
            row: number;
            column: number;
        };
        to: {
            row: number;
            column: number;
        };
        expression: string;
        names?: XmlNode[];
    } {
        const raw = formula.replace(/^=/, '');
        let match = /^(?:'((?:[^']|'')+)'|([^'!]+))!(\$?[A-Z]{1,3}\$?[1-9][0-9]*)(?::(\$?[A-Z]{1,3}\$?[1-9][0-9]*))?$/i.exec(raw);
        if (!match) {
            if (seen.has(raw) || seen.size > 32)
                error('A chart defined-name binding is cyclic.');
            seen.add(raw);
            const name = elements(child(this.workbook, 'definedNames', S), 'definedName', S).find(node => attr(node, 'name', '')?.toLowerCase() === raw.toLowerCase() && attr(node, 'localSheetId', '') === undefined);
            if (!name)
                error('This chart formula is not a static native data range.');
            const nested = this.range(text(name!), seen);
            return { ...nested, names: [name!, ...(nested.names ?? [])] };
        }
        const sheetName = (match[1]?.replaceAll("''", "'") ?? match[2]!).replace(/^\[0\]/, ''), sheet = elements(child(this.workbook, 'sheets', S), 'sheet', S).find(node => attr(node, 'name', '')?.toLowerCase() === sheetName.toLowerCase());
        if (!sheet)
            error('A chart data worksheet is missing.');
        const relation = this.opc.relationships(this.workbookPart).get(attr(sheet, 'id', R) ?? '');
        if (!relation || relation.external || !relation.type.endsWith('/worksheet'))
            error('A chart data worksheet relationship is invalid.');
        const from = address(match[3]!), to = address(match[4] ?? match[3]!);
        if (to.row < from.row || to.column < from.column || to.row !== from.row && to.column !== from.column)
            error('A chart series must bind one worksheet row or column.');
        return { sheet: relation!.target, from, to, expression: raw };
    }
    validate(formula: string): void { this.range(formula); }
    write(formula: string, values: readonly Cell[], previous?: readonly Cell[]): string | undefined {
        const range = this.range(formula), vertical = range.from.column === range.to.column, oldCount = vertical ? range.to.row - range.from.row + 1 : range.to.column - range.from.column + 1, count = Math.max(oldCount, values.length);
        if (count > this.limits.maxNodes || range.from.row + (vertical ? count - 1 : 0) > 1048576 || range.from.column + (vertical ? 0 : count - 1) > 16384)
            throw new DocumentError('ResourceLimitExceeded', 'A chart vector exceeds its bounded worksheet range.');
        let root = this.changed.get(range.sheet);
        if (!root) {
            root = this.opc.read(range.sheet);
            this.changed.set(range.sheet, root);
        }
        let data = child(root, 'sheetData', S);
        if (!data) {
            data = { uri: S, name: 'sheetData', attrs: new Map(), namespaces: {}, content: [], part: range.sheet };
            root.content.push(data);
        }
        const rows = this.indexSheet(root), changedRows = new Set<XmlNode>();
        for (let index = 0; index < count; index++) {
            const samePoint=!!previous&&index<oldCount&&index<previous.length&&index<values.length&&previous[index]===values[index];
            const r = range.from.row + (vertical ? index : 0), c = range.from.column + (vertical ? 0 : index), ref = column(c) + r;
            let row = rows.get(r);
            if (!row) {
                row = { uri: S, name: 'row', attrs: new Map([[' r', String(r)]]), namespaces: {}, content: [], part: range.sheet };
                rows.set(r, row);
                data.content.push(row);
            }
            const cells = this.cellIndexes.get(row) ?? new Map<string, XmlNode>();
            this.cellIndexes.set(row, cells);
            let cell = cells.get(ref);
            if(samePoint)continue;
            if(cell&&!samePoint)this.separateSharedFormula(root,cell);
            changedRows.add(row);
            if (!cell) {
                cell = { uri: S, name: 'c', attrs: new Map([[' r', ref]]), namespaces: {}, content: [], part: range.sheet };
                row.content.push(cell);
                cells.set(ref, cell);
            }
            cell.content.splice(0, cell.content.length, ...cell.content.filter(node => typeof node === 'string' || node.uri !== S || !(samePoint?['v','is']:['f','v','is']).includes(node.name)));
            (cell.attrs as Map<string, string>).delete(' t');
            const value = values[index];
            const make = (name: string, content: (XmlNode | string)[]): XmlNode => ({ uri: S, name, attrs: new Map(), namespaces: {}, content, part: range.sheet });
            if (typeof value === 'string') {
                (cell.attrs as Map<string, string>).set(' t', 'inlineStr');
                const token = make('t', [value]);
                (token.attrs as Map<string, string>).set('http://www.w3.org/XML/1998/namespace space', 'preserve');
                cell.content.push(make('is', [token]));
            }
            else if (typeof value === 'number')
                cell.content.push(make('v', [String(value)]));
            const extra = cell.content.filter(node => typeof node !== 'string' && node.uri === S && node.name === 'extLst');
            if (extra.length)
                cell.content.splice(0, cell.content.length, ...cell.content.filter(node => !extra.includes(node as XmlNode)), ...extra);
        }
        for (const row of changedRows) {
            const cells = [...this.cellIndexes.get(row)!.values()].sort((a, b) => address(attr(a, 'r', '')!).column - address(attr(b, 'r', '')!).column), cellSet = new Set(cells), other = row.content.filter(node => !cellSet.has(node as XmlNode));
            row.content.splice(0, row.content.length, ...cells, ...other);
            (row.attrs as Map<string, string>).delete(' spans');
        }
        const ordered = [...rows].sort((a, b) => a[0] - b[0]).map(([, row]) => row), rowSet = new Set(ordered);
        data.content.splice(0, data.content.length, ...ordered, ...data.content.filter(node => !rowSet.has(node as XmlNode)));
        root.content.splice(0, root.content.length, ...root.content.filter(node => typeof node === 'string' || node.uri !== S || node.name !== 'dimension'));
        const end = { row: range.from.row + (vertical ? values.length - 1 : 0), column: range.from.column + (vertical ? 0 : values.length - 1) };
        if (!values.length)
            return undefined;
        const expression = range.expression.replace(/(\$?[A-Z]{1,3}\$?[1-9][0-9]*)(?::\$?[A-Z]{1,3}\$?[1-9][0-9]*)?$/i, () => '$' + column(range.from.column) + '$' + range.from.row + (values.length === 1 ? '' : ':$' + column(end.column) + '$' + end.row));
        if (range.names?.length) {
            const name = range.names.at(-1)!;
            name.content.splice(0, name.content.length, expression);
            this.changed.set(this.workbookPart, this.workbook);
            return formula;
        }
        return expression;
    }
    append(header: string, values: readonly Cell[]): {
        nameFormula: string;
        valueFormula?: string;
    } {
        const sheet = elements(child(this.workbook, 'sheets', S), 'sheet', S)[0];
        if (!sheet)
            error('The native chart workbook has no worksheet.');
        const name = attr(sheet, 'name', '')!, relation = this.opc.relationships(this.workbookPart).get(attr(sheet, 'id', R) ?? '');
        if (!relation || relation.external)
            error('The native chart data sheet is unavailable.');
        const root = this.changed.get(relation!.target) ?? this.opc.read(relation!.target);
        this.indexSheet(root);
        let maximum = 0;
        for (const row of elements(child(root, 'sheetData', S), 'row', S))
            for (const cell of elements(row, 'c', S))
                maximum = Math.max(maximum, address(attr(cell, 'r', '') ?? 'A1').column);
        if (maximum >= 16384)
            throw new DocumentError('ResourceLimitExceeded', 'There is no free worksheet column for chart data.');
        const prefix = "'" + name.replaceAll("'", "''") + "'!$" + column(maximum + 1) + '$', nameFormula = prefix + '1';
        this.write(nameFormula, [header]);
        return { nameFormula, ...(values.length ? { valueFormula: this.write(prefix + '2', values)! } : {}) };
    }
    finish(): Uint8Array {
        const files = new Map(this.opc.files);
        if (this.changed.size) {
            const chain = this.opc.related(this.workbookPart, 'calcChain');
            if (chain) {
                files.delete(chain);
                const relation = [...this.opc.relationships(this.workbookPart).values()].find(relation => relation.target === chain), slash = this.workbookPart.lastIndexOf('/'), path = this.workbookPart.slice(0, slash + 1) + '_rels/' + this.workbookPart.slice(slash + 1) + '.rels';
                if (relation && files.has(path)) {
                    const root = this.opc.read(path);
                    root.content.splice(0, root.content.length, ...root.content.filter(node => typeof node === 'string' || attr(node, 'Id', '') !== relation.id));
                    this.changed.set(path, root);
                }
                const types = this.opc.read('[Content_Types].xml');
                types.content.splice(0, types.content.length, ...types.content.filter(node => typeof node === 'string' || attr(node, 'PartName', '') !== '/' + chain));
                this.changed.set('[Content_Types].xml', types);
            }
            let calc = child(this.workbook, 'calcPr', S);
            if (!calc) {
                calc = { name: 'calcPr', uri: S, attrs: new Map(), namespaces: {}, content: [], part: this.workbookPart };
                const after = new Set(['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst']), index = this.workbook.content.findIndex(node => typeof node !== 'string' && node.uri === S && after.has(node.name));
                this.workbook.content.splice(index < 0 ? this.workbook.content.length : index, 0, calc);
            }
            for (const [key, value] of [['fullCalcOnLoad', '1'], ['forceFullCalc', '1'], ['calcId', '0']])
                (calc.attrs as Map<string, string>).set(' ' + key, value!);
            this.changed.set(this.workbookPart, this.workbook);
        }
        const codecs = globalThis as unknown as {
            TextEncoder: new () => {
                encode(text: string): Uint8Array;
            };
        };
        for (const [path, root] of this.changed)
            files.set(path, new codecs.TextEncoder().encode(serializeXml(root)));
        this.sharedBudget.nodes += this.opc.xmlBudget.nodes;
        if (this.sharedBudget.nodes > this.limits.maxNodes)
            throw new DocumentError('ResourceLimitExceeded', 'Embedded chart workbook XML exceeds its aggregate budget.');
        return writeOpcZip(files, { limits: this.limits });
    }
}

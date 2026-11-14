import type { DocumentReadLimits } from "../runtime.js";
import type { DocumentEdit } from "../document/edit.js";
import type { ParsedHtml } from "../document/html.js";
import type { TableChange } from "../contracts/table-operations.js";
import { checkLimit, invalid } from "../document/validation.js";
interface Cell {
    id: string;
    row: number;
    column: number;
    rows: number;
    columns: number;
}
const escape = (text: string): string => text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\r\n?/g, "\n").replace(/\n/g, "<br>");
/** Table occupancy is transient; persistent cells and grid remain HTML and metadata. */
export interface SourceTableEditContext { readonly edit: DocumentEdit; readonly parse: () => ParsedHtml; readonly limits: DocumentReadLimits; readonly newCellHtml?: string }
export function editSourceTable(ctx: SourceTableEditContext, tableId: string, initialGrid: readonly number[], geometry: { width: number; height: number }, change: TableChange): { grid: number[]; width: number; height: number; cells: Cell[] } {
    const { edit } = ctx, object = { grid: [...initialGrid] }, html = ctx.parse(), table = html.peekNode(tableId)!;
    if (!table || table.tagName !== "table") invalid("Table operations require an HTML table source.");
    const rows = table.childNodeIds.flatMap(id => { const node = html.peekNode(id)!; return node.tagName === "tr" ? [id] : node.childNodeIds; });
    const originalRows = [...rows], grid = [...object.grid], heights = rows.map(id => {
        const style = html.peekNode(id)!.attributes.style ?? "", match = /(?:^|;)\s*height\s*:\s*([\d.]+)pt\s*(?:;|$)/i.exec(style);
        return match ? Number(match[1]) : geometry.height / rows.length;
    }), cells: Cell[] = [], occupied: Cell[][] = Array.from({ length: rows.length }, () => []);
    checkLimit(rows.length * grid.length, ctx.limits.maxNodes);
    for (const [r, id] of rows.entries())
        for (const child of html.peekNode(id)!.childNodeIds) {
            let c = 0;
            while (occupied[r]![c])
                c++;
            const attrs = html.peekNode(child)!.attributes, cell = { id: child, row: r, column: c, rows: Number(attrs.rowspan ?? 1) || rows.length - r, columns: Number(attrs.colspan ?? 1) };
            if (r + cell.rows > rows.length || c + cell.columns > grid.length) invalid("A merged cell exceeds the table grid.");
            cells.push(cell);
            for (let y = r; y < r + cell.rows; y++)
                for (let x = c; x < c + cell.columns; x++)
                    { if (occupied[y]![x]) invalid("Merged table cells overlap."); occupied[y]![x] = cell; }
        }
    const removed = new Set<string>();
    let nextWidth = geometry.width, nextHeight = geometry.height;
    const area = (row: number, column: number, rowCount = 1, columnCount = 1): Cell[] => {
        if (row + rowCount > rows.length || column + columnCount > grid.length)
            invalid("The table selection exceeds its grid.");
        return [...new Set(occupied.slice(row, row + rowCount).flatMap(line => line.slice(column, column + columnCount)))];
    };
    const newCell = (row: number, column: number): Cell => ({ id: edit.insertSource(rows[row]!, ctx.newCellHtml ?? "<td><p></p></td>").rootNodeIds[0]!, row, column, rows: 1, columns: 1 });
    if (change.kind === "resize") {
        if (change.grid) {
            if (change.grid.length !== grid.length)
                invalid("Column resizing must preserve the table grid.");
            object.grid = [...change.grid];
            nextWidth = change.grid.reduce((a, b) => a + b, 0);
        }
        if (change.rowHeights) {
            if (change.rowHeights.length !== rows.length)
                invalid("Row resizing must preserve the table grid.");
            heights.splice(0, heights.length, ...change.rowHeights);
            nextHeight = heights.reduce((a, b) => a + b, 0);
            for (const [index, id] of rows.entries())
                edit.sourceStyle(id, "height", heights[index] + "pt");
        }
    }
    else if (change.kind === "setCells") {
        const width = change.values[0]!.length;
        if (change.values.some(row => row.length !== width))
            invalid("Pasted table values must be rectangular.");
        area(change.row, change.column, change.values.length, width);
        for (const [r, values] of change.values.entries())
            for (const [c, value] of values.entries()) {
                const cell = occupied[change.row + r]![change.column + c]!;
                if (cell.rows !== 1 || cell.columns !== 1)
                    invalid("Pasting a value matrix requires unmerged target cells.");
                if (html.getNode(cell.id)!.text === value)
                    continue;
                edit.removeSource(cell.id, true);
                edit.insertSource(cell.id, "<p>" + escape(value) + "</p>");
            }
    }
    else if (change.kind === "mergeCells") {
        const selected = area(change.row, change.column, change.rows, change.columns);
        if (selected.some(cell => html.directText(cell.id)?.trim()))
            invalid("Merge cells requires paragraph-structured text; normalize direct cell text before merging.");
        if (selected.some(cell => cell.row < change.row || cell.column < change.column || cell.row + cell.rows > change.row + change.rows || cell.column + cell.columns > change.column + change.columns))
            invalid("Merge selection cannot cut through an existing merged cell.");
        const first = occupied[change.row]![change.column]!;
        for (const cell of selected)
            if (cell !== first) {
                const children = html.peekNode(cell.id)!.childNodeIds;
                if (children.length)
                    edit.moveSource(children, first.id);
                removed.add(cell.id);
            }
        first.rows = change.rows;
        first.columns = change.columns;
    }
    else if (change.kind === "splitCell") {
        const cell = area(change.row, change.column)[0]!;
        if (cell.row !== change.row || cell.column !== change.column)
            invalid("Split must identify the starting cell of a merged area.");
        if(change.rows===undefined&&change.columns===undefined){
          for (let r = cell.row; r < cell.row + cell.rows; r++)
              for (let c = cell.column; c < cell.column + cell.columns; c++)
                  if (r !== cell.row || c !== cell.column)cells.push(newCell(r,c));
          cell.rows=cell.columns=1;
        }else{
          const subdivisions={rows:change.rows??1,columns:change.columns??1};
          checkLimit((rows.length+subdivisions.rows)*(grid.length+subdivisions.columns),ctx.limits.maxNodes);
          const refine=(rowAxis:boolean,count:number):number[]=>{
            const sizes=rowAxis?heights:grid,start=rowAxis?cell.row:cell.column,span=rowAxis?cell.rows:cell.columns;
            const from=sizes.slice(0,start).reduce((a,b)=>a+b,0),extent=sizes.slice(start,start+span).reduce((a,b)=>a+b,0),boundaries=Array.from({length:count+1},(_,n)=>from+extent*n/count);
            for(const boundary of boundaries.slice(1,-1)){
              let at=0,position=0;while(at<sizes.length&&position+sizes[at]!<boundary-1e-6){position+=sizes[at]!;at++;}
              if(at===sizes.length||Math.abs(position-boundary)<1e-6||Math.abs(position+sizes[at]!-boundary)<1e-6)continue;
              const first=boundary-position,last=sizes[at]!-first;sizes.splice(at,1,first,last);
              for(const item of cells){const begin=rowAxis?item.row:item.column,length=rowAxis?item.rows:item.columns;if(begin>at){if(rowAxis)item.row++;else item.column++;}else if(begin+length>at){if(rowAxis)item.rows++;else item.columns++;}}
              if(rowAxis){const current=rows[at]!,reference=rows[at+1],parent=edit.sourceParent(current),id=edit.insertSource(parent,'<tr></tr>',reference&&edit.sourceParent(reference)===parent?reference:undefined).rootNodeIds[0]!;rows.splice(at+1,0,id);edit.sourceStyle(current,'height',first+'pt');edit.sourceStyle(id,'height',last+'pt');}
            }
            const positions=[0];for(const size of sizes)positions.push(positions.at(-1)!+size);return boundaries.map(boundary=>{const index=positions.findIndex(value=>Math.abs(value-boundary)<1e-5);if(index<0)invalid('The refined table boundary is missing.');return index;});
          };
          const columnBounds=refine(false,subdivisions.columns),rowBounds=refine(true,subdivisions.rows);
          for(let r=0;r<subdivisions.rows;r++)for(let c=0;c<subdivisions.columns;c++){
            const target=r===0&&c===0?cell:newCell(rowBounds[r]!,columnBounds[c]!);target.row=rowBounds[r]!;target.column=columnBounds[c]!;target.rows=rowBounds[r+1]!-target.row;target.columns=columnBounds[c+1]!-target.column;if(target!==cell)cells.push(target);
          }
          object.grid=grid;
        }
    }
    else {
        const rowChange = change.kind.endsWith("Rows"), insert = change.kind.startsWith("insert"), size = rowChange ? rows.length : grid.length;
        if (change.at > size || !insert && change.at + change.count > size || !insert && change.count === size)
            invalid("Row/column operation must retain a nonempty table and stay within its grid.");
        const result = size + (insert ? change.count : -change.count);
        checkLimit(result * (rowChange ? grid.length : rows.length), ctx.limits.maxNodes);
        for (const cell of cells) {
            const start = rowChange ? cell.row : cell.column, span = rowChange ? cell.rows : cell.columns;
            let next = start, length = span;
            if (insert) {
                if (start >= change.at)
                    next += change.count;
                else if (start + span > change.at)
                    length += change.count;
            }
            else {
                const end = change.at + change.count, overlap = Math.max(0, Math.min(start + span, end) - Math.max(start, change.at));
                length -= overlap;
                if (!length) {
                    removed.add(cell.id);
                    continue;
                }
                next = start >= end ? start - change.count : start >= change.at ? change.at : start;
            }
            if (rowChange) {
                cell.row = next;
                cell.rows = length;
            }
            else {
                cell.column = next;
                cell.columns = length;
            }
        }
        if (rowChange) {
            if (insert) {
                const reference = rows[change.at], parent = edit.sourceParent(reference ?? rows.at(-1)!), height = heights[Math.min(change.at, heights.length - 1)]!, newRows = Array.from({ length: change.count }, () => edit.insertSource(parent, '<tr style="height:' + height + 'pt"></tr>', reference).rootNodeIds[0]!);
                rows.splice(change.at, 0, ...newRows);
                heights.splice(change.at, 0, ...newRows.map(() => height));
                nextHeight += height * change.count;
            }
            else {
                nextHeight -= heights.slice(change.at, change.at + change.count).reduce((a, b) => a + b, 0);
                rows.splice(change.at, change.count);
                heights.splice(change.at, change.count);
            }
        }
        else {
            if (insert) {
                const width = grid[Math.min(change.at, grid.length - 1)]!;
                grid.splice(change.at, 0, ...Array.from({ length: change.count }, () => width));
                nextWidth += width * change.count;
            }
            else {
                nextWidth -= grid.slice(change.at, change.at + change.count).reduce((a, b) => a + b, 0);
                grid.splice(change.at, change.count);
            }
            object.grid = grid;
        }
        const coverage: boolean[][] = Array.from({ length: rows.length }, () => []);
        for (const cell of cells)
            if (!removed.has(cell.id))
                for (let r = cell.row; r < cell.row + cell.rows; r++)
                    for (let c = cell.column; c < cell.column + cell.columns; c++)
                        coverage[r]![c] = true;
        for (let r = 0; r < rows.length; r++)
            for (let c = 0; c < grid.length; c++)
                if (!coverage[r]![c])
                    cells.push(newCell(r, c));
    }
    for (const id of removed)
        edit.removeSource(id);
    const arranged = ctx.parse();
    for (const [r, id] of rows.entries()) {
        const ordered = cells.filter(cell => !removed.has(cell.id) && cell.row === r).sort((a, b) => a.column - b.column), current = arranged.peekNode(id)!.childNodeIds;
        if (current.length !== ordered.length || ordered.some((cell, index) => cell.id !== current[index]))
            edit.moveSource(ordered.map(cell => cell.id), id);
        for (const cell of ordered) {
            const attrs = arranged.peekNode(cell.id)!.attributes;
            edit.attributes(cell.id, Object.fromEntries(([["rowspan", cell.rows], ["colspan", cell.columns]] as const).map(([name, value]) => [name, Number(attrs[name!] ?? 1) === value ? attrs[name!] : Number(value) === 1 ? undefined : String(value)])));
        }
    }
    for (const id of originalRows)
        if (!rows.includes(id))
            edit.removeSource(id);
    if (nextWidth !== geometry.width || nextHeight !== geometry.height)
        return { grid: object.grid, width: Math.max(1, nextWidth), height: Math.max(1, nextHeight), cells: cells.filter(cell => !removed.has(cell.id)) };
    return { grid: object.grid, width: geometry.width, height: geometry.height, cells: cells.filter(cell => !removed.has(cell.id)) };
}

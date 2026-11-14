import { DocumentError } from "@doffi/core";
import type { DeepReadonly, DocumentReadLimits, PresentationDataSet } from "@doffi/core";
import type { DocumentZipOptions } from "../document-zip.js";
import { writeOpcZip } from "../document-zip.js";
import { OutputOpc } from "../export/opc.js";
import { attributes, declaration, empty, escapeXml, XmlParts } from "../export/xml.js";
import type { XmlPartBudget } from "../export/xml.js";
import { R, S, xml } from "../ooxml/xml.js";
import type { XmlBudget } from "../ooxml/xml.js";
export interface ExportedConstantWorkbook {
    readonly bytes: Uint8Array;
    readonly categoryFormula?: string;
    readonly series: readonly {
        readonly seriesId: string;
        readonly nameFormula: string;
        readonly valueFormula?: string;
    }[];
}
function column(index: number): string {
    let result = "";
    for (let value = index + 1; value; value = Math.floor((value - 1) / 26))
        result = String.fromCharCode(65 + (value - 1) % 26) + result;
    return result;
}
/** A bounded, first-party constant table, derived from the current data set.
 * The enclosing PPT and every embedded workbook share both output and XML quotas. */
export function exportConstantWorkbook(data: DeepReadonly<PresentationDataSet>, limits: DocumentReadLimits, budget: XmlPartBudget, xmlBudget: XmlBudget, options: DocumentZipOptions): ExportedConstantWorkbook {
    if (data.series.length + 1 > 16384 || data.categories.length + 1 > 1048576)
        throw new DocumentError("UnsupportedFeature", "The chart constants exceed native worksheet dimensions.");
    const opc = new OutputOpc(new XmlParts(limits, budget));
    const sheet = opc.part("xl/worksheets/sheet1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
    sheet.add(declaration + '<worksheet xmlns="' + S + '"><sheetData>');
    const stringCell = (address: string, value: string): void => sheet.add('<c' + attributes({ r: address, t: "inlineStr" }) + '><is><t>' + escapeXml(value) + '</t></is></c>');
    sheet.add('<row r="1">');
    stringCell("A1", "Category");
    for (const [index, series] of data.series.entries())
        stringCell(column(index + 1) + "1", series.name);
    sheet.add('</row>');
    for (const [index, category] of data.categories.entries()) {
        const row = index + 2;
        sheet.add('<row' + attributes({ r: row }) + '>');
        stringCell("A" + row, category);
        for (const [seriesIndex, series] of data.series.entries()) {
            const address = column(seriesIndex + 1) + row, value = series.values[index];
            if (value === null)
                sheet.add(empty("c", { r: address }));
            else if (value === undefined || !Number.isFinite(value))
                throw new DocumentError("InvalidDocument", "A chart constant is missing or not finite.");
            else
                sheet.add('<c' + attributes({ r: address }) + '><v>' + String(value) + '</v></c>');
        }
        sheet.add('</row>');
    }
    sheet.add('</sheetData></worksheet>');
    sheet.finish();
    const relation = opc.relate("xl/workbook.xml", "worksheet", "worksheets/sheet1.xml"), workbook = opc.part("xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
    workbook.add(declaration + '<workbook xmlns="' + S + '" xmlns:r="' + R + '"><sheets>' + empty("sheet", { name: "Data", sheetId: 1, "r:id": relation }) + '</sheets></workbook>');
    workbook.finish();
    opc.relate("", "officeDocument", "xl/workbook.xml");
    opc.finish();
    for (const [path, bytes] of opc.parts.files)
        xml(bytes, "embedded-workbook/" + path, limits, xmlBudget);
    const formula = (col: string, from: number, to: number): string => "'Data'!$" + col + "$" + from + (from === to ? "" : ":$" + col + "$" + to);
    return {
        bytes: writeOpcZip(opc.parts.files, options),
        ...(data.categories.length ? { categoryFormula: formula("A", 2, data.categories.length + 1) } : {}),
        series: data.series.map((series, index) => ({ seriesId: series.seriesId, nameFormula: formula(column(index + 1), 1, 1), ...(data.categories.length ? { valueFormula: formula(column(index + 1), 2, data.categories.length + 1) } : {}) })),
    };
}

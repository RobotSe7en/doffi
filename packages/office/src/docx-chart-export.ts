import { chartRunStyle, writeChartTextStyle } from './docx-chart-text.js';
import { fontName } from './export/values.js';
import { fillXml, lineXml, effectXml } from './export/drawing-paint.js';
import type { Css } from './export/css.js';
import { patchWordChart } from './docx-chart-patch.js';
import { DocumentError } from '@doffi/core';
import type { DeepReadonly, DocumentPackage, DocumentReadLimits, WordChart } from '@doffi/core';
import { A, C, R, attr, child, elements, text, xml } from './ooxml/xml.js';
import type { XmlNode } from './ooxml/xml.js';
import { serializeXml } from './ooxml/serialize.js';
import { attributes, declaration, empty, escapeXml } from './export/xml.js';
import type { XmlPartBudget } from './export/xml.js';
import type { OutputOpc } from './export/opc.js';
import { exportConstantWorkbook } from './export/chart-workbook.js';
import { writeChartGraph } from './chart-native-graph.js';
import { ChartWorkbookPatch } from './chart-workbook-patch.js';
import { chartText } from './docx-chart-import.js';
const cv = (name: string, value: string | number | boolean | undefined): string => value === undefined ? '' : empty('c:' + name, { val: typeof value === 'boolean' ? value ? 1 : 0 : value });
const raw = (node: XmlNode | undefined): string => node ? serializeXml(node) : '';
const encode = (value: string): Uint8Array => new (globalThis as unknown as {
    TextEncoder: new () => {
        encode(text: string): Uint8Array;
    };
}).TextEncoder().encode(value);
const typeName = (kind: string): string => kind === 'column' || kind === 'bar' ? 'barChart' : kind === 'line' ? 'lineChart' : kind === 'area' ? 'areaChart' : kind === 'doughnut' ? 'doughnutChart' : kind === 'scatter' ? 'scatterChart' : 'pieChart';
const cacheValues = (node: XmlNode | undefined, numeric: boolean, maximum: number): readonly (string | number | null)[] => {
    const ref = child(node, numeric ? 'numRef' : 'strRef', C), cache = child(node, numeric ? 'numLit' : 'strLit', C) ?? child(ref, numeric ? 'numCache' : 'strCache', C);
    if (!cache)
        return [];
    const count = Number(attr(child(cache, 'ptCount', C), 'val', '') ?? 0);
    if (!Number.isSafeInteger(count) || count < 0 || count > maximum)
        throw new DocumentError('ResourceLimitExceeded', 'A native chart cache exceeds its bounded vector budget.');
    const values: (string | number | null)[] = Array(count).fill(numeric ? null : '');
    for (const point of elements(cache, 'pt', C)) {
        const value = text(child(point, 'v', C) ?? point);
        const index = Number(attr(point, 'idx', ''));
        if (!Number.isSafeInteger(index) || index < 0 || index >= count)
            throw new DocumentError('InvalidDocument', 'A native chart cache point is invalid.');
        values[index] = numeric ? Number(value) : value;
    }
    return values;
};
function chartTitle(value: string | undefined, original?: XmlNode, style?: DeepReadonly<NonNullable<WordChart['textStyle']>>): string {
    if (!value)
        return '';
    if (original && chartText(original) === value && !style)
        return raw(original);
    if (original && !style) {
        const copy = (node: XmlNode): XmlNode => ({ ...node, attrs: new Map(node.attrs), content: node.content.map(node => typeof node === 'string' ? node : copy(node)) }), title = copy(original), rich = child(child(title, 'tx', C), 'rich', C), tokens: XmlNode[] = [];
        const collect = (node: XmlNode): void => { if (node.uri === A && node.name === 't')
            tokens.push(node);
        else
            for (const child of node.content)
                if (typeof child !== 'string')
                    collect(child); };
        if (rich) {
            collect(rich);
            if (tokens.length) {
                tokens.forEach((node, index) => node.content.splice(0, node.content.length, ...(index === 0 ? [value] : [])));
                return raw(title);
            }
        }
    }
    return '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>' + (style ? '<a:pPr><a:defRPr' + chartRunStyle(style) + '</a:defRPr></a:pPr>' : '') + '<a:r>' + (style ? '<a:rPr' + chartRunStyle(style) + '</a:rPr>' : '') + '<a:t>' + escapeXml(value) + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>';
}
function seriesAppearance(series: DeepReadonly<WordChart['series'][number]>, original: XmlNode | undefined, baselineColor: string | undefined, kind: string, scatterStyle: WordChart['scatterStyle']): string {
    const source = child(original, 'spPr', C), copy = (node: XmlNode): XmlNode => ({ ...node, attrs: new Map(node.attrs), content: node.content.map(node => typeof node === 'string' ? node : copy(node)) }), properties: XmlNode = source ? copy(source) : { uri: C, name: 'spPr', attrs: new Map(), namespaces: { c: C, a: A }, content: [], part: 'chart-style' }, make = (name: string, content: (XmlNode | string)[] = [], attrs: ReadonlyMap<string, string> = new Map()): XmlNode => ({ uri: A, name, attrs, namespaces: { a: A }, content, part: properties.part });
    const fills = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);
    const paint = (node: XmlNode, color: string | undefined, visible = true): void => { node.content.splice(0, node.content.length, ...node.content.filter(node => typeof node === 'string' || node.uri !== A || !fills.has(node.name))); const index = node.content.findIndex(node => typeof node !== 'string' && node.uri === A && node.name === 'xfrm') + 1; if (!visible)
        node.content.splice(index, 0, make('noFill'));
    else if (color)
        node.content.splice(index, 0, make('solidFill', [make('srgbClr', [], new Map([[' val', color.slice(1)]]))])); };
    if (series.color && series.color !== baselineColor)
        paint(properties, series.color);
    const lineKind = kind === 'line' || kind === 'scatter', visible = series.lineVisible ?? (kind === 'scatter' ? (scatterStyle ?? 'marker') !== 'marker' : undefined);
    if (lineKind && (visible !== undefined || series.color && series.color !== baselineColor)) {
        let line = child(properties, 'ln', A);
        if (!line) {
            line = make('ln');
            const index = properties.content.findIndex(node => typeof node !== 'string' && node.uri === A && ['effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst'].includes(node.name));
            properties.content.splice(index < 0 ? properties.content.length : index, 0, line);
        }
        if (visible === false)
            paint(line, undefined, false);
        else if (series.color && series.color !== baselineColor)
            paint(line, series.color);
        else if (visible === true && child(line, 'noFill', A))
            paint(line, series.color);
    }
    return properties.content.length || properties.attrs.size ? raw(properties) : '';
}
export class WordChartWriter {
    private count = 0;
    private readonly xmlBudget = { nodes: 0 };
    private readonly workbookBudget = { bytes: 0, nodes: 0 };
    private readonly constantBookXmlBudget = { nodes: 0 };
    private readonly workLimits: DocumentReadLimits;
    constructor(private readonly opc: OutputOpc, private readonly files: DocumentPackage, private readonly limits: DocumentReadLimits, private readonly budget: XmlPartBudget) { this.workLimits = { ...limits, maxNodes: Math.min(Number.MAX_SAFE_INTEGER, limits.maxNodes * 4) }; }
    write(chart: DeepReadonly<WordChart>, ownerPart: string, css: Css = {}, specified: Css = {}): string {
        const index = ++this.count, part = 'word/charts/chart-' + index + '.xml';
        let template: XmlNode | undefined, book: ChartWorkbookPatch | undefined, graph = chart.native ? JSON.parse(JSON.stringify(chart.native)) as NonNullable<WordChart['native']> : undefined;
        const sourcePart = graph?.parts.find(part => part.sourcePart === graph.chartPart);
        if (sourcePart) {
            const bytes = this.files.get(sourcePart.assetPath);
            if (!bytes)
                throw new DocumentError('AssetUnavailable', 'A native chart template is missing.');
            template = xml(bytes, sourcePart.sourcePart, this.workLimits, this.xmlBudget);
        }
        const originalPlot = child(child(template, 'chart', C), 'plotArea', C), originalSeries = elements(originalPlot, undefined, C).filter(node => node.name.endsWith('Chart')).flatMap(group => elements(group, 'ser', C)), nativeById = new Map((graph?.bindings ?? []).map((binding, index) => [binding.seriesId, { binding, source: originalSeries[binding.sourceIndex ?? index] }]));
        const getBook = (): ChartWorkbookPatch => {
            if (book)
                return book;
            const native = graph?.parts.find(part => part.sourcePart === graph.workbookPart), bytes = native ? this.files.get(native.assetPath) : undefined;
            if (!bytes)
                throw new DocumentError('ExportFailed', 'The original embedded chart workbook is missing.');
            book = new ChartWorkbookPatch(bytes, this.limits, this.workbookBudget);
            return book;
        };
        const formulas = new Map<string, {
            name?: string | undefined;
            category?: string | undefined;
            value?: string | undefined;
            x?: string | undefined;
        }>();
        let bookBytes: Uint8Array | undefined, bookPart: string | undefined;
        if (!graph?.workbookPart) {
            const extra = chart.series.filter(series => series.xValues).map(series => ({ seriesId: series.seriesId + '-x', name: series.name + ' X', values: series.xValues! })), embedded = exportConstantWorkbook({ dataSetId: 'word-chart', categories: chart.categories, series: [...chart.series, ...extra] }, this.limits, this.budget, this.constantBookXmlBudget, {});
            bookBytes = embedded.bytes;
            chart.series.forEach((series, index) => { const ranges = embedded.series[index]!, x = series.xValues ? embedded.series.find(item => item.seriesId === series.seriesId + '-x')?.valueFormula : undefined; formulas.set(series.seriesId, { name: ranges.nameFormula, category: embedded.categoryFormula, value: ranges.valueFormula, ...(x ? { x } : {}) }); });
        }
        else
            for (const series of chart.series) {
                const original = nativeById.get(series.seriesId), binding = original?.binding, native = original?.source, kind = chart.kind === 'combo' ? series.kind : chart.kind, scatter = kind === 'scatter', ranges: {
                    name?: string | undefined;
                    category?: string | undefined;
                    value?: string | undefined;
                    x?: string | undefined;
                } = { name: binding?.nameFormula, category: binding?.categoryFormula, value: binding?.valueFormula, x: binding?.xFormula };
                const changed = (actual: readonly (string | number | null)[], cached: readonly (string | number | null)[]): boolean => actual.length !== cached.length || actual.some((value, index) => value !== cached[index]);
                if (binding) {
                    if (ranges.name && series.name !== chartText(child(native, 'tx', C)))
                        ranges.name = getBook().write(ranges.name, [series.name]);
                    if (ranges.value && changed(series.values, cacheValues((child(native, 'yVal', C) ?? child(native, 'val', C)), true, this.limits.maxNodes)))
                        ranges.value = getBook().write(ranges.value, series.values, cacheValues(child(native, 'yVal', C) ?? child(native, 'val', C), true, this.limits.maxNodes) as (number | null)[]);
                    if (ranges.category && changed(chart.categories, (child(child(native, 'cat', C), 'numRef', C) || child(child(native, 'cat', C), 'numLit', C) ? cacheValues(child(native, 'cat', C), true, this.limits.maxNodes).map(value => String(value ?? '')) : cacheValues(child(native, 'cat', C), false, this.limits.maxNodes))))
                        ranges.category = getBook().write(ranges.category, chart.categories, (child(child(native, 'cat', C), 'numRef', C) || child(child(native, 'cat', C), 'numLit', C) ? cacheValues(child(native, 'cat', C), true, this.limits.maxNodes).map(value => String(value ?? '')) : cacheValues(child(native, 'cat', C), false, this.limits.maxNodes)) as string[]);
                    if (scatter && series.xValues && !ranges.x)
                        ranges.x = getBook().append(series.name + ' X', series.xValues).valueFormula;
                    if (!scatter && !ranges.category)
                        ranges.category = getBook().append('Category', chart.categories).valueFormula;
                    if (ranges.x && series.xValues && changed(series.xValues, cacheValues(child(native, 'xVal', C), true, this.limits.maxNodes)))
                        ranges.x = getBook().write(ranges.x, series.xValues, cacheValues(child(native, 'xVal', C), true, this.limits.maxNodes) as (number | null)[]);
                }
                else {
                    const values = getBook().append(series.name, series.values);
                    ranges.name = values.nameFormula;
                    ranges.value = values.valueFormula;
                    ranges.category = graph.bindings.find(binding => binding.categoryFormula)?.categoryFormula ?? getBook().append('Category', chart.categories).valueFormula;
                    if (series.xValues)
                        ranges.x = getBook().append(series.name + ' X', series.xValues).valueFormula;
                }
                formulas.set(series.seriesId, ranges);
            }
        if (book)
            bookBytes = (book as ChartWorkbookPatch).finish();
        const vector = (values: readonly (string | number | null)[], numeric: boolean, formula?: string): string => { const type = numeric ? 'num' : 'str', cache = (numeric ? '<c:formatCode>' + escapeXml(chart.numberFormat ?? 'General') + '</c:formatCode>' : '') + cv('ptCount', values.length) + values.flatMap((value, index) => value === null ? [] : ['<c:pt' + attributes({ idx: index }) + '><c:v>' + escapeXml(String(value)) + '</c:v></c:pt>']).join(''); return formula ? '<c:' + type + 'Ref><c:f>' + escapeXml(formula) + '</c:f><c:' + type + 'Cache>' + cache + '</c:' + type + 'Cache></c:' + type + 'Ref>' : '<c:' + type + 'Lit>' + cache + '</c:' + type + 'Lit>'; };
        const labelContents = chart.labelContents ?? ['value'], labels = '<c:dLbls>' + writeChartTextStyle(chart.labelStyle) + ['showLegendKey', 'showVal', 'showCatName', 'showSerName', 'showPercent', 'showBubbleSize'].map((name, index) => cv(name, chart.showLabels && (['none', 'value', 'category', 'series', 'percent', 'none'][index] !== undefined && labelContents.includes(['none', 'value', 'category', 'series', 'percent', 'none'][index] as 'value')))).join('') + '</c:dLbls>';
        const groups = new Map<string, {
            kind: string;
            secondary: boolean;
            series: typeof chart.series[number][];
        }>();
        for (const series of chart.series) {
            const kind = chart.kind === 'combo' ? series.kind! : chart.kind, key = kind + ' ' + (series.axis ?? 'primary');
            let group = groups.get(key);
            if (!group) {
                group = { kind, secondary: series.axis === 'secondary', series: [] };
                groups.set(key, group);
            }
            group.series.push(series);
        }
        let ordinal = 0;
        const plots = [...groups.values()].map(group => {
            const name = typeName(group.kind), originalGroup = elements(originalPlot, name, C)[0], circular = group.kind === 'pie' || group.kind === 'doughnut', axis = group.secondary ? [3, 4] : [1, 2];
            let result = '<c:' + name + '>' + (name === 'barChart' ? cv('barDir', group.kind === 'bar' ? 'bar' : 'col') + cv('grouping', chart.grouping === 'standard' || !chart.grouping ? 'clustered' : chart.grouping) : name === 'lineChart' || name === 'areaChart' ? cv('grouping', chart.grouping === 'clustered' || !chart.grouping ? 'standard' : chart.grouping) : name === 'scatterChart' ? cv('scatterStyle', chart.scatterStyle?.startsWith('smooth') ? 'smoothMarker' : 'lineMarker') : '') + cv('varyColors', circular);
            for (const series of group.series) {
                const order = ordinal++, native = nativeById.get(series.seriesId)?.source, ranges = formulas.get(series.seriesId)!, appearance = seriesAppearance(series, native, nativeById.get(series.seriesId)?.binding.color, group.kind, chart.scatterStyle);
                result += '<c:ser>' + cv('idx', order) + cv('order', order) + '<c:tx>' + (ranges.name ? '<c:strRef><c:f>' + escapeXml(ranges.name) + '</c:f><c:strCache>' + cv('ptCount', 1) + '<c:pt idx="0"><c:v>' + escapeXml(series.name) + '</c:v></c:pt></c:strCache></c:strRef>' : '<c:v>' + escapeXml(series.name) + '</c:v>') + '</c:tx>' + appearance;
                if (group.kind === 'line' || group.kind === 'scatter')
                    result += '<c:marker>' + cv('symbol', series.marker ?? (group.kind === 'scatter' && !['line', 'smooth'].includes(chart.scatterStyle ?? 'marker') ? 'circle' : 'none')) + cv('size', series.markerSize ?? 5) + '</c:marker>';
                const pointStyles = new Map(elements(native, 'dPt', C).map(point => [Number(attr(child(point, 'idx', C), 'val', '')), point]));
                for (let index = 0; index < Math.max(series.pointColors?.length ?? 0, ...[...pointStyles.keys()].map(index => index + 1)); index++) {
                    const color = series.pointColors?.[index] ?? undefined, originalColor = nativeById.get(series.seriesId)?.binding.pointColors?.[index] ?? undefined, point = pointStyles.get(index);
                    if (color === originalColor) {
                        if (point)
                            result += raw(point);
                    }
                    else if (color)
                        result += '<c:dPt>' + cv('idx', index) + '<c:spPr><a:solidFill><a:srgbClr val="' + color.slice(1) + '"/></a:solidFill></c:spPr>' + elements(point, undefined, C).filter(node => !['idx', 'spPr'].includes(node.name)).map(raw).join('') + '</c:dPt>';
                }
                for (const childNode of elements(native, undefined, C))
                    if(!circular&&['trendline','errBars'].includes(childNode.name))
                        result += raw(childNode);
                if (group.kind === 'scatter')
                    result += '<c:xVal>' + vector(series.xValues!, true, ranges.x) + '</c:xVal><c:yVal>' + vector(series.values, true, ranges.value) + '</c:yVal>';
                else
                    result += '<c:cat>' + vector(chart.categories, false, ranges.category) + '</c:cat><c:val>' + vector(series.values, true, ranges.value) + '</c:val>';
                if (group.kind === 'line' || group.kind === 'scatter')
                    result += cv('smooth', series.smooth ?? chart.scatterStyle?.startsWith('smooth') ?? false);
                result += raw(child(native, 'extLst', C)) + '</c:ser>';
            }
            result += labels;
            if (name === 'barChart')
                result += cv('gapWidth', chart.gapWidth ?? 150) + cv('overlap', chart.grouping === 'stacked' || chart.grouping === 'percentStacked' ? chart.overlap ?? 100 : chart.overlap ?? 0);
            else if (circular)
                result += cv('firstSliceAng', chart.firstSliceAngle ?? 0) + (group.kind === 'doughnut' ? cv('holeSize', chart.holeSize ?? 50) : '');
            if (!circular)
                result += axis.map(id => cv('axId', id)).join('');
            return result + raw(child(originalGroup, 'extLst', C)) + '</c:' + name + '>';
        }).join('');
        const axis = (id: number, other: number, kind: 'catAx' | 'valAx', position: string, secondary = false, x = false): string => {
            const options = secondary ? (id === 3 ? chart.axes?.secondaryCategory : chart.axes?.secondaryValue) : (id === 1 ? chart.axes?.category : chart.axes?.value);
            const native = elements(originalPlot, kind, C)[secondary ? 1 : 0], range = x ? secondary ? { minimum: chart.secondaryXMinimum, maximum: chart.secondaryXMaximum, majorUnit: chart.secondaryXMajorUnit } : { minimum: chart.xMinimum, maximum: chart.xMaximum, majorUnit: chart.xMajorUnit } : secondary ? { minimum: chart.secondaryMinimum, maximum: chart.secondaryMaximum, majorUnit: chart.secondaryMajorUnit } : chart;
            const scale = '<c:scaling>' + cv('orientation', options?.reverse ? 'maxMin' : 'minMax') + cv('max', range.maximum) + cv('min', range.minimum) + '</c:scaling>', common = cv('axId', id) + scale + cv('delete', options?.visible === undefined ? secondary && id === 3 : !options.visible) + cv('axPos', options?.position ?? position) + (kind === 'valAx' && (options?.majorGridlines ?? (!secondary && !x)) ? (raw(child(native, 'majorGridlines', C)) || '<c:majorGridlines/>') : '') + chartTitle(options?.title, child(native, 'title', C)) + empty('c:numFmt', { formatCode: options?.numberFormat ?? chart.numberFormat ?? 'General', sourceLinked: 0 }) + cv('majorTickMark', 'none') + cv('minorTickMark', 'none') + cv('tickLblPos', options?.labelPosition ?? 'nextTo') + raw(child(native, 'spPr', C)) + (writeChartTextStyle(options?.textStyle) || raw(child(native, 'txPr', C))) + cv('crossAx', other) + (options?.crossesAt !== undefined ? cv('crossesAt', options.crossesAt) : cv('crosses', options?.crosses ?? (secondary ? 'max' : 'autoZero')));
            return '<c:' + kind + '>' + common + (kind === 'catAx' ? cv('auto', true) + cv('lblAlgn', 'ctr') + cv('lblOffset', 100) : cv('crossBetween', options?.crossBetween ?? 'between') + cv('majorUnit', range.majorUnit)) + raw(child(native, 'extLst', C)) + '</c:' + kind + '>';
        };
        const circular = chart.kind === 'pie' || chart.kind === 'doughnut', horizontal = chart.kind === 'bar' || chart.kind === 'combo' && chart.series.every(series => series.kind === 'bar'), scatter = chart.kind === 'scatter' || chart.kind === 'combo' && chart.series.some(series => series.axis !== 'secondary' && series.kind === 'scatter'), secondaryScatter = chart.kind === 'scatter' || chart.kind === 'combo' && chart.series.some(series => series.axis === 'secondary' && series.kind === 'scatter'), axes = circular ? '' : axis(1, 2, scatter ? 'valAx' : 'catAx', horizontal ? 'l' : 'b', false, scatter) + axis(2, 1, 'valAx', horizontal ? 'b' : 'l') + (chart.series.some(series => series.axis === 'secondary') ? axis(3, 4, secondaryScatter ? 'valAx' : 'catAx', horizontal ? 'l' : 'b', true, secondaryScatter) + axis(4, 3, 'valAx', horizontal ? 't' : 'r', true) : '');
        const originalChart = child(template, 'chart', C), originalTitle = child(originalChart, 'title', C), title = chartTitle(chart.title, originalTitle, chart.titleStyle ?? (!template ? { fontSize: Math.min(14, (chart.textStyle?.fontSize ?? 11) * 1.35) } : undefined)), legend = chart.legend === 'none' ? '' : '<c:legend>' + cv('legendPos', ({ left: 'l', right: 'r', top: 't', bottom: 'b' } as const)[chart.legend]) + '<c:layout/>' + cv('overlay', false) + raw(child(child(originalChart, 'legend', C), 'spPr', C)) + (writeChartTextStyle(chart.legendStyle) || raw(child(child(originalChart, 'legend', C), 'txPr', C))) + '</c:legend>';
        let external = raw(child(template, 'externalData', C));
        if (bookBytes) {
            if (graph) {
                bookPart = graph.workbookPart;
                if (!bookPart) {
                    bookPart = 'doffi/chart-workbook.xlsx';
                    graph.workbookPart = bookPart;
                    graph.parts.push({ sourcePart: bookPart, assetPath: graph.parts[0]!.assetPath, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                    const id = 'doffiWorkbook';
                    graph.relationships.push({ sourcePart: graph.chartPart, id, type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package', target: bookPart, external: false });
                    external = '<c:externalData r:id="' + id + '">' + cv('autoUpdate', false) + '</c:externalData>';
                }
            }
            else {
                const target = 'word/embeddings/chart-' + index + '.xlsx';
                this.opc.binary(target, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', bookBytes);
                external = '<c:externalData r:id="' + this.opc.relate(part, 'package', '/' + target) + '">' + cv('autoUpdate', false) + '</c:externalData>';
            }
        }
        const extras = (node: XmlNode | undefined, names: readonly string[]): string => (node?.content ?? []).filter((node): node is XmlNode => typeof node !== 'string' && !(node.uri === C && names.includes(node.name))).map(raw).join('');
        const inheritedFonts = Object.fromEntries(['ascii', 'hAnsi', 'eastAsia', 'cs'].flatMap(role => { const value = css['--word-font-' + role]; return value ? [[role, fontName(value)]] : []; })), defaultTextStyle = chart.textStyle?{...chart.textStyle}:(!template ? { fontSize: 11, ...(Object.keys(inheritedFonts).length ? { fonts: inheritedFonts } : {}) } : undefined);
        if(!chart.textStyle&&defaultTextStyle?.fonts?.ascii)defaultTextStyle.fonts={...defaultTextStyle.fonts,hAnsi:defaultTextStyle.fonts.ascii};
        if (!template&&defaultTextStyle && !defaultTextStyle.fonts && css['font-family']) {
            const family = fontName(css['font-family']);
            (defaultTextStyle as NonNullable<WordChart['textStyle']>).fonts = { ascii: family, hAnsi: family, eastAsia: family, cs: family };
        }
        const output = declaration + '<c:chartSpace xmlns:c="' + C + '" xmlns:a="' + A + '" xmlns:r="' + R + '">' + extras(template, ['chart', 'externalData', 'spPr', 'txPr', 'printSettings', 'userShapes', 'extLst']) + '<c:chart>' + title + cv('autoTitleDeleted', !chart.title) + ['pivotFmts', 'view3D', 'floor', 'sideWall', 'backWall'].map(name => raw(child(originalChart, name, C))).join('') + '<c:plotArea>' + raw(child(originalPlot, 'layout', C)) + plots + axes + raw(child(originalPlot, 'spPr', C)) + raw(child(originalPlot, 'extLst', C)) + '</c:plotArea>' + legend + cv('plotVisOnly', true) + cv('dispBlanksAs', chart.blankAs ?? 'gap') + cv('showDLblsOverMax', false) + raw(child(originalChart, 'extLst', C)) + '</c:chart>' + raw(child(template, 'spPr', C)) + (writeChartTextStyle(defaultTextStyle) || raw(child(template, 'txPr', C))) + external + raw(child(template, 'printSettings', C)) + raw(child(template, 'userShapes', C)) + raw(child(template, 'extLst', C)) + '</c:chartSpace>';
        const compiled = xml(encode(output), part, this.workLimits, this.xmlBudget), root = template ? patchWordChart(template, { ...compiled, attrs: new Map([...template.attrs, ...compiled.attrs]), namespaces: { ...template.namespaces, ...compiled.namespaces } }, chart) : compiled;
        const fill = css['--office-fill-kind'] !== undefined || css['--office-fill-color'] !== undefined, line = css['--office-line-kind'] !== undefined || css['--office-line-color'] !== undefined, shadow = css['--office-shadow-kind'] !== undefined;
        if (fill || line || shadow) {
            const paint = xml(encode('<c:spPr xmlns:c="' + C + '" xmlns:a="' + A + '">' + (fill ? fillXml(css, '--office-fill', specified) : '') + (line ? lineXml(css, '--office-line', specified) : '') + (shadow ? effectXml(css, specified) : '') + '</c:spPr>'), part, this.workLimits, this.xmlBudget);
            let properties = child(root, 'spPr', C);
            if (!properties) {
                properties = { name: 'spPr', uri: C, attrs: new Map(), namespaces: { c: C, a: A }, content: [], part };
                const index = root.content.findIndex(node => typeof node !== 'string' && node.uri === C && ['txPr', 'externalData', 'printSettings', 'userShapes', 'extLst'].includes(node.name));
                root.content.splice(index < 0 ? root.content.length : index, 0, properties);
            }
            const names = [...(fill ? ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'] : []), ...(line ? ['ln'] : []), ...(shadow ? ['effectLst', 'effectDag'] : [])], order = ['xfrm', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'ln', 'effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst'];
            properties.content.splice(0, properties.content.length, ...properties.content.filter(node => typeof node === 'string' || node.uri !== A || !names.includes(node.name)));
            for (const node of elements(paint, undefined, A)) {
                const index = properties.content.findIndex(other => typeof other !== 'string' && other.uri === A && order.indexOf(other.name) > order.indexOf(node.name));
                properties.content.splice(index < 0 ? properties.content.length : index, 0, node);
            }
        }
        const bytes = encode(declaration + serializeXml(root));
        if (graph) {
            const overrides = new Map<string, Uint8Array>([[graph.chartPart, bytes]]);
            if (bookBytes && bookPart)
                overrides.set(bookPart, bookBytes);
            writeChartGraph(this.opc, graph, this.files, part, overrides, this.workLimits, this.xmlBudget);
        }
        else
            this.opc.binary(part, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml', bytes);
        return this.opc.relate(ownerPart, 'chart', '/' + part);
    }
}

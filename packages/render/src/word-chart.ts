import { createWorkbookNumberFormatter } from '@doffi/core';
import type { DeepReadonly, WordChart } from '@doffi/core';
import { chartValueScale, chartValueFraction, chartAxisTicks, chartAxisPlacement, chartCategoryAxis } from './chart-axis.js';
import { svgElement } from './presentation/svg.js';
import { appendChartMarker } from './chart-marker.js';
import { markViewNode } from './view-nodes.js';
import { cssStringValue } from './presentation/appearance.js';
import { officeFontSegments } from './font-scripts.js';
/** Word adapters use the same scalar, tick and marker primitives as PPT and
 * workbook charts; only the extended Word plot families are composed here. */
export function renderWordChart(node: HTMLElement, chart: DeepReadonly<WordChart>, width: number, height: number, id: string, account: (amount: number) => void, locale?: string): SVGSVGElement {
    const document = node.ownerDocument, computed = document.defaultView!.getComputedStyle(node), palette = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'].map((color, index) => computed.getPropertyValue('--office-theme-accent' + (index + 1)).trim() || color);
    const make: typeof svgElement = (document, tag, attributes) => { account(1); return svgElement(document, tag, attributes); }, svg = markViewNode(make(document, 'svg', { viewBox: `0 0 ${width} ${height}`, width: '100%', height: '100%', role: 'img', 'aria-label': chart.title ?? chart.kind, 'data-sdk-word-chart': chart.kind }));
    svg.style.cssText = 'position:absolute;inset:0;overflow:visible;pointer-events:none';
    const formatter = createWorkbookNumberFormatter({ dateSystem: '1900', ...(locale ? { documentLocale: locale } : {}) });
    const label = (series: DeepReadonly<WordChart['series'][number]>, index: number, value: number): string => {
        const kinds = chart.labelContents ?? ['value'];
        let percent = '';
        if (kinds.includes('percent')) {
            const largest = series.values.reduce<number>((largest, value) => Math.max(largest, Math.abs(value ?? 0)), 0), sum = largest ? series.values.reduce<number>((sum, value) => sum + Math.abs(value ?? 0) / largest, 0) : 0;
            percent = (sum ? Math.abs(value) / largest / sum * 100 : 0).toFixed(0) + '%';
        }
        return kinds.map(kind => kind === 'category' ? chart.categories[index] ?? '' : kind === 'series' ? series.name : kind === 'percent' ? percent : format(value)).join(' ');
    };
    const format = (value: number, code = chart.numberFormat ?? 'General'): string => { const display = formatter.format(value, { code }); return display.kind === 'parts' ? display.parts.map(part => part.kind === 'text' ? part.text : ' ').join('') : '####'; };
    const font = chart.textStyle?.fontSize ?? 11, color = (index: number): string => chart.series[index]?.color ?? palette[index % palette.length]!;
    const inheritedFonts = Object.fromEntries(['ascii', 'hAnsi', 'eastAsia', 'cs'].flatMap(role => { const value = computed.getPropertyValue('--word-font-' + role).trim(); return value ? [[role, cssStringValue(value)]] : []; }));
    if(inheritedFonts['ascii'])inheritedFonts['hAnsi']=inheritedFonts['ascii'];
    const text = (parent: SVGElement, value: string, x: number, y: number, anchor = 'middle', size = font, override?: DeepReadonly<NonNullable<WordChart['textStyle']>>): void => {
        const style = { ...chart.textStyle, ...override, fonts: { ...inheritedFonts, ...chart.textStyle?.fonts, ...override?.fonts } }, label = make(document, 'text', { x, y, 'text-anchor': anchor, 'font-size': override?.fontSize ?? size, fill: style.color && style.color !== 'auto' ? style.color : '#333333', 'font-family': computed.fontFamily, ...(style.bold !== undefined ? { 'font-weight': style.bold ? 'bold' : 'normal' } : {}), ...(style.italic !== undefined ? { 'font-style': style.italic ? 'italic' : 'normal' } : {}) });
        for (const part of officeFontSegments(value)) {
            const role = part.role === 'east-asia' ? 'eastAsia' : part.role === 'h-ansi' ? 'hAnsi' : part.role === 'complex-script' ? 'cs' : 'ascii', span = make(document, 'tspan', { 'data-word-script': role }), family = style.fonts[role] ?? (role === 'hAnsi' ? style.fonts.ascii : role === 'ascii' ? style.fonts.hAnsi : undefined);
            if (family)
                span.style.fontFamily = JSON.stringify(family);
            span.textContent = part.text;
            label.append(span);
        }
        parent.append(label);
    };
    const legendWidth = chart.legend === 'left' || chart.legend === 'right' ? Math.min(100, width * .27) : 0, legendHeight = chart.legend === 'top' || chart.legend === 'bottom' ? 22 : 0, left = 42 + (chart.legend === 'left' ? legendWidth : 0), right = chart.series.some(series => series.axis === 'secondary') ? 42 : 15, top = (chart.title ? 30 : 12) + (chart.legend === 'top' ? legendHeight : 0), bottom = 32 + (chart.legend === 'bottom' ? legendHeight : 0), plotWidth = Math.max(1, width - left - right - (chart.legend === 'right' ? legendWidth : 0)), plotHeight = Math.max(1, height - top - bottom), horizontal = chart.kind === 'bar' || chart.kind === 'combo' && chart.series.every(series => series.kind === 'bar');
    const titleY = 18;
    if (chart.title)
        text(svg, chart.title, width / 2, titleY, 'middle', Math.min(14, font * 1.35), chart.titleStyle);
    const circle = chart.kind === 'pie' || chart.kind === 'doughnut';
    if (chart.legend !== 'none') {
        const entries = circle && chart.kind === 'pie' ? chart.categories.map((name, index) => ({ name, color: chart.series[0]?.pointColors?.[index] ?? palette[index % palette.length]! })) : chart.series.map((series, index) => ({ name: series.name, color: color(index) }));
        const vertical = chart.legend === 'left' || chart.legend === 'right', x = chart.legend === 'left' ? 4 : chart.legend === 'right' ? left + plotWidth + right : Math.max(4, (width - Math.min(width, entries.length * 90)) / 2), y = chart.legend === 'top' ? (chart.title ? 35 : 15) : chart.legend === 'bottom' ? height - 10 : top + font;
        entries.forEach((entry, index) => { const px = x + (vertical ? 0 : index * 90), py = y + (vertical ? index * (font + 7) : 0); svg.append(make(document, 'rect', { x: px, y: py - font + 1, width: font, height: font, fill: entry.color })); text(svg, entry.name, px + font + 4, py, 'start', font, chart.legendStyle); });
    }
    const clip = make(document, 'clipPath', { id }), defs = make(document, 'defs', {});
    clip.append(make(document, 'rect', { x: left, y: top, width: plotWidth, height: plotHeight }));
    defs.append(clip);
    svg.append(defs);
    const plots = make(document, 'g', { 'clip-path': 'url(#' + id + ')' });
    svg.append(plots);
    if (circle) {
        const cx = left + plotWidth / 2, cy = top + plotHeight / 2, r = Math.max(1, Math.min(plotWidth, plotHeight) / 2 - 8), series = chart.kind === 'pie' ? chart.series.slice(0, 1) : chart.series, hole = chart.kind === 'doughnut' ? (chart.holeSize ?? 50) / 100 : 0, ring = (r - r * hole) / Math.max(1, series.length);
        series.forEach((series, seriesIndex) => {
            const magnitude = series.values.map(value => Math.abs(value ?? 0)), largest = magnitude.reduce((largest, value) => Math.max(largest, value), 0), scaled = magnitude.map(value => largest ? value / largest : 0), sum = scaled.reduce((a, b) => a + b, 0), outer = r - ring * seriesIndex, inner = chart.kind === 'pie' ? 0 : outer - ring;
            let angle = (chart.firstSliceAngle ?? 0) * Math.PI / 180 - Math.PI / 2;
            scaled.forEach((value, index) => {
                if (!value || !sum)
                    return;
                const sweep = value / sum * Math.PI * 2, end = angle + sweep, start = { x: cx + outer * Math.cos(angle), y: cy + outer * Math.sin(angle) }, finish = { x: cx + outer * Math.cos(end), y: cy + outer * Math.sin(end) }, large = sweep > Math.PI ? 1 : 0;
                let d: string;
                if (sweep >= Math.PI * 2 - 1e-9)
                    d = `M${cx} ${cy - outer}A${outer} ${outer} 0 1 1 ${cx} ${cy + outer}A${outer} ${outer} 0 1 1 ${cx} ${cy - outer}` + (inner ? `M${cx} ${cy - inner}A${inner} ${inner} 0 1 0 ${cx} ${cy + inner}A${inner} ${inner} 0 1 0 ${cx} ${cy - inner}` : 'Z');
                else
                    d = `M${start.x} ${start.y}A${outer} ${outer} 0 ${large} 1 ${finish.x} ${finish.y}` + (inner ? `L${cx + inner * Math.cos(end)} ${cy + inner * Math.sin(end)}A${inner} ${inner} 0 ${large} 0 ${cx + inner * Math.cos(angle)} ${cy + inner * Math.sin(angle)}Z` : `L${cx} ${cy}Z`);
                plots.append(make(document, 'path', { d, fill: series.pointColors?.[index] ?? palette[index % palette.length]!, stroke: '#ffffff', 'stroke-width': .5, 'fill-rule': 'evenodd', 'data-series': series.seriesId, 'data-point': index }));
                if (chart.showLabels)
                    text(svg, label(series, index, series.values[index] ?? 0), cx + (outer + inner) / 2 * Math.cos(angle + sweep / 2), cy + (outer + inner) / 2 * Math.sin(angle + sweep / 2) + font / 3, 'middle', font, chart.labelStyle);
                angle = end;
            });
        });
        formatter.dispose();
        return svg;
    }
    const count = chart.categories.length, stacked = chart.grouping === 'stacked' || chart.grouping === 'percentStacked', percent = chart.grouping === 'percentStacked', kind = (series: DeepReadonly<WordChart['series'][number]>): string => chart.kind === 'combo' ? series.kind! : chart.kind;
    const stackKey = (series: DeepReadonly<WordChart['series'][number]>, value: number): string => (series.axis ?? 'primary') + ' ' + kind(series) + ' ' + (value < 0 ? '-' : '+'), accum = new Map<string, number[]>();
    const values = chart.series.map(series => series.values.map((raw, index) => {
        if (raw === null)
            return chart.blankAs === 'zero' ? 0 : null;
        if (!percent)
            return raw;
        const total = chart.series.filter(other => (other.axis ?? 'primary') === (series.axis ?? 'primary') && kind(other) === kind(series)).reduce((sum, other) => sum + Math.abs(other.values[index] ?? 0), 0);
        return total ? raw / total : 0;
    }));
    const ranges = values.map((vector, ordinal) => vector.map((value, index) => {
        if (value === null)
            return null;
        const series = chart.series[ordinal]!, key = stackKey(series, value), previous = accum.get(key) ?? Array(count).fill(0) as number[], base = stacked && kind(series) !== 'scatter' ? previous[index]! : 0, end = base + value;
        if (stacked) {
            previous[index] = end;
            accum.set(key, previous);
        }
        return { base, end };
    }));
    const scale = (secondary = false) => chartValueScale({ kind: 'column', legend: 'none', showLabels: false, ...(secondary ? { minimum: chart.secondaryMinimum, maximum: chart.secondaryMaximum, majorUnit: chart.secondaryMajorUnit } : { minimum: chart.minimum, maximum: chart.maximum, majorUnit: chart.majorUnit }) }, chart.series.flatMap((series, index) => (series.axis === 'secondary') === secondary ? [{ values: ranges[index]!.flatMap(value => value ? [value.base, value.end] : []) }] : []));
    const primary = scale(), secondary = scale(true), xScale = chartValueScale({ kind: 'column', legend: 'none', showLabels: false, minimum: chart.xMinimum, maximum: chart.xMaximum, majorUnit: chart.xMajorUnit }, chart.series.filter(series => kind(series) === 'scatter' && series.axis !== 'secondary').map(series => ({ values: series.xValues! }))), secondaryXScale = chartValueScale({ kind: 'column', legend: 'none', showLabels: false, minimum: chart.secondaryXMinimum, maximum: chart.secondaryXMaximum, majorUnit: chart.secondaryXMajorUnit }, chart.series.filter(series => kind(series) === 'scatter' && series.axis === 'secondary').map(series => ({ values: series.xValues! })));
    const axisOptions = (secondaryAxis = false) => ({ category: secondaryAxis ? chart.axes?.secondaryCategory : chart.axes?.category, value: secondaryAxis ? chart.axes?.secondaryValue : chart.axes?.value });
    const fraction = (value: number, secondaryAxis = false): number => { const range = secondaryAxis ? secondary : primary; return chartValueFraction(value, range.minimum, range.maximum, axisOptions(secondaryAxis).value?.reverse ?? false); };
    const xFraction = (value: number, secondaryAxis = false): number => { const range = secondaryAxis ? secondaryXScale : xScale; return chartValueFraction(value, range.minimum, range.maximum, axisOptions(secondaryAxis).category?.reverse ?? false); };
    const categories = (secondaryAxis = false) => { const options = axisOptions(secondaryAxis); return chartCategoryAxis(Math.max(1, count), 0, 1, options.value?.crossBetween !== 'midCat', options.category?.reverse ?? false); };
    const coord = (index: number, value: number, secondaryAxis = false, scatterX?: number): {
        x: number;
        y: number;
    } => horizontal ? { x: left + fraction(value, secondaryAxis) * plotWidth, y: top + (1 - categories(secondaryAxis).position(index)) * plotHeight } : { x: left + (scatterX === undefined ? categories(secondaryAxis).position(index) : xFraction(scatterX, secondaryAxis)) * plotWidth, y: top + (1 - fraction(value, secondaryAxis)) * plotHeight };
    const numericCategories = (secondaryAxis = false): boolean => chart.kind === 'scatter' || chart.kind === 'combo' && chart.series.some(series => (series.axis === 'secondary') === secondaryAxis && series.kind === 'scatter');
    const axisStyle = (options: ReturnType<typeof axisOptions>['category'], secondaryAxis = false): CSSStyleDeclaration => { const style = document.createElement('span').style; for (const [name, value] of Object.entries({ 'crosses': options?.crosses ?? (secondaryAxis ? 'max' : 'autoZero'), 'crosses-at': options?.crossesAt, 'label-position': options?.labelPosition, 'cross-between': options?.crossBetween, 'position': options?.position }))
        if (value !== undefined)
            style.setProperty('--office-axis-' + name, String(value)); return style; };
    const placement = (secondaryAxis = false) => {
        const options = axisOptions(secondaryAxis), range = secondaryAxis ? secondary : primary;
        if (!numericCategories(secondaryAxis))
            return chartAxisPlacement({ kind: 'column', legend: 'none', showLabels: false, reverseCategories: options.category?.reverse ?? false, reverseValues: options.value?.reverse ?? false }, axisStyle(options.category, secondaryAxis), axisStyle(options.value, secondaryAxis), range.minimum, range.maximum, Math.max(1, count));
        const xRange = secondaryAxis ? secondaryXScale : xScale, cross = (options: ReturnType<typeof axisOptions>['category'], min: number, max: number): number => Math.min(max, Math.max(min, options?.crossesAt ?? ((options?.crosses ?? (secondaryAxis ? 'max' : 'autoZero')) === 'max' ? max : options?.crosses === 'min' ? min : 0))), x = xFraction(cross(options.value, xRange.minimum, xRange.maximum), secondaryAxis), y = 1 - fraction(cross(options.category, range.minimum, range.maximum), secondaryAxis), label = (position: string | undefined, at: number, low: number, high: number): number => position === 'low' ? low : position === 'high' ? high : at, categoryLabel = label(options.category?.labelPosition, y, options.value?.reverse ? 0 : 1, options.value?.reverse ? 1 : 0), valueLabel = label(options.value?.labelPosition, x, options.category?.reverse ? 1 : 0, options.category?.reverse ? 0 : 1);
        return { x, y, categoryLabel, valueLabel, categoryTop: categoryLabel < .5, valueRight: valueLabel > .5 };
    };
    const axesLayer = make(document, 'g', {});
    svg.insertBefore(axesLayer, plots);
    const drawAxes = (secondaryAxis = false): void => {
        const options = axisOptions(secondaryAxis), range = secondaryAxis ? secondary : primary, place = placement(secondaryAxis), categoryVisible = options.category?.visible ?? !secondaryAxis, valueVisible = options.value?.visible ?? true, categoryLabels = categoryVisible && options.category?.labelPosition !== 'none', valueLabels = valueVisible && options.value?.labelPosition !== 'none';
        const catLine = horizontal ? left + (1 - place.y) * plotWidth : top + place.y * plotHeight, valueLine = horizontal ? top + (1 - place.x) * plotHeight : left + place.x * plotWidth, catLabel = horizontal ? left + (1 - place.categoryLabel) * plotWidth : top + place.categoryLabel * plotHeight, valueLabel = horizontal ? top + (1 - place.valueLabel) * plotHeight : left + place.valueLabel * plotWidth;
        if (categoryVisible)
            axesLayer.append(make(document, 'line', horizontal ? { x1: catLine, x2: catLine, y1: top, y2: top + plotHeight, stroke: '#888888', 'stroke-width': .6 } : { x1: left, x2: left + plotWidth, y1: catLine, y2: catLine, stroke: '#888888', 'stroke-width': .6 }));
        if (valueVisible)
            axesLayer.append(make(document, 'line', horizontal ? { x1: left, x2: left + plotWidth, y1: valueLine, y2: valueLine, stroke: '#888888', 'stroke-width': .6 } : { x1: valueLine, x2: valueLine, y1: top, y2: top + plotHeight, stroke: '#888888', 'stroke-width': .6 }));
        for (const value of chartAxisTicks(range.minimum, range.maximum, range.unit, 1000)) {
            const p = coord(0, value, secondaryAxis);
            if (options.value?.majorGridlines ?? !secondaryAxis)
                axesLayer.append(make(document, 'line', horizontal ? { x1: p.x, x2: p.x, y1: top, y2: top + plotHeight, stroke: '#d9dfe7', 'stroke-width': .5 } : { x1: left, x2: left + plotWidth, y1: p.y, y2: p.y, stroke: '#d9dfe7', 'stroke-width': .5 }));
            if (valueLabels)
                text(svg, format(value, options.value?.numberFormat ?? chart.numberFormat ?? 'General'), horizontal ? p.x : valueLabel + (place.valueRight ? 6 : -6), horizontal ? valueLabel + (place.valueRight ? -5 : font + 4) : p.y + font / 3, horizontal ? 'middle' : place.valueRight ? 'start' : 'end', font, options.value?.textStyle);
        }
        const axisTitle = (value: string | undefined, x: number, y: number, rotate = false): void => { if (!value)
            return; const group = make(document, 'g', { transform: 'translate(' + x + ' ' + y + ')' + (rotate ? ' rotate(-90)' : '') }); svg.append(group); text(group, value, 0, 0); };
        if (valueVisible)
            axisTitle(options.value?.title, horizontal ? left + plotWidth / 2 : secondaryAxis ? left + plotWidth + 30 : left - 32, horizontal ? top + plotHeight + font * 2.8 : top + plotHeight / 2, !horizontal);
        if (categoryVisible)
            axisTitle(options.category?.title, horizontal ? left - 32 : left + plotWidth / 2, horizontal ? top + plotHeight / 2 : top + plotHeight + font * 2.8, horizontal);
        if (!categoryLabels)
            return;
        if (numericCategories(secondaryAxis)) {
            const range = secondaryAxis ? secondaryXScale : xScale;
            for (const value of chartAxisTicks(range.minimum, range.maximum, range.unit, 1000))
                text(svg, format(value, options.category?.numberFormat ?? chart.numberFormat ?? 'General'), left + xFraction(value, secondaryAxis) * plotWidth, catLabel + (place.categoryTop ? -5 : font + 4), 'middle', font, options.category?.textStyle);
        }
        else {
            const step = Math.max(1, Math.ceil(count / Math.max(1, (horizontal ? plotHeight : plotWidth) / (horizontal ? 18 : 50))));
            chart.categories.forEach((category, index) => { if (index % step)
                return; const p = coord(index, range.minimum, secondaryAxis); text(svg, category, horizontal ? catLabel + (place.categoryTop ? 6 : -6) : p.x, horizontal ? p.y + font / 3 : catLabel + (place.categoryTop ? -5 : font + 4), horizontal ? (place.categoryTop ? 'start' : 'end') : 'middle', font, options.category?.textStyle); });
        }
    };
    drawAxes();
    if (chart.series.some(series => series.axis === 'secondary'))
        drawAxes(true);
    chart.series.forEach((series, ordinal) => {
        const type = kind(series), secondaryAxis = series.axis === 'secondary', fill = color(ordinal), isBar = type === 'column' || type === 'bar', peers = chart.series.filter(other => kind(other) === type && (other.axis ?? 'primary') === (series.axis ?? 'primary')), slot = stacked ? 0 : peers.indexOf(series), slots = stacked ? 1 : peers.length, spacing = (horizontal ? plotHeight : plotWidth) / Math.max(1, count), groupWidth = spacing / (1 + (chart.gapWidth ?? 150) / 100), overlap = (chart.overlap ?? 0) / 100, barWidth = groupWidth / Math.max(1, slots - (slots - 1) * overlap);
        let segments: {
            x: number;
            y: number;
            baseX: number;
            baseY: number;
            index: number;
        }[][] = [[]];
        ranges[ordinal]!.forEach((range, index) => {
            const x = type === 'scatter' ? series.xValues?.[index] : undefined;
            if (!range || type === 'scatter' && (x === undefined || x === null)) {
                if (chart.blankAs !== 'span' && segments.at(-1)!.length)
                    segments.push([]);
                return;
            }
            const p = coord(index, range.end, secondaryAxis, x ?? undefined), base = coord(index, range.base, secondaryAxis, x ?? undefined);
            if (isBar) {
                const offset = -groupWidth / 2 + slot * barWidth * (1 - overlap), rect = horizontal ? { x: Math.min(p.x, base.x), y: p.y + offset, width: Math.abs(p.x - base.x), height: barWidth } : { x: p.x + offset, y: Math.min(p.y, base.y), width: barWidth, height: Math.abs(p.y - base.y) };
                plots.append(make(document, 'rect', { ...rect, fill: series.pointColors?.[index] ?? fill, 'data-series': series.seriesId, 'data-point': index }));
            }
            else
                segments.at(-1)!.push({ ...p, baseX: base.x, baseY: base.y, index });
            if (chart.showLabels)
                text(plots, label(series, index, series.values[index] ?? 0), horizontal ? p.x + (range.end >= range.base ? 4 : -4) : p.x, horizontal ? p.y + font / 3 : p.y - 4, horizontal ? (range.end >= range.base ? 'start' : 'end') : 'middle', font, chart.labelStyle);
        });
        if (isBar)
            return;
        for (const points of segments) {
            if (!points.length)
                continue;
            const smooth = series.smooth ?? (type === 'scatter' && chart.scatterStyle?.startsWith('smooth')), path = points.map((p, index) => {
                if (!index)
                    return 'M' + p.x + ' ' + p.y;
                if (!smooth)
                    return 'L' + p.x + ' ' + p.y;
                const start = points[index - 1]!, previous = points[Math.max(0, index - 2)]!, next = points[Math.min(points.length - 1, index + 1)]!;
                return 'C' + (start.x + (p.x - previous.x) / 6) + ' ' + (start.y + (p.y - previous.y) / 6) + ' ' + (p.x - (next.x - start.x) / 6) + ' ' + (p.y - (next.y - start.y) / 6) + ' ' + p.x + ' ' + p.y;
            }).join(''), area = type === 'area', drawLine = series.lineVisible ?? (type !== 'scatter' || (chart.scatterStyle ?? 'marker') !== 'marker');
            if (area || drawLine) {
                const d = area ? path + [...points].reverse().map(p => 'L' + p.baseX + ' ' + p.baseY).join('') + 'Z' : path;
                plots.append(make(document, 'path', { d, fill: area ? fill : 'none', stroke: fill, 'stroke-width': 1.5, 'data-series': series.seriesId }));
            }
            const marker = series.marker === 'auto' ? 'circle' : series.marker ?? (type === 'scatter' && !['line', 'smooth'].includes(chart.scatterStyle ?? 'marker') ? 'circle' : 'none');
            for (const p of points)
                appendChartMarker(plots, marker, series.markerSize ?? 5, p.x, p.y, series.pointColors?.[p.index] ?? fill, { stroke: fill, 'stroke-width': .75 }, make);
        }
    });
    formatter.dispose();
    return svg;
}

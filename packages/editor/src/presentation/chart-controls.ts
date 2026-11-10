import { presentationCommandIds } from "../commands.js";
import type { EditorCommandRegistry } from "../commands.js";
import type { DocumentEditingSession, DocumentOperation } from "@doffi/core";
type ChartData = Extract<DocumentOperation, {
    op: "setPresentationChartData";
}>["data"];
export interface ChartControlsHost {
    readonly session: DocumentEditingSession;
    readonly commands?: EditorCommandRegistry;
    target(): string | undefined;
    parent(): string;
    perform(operations: DocumentOperation[]): void;
    select(id: string): void;
    attempt(action: () => void): void;
}
export interface ChartControls {
    refresh(): void;
    dispose(): void;
}
export function createChartControls(document: Document, parent: HTMLElement, host: ChartControlsHost): ChartControls {
    const pane = document.createElement("fieldset"), legend = document.createElement("legend"), grid = document.createElement("div"), tools = document.createElement("div"), controls = new AbortController();
    legend.textContent = "图表与常量数据";
    pane.append(legend);
    parent.append(pane);
    let gridControls = new AbortController();
    const read = () => { const p = host.session.snapshot().metadata.presentation!, chart = p.charts.find(c => c.nodeId === host.target()); return { chart, data: chart ? p.dataSets.find(d => d.dataSetId === chart.dataSetId) : undefined }; };
    const button = (label: string, action: () => void, container: HTMLElement = pane, signal = controls.signal): HTMLButtonElement => { const node = document.createElement("button"); node.type = "button"; node.textContent = label; node.addEventListener("click", () => host.attempt(action), { signal }); container.append(node); const id = presentationCommandIds[label]; if (id) host.commands?.register(id, label, action, () => !node.disabled); return node; };
    const choice = (label: string, values: readonly (readonly [
        string,
        string
    ])[]): HTMLSelectElement => { const wrap = document.createElement("label"), node = document.createElement("select"); node.setAttribute("aria-label", label); wrap.append(label, node); for (const [value, label] of values) {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = label;
        node.append(option);
    } tools.append(wrap); return node; };
    const field = (label: string, type = "text"): HTMLInputElement => { const wrap = document.createElement("label"), node = document.createElement("input"); node.type = type; node.setAttribute("aria-label", label); wrap.append(label, node); tools.append(wrap); return node; };
    button("插入图表", () => { const before = new Set(host.session.snapshot().metadata.presentation!.charts.map(c => c.nodeId)); host.perform([{ op: "insertPresentationChart", parentNodeId: host.parent(), geometry: { x: 100, y: 100, width: 480, height: 280, rotation: 0, flipH: false, flipV: false }, kind: "column", title: "图表标题", data: { categories: ["分类 1", "分类 2"], series: [{ name: "系列 1", values: [10, 20] }] } }]); host.select(host.session.snapshot().metadata.presentation!.charts.find(c => !before.has(c.nodeId))!.nodeId); });
    pane.append(tools, grid);
    const kind = choice("图表类型", [["column", "柱形图"], ["line", "折线图"], ["pie", "饼图"]]), legendChoice = choice("图例位置", [["none", "无图例"], ["left", "左侧"], ["right", "右侧"], ["top", "上方"], ["bottom", "下方"]]), title = field("图表标题"), labels = field("显示数据标签", "checkbox"), minimum = field("纵轴最小值", "number"), maximum = field("纵轴最大值", "number"), numberFormat = field("图表数字格式");
    const update = (changes: Extract<DocumentOperation, {
        op: "setPresentationChart";
    }>["changes"], extra: Partial<Extract<DocumentOperation, {
        op: "setPresentationChart";
    }>> = {}): void => { const nodeId = host.target(); if (nodeId)
        host.perform([{ op: "setPresentationChart", nodeId, changes, ...extra }]); };
    kind.addEventListener("change", () => host.attempt(() => update({ kind: kind.value as "column" | "line" | "pie" })), { signal: controls.signal });
    legendChoice.addEventListener("change", () => host.attempt(() => update({ legend: legendChoice.value as "none" | "left" | "right" | "top" | "bottom" })), { signal: controls.signal });
    title.addEventListener("change", () => host.attempt(() => update({}, { title: title.value || null })), { signal: controls.signal });
    labels.addEventListener("change", () => host.attempt(() => update({ showLabels: labels.checked })), { signal: controls.signal });
    for (const [key, node] of [["minimum", minimum], ["maximum", maximum]] as const)
        node.addEventListener("change", () => host.attempt(() => node.value.trim() ? update({ [key]: Number(node.value) }) : update({}, { clear: [key] })), { signal: controls.signal });
    numberFormat.addEventListener("change", () => host.attempt(() => numberFormat.value ? update({ numberFormat: numberFormat.value }) : update({}, { clear: ["numberFormat"] })), { signal: controls.signal });
    const dataEdit = (change: (data: ChartData) => void): void => { const { chart, data } = read(); if (!chart || !data)
        return; const next: ChartData = { categories: [...data.categories], series: data.series.map(s => ({ ...s, values: [...s.values] })) }; change(next); host.perform([{ op: "setPresentationChartData", nodeId: chart.nodeId, data: { categories: next.categories, series: next.series } }]); };
    button("新增分类", () => dataEdit(data => { data.categories.push("新分类"); for (const series of data.series)
        series.values.push(null); }), tools);
    button("新增系列", () => dataEdit(data => { data.series.push({ name: "新系列", values: data.categories.map(() => null) }); }), tools);
    // Adding a series allocates its identity in the core; do not send a guessed ID.
    const series = choice("外观系列", []), color = field("系列新颜色", "color");
    button("设置系列颜色", () => { const nodeId = host.target(); if (nodeId)
        host.perform([{ op: "setPresentationChartStyle", nodeId, target: "series", seriesId: series.value, properties: { "--office-fill-kind": "solid", "--office-fill-color": color.value, "--office-line-kind": "solid", "--office-line-color": color.value } }]); }, tools);
    return { refresh() {
            const { chart, data } = read();
            tools.hidden = grid.hidden = !chart;
            if (!chart || !data)
                return;
            kind.value = chart.kind;
            legendChoice.value = chart.legend;
            title.value = chart.titleNodeId ? host.session.getNode(chart.titleNodeId)!.text : "";
            labels.checked = chart.showLabels;
            minimum.value = chart.minimum === undefined ? "" : String(chart.minimum);
            maximum.value = chart.maximum === undefined ? "" : String(chart.maximum);
            numberFormat.value = chart.numberFormat ?? "";
            const previous = series.value;
            series.replaceChildren(...data.series.map(s => { const option = document.createElement("option"); option.value = s.seriesId; option.textContent = s.name; return option; }));
            if (data.series.some(s => s.seriesId === previous))
                series.value = previous;
            gridControls.abort();
            gridControls = new AbortController();
            grid.replaceChildren();
            const table = document.createElement("table"), head = document.createElement("tr");
            table.setAttribute("aria-label", "图表数据");
            const heading = document.createElement("th");
            heading.textContent = "系列／分类";
            head.append(heading);
            table.append(head);
            grid.append(table);
            const input = (label: string, value: string, change: (value: string) => void, container: HTMLElement, type = "text"): void => { const node = document.createElement("input"); node.type = type; node.value = value; node.setAttribute("aria-label", label); node.addEventListener("change", () => host.attempt(() => change(node.value)), { signal: gridControls.signal }); container.append(node); };
            for (const [c, category] of data.categories.entries()) {
                const cell = document.createElement("th");
                head.append(cell);
                input("分类 " + (c + 1), category, value => dataEdit(data => { data.categories[c] = value; }), cell);
                button("删除分类 " + (c + 1), () => dataEdit(data => { data.categories.splice(c, 1); for (const s of data.series)
                    s.values.splice(c, 1); }), cell, gridControls.signal);
            }
            for (const [r, s] of data.series.entries()) {
                const row = document.createElement("tr"), name = document.createElement("th");
                row.append(name);
                table.append(row);
                input("系列 " + (r + 1), s.name, value => dataEdit(data => { data.series[r]!.name = value; }), name);
                const remove = button("删除系列 " + (r + 1), () => dataEdit(data => { data.series.splice(r, 1); }), name, gridControls.signal);
                remove.disabled = data.series.length === 1;
                for (const [c, value] of s.values.entries()) {
                    const cell = document.createElement("td");
                    row.append(cell);
                    input("系列 " + (r + 1) + " 分类 " + (c + 1), value === null ? "" : String(value), value => dataEdit(data => { const n = value.trim() ? Number(value) : null; if (n !== null && !Number.isFinite(n))
                        throw new Error("请输入有限数值或留空。"); data.series[r]!.values[c] = n; }), cell, "number");
                }
            }
        }, dispose() { controls.abort(); gridControls.abort(); pane.remove(); } };
}

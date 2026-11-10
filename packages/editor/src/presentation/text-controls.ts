import { presentationCommandIds } from "../commands.js";
import type { EditorCommandRegistry } from "../commands.js";
import type { DocumentEditingSession, DocumentOperation } from "@doffi/core";
import type { PresentationInput } from "./input.js";
export interface TextControlsHost {
    readonly session: DocumentEditingSession;
    readonly commands?: EditorCommandRegistry;
    input(): PresentationInput | undefined;
    target(): string | undefined;
    edit(): void;
    notes(): void;
    finish(): void;
    perform(operations: DocumentOperation[]): void;
    attempt(action: () => void): void;
}
export interface TextControls {
    refresh(): void;
    dispose(): void;
}
export function createTextControls(document: Document, parent: HTMLElement, host: TextControlsHost): TextControls {
    const controls = new AbortController(), bar = document.createElement("fieldset"), legend = document.createElement("legend");
    legend.textContent = "文字与备注";
    bar.append(legend);
    parent.append(bar);
    const formatting: HTMLButtonElement[] = [];
    const button = (label: string, action: () => void, format = false): HTMLButtonElement => {
        const node = document.createElement("button");
        node.type = "button";
        node.textContent = label;
        node.addEventListener("click", () => host.attempt(action), { signal: controls.signal });
        bar.append(node);
        const id = presentationCommandIds[label]; if (id) host.commands?.register(id, label, action, () => !node.disabled);
        if (format)
            formatting.push(node);
        return node;
    };
    const field = (label: string, action: (value: string) => void, type = "text"): HTMLInputElement => { const wrap = document.createElement("label"), input = document.createElement("input"); input.type = type; input.setAttribute("aria-label", label); wrap.append(label, input); bar.append(wrap); input.addEventListener("change", () => host.attempt(() => action(input.value)), { signal: controls.signal }); return input; };
    const edit = button("编辑文字", host.edit), notes = button("编辑备注", host.notes), finish = button("完成输入", host.finish);
    for (const [label, property, value] of [["加粗", "font-weight", "700"], ["常规字重", "font-weight", "400"], ["斜体", "font-style", "italic"], ["下划线", "text-decoration", "underline"]] as const)
        button(label, () => { const input = host.input(); if (label === "加粗") input?.toggle("bold"); else if (label === "斜体") input?.toggle("italic"); else if (label === "下划线") input?.toggle("underline"); else input?.format(property, value); }, true);
    const size = field("字号（pt）", value => {
        const n = Number(value);
        if (!Number.isFinite(n) || n <= 0 || n > 4000)
            throw new Error("请输入有效字号。");
        host.input()?.format("font-size", n + "pt");
    }, "number");
    const eastAsia = field("东亚字体", value => host.input()?.format("--office-font-east-asia", JSON.stringify(value)));
    const latin = field("西文字体", value => host.input()?.format("--office-font-ascii", JSON.stringify(value)));
    const color = field("文字颜色", value => host.input()?.format("color", value), "color");
    for (const [label, value] of [["段落左对齐", "left"], ["段落居中", "center"], ["段落右对齐", "right"]] as const)
        button(label, () => host.input()?.paragraph("text-align", value), true);
    for (const [label, value] of [["无列表", "none"], ["项目符号", "bullet"], ["编号列表", "number"]] as const)
        button(label, () => host.input()?.list(value), true);
    const lineHeight = field("行距倍数", value => {
        const n = Number(value);
        if (!Number.isFinite(n) || n <= 0 || n > 20)
            throw new Error("请输入有效行距。");
        host.input()?.paragraph("line-height", String(n));
    }, "number");
    const wrap = document.createElement("label"), fit = document.createElement("select");
    fit.setAttribute("aria-label", "文本自适应");
    for (const [value, label] of [["inherit", "继承布局"], ["none", "固定文本框"], ["normal", "溢出时缩小文字"], ["shape", "随文字扩展文本框"]]) {
        const option = document.createElement("option");
        option.value = value!;
        option.textContent = label!;
        fit.append(option);
    }
    wrap.append("文本自适应", fit);
    bar.append(wrap);
    fit.addEventListener("change", () => host.attempt(() => {
        const nodeId = host.target();
        if (nodeId)
            host.perform([{ op: "setPresentationTextProperties", nodeId, changes: fit.value === "inherit" ? { autoFit: null, fontScale: null, lineSpacingReduction: null } : { autoFit: fit.value as "none" | "normal" | "shape" } }]);
    }), { signal: controls.signal });
    const wrapLabel = document.createElement("label"), wrapping = document.createElement("select");
    wrapping.setAttribute("aria-label", "文本换行");
    for (const [value, label] of [["inherit", "继承换行"], ["true", "自动换行"], ["false", "不换行"]]) {
        const option = document.createElement("option");
        option.value = value!;
        option.textContent = label!;
        wrapping.append(option);
    }
    wrapLabel.append("文本换行", wrapping);
    bar.append(wrapLabel);
    wrapping.addEventListener("change", () => host.attempt(() => { const nodeId = host.target(); if (nodeId)
        host.perform([{ op: "setPresentationTextProperties", nodeId, changes: { wrap: wrapping.value === "inherit" ? null : wrapping.value === "true" } }]); }), { signal: controls.signal });
    return { refresh() {
            const input = host.input(), id = host.target(), object = host.session.snapshot().metadata.presentation!.objects.find(o => o.nodeId === id);
            edit.disabled = !id;
            notes.disabled = !!input?.composing();
            finish.disabled = !input;
            for (const node of [...formatting, size, eastAsia, latin, color, lineHeight])
                node.disabled = !input || input.composing();
            fit.disabled = wrapping.disabled = object?.kind !== "text" && object?.kind !== "geometry";
            wrapping.value = object && (object.kind === "text" || object.kind === "geometry") && object.text?.wrap !== undefined ? String(object.text.wrap) : "inherit";
            fit.value = object && (object.kind === "text" || object.kind === "geometry") ? object.text?.autoFit ?? "inherit" : "inherit";
        }, dispose() { controls.abort(); bar.remove(); } };
}

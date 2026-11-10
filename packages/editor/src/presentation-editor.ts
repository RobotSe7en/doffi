import { captureZoomAnchor, restoreZoomAnchor } from "./viewport-zoom.js";
import { readTextFormat } from "./text-format-state.js";
import { createEditorCommands, presentationCommandIds } from "./commands.js";
import type { EditorCommandState, EditorCommandTarget } from "./commands.js";
import type { DocumentTextCommand } from "./rich-text.js";
import type { TextFormatState } from "./text-format-state.js";
import { createAssetDescriptor, createOpaqueId, DocumentError, PresentationPresetSchema, resolvePresentationGeometry } from "@doffi/core";
import type { Anchor, Rectangle, DocumentEditingSession, DocumentNode, DocumentOperation, DocumentSelection, PresentationGeometry, PresentationObjectInput, TransactionPrepareOptions, TransactionReceipt } from "@doffi/core";
import { mountPreview } from "@doffi/render";
import type { DocumentPreviewOptions, DocumentRenderReport, FontResolution } from "@doffi/render";
import { bindObjectGestures } from "./presentation/gestures.js";
import type { ObjectGestures } from "./presentation/gestures.js";
import { createAppearanceControls } from "./presentation/appearance.js";
import { createConnectorControls } from "./presentation/connectors.js";
import { bindPresentationInput } from "./presentation/input.js";
import type { PresentationInput } from "./presentation/input.js";
import type { SelectionBookmark } from "./rich-text/dom.js";
import { mapSelectionBookmark } from "./rich-text/dom.js";
import { createTextControls } from "./presentation/text-controls.js";
import { createTableControls } from "./presentation/table-controls.js";
import { bindPresentationRegions } from "./presentation/regions.js";
import type { PresentationRegions } from "./presentation/regions.js";
import { createCommentControls } from "./presentation/comment-controls.js";
import { createChartControls } from "./presentation/chart-controls.js";
import { createDefinitionControls } from "./presentation/definition-controls.js";
import { annotationPopoverStyles, workspaceControls } from "./workspace.js";
export interface PresentationEditorOptions extends DocumentPreviewOptions {
    readonly appearance?: "default" | "workspace";
    readonly annotationTrigger?: "selection" | "command";
    readonly onCommandStateChange?: (commands: readonly EditorCommandState[]) => void;
    readonly onFormatStateChange?: (state: TextFormatState) => void;
    readonly onSelectionChange?: (selection: DocumentSelection | undefined) => void;
    readonly onError?: (error: Error) => void;
    readonly onInputRecovery?: (text: string) => void;
}
export interface PresentationEditorView extends EditorCommandTarget {
    executeTextCommand(command: DocumentTextCommand): void;
    formatState(): TextFormatState;
    clipboard(action: "copy" | "cut" | "paste"): Promise<void>;
    readonly element: HTMLElement;
    readonly ready: Promise<void>;
    selection(): DocumentSelection | undefined;
    selectObjects(nodeIds: readonly string[]): void;
    selectRegion(slideId: string, rect: Rectangle): void;
    regionLayoutKey(slideId: string): string;
    setDefinitionSurface(rootNodeId?: string): void;
    editText(nodeId?: string): void;
    editNotes(slideId?: string): void;
    editComment(annotationId: string, replyId?: string): void;
    finishInput(): void;
    activeSlide(): string;
    setActiveSlide(slideId: string): void;
    perform(operations: DocumentOperation[], options?: TransactionPrepareOptions): TransactionReceipt;
    insertImage(bytes: Uint8Array, mimeType: "image/png" | "image/jpeg" | "image/webp", alt?: string): TransactionReceipt;
    layout(): DocumentRenderReport;
    fonts(): readonly FontResolution[];
    setZoom(value: number): void;
    setMode(mode: "text" | "node" | "region"): void;
    focusAnnotation(annotationId: string): void;
    startAnnotation(): void;
    dispose(): void;
}
const labels: Readonly<Record<string, string>> = { text: "文本框", geometry: "形状", image: "图片", connector: "连接线", table: "表格", chart: "图表", group: "组合" };
const shapeLabels: Readonly<Record<string, string>> = { rect: "矩形", roundRect: "圆角矩形", ellipse: "椭圆", triangle: "三角形", diamond: "菱形", parallelogram: "平行四边形", trapezoid: "梯形", pentagon: "五边形", hexagon: "六边形", rightArrow: "右箭头", leftArrow: "左箭头", upArrow: "上箭头", downArrow: "下箭头", leftRightArrow: "双向箭头", chevron: "燕尾箭头", line: "直线" };
/** Host-owned PPT session; source, history and authorization stay in the core. */
export function mountPresentationEditor(container: HTMLElement, session: DocumentEditingSession, options: PresentationEditorOptions = {}): PresentationEditorView {
    if (!session.snapshot().metadata.presentation)
        throw new DocumentError("UnsupportedFormat", "Presentation editing requires a presentation session.");
    const document = container.ownerDocument, element = document.createElement("div"), shadow = element.attachShadow({ mode: "open" }), controller = new AbortController();
    element.tabIndex = 0;
    const style = document.createElement("style");
    style.textContent = ":host{display:block;font:14px system-ui;color:#20252b}*{box-sizing:border-box}button,input,select{font:inherit}button,select{padding:5px 8px}button{cursor:pointer}button:disabled{cursor:default}button[aria-pressed=true]{background:#deebff;border-color:#245aa5}button:focus-visible,input:focus-visible,select:focus-visible{outline:2px solid #245aa5;outline-offset:2px}.inspector{max-height:30vh;overflow:auto}.tools{display:flex;flex-wrap:wrap;gap:6px;margin:8px 0}.body{display:grid;grid-template-columns:180px minmax(0,1fr);gap:12px}.nav{max-height:75vh;overflow:auto}.nav button{display:block;width:100%;text-align:left;margin:4px 0}.stage{overflow:auto;max-height:75vh;border:1px solid #ccd3dc}.plane{position:relative;width:max-content;min-width:100%}fieldset{display:flex;flex-wrap:wrap;gap:8px;border:1px solid #ccd3dc;margin:8px 0;padding:8px}label{display:inline-flex;align-items:center;gap:4px}input[type=number]{width:80px}.error{color:#a00000;white-space:pre-wrap}.status{margin:6px 0}h2{font-size:14px;margin:12px 0 6px}small{color:#555}";
    const workspace = options.appearance === "workspace";
    if (workspace) style.textContent += workspaceControls + annotationPopoverStyles + `
      :host>.tools{display:var(--doffi-presentation-controls,flex)}
      .tools{margin:0;padding:10px 17px;background:#f8faff;border-bottom:1px solid #e0e6f0;gap:5px}.tools button,.tools select{font-size:10px;padding:5px 7px}.tools label{font-size:10px;color:#8193ae}
      .advanced-tools{padding:9px 18px;background:#f9fbfe;border-bottom:1px solid #e0e6f0;font-size:10px;color:#8193ae}.advanced-tools summary{cursor:pointer}.inspector{max-height:230px;overflow:auto;padding-top:7px}
      .status{margin:0;padding:7px 18px;background:#f0f4fb;color:#8193ae;font-size:10px}
      .body{grid-template-columns:135px minmax(0,1fr) var(--doffi-presentation-comments-width,240px);gap:0;background:#eceef1}
      .nav{padding:5px 9px 16px;border-right:1px solid #e0e6f0;background:#f6f8fc;max-height:65vh}.nav h2{font-size:10px;color:#8193ae;font-weight:550;margin:13px 5px 9px}.nav button{font-size:10px;white-space:normal;border:1px solid #e0e6f0;padding:10px 8px;line-height:1.7;margin:6px 0}
      .stage{border:0;min-height:410px;max-height:65vh}.comments-pane{display:var(--doffi-presentation-comments-display,block);max-height:65vh;overflow:auto;padding:14px 10px;background:#f9fbfe;border-left:1px solid #e0e6f0}
      .comments-pane>fieldset{display:block;padding:0;margin:0;border:0}.comments-pane legend{font-size:12px;font-weight:600;margin-bottom:13px;color:#5274a8}.comments-pane section{border:1px solid #dfe6f1;border-radius:7px;background:white;padding:11px;margin-bottom:10px;overflow-wrap:anywhere}
      .comments-pane section p{font-size:11px;line-height:1.8;color:#7185a2;margin-top:0}.comments-pane section button{font-size:9px;padding:4px 6px;margin:2px;border-color:transparent;background:#f0f4fb}.comments-pane input{max-width:100%;margin-top:8px}
      .annotation-popover label{display:flex;flex-wrap:wrap;gap:5px;margin-bottom:9px}.annotation-popover label input{min-width:0;max-width:100%}.annotation-popover select{font-size:10px;margin-bottom:9px}.annotation-popover .composer button{font-size:11px}
      fieldset{border-color:#dfe6f1;border-radius:5px;font-size:11px}small{color:#8193ae}
      .advanced-tools[open]{position:fixed;z-index:40;right:14px;top:208px;width:min(420px,calc(100vw - 28px));max-height:calc(100dvh - 235px);overflow:auto;border:1px solid #d6e1f1;border-radius:6px;background:white;box-shadow:0 8px 28px #243f6920}
      .advanced-tools[open] .inspector{max-height:none}.advanced-tools:not([open]){display:var(--doffi-presentation-controls,block)}
      @media(max-width:1100px){.body{grid-template-columns:105px minmax(0,1fr) var(--doffi-presentation-comments-width,210px)}.tools{padding:8px 12px;gap:4px}.tools button{font-size:9px;padding:5px}.comments-pane{padding:12px 8px}}
      @media(max-width:800px){.body{grid-template-columns:100px minmax(0,1fr)}.comments-pane{grid-column:1/-1;max-height:280px;border-top:1px solid #e0e6f0;border-left:0}.stage{min-height:330px}.tools{max-height:130px;overflow:auto}}
    `;
    const toolbar = document.createElement("div");
    toolbar.className = "tools";
    toolbar.setAttribute("role", "toolbar");
    toolbar.setAttribute("aria-label", "幻灯片与对象工具");
    style.textContent += "[hidden]{display:none!important}";
    const inspector = document.createElement("div"), status = document.createElement("p"), error = document.createElement("p"), body = document.createElement("div"), nav = document.createElement("nav"), slideList = document.createElement("div"), objectList = document.createElement("div"), stage = document.createElement("div"), plane = document.createElement("div"), previewHost = document.createElement("div");
    inspector.className = "inspector";
    status.className = "status";
    status.setAttribute("aria-live", "polite");
    error.className = "error";
    error.setAttribute("role", "alert");
    error.hidden = true;
    body.className = "body";
    nav.className = "nav";
    nav.setAttribute("aria-label", "幻灯片与对象导航");
    stage.className = "stage";
    plane.className = "plane"; plane.style.cssText = "min-width:0;margin-inline:auto";
    const slideHeading = document.createElement("h2");
    slideHeading.textContent = "幻灯片";
    const objectHeading = document.createElement("h2");
    objectHeading.textContent = "当前页对象";
    nav.append(slideHeading, slideList, objectHeading, objectList);
    plane.append(previewHost);
    stage.append(plane);
    body.append(nav, stage);
    const commentsPane = document.createElement("aside"); commentsPane.className = "comments-pane"; commentsPane.setAttribute("aria-label", "幻灯片批注");
    if (workspace) {
        const advanced = document.createElement("details"), summary = document.createElement("summary"); advanced.className = "advanced-tools"; summary.textContent = "对象属性与高级编辑"; advanced.append(summary, inspector);
        body.append(commentsPane); shadow.append(style, toolbar, advanced, status, error, body);
    } else shadow.append(style, toolbar, inspector, status, error, body);
    let disposed = false, busy = true, selected: string[] = [], active = session.snapshot().metadata.presentation!.slides[0]!.slideId, zoom = 1, generation = 0, gestures: ObjectGestures | undefined, frameControls: AbortController | undefined, navControls = new AbortController();
    let input: PresentationInput | undefined, inputScope: {
        kind: "text";
        nodeId: string;
    } | {
        kind: "notes";
        slideId: string;
    } | {
        kind: "comment";
        annotationId: string;
        replyId?: string;
    } | undefined, inputSelection: SelectionBookmark | undefined, draftNotes: HTMLElement | undefined, deferredRefresh = false;
    const commandRegistry = createEditorCommands(() => !disposed, options.onCommandStateChange, () => !busy && !input?.composing());
    const typingStyle = new Map<string, string>();
    let definitionSurface = options.presentationSurfaceId, previewLoaded = false;
    let regionSelection: Extract<Anchor, {
        kind: "region";
    }> | undefined, regionMode = false, regions: PresentationRegions | undefined, focusedAnnotation: string | undefined;
    const regionMessage = document.createElement("p");
    regionMessage.hidden = true;
    regionMessage.setAttribute("role", "status");
    status.after(regionMessage);
    let cacheRevision = -1;
    const nodeCache = new Map<string, Pick<DocumentNode, "nodeId" | "parentNodeId" | "childNodeIds" | "attributes" | "text"> | undefined>();
    const excerpt = (source: string): string => { let text = "", count = 0; for (const character of source) {
        const code = character.codePointAt(0)!, attached = /\p{Mark}/u.test(character) || code >= 0xfe00 && code <= 0xfe0f || code >= 0xe0100 && code <= 0xe01ef;
        if (count >= 20 && !attached || text.length >= 128)
            break;
        text += character;
        if (!attached)
            count++;
    } return text; };
    const readNode = (id: string): ReturnType<typeof nodeCache.get> => {
        const revision = session.snapshot().manifest.revision;
        if (revision !== cacheRevision) {
            nodeCache.clear();
            cacheRevision = revision;
        }
        if (!nodeCache.has(id)) {
            const node = session.getNode(id);
            nodeCache.set(id, node ? { nodeId: node.nodeId, parentNodeId: node.parentNodeId, childNodeIds: node.childNodeIds, attributes: node.attributes, text: excerpt(node.text) } : undefined);
        }
        return nodeCache.get(id);
    };
    const p = () => session.snapshot().metadata.presentation!;
    const orderedSlides = () => { const slides = p().slides, root = readNode(slides[0]!.rootNodeId)!.parentNodeId!, byRoot = new Map(slides.map(slide => [slide.rootNodeId, slide])); return readNode(root)!.childNodeIds.flatMap(id => byRoot.has(id) ? [byRoot.get(id)!] : []); };
    const currentSlide = () => p().slides.find(slide => slide.slideId === active)!;
    const surfaceRoot = (nodeId: string): string | undefined => { for (let id: string | null = nodeId; id !== null; id = readNode(id)?.parentNodeId ?? null)
        if (["slide", "master", "layout", "notes-master"].includes(readNode(id)?.attributes["data-office-kind"] ?? ""))
            return id; return undefined; };
    const currentSurface = (): string => definitionSurface ?? currentSlide().rootNodeId;
    const ensure = (): void => { if (disposed)
        throw new DocumentError("InvalidOperation", "The presentation editor has been disposed."); };
    const report = (reason: unknown): void => { if (disposed)
        return; const value = reason instanceof Error ? reason : new Error("演示文稿操作失败。"); error.textContent = value.message; error.hidden = false; try {
        options.onError?.(value);
    }
    catch { /* Host callback errors do not alter source state. */ } };
    const attempt = (action: () => unknown): void => { if (disposed)
        return; error.hidden = true; try {
        action();
    }
    catch (reason) {
        report(reason);
        renderControls();
    } };
    const button = (label: string, action: () => unknown, parent: HTMLElement = toolbar, signal = controller.signal): HTMLButtonElement => { const node = document.createElement("button"); node.type = "button"; node.textContent = label; node.addEventListener("click", () => attempt(action), { signal }); parent.append(node); const id = presentationCommandIds[label]; if (parent === toolbar && id) commandRegistry.register(id, label, action, () => !busy && !input?.composing() && !node.disabled); return node; };
    const field = (label: string, parent: HTMLElement, type = "number"): HTMLInputElement => { const wrap = document.createElement("label"), input = document.createElement("input"); wrap.append(label, input); input.type = type; input.setAttribute("aria-label", label); if (type === "number")
        input.step = "any"; parent.append(wrap); return input; };
    const select = (label: string, values: readonly {
        value: string;
        label: string;
    }[], parent: HTMLElement): HTMLSelectElement => { const wrap = document.createElement("label"), input = document.createElement("select"); input.setAttribute("aria-label", label); for (const value of values) {
        const option = document.createElement("option");
        option.value = value.value;
        option.textContent = value.label;
        input.append(option);
    } wrap.append(label, input); parent.append(wrap); return input; };
    const ownedSlide = (nodeId: string): string | undefined => { for (let id: string | null = nodeId; id !== null; id = readNode(id)?.parentNodeId ?? null) {
        const slide = p().slides.find(slide => slide.rootNodeId === id);
        if (slide)
            return slide.slideId;
    } return undefined; };
    const selection = (): DocumentSelection | undefined => {
        ensure();
        const { documentId, revision } = session.snapshot().manifest, source = preview.source(), native = preview.element.contentDocument?.getSelection();
        const text = inputScope?.kind === "text" && !input?.composing() && native?.rangeCount ? source?.textAnchor(native.getRangeAt(0), session) : undefined;
        const anchor = text ?? (regionSelection ? Object.freeze({ ...regionSelection, rect: Object.freeze({ ...regionSelection.rect }), nodeIds: Object.freeze([...regionSelection.nodeIds]) }) : undefined) ?? (selected.length === 1 ? Object.freeze({ kind: "node" as const, nodeId: selected[0]! }) : selected.length ? Object.freeze({ kind: "node-set" as const, nodeIds: Object.freeze([...selected]) }) : undefined);
        return anchor ? Object.freeze({ documentId, baseRevision: revision, anchor }) : undefined;
    };
    const notify = (): void => { try {
        options.onSelectionChange?.(selection());
    }
    catch { /* Hosts own their callback errors. */ } };
    const selectObjects = (ids: readonly string[]): void => {
        ensure();
        if (inputScope)
            finishInput();
        regionSelection = undefined;
        const unique = [...new Set(ids)], objects = new Set(p().objects.map(o => o.nodeId)), slides = new Set(unique.map(surfaceRoot));
        if (unique.some(id => !objects.has(id)) || slides.has(undefined) || slides.size > 1 || definitionSurface && [...slides].some(id => id !== definitionSurface))
            throw new DocumentError("InvalidOperation", "请选择同一张幻灯片上的源对象。");
        const requested = new Set(unique);
        selected = unique.filter(id => { for (let parent = readNode(id)!.parentNodeId; parent !== null; parent = readNode(parent)?.parentNodeId ?? null)
            if (requested.has(parent))
                return false; return true; });
        if (selected.length && !definitionSurface)
            active = ownedSlide(selected[0]!)!;
        gestures?.paint();
        regions?.paint();
        renderControls();
        notify();
        if (workspace && options.annotationTrigger !== "command") commentControls.selectionChanged();
    };
    const perform = (operations: DocumentOperation[], prepareOptions?: TransactionPrepareOptions): TransactionReceipt => {
        ensure();
        const { documentId, revision } = session.snapshot().manifest, candidate = session.prepare({ protocolVersion: 1, documentId, baseRevision: revision, operations }, prepareOptions);
        try {
            return session.commit(candidate);
        }
        catch (error) {
            session.discard(candidate);
            throw error;
        }
    };
    function selectRegion(slideId: string, rect: Rectangle): void {
        ensure();
        if (definitionSurface || !current())
            throw new DocumentError("InvalidOperation", "请在已就绪的幻灯片画布选择区域。");
        const slide = p().slides.find(s => s.slideId === slideId);
        if (!slide)
            throw new DocumentError("InvalidOperation", "幻灯片不存在。");
        finishInput();
        active = slideId;
        selected = [];
        const anchor = preview.source()!.slideRegion(slide.rootNodeId, rect);
        if (anchor.kind !== "region")
            throw new DocumentError("InvalidOperation", "区域选择失败。");
        regionSelection = anchor;
        regions?.paint();
        gestures?.paint();
        regions?.paint();
        renderControls();
        notify();
        if (workspace && options.annotationTrigger !== "command") commentControls.selectionChanged();
    }
    function regionLayoutKey(slideId: string): string { ensure(); if(definitionSurface)throw new DocumentError("InvalidOperation","请先返回幻灯片画布，再核对区域布局。"); const slide = p().slides.find(s => s.slideId === slideId); if (!slide || !current())
        throw new DocumentError("InvalidOperation", "请等待幻灯片画布就绪后检查区域布局。"); return preview.source()!.layoutKey(slide.rootNodeId); }
    const regionTool = button("框选区域", () => { finishInput(); regionMode = !regionMode; gestures?.cancel(); regions?.cancel(); gestures?.paint(); renderControls(); });
    const geometryBatch = (change: (geometry: Readonly<PresentationGeometry>) => PresentationGeometry): DocumentOperation[] => selected.map(nodeId => ({ op: "setPresentationGeometry", nodeId, geometry: change(resolvePresentationGeometry(session, nodeId)) }));
    const defaultGeometry = (): PresentationGeometry => ({ x: 48, y: 72, width: 240, height: 120, rotation: 0, flipH: false, flipV: false });
    const addObject = (object: PresentationObjectInput): TransactionReceipt => {
        const before = new Set(p().objects.map(o => o.nodeId)), receipt = perform([{ op: "insertPresentationObject", parentNodeId: currentSurface(), geometry: defaultGeometry(), object }]);
        selectObjects(p().objects.filter(o => !before.has(o.nodeId) && readNode(o.nodeId)!.parentNodeId === currentSurface()).map(o => o.nodeId));
        return receipt;
    };
    const insertImage = (bytes: Uint8Array, mimeType: "image/png" | "image/jpeg" | "image/webp", alt = "图片"): TransactionReceipt => {
        ensure();
        const path = "assets/" + createOpaqueId() + ({ "image/png": ".png", "image/jpeg": ".jpg", "image/webp": ".webp" }[mimeType]), descriptor = createAssetDescriptor(path, mimeType, bytes), before = new Set(p().objects.map(o => o.nodeId));
        const receipt = perform([{ op: "insertPresentationObject", parentNodeId: currentSurface(), geometry: defaultGeometry(), object: { kind: "image", assetPath: path, alt } }], { assets: [{ descriptor, bytes }] });
        selectObjects(p().objects.filter(o => !before.has(o.nodeId)).map(o => o.nodeId));
        return receipt;
    };
    const navigateHistory = (direction: "undo" | "redo"): void => { const before = input?.selection(), receipt = direction === "undo" ? session.undo() : session.redo(); inputSelection = mapSelectionBookmark(before, receipt.mapping); input?.focus(inputSelection); notify(); };
    const undo = button("撤销", () => navigateHistory("undo")), redo = button("重做", () => navigateHistory("redo"));
    const layoutChoice = select("新页布局", p().layouts.map(layout => ({ value: layout.layoutId, label: layout.name ?? "布局" })), toolbar);
    button("新增幻灯片", () => { const before = new Set(p().slides.map(slide => slide.slideId)); perform([{ op: "insertSlide", layoutId: layoutChoice.value, afterSlideId: active }]); setActiveSlide(p().slides.find(slide => !before.has(slide.slideId))!.slideId); });
    button("复制幻灯片", () => { const before = new Set(p().slides.map(slide => slide.slideId)); perform([{ op: "duplicateSlide", slideId: active }]); setActiveSlide(p().slides.find(slide => !before.has(slide.slideId))!.slideId); });
    const removeSlide = button("删除幻灯片", () => perform([{ op: "removeSlide", slideId: active }]));
    button("幻灯片上移", () => { const slides = orderedSlides(), index = slides.findIndex(slide => slide.slideId === active); if (index > 0)
        perform([{ op: "moveSlide", slideId: active, beforeSlideId: slides[index - 1]!.slideId }]); });
    button("幻灯片下移", () => { const slides = orderedSlides(), index = slides.findIndex(slide => slide.slideId === active); if (index < slides.length - 1)
        perform([{ op: "moveSlide", slideId: active, ...(slides[index + 2] ? { beforeSlideId: slides[index + 2]!.slideId } : {}) }]); });
    button("文本框", () => addObject({ kind: "text", text: "请输入文字" }));
    button("形状", () => addObject({ kind: "geometry", preset: "rect", text: "" }));
    button("表格", () => addObject({ kind: "table", rows: 2, columns: 3 }));
    button("连接线", () => addObject({ kind: "connector", routing: "straight", ...(selected[0] ? { startNodeId: selected[0] } : {}), ...(selected[1] ? { endNodeId: selected[1] } : {}) }));
    const imageInput = document.createElement("input");
    imageInput.type = "file";
    imageInput.accept = "image/png,image/jpeg,image/webp";
    imageInput.hidden = true;
    toolbar.append(imageInput);
    button("插入图片", () => imageInput.click());
    imageInput.addEventListener("change", () => { const file = imageInput.files?.[0]; imageInput.value = ""; if (!file)
        return; void file.arrayBuffer().then(buffer => { if (!disposed)
        attempt(() => insertImage(new Uint8Array(buffer), file.type as "image/png" | "image/jpeg" | "image/webp", file.name)); }).catch(report); }, { signal: controller.signal });
    const duplicate = button("复制对象", () => { const before = new Set(p().objects.map(o => o.nodeId)); perform([{ op: "duplicatePresentationObjects", nodeIds: selected, offset: { x: 12, y: 12 } }]); selectObjects(p().objects.filter(o => !before.has(o.nodeId)).map(o => o.nodeId)); });
    const removeObject = button("删除对象", () => perform(selected.map(nodeId => ({ op: "removeNode", nodeId }))));
    const group = button("组合", () => { const before = new Set(p().objects.map(o => o.nodeId)); perform([{ op: "groupPresentationObjects", nodeIds: selected }]); selectObjects(p().objects.filter(o => !before.has(o.nodeId)).map(o => o.nodeId)); });
    const ungroup = button("取消组合", () => { const children = readNode(selected[0]!)!.childNodeIds; perform([{ op: "ungroupPresentationObject", nodeId: selected[0]! }]); selectObjects(children); });
    button("选择上层组合", () => { const parent = selected.length === 1 ? readNode(selected[0]!)!.parentNodeId : null; if (parent && p().objects.find(o => o.nodeId === parent)?.kind === "group")
        selectObjects([parent]); });
    const objectButtons = [duplicate, removeObject, group, ungroup];
    for (const [label, position] of [["置顶", "front"], ["置底", "back"], ["上移一层", "forward"], ["下移一层", "backward"]] as const)
        objectButtons.push(button(label, () => perform([{ op: "reorderPresentationObjects", nodeIds: selected, position }])));
    const relative = select("对齐范围", [{ value: "selection", label: "所选对象" }, { value: "slide", label: "整张幻灯片" }], toolbar);
    for (const [label, alignment] of [["左对齐", "left"], ["水平居中", "center"], ["右对齐", "right"], ["顶端对齐", "top"], ["垂直居中", "middle"], ["底端对齐", "bottom"]] as const)
        objectButtons.push(button(label, () => perform([{ op: "alignPresentationObjects", nodeIds: selected, alignment, relativeTo: relative.value as "selection" | "slide" }])));
    const horizontal = button("水平分布", () => perform([{ op: "distributePresentationObjects", nodeIds: selected, axis: "horizontal" }])), vertical = button("垂直分布", () => perform([{ op: "distributePresentationObjects", nodeIds: selected, axis: "vertical" }]));
    const zoomChoice = select("缩放", [.5, .75, 1, 1.25, 1.5, 2].map(value => ({ value: String(value), label: Math.round(value * 100) + "%" })), toolbar);
    zoomChoice.value = "1";
    zoomChoice.addEventListener("change", () => attempt(() => setZoom(Number(zoomChoice.value))), { signal: controller.signal });
    const slideFields = document.createElement("fieldset"), geometryFields = document.createElement("fieldset"), cropFields = document.createElement("fieldset");
    inspector.append(slideFields, geometryFields, cropFields);
    const name = field("幻灯片名称", slideFields, "text"), hidden = field("隐藏幻灯片", slideFields, "checkbox");
    name.addEventListener("change", () => attempt(() => perform([{ op: "setSlideProperties", slideId: active, changes: { name: name.value } }])), { signal: controller.signal });
    hidden.addEventListener("change", () => attempt(() => perform([{ op: "setSlideProperties", slideId: active, changes: { hidden: hidden.checked } }])), { signal: controller.signal });
    const numeric = new Map<"x" | "y" | "width" | "height" | "rotation", HTMLInputElement>();
    for (const [key, label] of [["x", "X（pt）"], ["y", "Y（pt）"], ["width", "宽度（pt）"], ["height", "高度（pt）"], ["rotation", "旋转（度）"]] as const) {
        const input = field(label, geometryFields);
        numeric.set(key, input);
        input.addEventListener("change", () => attempt(() => { const value = Number(input.value); if (!input.value.trim() || !Number.isFinite(value))
            throw new DocumentError("InvalidOperation", "请输入有限的几何数值。"); perform(geometryBatch(g => ({ ...g, [key]: value }))); }), { signal: controller.signal });
    }
    objectButtons.push(button("水平翻转", () => perform(geometryBatch(g => ({ ...g, flipH: !g.flipH }))), geometryFields), button("垂直翻转", () => perform(geometryBatch(g => ({ ...g, flipV: !g.flipV }))), geometryFields));
    const preset = select("形状预设", PresentationPresetSchema.options.map(value => ({ value, label: shapeLabels[value] ?? value })), geometryFields);
    preset.addEventListener("change", () => attempt(() => perform([{ op: "setPresentationPreset", nodeId: selected[0]!, preset: preset.value as typeof PresentationPresetSchema.options[number] }])), { signal: controller.signal });
    const crops = new Map<"left" | "top" | "right" | "bottom", HTMLInputElement>();
    for (const [key, label] of [["left", "左裁剪（%）"], ["top", "上裁剪（%）"], ["right", "右裁剪（%）"], ["bottom", "下裁剪（%）"]] as const) {
        const input = field(label, cropFields);
        input.min = "0";
        input.max = "99.9";
        crops.set(key, input);
        input.addEventListener("change", () => attempt(() => perform([{ op: "setPresentationCrop", nodeId: selected[0]!, crop: Object.fromEntries([...crops].map(([key, input]) => [key, Number(input.value) / 100])) as {
                    left: number;
                    top: number;
                    right: number;
                    bottom: number;
                } }])), { signal: controller.signal });
    }
    const appearance = createAppearanceControls(document, inspector, { session, selected: () => selected, element: id => preview.source()?.elements.get(id)?.element, current: () => current(), perform: operations => perform(operations), attempt });
    const connectors = createConnectorControls(document, inspector, { session, node: readNode, selected: () => selected, current: () => current(), perform: operations => perform(operations), attempt });
    const textControls = createTextControls(document, inspector, { session, commands: commandRegistry, input: () => input, target: () => { const id = inputScope?.kind === "text" ? inputScope.nodeId : selected.length === 1 ? selected[0] : undefined; const object = p().objects.find(o => o.nodeId === id); return object && ["text", "geometry", "table"].includes(object.kind) ? id : undefined; }, edit: () => editText(), notes: () => editNotes(), finish: () => finishInput(), perform: operations => { perform(operations); }, attempt });
    if (workspace) {
        for (const label of ["编辑文字", "编辑备注", "完成输入"]) {
            const control = [...inspector.querySelectorAll("button")].find(node => node.textContent === label);
            if (control) toolbar.append(control);
        }
    }
    const tableControls = createTableControls(document, inspector, { session, target: () => selected.length === 1 && p().objects.find(o => o.nodeId === selected[0])?.kind === "table" ? selected[0] : undefined, perform: operations => { perform(operations); }, attempt });
    let preview = mountPreview(previewHost, session, options);
    const definitions = createDefinitionControls(document, inspector, { session, activeSlide: () => active, surface: () => definitionSurface, object: () => selected.length === 1 ? selected[0] : undefined, selectSurface: id => setDefinitionSurface(id), attempt });
    const chartControls = createChartControls(document, inspector, { session, commands: commandRegistry, target: () => selected.length === 1 ? selected[0] : undefined, parent: currentSurface, perform: operations => { perform(operations); }, select: id => selectObjects([id]), attempt });
    const commentControls = createCommentControls(document, workspace ? commentsPane : inspector, { session, selection, slideRoot: () => currentSlide().rootNodeId, perform: operations => { perform(operations); }, edit: editComment, locate: focusAnnotation, attempt }, workspace ? { root: shadow, toolbar, bounds: () => stage.getBoundingClientRect(), rect: chosen => {
        const rect = preview.source()?.rectangles(chosen.anchor).at(-1), frame = preview.element.getBoundingClientRect();
        return rect ? new DOMRect(frame.left + rect.left * zoom, frame.top + rect.top * zoom, rect.width * zoom, rect.height * zoom) : undefined;
    } } : undefined);
    container.append(element);
    const current = (): boolean => !disposed && !busy && preview.layout().revision === session.snapshot().manifest.revision;
    const keyboard = (event: KeyboardEvent): void => {
        if (event.key === "Escape" && inputScope && !input?.composing()) {
            finishInput();
            event.preventDefault();
            return;
        }
        const target = event.target as Element | null;
        if (target?.closest?.('input,textarea,select,[contenteditable="true"]'))
            return;
        const command = event.ctrlKey || event.metaKey;
        if (event.key === "Escape") {
            gestures?.cancel();
            selectObjects([]);
            event.preventDefault();
            return;
        }
        if (command && event.key.toLowerCase() === "z") {
            event.preventDefault();
            attempt(() => navigateHistory(event.shiftKey ? "redo" : "undo"));
            return;
        }
        if (command && event.key.toLowerCase() === "y") {
            event.preventDefault();
            attempt(() => navigateHistory("redo"));
            return;
        }
        if (!current())
            return;
        if (command && event.key.toLowerCase() === "a") {
            event.preventDefault();
            selectObjects(readNode(currentSurface())!.childNodeIds);
            return;
        }
        if (!selected.length)
            return;
        if (event.key === "Delete" || event.key === "Backspace") {
            event.preventDefault();
            attempt(() => removeObject.click());
            return;
        }
        if (command && event.key.toLowerCase() === "d") {
            event.preventDefault();
            attempt(() => duplicate.click());
            return;
        }
        if (command && event.key.toLowerCase() === "g") {
            event.preventDefault();
            attempt(() => event.shiftKey ? ungroup.click() : group.click());
            return;
        }
        const delta: Readonly<Record<string, readonly [
            number,
            number
        ]>> = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }, direction = delta[event.key];
        if (direction) {
            event.preventDefault();
            const step = event.shiftKey ? 10 : 1;
            attempt(() => perform(geometryBatch(g => ({ ...g, x: g.x + direction[0] * step, y: g.y + direction[1] * step }))));
        }
    };
    element.addEventListener("keydown", keyboard, { signal: controller.signal });
    let lastTextFormat = "";
    function notifyTextFormat(): void {
        if (disposed) return; commandRegistry.notify(); const object = selected.length === 1 ? preview.source()?.elements.get(selected[0]!)?.element : undefined;
        const range = object?.ownerDocument.createRange(); if (object && range) range.selectNodeContents(object);
        const state = input?.formatState() ?? (object && range && p().objects.some(item => item.nodeId === selected[0] && ["text", "geometry", "table"].includes(item.kind)) ? { ...readTextFormat(object, range), hasSelection: false, editingText: false } : { enabled: false }), key = JSON.stringify(state);
        if (key === lastTextFormat) return; lastTextFormat = key; try { options.onFormatStateChange?.(state); } catch { /* Host UI only. */ }
    }
    function renderControls(): void {
        if (disposed)
            return;
        const slides = orderedSlides();
        if (!p().slides.some(slide => slide.slideId === active))
            active = slides[0]!.slideId;
        regionTool.disabled = !!definitionSurface;
        regionTool.setAttribute("aria-pressed", String(regionMode));
        textControls.refresh();
        tableControls.refresh();
        definitions.refresh();
        chartControls.refresh();
        commentControls.refresh();
        selected = selected.filter(id => p().objects.some(o => o.nodeId === id) && (definitionSurface ? surfaceRoot(id) === definitionSurface : ownedSlide(id) === active));
        undo.disabled = !session.history().undoSteps;
        redo.disabled = !session.history().redoSteps;
        removeSlide.disabled = slides.length === 1;
        for (const button of objectButtons)
            button.disabled = busy || !selected.length;
        group.disabled = busy || selected.length < 2;
        ungroup.disabled = busy || selected.length !== 1 || p().objects.find(o => o.nodeId === selected[0])?.kind !== "group";
        horizontal.disabled = vertical.disabled = busy || selected.length < 3;
        slideFields.hidden = !!definitionSurface;
        layoutChoice.disabled = !!definitionSurface;
        for (const node of toolbar.querySelectorAll<HTMLButtonElement>("button"))
            if (["新增幻灯片", "复制幻灯片", "删除幻灯片", "幻灯片上移", "幻灯片下移"].includes(node.textContent!))
                node.disabled = !!definitionSurface || node === removeSlide && slides.length === 1;
        name.value = currentSlide().name ?? "";
        hidden.checked = currentSlide().hidden ?? false;
        const object = selected.length === 1 ? p().objects.find(o => o.nodeId === selected[0]) : undefined, g = object ? resolvePresentationGeometry(session, object.nodeId) : undefined;
        for (const [key, input] of numeric) {
            input.disabled = busy || !g;
            input.value = g ? String(Math.round(g[key] * 10000) / 10000) : "";
        }
        preset.disabled = busy || object?.kind !== "geometry";
        if (object?.kind === "geometry")
            preset.value = object.preset;
        cropFields.hidden = object?.kind !== "image";
        for (const [key, input] of crops)
            input.value = object?.kind === "image" ? String(Number(readNode(object.nodeId)!.attributes["data-office-crop-" + key] ?? 0) * 100) : "0";
        appearance.refresh();
        connectors.refresh();
        commandRegistry.notify(); notifyTextFormat();
        status.textContent = `第 ${slides.findIndex(slide => slide.slideId === active) + 1} / ${slides.length} 页 · ${selected.length ? "已选 " + selected.length + " 个对象" : "未选择对象"}${busy ? " · 正在更新画布" : ""}`;
        navControls.abort();
        navControls = new AbortController();
        slideList.replaceChildren();
        objectList.replaceChildren();
        for (const [index, slide] of slides.entries()) {
            const node = button(`${index + 1}. ${slide.name ?? "幻灯片"}${slide.hidden ? "（隐藏）" : ""}`, () => setActiveSlide(slide.slideId), slideList, navControls.signal);
            node.setAttribute("aria-pressed", String(slide.slideId === active));
        }
        for (const object of p().objects.filter(o => definitionSurface ? surfaceRoot(o.nodeId) === definitionSurface : ownedSlide(o.nodeId) === active)) {
            const text = readNode(object.nodeId)!.text, node = button((object.name ?? labels[object.kind]) + (text ? " · " + text : ""), () => selectObjects([object.nodeId]), objectList, navControls.signal);
            node.setAttribute("aria-pressed", String(selected.includes(object.nodeId)));
        }
    }
    function setActiveSlide(slideId: string): void {
        ensure();
        commentControls.clearSelection();
        regionSelection = undefined;
        if (inputScope)
            finishInput();
        if (definitionSurface)
            setDefinitionSurface();
        if (!p().slides.some(slide => slide.slideId === slideId))
            throw new DocumentError("InvalidOperation", "The requested slide is missing.");
        active = slideId;
        selectObjects([]);
        const version = generation;
        void ready.then(() => { if (disposed || active !== slideId || version !== generation)
            return; const node = preview.source()?.elements.get(currentSurface())?.element; if (node)
            stage.scrollTop += preview.element.getBoundingClientRect().top + node.getBoundingClientRect().top * zoom - stage.getBoundingClientRect().top - 12; }).catch(reason => { if (version === generation)
            report(reason); });
    }
    function setZoom(value: number): void {
        ensure();
        if (!Number.isFinite(value) || value < .5 || value > 2)
            throw new DocumentError("InvalidOperation", "缩放范围为 50% 至 200%。");
        const anchor = captureZoomAnchor(stage, plane, zoom); zoom = value;
        if (![...zoomChoice.options].some(option => option.value === String(value))) {
            const option = document.createElement("option"); option.value = String(value); option.textContent = Math.round(value * 100) + "%"; zoomChoice.append(option);
        }
        zoomChoice.value = String(value);
        const frame = preview.element.contentDocument, body = frame?.body, computed = body ? frame!.defaultView!.getComputedStyle(body) : undefined;
        let width = p().size.width * 4 / 3 + (parseFloat(computed?.paddingLeft ?? "0") || 0) + (parseFloat(computed?.paddingRight ?? "0") || 0);
        previewHost.style.width = width + "px";
        if (frame)
            width = Math.max(width, frame.documentElement.scrollWidth);
        previewHost.style.width = width + "px";
        previewHost.style.transform = `scale(${zoom})`;
        previewHost.style.transformOrigin = "0 0";
        plane.style.width = width * zoom + "px";
        plane.style.height = preview.element.offsetHeight * zoom + "px";
        gestures?.paint();
        regions?.paint();
        restoreZoomAnchor(stage, plane, zoom, anchor);
    }
    const disposeInput = (): void => { input?.dispose(); input = undefined; draftNotes?.remove(); draftNotes = undefined; };
    const bindInput = (): void => {
        if (!inputScope || disposed)
            return;
        const source = preview.source()!, frame = source.root.ownerDocument, raw = new DOMParser().parseFromString(session.snapshot().html, "text/html");
        let root: HTMLElement | undefined, canonical: HTMLElement | undefined, editable: HTMLElement | undefined, instance = false, newNotesSlideId: string | undefined;
        const scope = inputScope;
        if (scope.kind === "text") {
            const object = p().objects.find(o => o.nodeId === scope.nodeId);
            if (!object || !["text", "geometry", "table"].includes(object.kind)) {
                inputScope = undefined;
                return;
            }
            root = source.elements.get(object.nodeId)?.element;
            canonical = raw.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(object.nodeId) + '"]') ?? undefined;
            editable = root?.querySelector<HTMLElement>(object.kind === "table" ? "table" : ".sdk-ppt-text") ?? undefined;
        }
        else if (scope.kind === "notes") {
            const story = p().stories.find(s => s.kind === "notes" && s.slideId === scope.slideId);
            if (story && story.kind === "notes") {
                const id = story.bodyNodeId ?? story.rootNodeId;
                root = source.instanceElements(id).find(node => !!node.closest(".sdk-ppt-notes"));
                canonical = raw.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(id) + '"]') ?? undefined;
                editable = root;
                instance = true;
                root?.closest<HTMLDetailsElement>("details")?.setAttribute("open", "");
            }
            else {
                newNotesSlideId = scope.slideId;
                const panel = frame.createElement("details"), heading = frame.createElement("summary");
                panel.className = "sdk-ppt-notes";
                panel.open = true;
                heading.textContent = "演讲者备注";
                root = frame.createElement("div");
                canonical = frame.createElement("div");
                editable = root;
                const slide = p().slides.find(s => s.slideId === newNotesSlideId)!;
                source.elements.get(slide.rootNodeId)!.element.after(panel);
                panel.append(heading, root);
                draftNotes = panel;
            }
        }
        if (scope.kind === "comment") {
            const annotation = session.getAnnotations().find(a => a.annotationId === scope.annotationId), body = scope.replyId ? annotation?.replies?.find(r => r.replyId === scope.replyId) : annotation, id = body?.bodyNodeId;
            if (id) {
                root = source.instanceElements(id).find(node => !!node.closest(".sdk-ppt-comments"));
                canonical = raw.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(id) + '"]') ?? undefined;
                editable = root;
                instance = true;
            }
        }
        if (!root || !editable || !canonical) {
            inputScope = undefined;
            return;
        }
        input = bindPresentationInput({ session, root, editable, source: canonical, instance, typingStyle, history: navigateHistory, ...(newNotesSlideId ? { newNotesSlideId } : {}), label: inputScope.kind === "text" ? "幻灯片文字输入" : inputScope.kind === "notes" ? "演讲者备注输入" : "批注正文输入", save: value => { inputSelection = value; queueMicrotask(() => { if (!disposed) notifyTextFormat(); }); }, refresh: value => { inputSelection = value; if (!disposed && !deferredRefresh)
                refresh(); }, compositionEnded: () => { if (disposed)
                return; if (deferredRefresh) {
                deferredRefresh = false;
                refresh();
            } renderControls(); }, error: (reason, recovery) => { report(reason); if (recovery)
                try {
                    options.onInputRecovery?.(recovery);
                }
                catch { /* Recovery belongs to the host. */ } } });
        input.focus(inputSelection);
    };
    function finishInput(): void { if (!inputScope)
        return; input?.finish(); disposeInput(); inputScope = undefined; inputSelection = undefined; typingStyle.clear(); gestures?.paint(); regions?.paint(); renderControls(); }
    function editText(nodeId = selected.length === 1 ? selected[0] : undefined): void {
        ensure();
        if (!nodeId || !p().objects.some(o => o.nodeId === nodeId && ["text", "geometry", "table"].includes(o.kind)) || !(definitionSurface ? surfaceRoot(nodeId) === definitionSurface : ownedSlide(nodeId)))
            throw new DocumentError("InvalidOperation", "请选择当前幻灯片中的文字、形状或表格。");
        finishInput();
        regionSelection = undefined;
        selected = [nodeId];
        if (!definitionSurface)
            active = ownedSlide(nodeId)!;
        inputScope = { kind: "text", nodeId };
        gestures?.cancel();
        gestures?.paint();
        bindInput();
        renderControls();
        notify();
    }
    function editNotes(slideId = active): void { ensure(); regionSelection = undefined; if (definitionSurface)
        throw new DocumentError("InvalidOperation", "请先返回幻灯片编辑备注。"); if (!p().slides.some(s => s.slideId === slideId))
        throw new DocumentError("InvalidOperation", "备注必须属于现有幻灯片。"); finishInput(); active = slideId; inputScope = { kind: "notes", slideId }; selected = []; gestures?.cancel(); gestures?.paint(); bindInput(); renderControls(); }
    function editComment(annotationId: string, replyId?: string): void {
        ensure();
        finishInput();
        const annotation = session.getAnnotations().find(a => a.annotationId === annotationId), target = replyId ? annotation?.replies?.find(r => r.replyId === replyId) : annotation;
        if (!target)
            throw new DocumentError("InvalidOperation", "批注或回复不存在。");
        if (!target.bodyNodeId) {
            const text = (target.instruction ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
            perform([{ op: "setPresentationCommentBody", annotationId, ...(replyId ? { replyId } : {}), html: "<p>" + text + "</p>" }]);
        }
        if (definitionSurface)
            setDefinitionSurface();
        inputScope = { kind: "comment", annotationId, ...(replyId ? { replyId } : {}) };
        selected = [];
        gestures?.cancel();
        gestures?.paint();
        if (previewLoaded)
            bindInput();
        renderControls();
    }
    function focusAnnotation(annotationId: string): void {
        ensure();
        const annotation = session.getAnnotations().find(a => a.annotationId === annotationId);
        if (!annotation || annotation.status === "needs_reanchor")
            throw new DocumentError("InvalidOperation", "此批注需要重新定位。");
        focusedAnnotation = annotationId;
        if (definitionSurface)
            setDefinitionSurface();
        const version = generation;
        void ready.then(() => { if (disposed || version !== generation || focusedAnnotation !== annotationId)
            return; regions?.paint(); const box = preview.source()?.rectangles(annotation.anchor)[0]; if (box) {
            stage.scrollTop += preview.element.getBoundingClientRect().top + box.top * zoom - stage.getBoundingClientRect().top - 30;
            stage.scrollLeft += preview.element.getBoundingClientRect().left + box.left * zoom - stage.getBoundingClientRect().left - 30;
        } }).catch(reason => { if (version === generation)
            report(reason); });
    }
    function setMode(mode: "text" | "node" | "region"): void {
        ensure();
        commentControls.clearSelection();
        if (!["text", "node", "region"].includes(mode))
            throw new DocumentError("InvalidOperation", "Unsupported presentation interaction mode.");
        finishInput();
        regionMode = mode === "region";
        if (mode === "text")
            editText();
        gestures?.cancel();
        regions?.cancel();
        gestures?.paint();
        renderControls();
    }
    const bind = (): void => {
        if (disposed)
            return;
        busy = false;
        frameControls?.abort();
        gestures?.dispose();
        regions?.dispose();
        frameControls = new AbortController();
        const source = preview.source()!, frame = source.root.ownerDocument;
        gestures = bindObjectGestures(frame, { session, source, selected: () => selected, select: selectObjects, current: () => current() && !inputScope && !regionMode, perform: operations => { perform(operations); }, error: report });
        frame.addEventListener("keydown", keyboard, { signal: frameControls.signal });
        if (workspace && options.annotationTrigger !== "command") frame.addEventListener("pointerup", () => { if (inputScope?.kind === "text" && !input?.composing() && frame.getSelection()?.isCollapsed === false) commentControls.selectionChanged(); }, { signal: frameControls.signal });
        frame.addEventListener("dblclick", event => { const target = event.target as Element | null, node = target?.closest<HTMLElement>('[data-node-id][data-office-kind]'); if (node && source.elements.get(node.dataset.nodeId!)?.element === node && ["text", "geometry", "table"].includes(node.dataset.officeKind!))
            attempt(() => editText(node.dataset.nodeId)); }, { signal: frameControls.signal });
        regions = bindPresentationRegions(frame, { session, source, enabled: () => current() && regionMode && !definitionSurface, selected: () => regionSelection, focus: () => focusedAnnotation, select: anchor => { if (anchor.kind === "region") {
                const slide = p().slides.find(s => s.rootNodeId === anchor.surfaceId)!;
                selectRegion(slide.slideId, anchor.rect);
            } }, error: report, stale: value => { regionMessage.hidden = !value; regionMessage.textContent = value ? "区域布局或字体已变化，请重新选择区域定位批注。" : ""; } });
        bindInput();
        renderControls();
        gestures.paint();
        setZoom(zoom);
    };
    function refresh(): void {
        if (disposed)
            return;
        if (!previewLoaded) {
            busy = true;
            return;
        }
        if (definitionSurface && ![...p().masters, ...p().layouts, ...p().notesMasters].some(d => d.rootNodeId === definitionSurface)) {
            setDefinitionSurface();
            return;
        }
        if (input?.composing() && !input.committing()) {
            deferredRefresh = true;
            busy = true;
            renderControls();
            return;
        }
        inputSelection = input?.selection() ?? inputSelection;
        disposeInput();
        gestures?.cancel();
        regions?.cancel();
        busy = true;
        const version = ++generation;
        renderControls();
        notify();
        try {
            preview.update(session);
            bind();
            ready = preview.settled().then(() => { if (version === generation && !disposed) {
                setZoom(zoom);
                gestures?.paint();
                regions?.paint();
            } });
            void ready.catch(reason => { if (version === generation)
                report(reason); });
        }
        catch (reason) {
            ready = Promise.reject(reason);
            void ready.catch(report);
            report(reason);
        }
    }
    const loadPreview = (version: number): Promise<void> => preview.ready.then(() => { if (disposed || version !== generation)
        return; previewLoaded = true; if (preview.layout().revision !== session.snapshot().manifest.revision)
        preview.update(session); if (workspace && version === 0) zoom = stage.clientWidth >= p().size.width * 4 / 3 * .8 ? .75 : .5; bind(); return preview.settled().then(() => { if (!disposed && version === generation)
        setZoom(zoom); }); });
    let ready: Promise<void> = loadPreview(0);
    function setDefinitionSurface(rootNodeId?: string): void {
        ensure();
        commentControls.clearSelection();
        if (rootNodeId && ![...p().masters, ...p().layouts, ...p().notesMasters].some(d => d.rootNodeId === rootNodeId))
            throw new DocumentError("InvalidOperation", "请选择现有的母版、布局或备注母版。");
        finishInput();
        disposeInput();
        gestures?.dispose();
        regions?.dispose();
        regionSelection = undefined;
        regionMode = false;
        frameControls?.abort();
        preview.dispose();
        definitionSurface = rootNodeId;
        selected = [];
        busy = true;
        previewLoaded = false;
        const version = ++generation;
        const { presentationSurfaceId: _surface, ...previewOptions } = options;
        preview = mountPreview(previewHost, session, { ...previewOptions, ...(rootNodeId ? { presentationSurfaceId: rootNodeId } : {}) });
        ready = loadPreview(version);
        void ready.catch(reason => { if (version === generation)
            report(reason); });
        renderControls();
    }
    const unsubscribe = session.subscribe(event => { if (event.type === "documentChanged") {
        if (regionSelection && event.affectedSlideIds?.includes(p().slides.find(s => s.rootNodeId === regionSelection!.surfaceId)?.slideId ?? ""))
            regionSelection = undefined;
        refresh();
    } });
    renderControls();
    void ready.catch(reason => { if (generation === 0)
        report(reason); });
    const openInspector = (legend: string): void => {
        const advanced = shadow.querySelector<HTMLDetailsElement>(".advanced-tools"); if (advanced) advanced.open = true;
        const pane = [...inspector.querySelectorAll("fieldset")].find(pane => pane.querySelector("legend")?.textContent === legend);
        (pane ?? inspector).scrollIntoView({ block: "nearest" }); (pane ?? inspector).querySelector<HTMLElement>("input:not(:disabled),select:not(:disabled)")?.focus();
    };
    for (const [id, title] of [["formatObject", "对象外观"], ["slideLayout", "主题、母版与布局"], ["definitions", "主题、母版与布局"], ["charts", "图表与常量数据"], ["tableProperties", "表格结构"]]) commandRegistry.register(id!, title!, () => openInspector(title!), () => !busy && !input?.composing());
    commandRegistry.notify();
    return { element, commands: commandRegistry.commands, executeCommand: commandRegistry.executeCommand,
        clipboard(action) { ensure(); if (!input) return Promise.reject(new DocumentError("InvalidOperation", "请先选择文字。")); return input.clipboard(action); },
        formatState() {
            if (input) return input.formatState();
            const object = selected.length === 1 ? preview.source()?.elements.get(selected[0]!)?.element : undefined;
            if (!object || !p().objects.some(item => item.nodeId === selected[0] && ["text", "geometry", "table"].includes(item.kind))) return { enabled: false };
            const range = object.ownerDocument.createRange(); range.selectNodeContents(object); return { ...readTextFormat(object, range), hasSelection: false, editingText: false };
        }, executeTextCommand(command) {
            ensure(); if (!input && selected.length === 1) { editText(); const root = preview.source()?.elements.get(selected[0]!)?.element; if (root) { const range = root.ownerDocument.createRange(); range.selectNodeContents(root); const selection = root.ownerDocument.getSelection()!; selection.removeAllRanges(); selection.addRange(range); } }
            if (!input || input.composing()) throw new DocumentError("InvalidOperation", "请选择可编辑的文字。");
            if (command.type === "toggle") input.toggle(command.property);
            else if (command.type === "format") input.format(command.property, command.value);
            else input.paragraph("text-align", command.value);
            notifyTextFormat();
        }, get ready() { return ready; }, selection, selectObjects, selectRegion, regionLayoutKey, setDefinitionSurface, editText, editNotes, editComment, finishInput, activeSlide: () => { ensure(); return active; }, setActiveSlide, perform, insertImage, layout: () => { ensure(); return preview.layout(); }, fonts: () => { ensure(); return preview.fonts(); }, setZoom, setMode, focusAnnotation, startAnnotation: () => commentControls.start(),
        dispose() { if (disposed)
            return; disposed = true; ++generation; unsubscribe(); controller.abort(); navControls.abort(); frameControls?.abort(); gestures?.dispose(); regions?.dispose(); disposeInput(); textControls.dispose(); tableControls.dispose(); definitions.dispose(); chartControls.dispose(); commentControls.dispose(); appearance.dispose(); connectors.dispose(); nodeCache.clear(); preview.dispose(); element.remove(); } };
}

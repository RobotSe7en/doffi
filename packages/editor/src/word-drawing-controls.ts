import { wordImageResource } from "./word-image-resource.js";
import { DocumentError, DrawingPresetSchema } from "@doffi/core";
import type {
  DocumentEditingSession,
  DocumentOperation,
  TransactionPrepareOptions,
  WordDrawing,
  WordDrawingLayout,
} from "@doffi/core";
import type { SourceProjection } from "@doffi/render";
import { createOfficeDialog } from "./office-dialog.js";
export type WordDrawingAction =
  | "insertImage"
  | "insertShape"
  | "insertTextBox"
  | "insertConnector"
  | "drawingSettings"
  | "selectionPane";
export function createWordDrawingControls(
  parent: ShadowRoot,
  before: HTMLElement,
  session: DocumentEditingSession,
  selection: () => { paragraphId: string; offset: number } | undefined,
  editStory: (id: string, point?: { x: number; y: number }) => void,
  onError: (error: unknown) => void,
  selectionChanged: () => void,
): {
  open(action: WordDrawingAction): void;
  insertImage(file: File): Promise<void>;
  selectedNodeIds(): readonly string[];
  selectByNodeId(nodeId: string): void;
  reorderSelected(direction: "forward" | "backward"): void;
  groupOrUngroupSelected(): void;
  removeSelected(): void;
  editSelectedText(): void;
  bind(root: HTMLElement, projection: SourceProjection, signal: AbortSignal): void;
  dispose(): void;
} {
  const document = before.ownerDocument;
  let panel: HTMLElement | undefined,
    formatDialog: ReturnType<typeof createOfficeDialog> | undefined,
    controller: AbortController | undefined,
    selected: string[] = [],
    boundRoot: HTMLElement | undefined,
    selectionStyle: HTMLStyleElement | undefined,
    paintSelection = (): void => {},
    updateSelectionPane: (() => void) | undefined;
  const drawings = () => session.snapshot().metadata.word?.drawings ?? [];
  const storyDrawings = () => {
    const all = drawings();
    const storyRootId = boundRoot?.dataset.nodeId;
    const inStory = all.filter((item) => {
      if (storyRootId) {
        for (
          let current: string | null | undefined = item.nodeId;
          current;
          current = session.getNode(current)?.parentNodeId
        )
          if (current === storyRootId) return true;
        return false;
      }
      return Boolean(boundRoot?.querySelector('[data-sdk-word-object="' + CSS.escape(item.nodeId) + '"]'));
    });
    const byId = new Map(inStory.map((item) => [item.nodeId, item]));
    const childrenOf = new Map<string, typeof inStory>();
    const roots: typeof inStory = [];
    for (const item of inStory) {
      let parentDrawing: string | undefined;
      for (
        let parentId = session.getNode(item.nodeId)?.parentNodeId;
        parentId;
        parentId = session.getNode(parentId)?.parentNodeId
      ) {
        if (byId.has(parentId)) {
          parentDrawing = parentId;
          break;
        }
      }
      if (parentDrawing) {
        const list = childrenOf.get(parentDrawing) ?? [];
        list.push(item);
        childrenOf.set(parentDrawing, list);
      } else roots.push(item);
    }
    const ordered: { item: (typeof inStory)[number]; depth: number }[] = [];
    const visit = (items: typeof inStory, depth: number) => {
      const sorted =
        depth === 0
          ? [...items].sort((a, b) => (b.layout.zOrder ?? 0) - (a.layout.zOrder ?? 0))
          : items;
      for (const item of sorted) {
        ordered.push({ item, depth });
        const kids = childrenOf.get(item.nodeId);
        if (kids?.length) visit(kids, depth + 1);
      }
    };
    visit(roots, 0);
    return ordered;
  };
  const selectByNodeId = (nodeId: string): void => {
    if (!drawings().some((item) => item.nodeId === nodeId)) return;
    selected = [nodeId];
    paintSelection();
    updateSelectionPane?.();
  };
  const reorderSelected = (direction: "forward" | "backward"): void => {
    const roots = selectedRoots()
      .map((id) => drawings().find((item) => item.nodeId === id))
      .filter((item): item is WordDrawing => Boolean(item));
    if (!roots.length)
      throw new DocumentError("InvalidOperation", "请先选择一个绘图对象再调整层序。");
    const siblings = storyDrawings()
      .filter((entry) => entry.depth === 0)
      .map((entry) => entry.item);
    const ops: DocumentOperation[] = [];
    for (const target of roots) {
      const currentZ = target.layout.zOrder ?? 0;
      const idx = siblings.findIndex((item) => item.nodeId === target.nodeId);
      const neighbor = direction === "forward" ? siblings[idx - 1] : siblings[idx + 1];
      let nextZ: number;
      if (neighbor) {
        const neighborZ = neighbor.layout.zOrder ?? 0;
        if (direction === "forward") {
          nextZ = Math.max(currentZ + 1, neighborZ + 1);
        } else {
          if (neighborZ > 0) {
            nextZ = Math.max(0, Math.min(currentZ - 1, neighborZ - 1));
          } else {
            nextZ = 0;
            ops.push({
              op: "setWordDrawingLayout",
              nodeId: neighbor.nodeId,
              layout: { ...(JSON.parse(JSON.stringify(neighbor.layout)) as WordDrawingLayout), zOrder: 1 },
            });
          }
        }
      } else {
        nextZ = direction === "forward" ? currentZ + 1 : Math.max(0, currentZ - 1);
      }
      if (nextZ !== currentZ || ops.length) {
        ops.push({
          op: "setWordDrawingLayout",
          nodeId: target.nodeId,
          layout: { ...(JSON.parse(JSON.stringify(target.layout)) as WordDrawingLayout), zOrder: nextZ },
        });
      }
    }
    if (ops.length) apply(ops);
    updateSelectionPane?.();
  };
  const groupOrUngroupSelected = (): void => {
    const roots = selectedRoots()
      .map((id) => drawings().find((item) => item.nodeId === id))
      .filter((item): item is WordDrawing => Boolean(item));
    if (roots.length === 1 && roots[0]!.kind === "group") {
      apply([
        ...materialize([roots[0]!]),
        { op: "ungroupWordDrawing", nodeId: roots[0]!.nodeId },
      ]);
      selected = [];
      paintSelection();
      updateSelectionPane?.();
      return;
    }
    if (roots.length < 2)
      throw new DocumentError("InvalidOperation", "请选择两个及以上浮动对象以组合，或选择一个组合以取消组合。");
    const paragraphs = roots.map((item) => {
      let current = session.getNode(item.nodeId)?.parentNodeId;
      while (current) {
        const node = session.getNode(current);
        if (node && /^p$|^h[1-6]$/.test(node.tagName)) return current;
        current = node?.parentNodeId;
      }
      return undefined;
    });
    const paragraphId = paragraphs[0];
    if (!paragraphId || paragraphs.some((id) => id !== paragraphId))
      throw new DocumentError("InvalidOperation", "仅支持组合同一段落下的浮动绘图对象。");
    apply([
      ...materialize(roots),
      { op: "groupWordDrawings", nodeId: paragraphId, nodeIds: roots.map((item) => item.nodeId) },
    ]);
    selected = [];
    paintSelection();
    updateSelectionPane?.();
  };
  const removeSelected = (): void => {
    const roots = selectedRoots();
    if (!roots.length) return;
    apply(roots.map((nodeId) => ({ op: "removeWordDrawing", nodeId })));
    selected = [];
    paintSelection();
    updateSelectionPane?.();
  };
  const selectedRoots = () => selected.filter(id => {
    for (let parentId = session.getNode(id)?.parentNodeId; parentId; parentId = session.getNode(parentId)?.parentNodeId)
      if (selected.includes(parentId)) return false;
    return true;
  });
  const apply = (operations: DocumentOperation[], options?: TransactionPrepareOptions): void => {
    const snapshot = session.snapshot();
    session.commit(
      session.prepare(
        {
          protocolVersion: 1,
          documentId: snapshot.manifest.documentId,
          baseRevision: snapshot.manifest.revision,
          operations,
        },
        options,
      ),
    );
  };
  const displayed = (object: ReturnType<typeof drawings>[number]) => {
    const node = [
      ...(boundRoot?.querySelectorAll<HTMLElement>(
        '[data-sdk-word-object="' + CSS.escape(object.nodeId) + '"]',
      ) ?? []),
    ].find(
      (node) =>
        node.dataset.sdkWordPositionX !== undefined || node.dataset.sdkWordWidth !== undefined,
    );
    return {
      ...object.geometry,
      x:
        node?.dataset.sdkWordPositionX !== undefined
          ? Number(node.dataset.sdkWordPositionX)
          : object.geometry.x,
      y:
        node?.dataset.sdkWordPositionY !== undefined
          ? Number(node.dataset.sdkWordPositionY)
          : object.geometry.y,
      width:
        node?.dataset.sdkWordWidth !== undefined
          ? Number(node.dataset.sdkWordWidth)
          : object.geometry.width,
      height:
        node?.dataset.sdkWordHeight !== undefined
          ? Number(node.dataset.sdkWordHeight)
          : object.geometry.height,
    };
  };
  const absoluteLayout = (
    input: ReturnType<typeof drawings>[number]["layout"],
  ): WordDrawingLayout => {
    const layout = JSON.parse(JSON.stringify(input)) as WordDrawingLayout;
    delete layout.alignH;
    delete layout.alignV;
    return layout;
  };
  const fixedSizeLayout = (
    input: ReturnType<typeof drawings>[number]["layout"],
  ): WordDrawingLayout => {
    const layout = JSON.parse(JSON.stringify(input)) as WordDrawingLayout;
    delete layout.relativeWidth;
    delete layout.relativeHeight;
    return layout;
  };
  const materialize = (
    objects: readonly ReturnType<typeof drawings>[number][],
    force = false,
  ): DocumentOperation[] =>
    objects.flatMap((object) => {
      if (!force && !object.layout.relativeWidth && !object.layout.relativeHeight) return [];
      return [
        {
          op: "setWordDrawingGeometry" as const,
          nodeId: object.nodeId,
          geometry: displayed(object),
        },
        {
          op: "setWordDrawingLayout" as const,
          nodeId: object.nodeId,
          layout: fixedSizeLayout(absoluteLayout(object.layout)),
        },
      ];
    });
  const close = (): void => {
    controller?.abort();
    formatDialog?.close();
    formatDialog?.dialog.remove();
    formatDialog = undefined;
    panel?.remove();
    panel = undefined;
    updateSelectionPane = undefined;
  };
  const attempt = (run: () => void): void => {
    try {
      run();
    } catch (error) {
      onError(error);
    }
  };
  async function insertImage(file: File): Promise<void> {
    const location = selection(),
      revision = session.snapshot().manifest.revision;
    if (!location) throw new DocumentError("InvalidOperation", "请先在段落中放置光标。");
    const resource = await wordImageResource(document, file);
    if (session.snapshot().manifest.revision !== revision)
      throw new DocumentError("RevisionConflict", "读取图片期间文档已更改，请重试。");
    apply(
      [
        {
          op: "insertWordDrawing",
          nodeId: location.paragraphId,
          offset: location.offset,
          geometry: {
            x: 0,
            y: 0,
            width: resource.width,
            height: resource.height,
            rotation: 0,
            flipH: false,
            flipV: false,
          },
          layout: { mode: "inline" },
          object: { kind: "image", assetPath: resource.path, alt: file.name },
        },
      ],
      resource.options,
    );
  }
  function editSelectedText(): void {
    const current = drawings().find(item => item.nodeId === selected[0]);
    if (selected.length !== 1 || !current || !(current.kind === "shape" || current.kind === "textBox"))
      throw new DocumentError("InvalidOperation", "请选择一个形状或文本框。");
    let storyId = current.textStoryId;
    if (!storyId) {
      apply([{ op: "setWordDrawingText", nodeId: current.nodeId, text: {} }]);
      const updated = drawings().find(item => item.nodeId === current.nodeId)!;
      if ("textStoryId" in updated) storyId = updated.textStoryId;
    }
    close();
    if (storyId) editStory(storyId);
  }
  function open(action: WordDrawingAction): void {
    close();
    if (action === "selectionPane") {
      controller = new AbortController();
      const signal = controller.signal;
      panel = document.createElement("section");
      panel.setAttribute("role", "region");
      panel.setAttribute("aria-label", "选择窗格");
      panel.style.cssText =
        "padding:12px;background:white;border:1px solid #dbe3ef;max-height:360px;overflow:auto;display:flex;flex-direction:column;gap:8px";
      const header = document.createElement("div");
      header.style.cssText = "display:flex;flex-wrap:wrap;align-items:center;gap:6px";
      const title = document.createElement("strong");
      title.textContent = "选择窗格";
      title.style.marginRight = "auto";
      const makeButton = (text: string, onClick: () => void): HTMLButtonElement => {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.textContent = text;
        btn.addEventListener("click", () => attempt(onClick), { signal });
        return btn;
      };
      const showAll = makeButton("全部显示", () => {
        const ops = storyDrawings()
          .filter(({ item }) => item.layout.hidden)
          .map(({ item }) => ({
            op: "setWordDrawingLayout" as const,
            nodeId: item.nodeId,
            layout: { ...(JSON.parse(JSON.stringify(item.layout)) as WordDrawingLayout), hidden: false },
          }));
        if (ops.length) apply(ops);
        updateSelectionPane?.();
      });
      const hideAll = makeButton("全部隐藏", () => {
        const ops = storyDrawings()
          .filter(({ item }) => !item.layout.hidden)
          .map(({ item }) => ({
            op: "setWordDrawingLayout" as const,
            nodeId: item.nodeId,
            layout: { ...(JSON.parse(JSON.stringify(item.layout)) as WordDrawingLayout), hidden: true },
          }));
        if (ops.length) apply(ops);
        updateSelectionPane?.();
      });
      const moveUp = makeButton("上移一层", () => reorderSelected("forward"));
      const moveDown = makeButton("下移一层", () => reorderSelected("backward"));
      let renamingId: string | undefined;
      const renameBtn = makeButton("重命名", () => {
        if (selected.length !== 1) return;
        renamingId = selected[0];
        updateSelectionPane?.();
      });
      const settings = makeButton("对象属性", () => open("drawingSettings"));
      const closeButton = makeButton("关闭选择窗格", close);
      const list = document.createElement("div");
      list.setAttribute("role", "listbox");
      list.setAttribute("aria-label", "绘图对象");
      list.setAttribute("aria-multiselectable", "true");
      const kinds = { image: "图片", shape: "形状", textBox: "文本框", chart: "图表", connector: "连接线", group: "组合" };
      const selectObject = (id: string, multiple: boolean, focus = false): void => {
        selected = multiple
          ? selected.includes(id) ? selected.filter(item => item !== id) : [...selected, id]
          : [id];
        paintSelection();
        updateSelectionPane?.();
        boundRoot?.querySelector('[data-sdk-word-object="' + CSS.escape(id) + '"]')?.scrollIntoView({ block: "nearest" });
        if (focus) list.querySelector<HTMLButtonElement>('[data-word-selection-id="' + CSS.escape(id) + '"]')?.focus();
      };
      list.addEventListener("click", event => {
        const targetEl = event.target as Element;
        if (targetEl.closest("[data-word-selection-toggle]") || targetEl.closest("[data-word-selection-rename]")) return;
        const id = targetEl.closest<HTMLElement>("[data-word-selection-id]")?.dataset.wordSelectionId;
        if (id) selectObject(id, event.shiftKey || event.ctrlKey || event.metaKey, true);
      }, { signal });
      list.addEventListener("dblclick", event => {
        const id = (event.target as Element).closest<HTMLElement>("[data-word-selection-id]")?.dataset.wordSelectionId;
        if (id) {
          selected = [id];
          renamingId = id;
          paintSelection();
          updateSelectionPane?.();
        }
      }, { signal });
      list.addEventListener("keydown", event => {
        if ((event.target as Element).closest("[data-word-selection-rename]")) return;
        if (event.key === "F2" && selected.length === 1) {
          event.preventDefault();
          renamingId = selected[0];
          updateSelectionPane?.();
        } else if (event.key === "Delete" || event.key === "Backspace") {
          event.preventDefault();
          attempt(() => removeSelected());
        } else if (["ArrowUp", "ArrowDown", "Home", "End"].includes(event.key)) {
          event.preventDefault();
          const options = [...list.querySelectorAll<HTMLButtonElement>("[data-word-selection-id]")];
          const current = options.indexOf((event.target as Element).closest<HTMLButtonElement>("[data-word-selection-id]")!);
          const index = event.key === "Home" ? 0 : event.key === "End" ? options.length - 1
            : Math.max(0, Math.min(options.length - 1, current + (event.key === "ArrowUp" ? -1 : 1)));
          const id = options[index]?.dataset.wordSelectionId;
          if (id) selectObject(id, event.shiftKey, true);
        }
      }, { signal });
      updateSelectionPane = () => {
        list.replaceChildren();
        const visible = storyDrawings();
        for (const [index, { item, depth }] of visible.entries()) {
          const row = document.createElement("div");
          row.style.cssText = "display:flex;align-items:center;gap:4px";
          const displayName = item.name || kinds[item.kind] + " " + (index + 1);
          if (renamingId === item.nodeId) {
            const input = document.createElement("input");
            input.type = "text";
            input.value = displayName;
            input.dataset.wordSelectionRename = item.nodeId;
            input.setAttribute("aria-label", "对象名称");
            input.style.cssText = "flex:1;padding:4px 6px;margin-left:" + (6 + depth * 16) + "px";
            let finished = false;
            const finish = (commit: boolean) => {
              if (finished) return;
              finished = true;
              const nextName = input.value.trim();
              renamingId = undefined;
              if (commit && nextName !== (item.name ?? ""))
                attempt(() => apply([{ op: "setWordDrawingName", nodeId: item.nodeId, name: nextName }]));
              updateSelectionPane?.();
            };
            input.addEventListener("keydown", (event) => {
              if (event.key === "Enter") { event.preventDefault(); finish(true); }
              else if (event.key === "Escape") { event.preventDefault(); finish(false); }
            }, { signal });
            input.addEventListener("blur", () => finish(true), { signal });
            row.append(input);
            queueMicrotask(() => { input.focus(); input.select(); });
          } else {
            const option = document.createElement("button");
            option.type = "button";
            option.setAttribute("role", "option");
            option.setAttribute("aria-selected", String(selected.includes(item.nodeId)));
            option.dataset.wordSelectionId = item.nodeId;
            option.textContent = displayName + (item.layout.hidden ? "（已隐藏）" : "");
            option.style.cssText =
              "flex:1;text-align:left;padding:6px;padding-left:" +
              (6 + depth * 16) +
              "px;" +
              (item.layout.hidden ? "opacity:.55;" : "") +
              (selected.includes(item.nodeId) ? "background:#dce9f9;color:#123456" : "");
            row.append(option);
          }
          const toggleVis = document.createElement("button");
          toggleVis.type = "button";
          toggleVis.dataset.wordSelectionToggle = item.nodeId;
          toggleVis.setAttribute("aria-label", (item.layout.hidden ? "显示 " : "隐藏 ") + displayName);
          toggleVis.setAttribute("aria-pressed", String(Boolean(item.layout.hidden)));
          toggleVis.textContent = item.layout.hidden ? "显示" : "隐藏";
          toggleVis.style.cssText = "padding:4px 8px;font-size:12px";
          toggleVis.addEventListener("click", () => {
            attempt(() => {
              const nextLayout = JSON.parse(JSON.stringify(item.layout)) as WordDrawingLayout;
              nextLayout.hidden = !item.layout.hidden;
              apply([{ op: "setWordDrawingLayout", nodeId: item.nodeId, layout: nextLayout }]);
              updateSelectionPane?.();
            });
          }, { signal });
          row.append(toggleVis);
          list.append(row);
        }
        if (!visible.length) list.textContent = "当前故事没有绘图对象。";
        const hasAny = visible.length > 0;
        const hasSelected = selected.some((id) => visible.some((entry) => entry.item.nodeId === id));
        settings.disabled = !hasSelected;
        moveUp.disabled = !hasSelected;
        moveDown.disabled = !hasSelected;
        renameBtn.disabled = selected.length !== 1 || !hasSelected;
        showAll.disabled = !visible.some((entry) => entry.item.layout.hidden);
        hideAll.disabled = !visible.some((entry) => !entry.item.layout.hidden) || !hasAny;
      };
      header.append(title, showAll, hideAll, moveUp, moveDown, renameBtn, settings, closeButton);
      panel.append(header, list);
      parent.insertBefore(panel, before);
      updateSelectionPane();
      return;
    }
    const location = selection();
    if (action !== "drawingSettings" && !location)
      throw new DocumentError("InvalidOperation", "请先在段落中放置光标。");
    selected = selected.filter((id) => drawings().some((item) => item.nodeId === id));
    if (action === "drawingSettings" && !selected.length)
      throw new DocumentError("InvalidOperation", "请先单击一个图片或绘图对象；Shift 单击可多选。");
    controller = new AbortController();
    const signal = controller.signal;
    panel = document.createElement("section");
    panel.setAttribute("role", "region");
    panel.setAttribute("aria-label", "绘图设置");
    panel.style.cssText =
      "display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px 20px";
    formatDialog = createOfficeDialog(document, action === "drawingSettings" ? "形状格式" : "插入绘图对象");
    formatDialog.content.style.cssText = "min-height:0;max-height:65dvh;overflow:auto;margin-top:12px";
    let currentGroup = "size";
    const groupItems: { node: HTMLElement; group: string }[] = [];
    const register = <T extends HTMLElement>(node: T, group = currentGroup): T => {
      node.dataset.wordTabGroup = group;
      groupItems.push({ node, group });
      return node;
    };
    let applyTabFilter = (_next: string): void => {};
    if (action === "drawingSettings") {
      const tabBar = document.createElement("div");
      tabBar.setAttribute("role", "tablist");
      tabBar.setAttribute("aria-label", "形状格式选项卡");
      tabBar.style.cssText = "display:flex;gap:4px;border-bottom:1px solid #dbe3ef;padding-bottom:6px;margin-bottom:10px";
      let activeTab = "all";
      const tabButtons: { id: string; btn: HTMLButtonElement }[] = [];
      applyTabFilter = (next: string): void => {
        activeTab = next;
        for (const item of tabButtons) {
          const on = item.id === activeTab;
          item.btn.setAttribute("aria-selected", String(on));
          item.btn.style.background = on ? "#e8f1ff" : "#f8fafc";
          item.btn.style.borderColor = on ? "#185abd" : "#d5dce6";
          item.btn.style.color = on ? "#185abd" : "#334155";
          item.btn.style.fontWeight = on ? "600" : "400";
        }
        for (const entry of groupItems) {
          entry.node.hidden = activeTab !== "all" && entry.group !== "common" && entry.group !== activeTab;
        }
      };
      for (const [id, label] of [
        ["all", "全部"],
        ["size", "大小与位置"],
        ["wrap", "文字环绕"],
        ["style", "填充与线条"],
        ["text", "文本与裁剪"],
      ] as const) {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.setAttribute("role", "tab");
        btn.textContent = label;
        btn.style.cssText = "padding:3px 10px;border-radius:4px;border:1px solid #d5dce6;font-size:12px;cursor:pointer";
        btn.addEventListener("click", () => applyTabFilter(id), { signal });
        tabButtons.push({ id, btn });
        tabBar.append(btn);
      }
      formatDialog.content.append(tabBar);
    }
    formatDialog.content.append(panel);
    parent.insertBefore(formatDialog.dialog, before);
    formatDialog.dialog.addEventListener("cancel", event => { event.preventDefault(); close(); }, { signal });
    formatDialog.show();
    const button = (label: string, run: () => void, group?: string): void => {
      const node = document.createElement("button");
      node.type = "button";
      node.textContent = label;
      node.addEventListener("click", () => attempt(run), { signal });
      if (label === "关闭绘图设置") formatDialog!.footer.append(node);
      else panel!.append(register(node, group));
    };
    const field = (name: string, value: number | string, type = "number"): HTMLInputElement => {
      const label = document.createElement("label"),
        input = document.createElement("input");
      label.textContent = name;
      input.setAttribute("aria-label", name);
      input.type = type;
      input.value = String(value);
      input.style.width = type === "number" ? "100px" : "100%";
      if (type === "number") input.step = "0.5";
      label.append(input);
      panel!.append(register(label));
      return input;
    };
    const choose = (
      name: string,
      options: readonly (readonly [string, string])[],
      value: string,
    ): HTMLSelectElement => {
      const label = document.createElement("label"),
        input = document.createElement("select");
      label.textContent = name;
      input.setAttribute("aria-label", name);
      for (const [key, text] of options) {
        const option = document.createElement("option");
        option.value = key;
        option.textContent = text;
        input.append(option);
      }
      input.value = value;
      label.append(input);
      panel!.append(register(label));
      return input;
    };
    button("关闭绘图设置", close);
    const current = drawings().find((item) => item.nodeId === selected[0]),
      g = current
        ? displayed(current)
        : {
            x: 0,
            y: 0,
            width: 180,
            height: 100,
            rotation: 0,
            flipH: false,
            flipV: false,
          },
      l = current?.layout ?? { mode: "inline" as const };
    const x = field("对象水平位置（磅）", g.x),
      y = field("对象垂直位置（磅）", g.y),
      width = field("对象宽度（磅）", g.width),
      height = field("对象高度（磅）", g.height),
      rotation = field("对象旋转角度", g.rotation);
    currentGroup = "wrap";
    const mode = choose(
        "对象布局",
        [
          ["inline", "嵌入正文"],
          ["floating", "浮动锚点"],
        ],
        l.mode,
      ),
      wrap = choose(
        "文字环绕",
        [
          ["square", "四周型"],
          ["tight", "紧密型"],
          ["through", "穿越型"],
          ["topAndBottom", "上下型"],
          ["none", "浮于文字上方"],
        ],
        l.wrap ?? "square",
      ),
      horizontal = choose(
        "水平参照",
        [
          ["column", "栏"],
          ["page", "页面"],
          ["margin", "页边距"],
          ["character", "字符"],
          ["leftMargin", "左页边距区域"],
          ["rightMargin", "右页边距区域"],
          ["insideMargin", "内页边距区域"],
          ["outsideMargin", "外页边距区域"],
        ],
        l.horizontal ?? "column",
      ),
      vertical = choose(
        "垂直参照",
        [
          ["paragraph", "段落"],
          ["page", "页面"],
          ["margin", "页边距"],
          ["line", "行"],
          ["topMargin", "上页边距区域"],
          ["bottomMargin", "下页边距区域"],
          ["insideMargin", "内页边距区域"],
          ["outsideMargin", "外页边距区域"],
        ],
        l.vertical ?? "paragraph",
      ),
      side = choose(
        "环绕侧",
        [
          ["bothSides", "两侧"],
          ["left", "左侧"],
          ["right", "右侧"],
          ["largest", "较宽一侧"],
        ],
        l.wrapSide ?? "bothSides",
      );
    const alignH = choose(
        "对象水平对齐",
        [
          ["", "绝对位置"],
          ["left", "左对齐"],
          ["center", "居中"],
          ["right", "右对齐"],
          ["inside", "内侧"],
          ["outside", "外侧"],
        ],
        l.alignH ?? "",
      ),
      alignV = choose(
        "对象垂直对齐",
        [
          ["", "绝对位置"],
          ["top", "顶部"],
          ["center", "居中"],
          ["bottom", "底部"],
          ["inside", "内侧"],
          ["outside", "外侧"],
        ],
        l.alignV ?? "",
      );
    const defaultDistance = current ? 0 : 3,
      wrapDistances = Object.fromEntries(
        (["left", "right", "top", "bottom"] as const).map((key, index) => [
          key,
          field(
            ["左环绕距离（磅）", "右环绕距离（磅）", "上环绕距离（磅）", "下环绕距离（磅）"][
              index
            ]!,
            l.distances?.[key] ?? defaultDistance,
          ),
        ]),
      ) as Record<"left" | "right" | "top" | "bottom", HTMLInputElement>,
      layer = field("对象叠放层次", l.zOrder ?? 0),
      behind = choose(
        "对象前后",
        [
          ["front", "文字前方"],
          ["behind", "文字后方"],
        ],
        l.behindText ? "behind" : "front",
      ),
      locked = choose(
        "锁定对象锚点",
        [
          ["0", "不锁定"],
          ["1", "锁定"],
        ],
        l.locked ? "1" : "0",
      ),
      allowOverlap = choose(
        "对象允许重叠",
        [
          ["1", "允许"],
          ["0", "不允许"],
        ],
        l.allowOverlap === false ? "0" : "1",
      ),
      layoutInCell = choose(
        "对象在单元格内布局",
        [
          ["1", "是"],
          ["0", "否"],
        ],
        l.layoutInCell === false ? "0" : "1",
      );
    for (const input of Object.values(wrapDistances)) input.min = "0";
    layer.min = "0";
    layer.step = "1";
    currentGroup = "size";
    const sizeModeW = choose(
        "对象宽度方式",
        [
          ["absolute", "固定尺寸"],
          ["relative", "百分比"],
        ],
        l.relativeWidth ? "relative" : "absolute",
      ),
      sizeModeH = choose(
        "对象高度方式",
        [
          ["absolute", "固定尺寸"],
          ["relative", "百分比"],
        ],
        l.relativeHeight ? "relative" : "absolute",
      ),
      sizeFromW = choose(
        "相对宽度参照",
        [
          ["page", "页面"],
          ["margin", "正文区域"],
          ["leftMargin", "左页边距区域"],
          ["rightMargin", "右页边距区域"],
          ["insideMargin", "内页边距区域"],
          ["outsideMargin", "外页边距区域"],
        ],
        l.relativeWidth?.relativeFrom ?? "page",
      ),
      sizeFromH = choose(
        "相对高度参照",
        [
          ["page", "页面"],
          ["margin", "正文区域"],
          ["topMargin", "上页边距区域"],
          ["bottomMargin", "下页边距区域"],
          ["insideMargin", "内页边距区域"],
          ["outsideMargin", "外页边距区域"],
        ],
        l.relativeHeight?.relativeFrom ?? "page",
      ),
      percentW = field("相对宽度（%，0 使用固定尺寸）", l.relativeWidth?.percent ?? 50),
      percentH = field("相对高度（%，0 使用固定尺寸）", l.relativeHeight?.percent ?? 50);
    for (const input of [percentW, percentH]) {
      input.min = "0";
      input.max = "2147483.647";
      input.step = "0.001";
    }
    const sizeHint = document.createElement("span");
    sizeHint.textContent = "移动保留百分比；拖动右下角缩放会改为固定尺寸，可撤销。";
    panel.append(register(sizeHint, "size"));
    const syncSizeFields = (): void => {
      const floating = mode.value === "floating",
        relativeW = floating && sizeModeW.value === "relative",
        relativeH = floating && sizeModeH.value === "relative";
      sizeModeW.disabled = sizeModeH.disabled = !floating;
      sizeFromW.disabled = percentW.disabled = !relativeW;
      sizeFromH.disabled = percentH.disabled = !relativeH;
      width.disabled = relativeW && Number(percentW.value) > 0;
      height.disabled = relativeH && Number(percentH.value) > 0;
    };
    for (const input of [mode, sizeModeW, sizeModeH, percentW, percentH])
      input.addEventListener("input", syncSizeFields, { signal });
    syncSizeFields();
    currentGroup = "wrap";
    const polygonText = (l.wrapPolygon ?? []).map((point) => point.x + "," + point.y).join("; "),
      polygon = field("环绕边界点（x,y；21600 为完整宽高）", polygonText, "text");
    polygon.style.width = "320px";
    polygon.maxLength = 100000;
    const contour = document.createElement("details"), contourTitle = document.createElement("summary");
    contourTitle.textContent = "自定义环绕边界";
    contour.style.gridColumn = "1 / -1";
    contour.append(contourTitle, polygon.parentElement!);
    panel.append(register(contour, "wrap"));
    const polygonLayout = (): Pick<WordDrawingLayout, "wrapPolygon" | "wrapPolygonEdited"> => {
      if (polygon.value === polygonText)
        return {
          wrapPolygon: l.wrapPolygon ? l.wrapPolygon.map((point) => ({ ...point })) : undefined,
          wrapPolygonEdited: l.wrapPolygonEdited,
        };
      if (!polygon.value.trim()) return { wrapPolygon: undefined, wrapPolygonEdited: undefined };
      const points = polygon.value.split(/[;；]/).map((value) => {
        const parts = value.trim().split(/[,，]/);
        if (
          parts.length !== 2 ||
          parts.some((part) => !part.trim() || !Number.isSafeInteger(Number(part)))
        )
          throw new DocumentError("InvalidOperation", "环绕边界须为整数 x,y 坐标，以分号分隔。");
        return { x: Number(parts[0]), y: Number(parts[1]) };
      });
      if (points.length < 3)
        throw new DocumentError("InvalidOperation", "环绕边界至少需要三个点。");
      return { wrapPolygon: points, wrapPolygonEdited: true };
    };
    const geometry = (): WordDrawing["geometry"] => ({
      ...g,
      x: Number(x.value),
      y: Number(y.value),
      width: Number(width.value),
      height: Number(height.value),
      rotation: Number(rotation.value),
    });
    const layout = (): WordDrawingLayout => ({
      ...fixedSizeLayout(absoluteLayout(l)),
      ...(mode.value === "floating" && sizeModeW.value === "relative"
        ? {
            relativeWidth: {
              relativeFrom: sizeFromW.value as NonNullable<
                WordDrawingLayout["relativeWidth"]
              >["relativeFrom"],
              percent: Number(percentW.value),
            },
          }
        : {}),
      ...(mode.value === "floating" && sizeModeH.value === "relative"
        ? {
            relativeHeight: {
              relativeFrom: sizeFromH.value as NonNullable<
                WordDrawingLayout["relativeHeight"]
              >["relativeFrom"],
              percent: Number(percentH.value),
            },
          }
        : {}),
      ...polygonLayout(),
      mode: mode.value as "inline" | "floating",
      wrap: wrap.value as NonNullable<WordDrawingLayout["wrap"]>,
      horizontal: horizontal.value as NonNullable<WordDrawingLayout["horizontal"]>,
      vertical: vertical.value as NonNullable<WordDrawingLayout["vertical"]>,
      wrapSide: side.value as NonNullable<WordDrawingLayout["wrapSide"]>,
      ...(alignH.value ? { alignH: alignH.value as NonNullable<WordDrawingLayout["alignH"]> } : {}),
      ...(alignV.value ? { alignV: alignV.value as NonNullable<WordDrawingLayout["alignV"]> } : {}),
      behindText: behind.value === "behind",
      locked: locked.value === "1",
      allowOverlap: allowOverlap.value === "1",
      layoutInCell: layoutInCell.value === "1",
      zOrder: Number(layer.value),
      distances: {
        top: Number(wrapDistances.top.value),
        right: Number(wrapDistances.right.value),
        bottom: Number(wrapDistances.bottom.value),
        left: Number(wrapDistances.left.value),
      },
    });
    if (action !== "drawingSettings") {
      const preset =
          action === "insertShape"
            ? choose(
                "形状类型",
                DrawingPresetSchema.options.map((key) => [key, key]),
                "rect",
              )
            : undefined,
        text =
          action === "insertShape" || action === "insertTextBox"
            ? field("对象文字", "文本框", "text")
            : undefined;
      const insert = (
        object: Extract<DocumentOperation, { op: "insertWordDrawing" }>["object"],
        options?: TransactionPrepareOptions,
      ): void => {
        const previous = new Set(drawings().map((item) => item.nodeId));
        apply(
          [
            {
              op: "insertWordDrawing",
              nodeId: location!.paragraphId,
              offset: location!.offset,
              geometry: geometry(),
              layout: layout(),
              object,
            },
          ],
          options,
        );
        selected = drawings()
          .filter((item) => !previous.has(item.nodeId))
          .map((item) => item.nodeId);
        open("drawingSettings");
      };
      if (action === "insertImage") {
        const input = field("选择图片", "", "file");
        input.accept = "image/png,image/jpeg,image/webp";
        input.addEventListener(
          "change",
          () => {
            void (async () => {
              const file = input.files?.[0];
              if (!file) return;
              const revision = session.snapshot().manifest.revision,
                resource = await wordImageResource(document, file);
              if (session.snapshot().manifest.revision !== revision)
                throw new DocumentError("RevisionConflict", "读取图片期间文档已更改，请重试。");
              width.value = String(resource.width);
              height.value = String(resource.height);
              insert({ kind: "image", assetPath: resource.path, alt: file.name }, resource.options);
            })().catch(onError);
          },
          { signal },
        );
      } else
        button("插入对象", () =>
          insert(
            action === "insertShape"
              ? {
                  kind: "shape",
                  preset: preset!.value as Extract<WordDrawing, { kind: "shape" }>["preset"],
                  text: text!.value,
                }
              : action === "insertTextBox"
                ? { kind: "textBox", text: text!.value }
                : { kind: "connector", connector: { routing: "straight" } },
          ),
        );
      return;
    }
    button("应用对象位置和布局", () => {
      apply(
        selected.flatMap((nodeId) => [
          { op: "setWordDrawingGeometry", nodeId, geometry: geometry() },
          { op: "setWordDrawingLayout", nodeId, layout: layout() },
        ]),
      );
      open("drawingSettings");
    }, "common");
    currentGroup = "size";
    button("水平翻转", () => {
      apply(
        selected.map((nodeId) => ({
          op: "setWordDrawingGeometry",
          nodeId,
          geometry: {
            ...drawings().find((item) => item.nodeId === nodeId)!.geometry,
            flipH: !g.flipH,
          },
        })),
      );
      open("drawingSettings");
    });
    button("垂直翻转", () => {
      apply(
        selected.map((nodeId) => ({
          op: "setWordDrawingGeometry",
          nodeId,
          geometry: {
            ...drawings().find((item) => item.nodeId === nodeId)!.geometry,
            flipV: !g.flipV,
          },
        })),
      );
      open("drawingSettings");
    });
    currentGroup = "style";
    const fill = field("对象填充颜色", "#4472c4", "color"),
      line = field("对象边框颜色", "#224466", "color"),
      lineWidth = field("对象边框宽度（磅）", 1);
    button("应用对象外观", () =>
      apply(
        selected.map((nodeId) => ({
          op: "setWordDrawingAppearance",
          nodeId,
          properties: {
            "--office-fill-kind": "solid",
            "--office-fill-color": fill.value,
            "--office-line-kind": "solid",
            "--office-line-color": line.value,
            "--office-line-width": lineWidth.value + "pt",
          },
        })),
      ),
    );
    currentGroup = "text";
    if (current?.kind === "image") {
      const crop = current.crop ?? { left: 0, top: 0, right: 0, bottom: 0 },
        fields = Object.fromEntries(
          (["left", "top", "right", "bottom"] as const).map((key, index) => [
            key,
            field(
              ["左裁剪百分比", "上裁剪百分比", "右裁剪百分比", "下裁剪百分比"][index]!,
              crop[key] * 100,
            ),
          ]),
        ) as Record<keyof typeof crop, HTMLInputElement>;
      button("应用图片裁剪", () =>
        apply([
          {
            op: "setWordDrawingCrop",
            nodeId: current.nodeId,
            crop: {
              left: Number(fields.left.value) / 100,
              top: Number(fields.top.value) / 100,
              right: Number(fields.right.value) / 100,
              bottom: Number(fields.bottom.value) / 100,
            },
          },
        ]),
      );
    }
    if (current?.kind === "shape") {
      const preset = choose(
        "形状类型",
        DrawingPresetSchema.options.map((key) => [key, key]),
        current.preset,
      );
      button("更改形状", () =>
        apply([
          {
            op: "setWordDrawingPreset",
            nodeId: current.nodeId,
            preset: preset.value as typeof current.preset,
          },
        ]),
      );
    }
    if (current?.kind === "shape" || current?.kind === "textBox") {
      const direction = choose(
          "对象文字方向",
          [
            ["horizontal", "横排"],
            ["vertical", "竖排"],
            ["vertical270", "反向竖排"],
          ],
          current.text?.direction ?? "horizontal",
        ),
        alignment = choose(
          "对象文字垂直对齐",
          [
            ["top", "顶部"],
            ["middle", "居中"],
            ["bottom", "底部"],
          ],
          current.text?.vertical ?? "top",
        ),
        padding = field("对象文字边距（磅）", current.text?.insets?.left ?? 7.2);
      button("应用对象文字属性", () =>
        apply([
          {
            op: "setWordDrawingText",
            nodeId: current.nodeId,
            text: {
              ...current.text,
              direction: direction.value as "horizontal" | "vertical" | "vertical270",
              vertical: alignment.value as "top" | "middle" | "bottom",
              insets: {
                top: Number(padding.value),
                right: Number(padding.value),
                bottom: Number(padding.value),
                left: Number(padding.value),
              },
            },
          },
        ]),
      );
      button("编辑对象文字", editSelectedText);
    }
    if (current?.kind === "connector") {
      const routing = choose(
          "连接线路径",
          [
            ["straight", "直线"],
            ["elbow", "肘形"],
          ],
          current.connector.routing,
        ),
        targets = drawings()
          .filter((item) => item.nodeId !== current.nodeId && item.kind !== "connector")
          .map((item) => [item.nodeId, item.name ?? item.kind + " " + item.nodeId] as const),
        start = choose(
          "连接线起点",
          [["", "自由端点"], ...targets],
          current.connector.startNodeId ?? "",
        ),
        end = choose(
          "连接线终点",
          [["", "自由端点"], ...targets],
          current.connector.endNodeId ?? "",
        ),
        startSite = field("起点连接位置（0–3）", current.connector.startSite ?? 0),
        endSite = field("终点连接位置（0–3）", current.connector.endSite ?? 0);
      button("应用连接线绑定", () =>
        apply([
          {
            op: "setWordDrawingConnector",
            nodeId: current.nodeId,
            connector: {
              routing: routing.value as "straight" | "elbow",
              ...(start.value ? { startNodeId: start.value } : {}),
              ...(end.value ? { endNodeId: end.value } : {}),
              startSite: Number(startSite.value),
              endSite: Number(endSite.value),
            },
          },
        ]),
      );
    }
    button("复制对象", () => {
      apply(selectedRoots().map((nodeId) => ({ op: "copyWordDrawing", nodeId })));
    }, "common");
    button("删除对象", () => {
      apply(selectedRoots().map((nodeId) => ({ op: "removeWordDrawing", nodeId })));
      selected = [];
      close();
    }, "common");
    if (selectedRoots().length > 1)
      button("组合对象", () => {
        const roots = selectedRoots(), id = roots[0]!,
          source = boundRoot?.querySelector<HTMLElement>(
            '[data-sdk-word-object="' + CSS.escape(id) + '"]',
          ),
          paragraphId =
            source?.dataset.sdkWordSourceParagraph ??
            source?.closest<HTMLElement>("p,h1,h2,h3,h4,h5,h6")?.dataset.nodeId;
        if (!paragraphId) throw new DocumentError("InvalidOperation", "组合对象须在同一段落中。");
        const objects = drawings().filter((item) => roots.includes(item.nodeId));
        apply([
          ...materialize(
            objects,
            objects.some((item) => !!item.layout.relativeWidth || !!item.layout.relativeHeight),
          ),
          { op: "groupWordDrawings", nodeId: paragraphId, nodeIds: roots },
        ]);
        const group = drawings().find(
          (item) =>
            item.kind === "group" && session.getNode(item.nodeId)!.childNodeIds.includes(id),
        );
        selected = group ? [group.nodeId] : [];
        open("drawingSettings");
      }, "common");
    if (current?.kind === "group")
      button("取消组合", () => {
        apply([...materialize([current]), { op: "ungroupWordDrawing", nodeId: current.nodeId }]);
        selected = [];
        close();
      }, "common");
    applyTabFilter("all");
  }
  function bind(root: HTMLElement, projection: SourceProjection, signal: AbortSignal): void {
    boundRoot = root;
    selected = selected.filter(
      (id) => !!root.querySelector('[data-sdk-word-object="' + CSS.escape(id) + '"]') ||
        root.closest<HTMLElement>("[data-sdk-word-object]")?.dataset.sdkWordObject === id,
    );
    selectionStyle?.remove();
    selectionStyle = root.ownerDocument.createElement("style");
    selectionStyle.textContent =
      "[data-sdk-word-selected]{outline:2px solid #2563eb!important;outline-offset:2px;cursor:move}";
    projection.markDerived(selectionStyle);
    root.ownerDocument.head.append(selectionStyle);
    const paint = (): void => {
      for (const node of root.querySelectorAll<HTMLElement>("[data-sdk-word-object]"))
        projection.setViewAttribute(
          node,
          "data-sdk-word-selected",
          selected.includes(node.dataset.sdkWordObject!) ? "" : null,
        );
      selectionChanged();
    };
    paintSelection = paint;
    paint();
    updateSelectionPane?.();
    root.addEventListener(
      "pointerdown",
      (event) => {
        let target = (event.target as Element).closest<HTMLElement>("[data-sdk-word-object]");
        if (!target || !root.contains(target)) return;
        if (!event.altKey) {
          for (
            let owner = target.parentElement?.closest<HTMLElement>("[data-sdk-word-object]");
            owner && root.contains(owner);
            owner = owner.parentElement?.closest<HTMLElement>("[data-sdk-word-object]")
          )
            target = owner;
        }
        const handle = target;
        event.preventDefault();
        event.stopImmediatePropagation();
        root.focus({ preventScroll: true });
        const id = target.dataset.sdkWordObject!;
        const paneOpen = updateSelectionPane !== undefined;
        if (event.shiftKey) {
          selected = selected.includes(id)
            ? selected.filter((item) => item !== id)
            : [...selected, id];
          paint();
          if (paneOpen) updateSelectionPane?.();
          return;
        }
        if (!selected.includes(id) || event.altKey && !selectedRoots().includes(id)) selected = [id];
        paint();
        const item = drawings().find((item) => item.nodeId === id)!;
        const sx = event.clientX,
          sy = event.clientY,
          rect = target.getBoundingClientRect(),
          resize = sx > rect.right - 12 && sy > rect.bottom - 12,
          initial = selectedRoots()
            .map((id) => drawings().find((item) => item.nodeId === id)!)
            .filter(Boolean)
            .map((object) => ({ ...object, geometry: displayed(object) })),
          baseRevision = session.snapshot().manifest.revision;
        let dx = 0,
          dy = 0;
        const move = (next: PointerEvent): void => {
          dx = (next.clientX - sx) * 0.75;
          dy = (next.clientY - sy) * 0.75;
          handle.style.outline = "2px dashed #2563eb";
          handle.style.translate = resize ? "0 0" : `${dx}pt ${dy}pt`;
          if (resize) {
            const measured = initial.find((object) => object.nodeId === item.nodeId)!.geometry;
            handle.style.width = Math.max(0.5, measured.width + dx) + "pt";
            handle.style.height = Math.max(0.5, measured.height + dy) + "pt";
          }
        };
        const end = (): void => {
          root.ownerDocument.removeEventListener("pointermove", move);
          root.ownerDocument.removeEventListener("pointerup", end);
          handle.style.removeProperty("translate");
          handle.style.removeProperty("outline");
          if (Math.abs(dx) + Math.abs(dy) > 1)
            attempt(() => {
              if (session.snapshot().manifest.revision !== baseRevision)
                throw new DocumentError("RevisionConflict", "拖动期间文档已更改，请重试。");
              apply(
                initial.flatMap((object) => [
                  {
                    op: "setWordDrawingGeometry" as const,
                    nodeId: object.nodeId,
                    geometry: {
                      ...(resize
                        ? drawings().find((item) => item.nodeId === object.nodeId)!.geometry
                        : object.geometry),
                      ...(resize
                        ? {
                            width: Math.max(0.5, object.geometry.width + dx),
                            height: Math.max(0.5, object.geometry.height + dy),
                          }
                        : { x: object.geometry.x + dx, y: object.geometry.y + dy }),
                    },
                  },
                  ...(resize
                    ? [
                        {
                          op: "setWordDrawingLayout" as const,
                          nodeId: object.nodeId,
                          layout: fixedSizeLayout(object.layout),
                        },
                      ]
                    : !resize
                      ? [
                          {
                            op: "setWordDrawingLayout" as const,
                            nodeId: object.nodeId,
                            layout:
                              object.layout.mode === "inline"
                                ? { mode: "floating" as const, wrap: "none" as const }
                                : absoluteLayout(object.layout),
                          },
                        ]
                      : []),
                ]),
              );
            });
          if (paneOpen) updateSelectionPane?.();
        };
        root.ownerDocument.addEventListener("pointermove", move, { signal });
        root.ownerDocument.addEventListener("pointerup", end, { once: true, signal });
      },
      { signal, capture: true },
    );
    root.addEventListener("dblclick", event => {
      const target = (event.target as Element).closest<HTMLElement>("[data-sdk-word-object]");
      if (!target || !root.contains(target)) return;
      const item = drawings().find(item => item.nodeId === target.dataset.sdkWordObject);
      if (!item || !("textStoryId" in item) || !item.textStoryId) return;
      event.preventDefault(); event.stopImmediatePropagation();
      close();
      editStory(item.textStoryId, { x: event.clientX, y: event.clientY });
    }, { signal, capture: true });
    root.addEventListener(
      "keydown",
      (event) => {
        if (
          root.dataset.sdkWordTextStory ||
          !selected.length ||
          ![
            "ArrowLeft",
            "ArrowRight",
            "ArrowUp",
            "ArrowDown",
            "Delete",
            "Backspace",
            "Escape",
          ].includes(event.key)
        )
          return;
        if (event.key === "Escape") {
          selected = [];
          paint();
          close();
          return;
        }
        event.preventDefault();
        event.stopImmediatePropagation();
        attempt(() => {
          if (event.key === "Delete" || event.key === "Backspace") {
            const paneOpen = updateSelectionPane !== undefined;
            apply(selectedRoots().map((nodeId) => ({ op: "removeWordDrawing", nodeId })));
            selected = [];
            if (paneOpen) updateSelectionPane?.();
            else close();
            return;
          }
          const delta = event.shiftKey ? 10 : 1;
          apply(
            selectedRoots().flatMap((nodeId) => {
              const object = drawings().find((item) => item.nodeId === nodeId)!,
                g = displayed(object);
              return [
                {
                  op: "setWordDrawingGeometry" as const,
                  nodeId,
                  geometry: {
                    ...g,
                    x:
                      g.x +
                      (event.key === "ArrowLeft" ? -delta : event.key === "ArrowRight" ? delta : 0),
                    y:
                      g.y +
                      (event.key === "ArrowUp" ? -delta : event.key === "ArrowDown" ? delta : 0),
                  },
                },
                {
                  op: "setWordDrawingLayout" as const,
                  nodeId,
                  layout: absoluteLayout(object.layout),
                },
              ];
            }),
          );
        });
      },
      { signal, capture: true },
    );
    root.addEventListener(
      "click",
      (event) => {
        if (!(event.target as Element).closest("[data-sdk-word-object]")) {
          selected = [];
          paint();
        }
      },
      { signal },
    );
  }
  return {
    selectedNodeIds: () => selected.filter((id) => drawings().some((d) => d.nodeId === id)),
    selectByNodeId,
    reorderSelected,
    groupOrUngroupSelected,
    removeSelected,
    editSelectedText,
    open,
    insertImage,
    bind,
    dispose() {
      close();
      selectionStyle?.remove();
    },
  };
}

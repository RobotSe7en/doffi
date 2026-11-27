import {
  createOpaqueId,
  DocumentError,
  type DocumentEditingSession,
  type DocumentOperation,
  type PageSetup,
  type WordBuildingBlock,
  type WordTemplate,
  type WordTheme,
} from "@doffi/core";
import { createOfficeDialog } from "./office-dialog.js";
import type { WordParitySelection } from "./word-parity-controls.js";

export const WORD_DEFAULT_TEMPLATE_STORAGE_KEY = "doffi.word.defaultTemplate.v1";
export const WORD_AUTOTEXT_STORAGE_KEY = "doffi.word.autoText.v1";

export interface StoredWordDefaultTemplate {
  readonly name?: string;
  readonly theme?: WordTheme;
  readonly pageSetup?: PageSetup;
}

export function loadStoredWordDefaultTemplate(): StoredWordDefaultTemplate | undefined {
  try {
    const raw = globalThis.localStorage?.getItem(WORD_DEFAULT_TEMPLATE_STORAGE_KEY);
    if (!raw) return undefined;
    const parsed = JSON.parse(raw) as StoredWordDefaultTemplate;
    return parsed && typeof parsed === "object" ? parsed : undefined;
  } catch {
    return undefined;
  }
}

export function saveStoredWordDefaultTemplate(data: StoredWordDefaultTemplate | null): void {
  try {
    if (!data) globalThis.localStorage?.removeItem(WORD_DEFAULT_TEMPLATE_STORAGE_KEY);
    else globalThis.localStorage?.setItem(WORD_DEFAULT_TEMPLATE_STORAGE_KEY, JSON.stringify(data));
  } catch {
    // Ignore storage quota / sandbox restrictions cleanly.
  }
}

function loadStoredAutoText(): WordBuildingBlock[] {
  try {
    const raw = globalThis.localStorage?.getItem(WORD_AUTOTEXT_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? (parsed as WordBuildingBlock[]) : [];
  } catch {
    return [];
  }
}

function saveStoredAutoText(blocks: readonly WordBuildingBlock[]): void {
  try {
    globalThis.localStorage?.setItem(WORD_AUTOTEXT_STORAGE_KEY, JSON.stringify(blocks.slice(0, 64)));
  } catch {
    // Ignore storage errors.
  }
}

export interface WordCoverPagePreset {
  readonly id: string;
  readonly label: string;
  readonly accent: string;
  readonly bg: string;
  readonly textColor: string;
  readonly desc: string;
}

export const WORD_COVER_PAGES: readonly WordCoverPagePreset[] = [
  { id: "austin", label: "奥斯汀 (Austin)", accent: "#2f5496", bg: "#eef4fc", textColor: "#1f3864", desc: "经典顶栏色块与居中主标题" },
  { id: "banded", label: "带状 (Banded)", accent: "#c55a11", bg: "#fff5ee", textColor: "#833c0c", desc: "横向色带视觉焦点与现代排版" },
  { id: "facet", label: "运动型 (Facet)", accent: "#548235", bg: "#f3f8ee", textColor: "#2e541c", desc: "清爽自然左对齐几何色块" },
  { id: "grid", label: "网格 (Grid)", accent: "#44546a", bg: "#f2f4f7", textColor: "#222a35", desc: "严谨商务边框与双栏元信息" },
  { id: "integral", label: "积分 (Integral)", accent: "#7030a0", bg: "#f7f1fb", textColor: "#461b69", desc: "典雅紫调左侧强调边线" },
  { id: "slice", label: "切片 (Slice)", accent: "#0070c0", bg: "#ecf6fc", textColor: "#00467a", desc: "现代报告大气主副标题版式" },
];

export interface WordArtPreset {
  readonly id: string;
  readonly label: string;
  readonly properties: Readonly<Record<string, string>>;
  readonly previewStyle: string;
}

export const WORDART_PRESETS: readonly WordArtPreset[] = [
  {
    id: "grad-blue-shadow",
    label: "蓝色渐变·阴影",
    properties: {
      "font-size": "28pt",
      "font-weight": "700",
      color: "#2f5496",
      "--word-text-fill-gradient": "90deg #5b9bd5 0% #1f4e79 100%",
      "--word-text-shadow": "2pt 2.5pt 45 #44546a",
    },
    previewStyle: "background:linear-gradient(90deg,#5b9bd5,#1f4e79);-webkit-background-clip:text;color:transparent;font-weight:700",
  },
  {
    id: "grad-sunset-bevel",
    label: "橙金渐变·三维棱台",
    properties: {
      "font-size": "28pt",
      "font-weight": "700",
      color: "#c55a11",
      "--word-text-fill-gradient": "135deg #ffc000 0% #ed7d31 55% #c00000 100%",
      "--word-text-bevel": "circle 3pt 3pt",
      "--word-text-3d": "orthographicFront",
    },
    previewStyle: "background:linear-gradient(135deg,#ffc000,#ed7d31,#c00000);-webkit-background-clip:text;color:transparent;font-weight:700",
  },
  {
    id: "outline-navy-glow",
    label: "深蓝轮廓·发光",
    properties: {
      "font-size": "28pt",
      "font-weight": "700",
      color: "#ffffff",
      "--word-text-outline": "1.25pt #1f4e79",
      "--word-text-glow": "6pt #5b9bd5",
    },
    previewStyle: "color:#fff;-webkit-text-stroke:1px #1f4e79;text-shadow:0 0 6px #5b9bd5;font-weight:700",
  },
  {
    id: "grad-emerald-reflect",
    label: "翠绿渐变·倒影",
    properties: {
      "font-size": "28pt",
      "font-weight": "700",
      color: "#385723",
      "--word-text-fill-gradient": "180deg #70ad47 0% #2e541c 100%",
      "--word-text-reflection": "1",
    },
    previewStyle: "background:linear-gradient(180deg,#70ad47,#2e541c);-webkit-background-clip:text;color:transparent;font-weight:700",
  },
  {
    id: "purple-3d-bevel",
    label: "皇家紫·三维浮雕",
    properties: {
      "font-size": "28pt",
      "font-weight": "700",
      color: "#7030a0",
      "--word-text-fill-gradient": "90deg #9b59b6 0% #4a235a 100%",
      "--word-text-bevel": "relaxedInset 4pt 4pt",
      "--word-text-3d": "perspectiveFront",
    },
    previewStyle: "background:linear-gradient(90deg,#9b59b6,#4a235a);-webkit-background-clip:text;color:transparent;font-weight:700",
  },
  {
    id: "gold-metallic",
    label: "鎏金金属·投影",
    properties: {
      "font-size": "28pt",
      "font-weight": "700",
      color: "#bf9000",
      "--word-text-fill-gradient": "180deg #ffe699 0% #bf9000 50% #7f6000 100%",
      "--word-text-outline": "0.75pt #7f6000",
      "--word-text-shadow": "1.5pt 2pt 45 #333333",
    },
    previewStyle: "background:linear-gradient(180deg,#ffe699,#bf9000,#7f6000);-webkit-background-clip:text;color:transparent;font-weight:700",
  },
];

const BUILTIN_QUICK_PARTS: readonly WordBuildingBlock[] = [
  {
    id: "builtin-signature",
    name: "商务签名块",
    gallery: "autoText",
    category: "内置自动图文集",
    description: "包含致礼、签名人及日期占位。",
    html: "<p>此致<br>敬礼！</p><p>签名人：____________<br>日期：2026年11月27日</p>",
  },
  {
    id: "builtin-confidential",
    name: "机密声明栏",
    gallery: "autoText",
    category: "内置自动图文集",
    description: "用于报告或合同首页的内部保密声明。",
    html: '<p style="border:0.75pt solid #b7472a;padding:6pt;background-color:#fdf2ef;color:#7f2a17"><strong>机密声明：</strong>本文档仅供指定收件人审阅，未经书面许可禁止复制、转发或向第三方披露。</p>',
  },
  {
    id: "builtin-meeting-header",
    name: "会议纪要抬头",
    gallery: "quickParts",
    category: "常用文档部件",
    description: "标准会议时间、地点、参会人与议题摘要。",
    html: '<p style="border-bottom:1.5pt solid #2f5496;padding-bottom:4pt"><strong>会议纪要</strong>｜时间：____________　地点：____________　主持人：____________</p>',
  },
  {
    id: "builtin-approval-note",
    name: "审阅批准意见栏",
    gallery: "quickParts",
    category: "常用文档部件",
    description: "标准公文或制度审批栏。",
    html: '<p style="border:0.5pt solid #44546a;padding:6pt">审核意见：________________________　审批人：________　日期：________</p>',
  },
];

function escapeHtml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

export function createWordTemplateControls(
  parent: ShadowRoot,
  before: HTMLElement,
  session: DocumentEditingSession,
  readSelection: () => WordParitySelection,
  callbacks: {
    onUpdateFields?(): void;
    onImportTemplateFile?(file: File): Promise<{
      name: string;
      theme?: WordTheme;
      pageSetup?: PageSetup;
    }>;
    onError(error: unknown): void;
  },
): {
  openCoverPageDialog(): void;
  removeCoverPage(): void;
  openQuickPartsDialog(): void;
  openTemplateDialog(): void;
  saveCurrentAsDefaultTemplate(): void;
  openWordArtDialog(): void;
  dispose(): void;
} {
  const document = before.ownerDocument;
  let activeDialog: ReturnType<typeof createOfficeDialog> | undefined;
  let controller: AbortController | undefined;

  const close = (): void => {
    controller?.abort();
    controller = undefined;
    activeDialog?.close();
    activeDialog?.dialog.remove();
    activeDialog = undefined;
  };

  const commit = (operations: DocumentOperation[]): void => {
    const snapshot = session.snapshot();
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: snapshot.manifest.documentId,
        baseRevision: snapshot.manifest.revision,
        operations,
      }),
    );
  };

  const documentRootId = (): string =>
    new DOMParser()
      .parseFromString(session.snapshot().html, "text/html")
      .querySelector<HTMLElement>('[data-office-kind="document"]')!.dataset.nodeId!;

  const showDialog = (title: string, width = "540px") => {
    close();
    controller = new AbortController();
    const ui = createOfficeDialog(document, title);
    ui.dialog.style.width = width;
    ui.content.style.minHeight = "0";
    activeDialog = ui;
    parent.insertBefore(ui.dialog, before);
    ui.dialog.addEventListener(
      "cancel",
      (event) => {
        event.preventDefault();
        close();
      },
      { signal: controller.signal },
    );
    const addButton = (label: string, run: () => void, primary = false) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = label;
      if (primary) btn.className = "primary";
      btn.addEventListener(
        "click",
        () => {
          try {
            run();
          } catch (error) {
            ui.error(error);
            callbacks.onError(error);
          }
        },
        { signal: controller!.signal },
      );
      ui.footer.append(btn);
      return btn;
    };
    return { ui, signal: controller.signal, addButton };
  };

  function firstBodyParagraphId(): string {
    const meta = session.snapshot().metadata;
    const sectionRoot = !meta.presentation ? meta.sections[0]?.rootNodeId : undefined;
    const doc = new DOMParser().parseFromString(session.snapshot().html, "text/html");
    const scope = sectionRoot
      ? doc.querySelector<HTMLElement>(`[data-node-id="${CSS.escape(sectionRoot)}"]`)
      : doc.querySelector<HTMLElement>('[data-office-story="main"]');
    const first = scope?.querySelector<HTMLElement>("p,h1,h2,h3,h4,h5,h6");
    if (!first?.dataset.nodeId)
      throw new DocumentError("InvalidOperation", "未找到正文首段。");
    return first.dataset.nodeId;
  }

  function removeCoverPage(): void {
    const doc = new DOMParser().parseFromString(session.snapshot().html, "text/html");
    const coverNodes = [...doc.querySelectorAll<HTMLElement>('[data-word-cover-page="true"]')];
    if (!coverNodes.length) throw new DocumentError("InvalidOperation", "当前文档未包含封面页。");
    const ops: DocumentOperation[] = coverNodes.map((node) => ({
      op: "removeNode",
      nodeId: node.dataset.nodeId!,
    }));
    commit(ops);
  }

  function openCoverPageDialog(): void {
    const meta = session.snapshot().metadata;
    const { ui, signal, addButton } = showDialog("内置封面页库", "580px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    let selectedId = WORD_COVER_PAGES[0]!.id;
    const gallery = document.createElement("div");
    gallery.style.cssText = "display:grid;grid-template-columns:repeat(3,1fr);gap:10px";

    const cards: HTMLButtonElement[] = [];
    for (const preset of WORD_COVER_PAGES) {
      const card = document.createElement("button");
      card.type = "button";
      card.dataset.coverPreset = preset.id;
      card.style.cssText =
        `display:flex;flex-direction:column;gap:4px;padding:10px;border-radius:6px;text-align:left;cursor:pointer;background:${preset.bg};border:2px solid ${preset.id === selectedId ? preset.accent : "#d5dce6"}`;
      card.innerHTML =
        `<div style="height:6px;width:100%;border-radius:2px;background:${preset.accent}"></div>` +
        `<strong style="font-size:13px;color:${preset.textColor}">${escapeHtml(preset.label)}</strong>` +
        `<span style="font-size:11px;color:#556">${escapeHtml(preset.desc)}</span>`;
      card.addEventListener(
        "click",
        () => {
          selectedId = preset.id;
          cards.forEach((c) => {
            const p = WORD_COVER_PAGES.find((x) => x.id === c.dataset.coverPreset)!;
            c.style.borderColor = p.id === selectedId ? p.accent : "#d5dce6";
          });
        },
        { signal },
      );
      cards.push(card);
      gallery.append(card);
    }

    const fields = document.createElement("div");
    fields.className = "dialog-options";
    const makeField = (labelText: string, value: string, ariaLabel = labelText) => {
      const l = document.createElement("label");
      l.textContent = labelText;
      const i = document.createElement("input");
      i.value = value;
      i.setAttribute("aria-label", ariaLabel);
      l.append(i);
      return { l, i };
    };
    const titleField = makeField("文档标题", meta.title || "年度项目报告", "封面文档标题");
    const subtitleField = makeField("副标题", "第一阶段成果与规划", "封面副标题");
    const authorField = makeField("作者", (!meta.presentation && meta.word?.properties?.creator) || "文档作者", "封面作者");
    const dateField = makeField("日期", new Date().toLocaleDateString("zh-CN"), "封面日期");
    fields.append(titleField.l, subtitleField.l, authorField.l, dateField.l);

    stack.append(gallery, fields);

    addButton("删除当前封面", () => {
      removeCoverPage();
      close();
    });
    addButton("取消", close);
    addButton(
      "插入封面",
      () => {
        const preset = WORD_COVER_PAGES.find((p) => p.id === selectedId) ?? WORD_COVER_PAGES[0]!;
        const firstId = firstBodyParagraphId();
        const coverHtml =
          `<p data-word-cover-page="true" style="background-color:${preset.bg};border-top:6pt solid ${preset.accent};padding:24pt 24pt 8pt 24pt;margin-bottom:0pt;font-size:28pt;font-weight:700;color:${preset.textColor}">${escapeHtml(titleField.i.value.trim() || "文档标题")}</p>` +
          `<p data-word-cover-page="true" style="background-color:${preset.bg};padding:4pt 24pt 36pt 24pt;margin-bottom:0pt;font-size:15pt;color:${preset.accent}">${escapeHtml(subtitleField.i.value.trim())}</p>` +
          `<p data-word-cover-page="true" style="background-color:${preset.bg};padding:48pt 24pt 24pt 24pt;margin-bottom:0pt;font-size:11pt;color:#44546a">${escapeHtml(authorField.i.value.trim())}　｜　${escapeHtml(dateField.i.value.trim())}<br data-word-break="page"></p>`;
        const ops: DocumentOperation[] = [
          {
            op: "insertFragment",
            nodeId: firstId,
            position: "before",
            html: coverHtml,
          },
          {
            op: "setWordOptions",
            nodeId: documentRootId(),
            title: titleField.i.value.trim() || meta.title || "",
            properties: {
              ...(!meta.presentation ? (meta.word?.properties ?? {}) : {}),
              creator: authorField.i.value.trim(),
            },
          },
        ];
        commit(ops);
        close();
      },
      true,
    );
    ui.show();
  }

  function openQuickPartsDialog(): void {
    const sel = readSelection();
    if (!sel.paragraphId) throw new DocumentError("InvalidOperation", "请先在正文段落中放置光标。");
    const docBlocks = session.snapshot().metadata.word?.buildingBlocks ?? [];
    const savedBlocks = loadStoredAutoText();
    const byId = new Map<string, WordBuildingBlock>();
    for (const b of [...BUILTIN_QUICK_PARTS, ...savedBlocks, ...docBlocks]) byId.set(b.id, b);
    const allBlocks = [...byId.values()];

    const { ui, signal, addButton } = showDialog("文档部件与自动图文集", "580px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const list = document.createElement("div");
    list.style.cssText = "display:grid;grid-template-columns:repeat(2,1fr);gap:8px;max-height:220px;overflow:auto";
    let selectedBlock: WordBuildingBlock = allBlocks[0]!;
    const buttons: HTMLButtonElement[] = [];

    const renderBlocks = () => {
      list.replaceChildren();
      buttons.length = 0;
      for (const block of allBlocks) {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.dataset.blockId = block.id;
        btn.style.cssText =
          `display:flex;flex-direction:column;gap:2px;padding:8px 10px;border-radius:6px;text-align:left;cursor:pointer;background:${block.id === selectedBlock.id ? "#e8f1ff" : "#fafbfd"};border:1px solid ${block.id === selectedBlock.id ? "#185abd" : "#d5dce6"}`;
        btn.innerHTML =
          `<strong style="font-size:12px">${escapeHtml(block.name)}</strong>` +
          `<span style="font-size:11px;color:#556">${escapeHtml(block.category ?? block.gallery)}</span>` +
          (block.description ? `<span style="font-size:11px;color:#667">${escapeHtml(block.description)}</span>` : "");
        btn.addEventListener("click", () => {
          selectedBlock = block;
          renderBlocks();
        }, { signal });
        buttons.push(btn);
        list.append(btn);
      }
    };
    renderBlocks();

    const propBar = document.createElement("div");
    propBar.style.cssText = "display:flex;flex-wrap:wrap;gap:6px;align-items:center";
    const propTitle = document.createElement("span");
    propTitle.style.cssText = "font-size:12px;font-weight:600";
    propTitle.textContent = "快速插入文档属性域：";
    propBar.append(propTitle);
    for (const [label, instr] of [
      ["作者", " AUTHOR "],
      ["标题", " TITLE "],
      ["主题", " SUBJECT "],
      ["关键字", " KEYWORDS "],
      ["备注", " COMMENTS "],
      ["模板", " TEMPLATE "],
    ] as const) {
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = label;
      b.style.cssText = "font-size:11px;padding:3px 8px";
      b.addEventListener("click", () => {
        commit([
          {
            op: "insertWordField",
            nodeId: sel.paragraphId!,
            offset: sel.to,
            instruction: instr,
            cachedText: label,
          },
        ]);
        callbacks.onUpdateFields?.();
        close();
      }, { signal });
      propBar.append(b);
    }

    const saveRow = document.createElement("div");
    saveRow.style.cssText = "display:flex;gap:8px;align-items:center;border-top:1px solid #e2e8f0;padding-top:8px";
    const nameInput = document.createElement("input");
    nameInput.placeholder = "新建自动图文集名称（保存当前所选文本）";
    nameInput.setAttribute("aria-label", "新建自动图文集名称");
    nameInput.style.flex = "1";
    const saveBtn = document.createElement("button");
    saveBtn.type = "button";
    saveBtn.textContent = "将所选内容保存到自动图文集";
    saveBtn.addEventListener("click", () => {
      const name = nameInput.value.trim();
      const content = sel.text.trim();
      if (!name) {
        ui.error("请输入自动图文集名称。");
        return;
      }
      if (!content) {
        ui.error("请先在文档中选中要保存的文本内容。");
        return;
      }
      const nextBlock: WordBuildingBlock = {
        id: "at-" + createOpaqueId(),
        name,
        gallery: "autoText",
        category: "自定义自动图文集",
        description: content.slice(0, 40),
        html: `<p>${escapeHtml(content)}</p>`,
      };
      const existingDoc = session.snapshot().metadata.word?.buildingBlocks ?? [];
      const mergedDoc = [...existingDoc, nextBlock];
      const mergedLocal = [...loadStoredAutoText(), nextBlock];
      saveStoredAutoText(mergedLocal);
      commit([
        {
          op: "setWordOptions",
          nodeId: documentRootId(),
          buildingBlocks: mergedDoc,
        },
      ]);
      allBlocks.push(nextBlock);
      selectedBlock = nextBlock;
      nameInput.value = "";
      ui.error("");
      renderBlocks();
    }, { signal });
    saveRow.append(nameInput, saveBtn);

    stack.append(list, propBar, saveRow);
    addButton("取消", close);
    addButton(
      "插入所选部件",
      () => {
        commit([
          {
            op: "insertFragment",
            nodeId: sel.paragraphId!,
            position: "after",
            html: selectedBlock.html,
          },
        ]);
        close();
      },
      true,
    );
    ui.show();
  }

  function saveCurrentAsDefaultTemplate(): void {
    const meta = session.snapshot().metadata;
    if (meta.presentation) return;
    saveStoredWordDefaultTemplate({
      name: meta.word?.template?.name ?? meta.title ?? "自定义默认模板",
      ...(meta.word?.theme ? { theme: JSON.parse(JSON.stringify(meta.word.theme)) as WordTheme } : {}),
      pageSetup: JSON.parse(JSON.stringify(meta.pageSetup)) as PageSetup,
    });
  }

  function openTemplateDialog(): void {
    const meta = session.snapshot().metadata;
    if (meta.presentation) throw new DocumentError("InvalidOperation", "文档模板仅适用于 Word 文档。");
    const tpl = meta.word?.template;
    const { ui, signal, addButton } = showDialog("文档模板和加载项", "520px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const makeField = (labelText: string, value: string, ariaLabel = labelText) => {
      const l = document.createElement("label");
      l.textContent = labelText;
      const i = document.createElement("input");
      i.value = value;
      i.setAttribute("aria-label", ariaLabel);
      l.append(i);
      return { l, i };
    };
    const nameField = makeField("文档模板名称", tpl?.name ?? "Normal.dotm", "文档模板名称");
    const targetField = makeField("附加模板路径 / 文件名", tpl?.target ?? "", "附加模板路径");
    const linkLabel = document.createElement("label");
    const linkBox = document.createElement("input");
    linkBox.type = "checkbox";
    linkBox.checked = Boolean(tpl?.linkStyles);
    linkBox.setAttribute("aria-label", "自动更新文档样式");
    linkLabel.append(linkBox, "自动更新文档样式 (w:linkStyles)");

    const status = document.createElement("p");
    status.style.cssText = "margin:0;font-size:12px;color:#445";
    const saved = loadStoredWordDefaultTemplate();
    status.textContent = saved
      ? `当前默认新建模板：${saved.name ?? "已自定义"}${saved.theme?.name ? `（主题：${saved.theme.name}）` : ""}`
      : "当前使用标准默认模板 (Normal.dotm)。";

    const toolsRow = document.createElement("div");
    toolsRow.style.cssText = "display:flex;gap:8px;flex-wrap:wrap";
    if (callbacks.onImportTemplateFile) {
      const fileInput = document.createElement("input");
      fileInput.type = "file";
      fileInput.accept = ".docx,.dotx";
      fileInput.hidden = true;
      const importBtn = document.createElement("button");
      importBtn.type = "button";
      importBtn.textContent = "从 .docx / .dotx 导入样式和主题…";
      importBtn.addEventListener("click", () => fileInput.click(), { signal });
      fileInput.addEventListener(
        "change",
        () => {
          const file = fileInput.files?.[0];
          if (!file) return;
          void callbacks
            .onImportTemplateFile!(file)
            .then((imported) => {
              nameField.i.value = imported.name;
              targetField.i.value = imported.name;
              const ops: DocumentOperation[] = [
                {
                  op: "setWordOptions",
                  nodeId: documentRootId(),
                  template: {
                    name: imported.name,
                    target: imported.name,
                    linkStyles: linkBox.checked,
                  },
                  ...(imported.theme ? { theme: imported.theme } : {}),
                },
              ];
              if (imported.pageSetup) {
                const latest = session.snapshot().metadata;
                if (!latest.presentation) {
                  for (const s of latest.sections) {
                    ops.push({
                      op: "setWordSection",
                      nodeId: s.rootNodeId,
                      changes: { pageSetup: imported.pageSetup },
                    });
                  }
                }
              }
              commit(ops);
              status.textContent = `已从模板“${imported.name}”应用主题和页面设置。`;
            })
            .catch((err) => {
              ui.error(err);
              callbacks.onError(err);
            });
        },
        { signal },
      );
      toolsRow.append(fileInput, importBtn);
    }

    const saveDefaultBtn = document.createElement("button");
    saveDefaultBtn.type = "button";
    saveDefaultBtn.textContent = "将当前主题与页面设为默认模板";
    saveDefaultBtn.addEventListener("click", () => {
      saveCurrentAsDefaultTemplate();
      status.textContent = "已保存为本地默认模板，新建 Word 文档时将自动应用。";
    }, { signal });

    const resetDefaultBtn = document.createElement("button");
    resetDefaultBtn.type = "button";
    resetDefaultBtn.textContent = "恢复出厂默认模板";
    resetDefaultBtn.addEventListener("click", () => {
      saveStoredWordDefaultTemplate(null);
      status.textContent = "已恢复标准默认模板 (Normal.dotm)。";
    }, { signal });

    toolsRow.append(saveDefaultBtn, resetDefaultBtn);
    stack.append(nameField.l, targetField.l, linkLabel, toolsRow, status);

    addButton("取消", close);
    addButton(
      "确定",
      () => {
        const name = nameField.i.value.trim();
        const target = targetField.i.value.trim();
        const nextTpl: WordTemplate | null =
          name || target || linkBox.checked
            ? {
                ...(name ? { name } : {}),
                ...(target ? { target } : {}),
                ...(linkBox.checked ? { linkStyles: true } : {}),
              }
            : null;
        commit([
          {
            op: "setWordOptions",
            nodeId: documentRootId(),
            template: nextTpl,
          },
        ]);
        close();
      },
      true,
    );
    ui.show();
  }

  function openWordArtDialog(): void {
    const sel = readSelection();
    if (!sel.paragraphId) throw new DocumentError("InvalidOperation", "请先在正文段落中放置光标。");
    const { ui, signal, addButton } = showDialog("插入艺术字 (WordArt)", "560px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    let chosen = WORDART_PRESETS[0]!;
    const grid = document.createElement("div");
    grid.style.cssText = "display:grid;grid-template-columns:repeat(3,1fr);gap:8px";

    const renderCards = () => {
      grid.replaceChildren();
      for (const preset of WORDART_PRESETS) {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.dataset.wordartPreset = preset.id;
        btn.style.cssText =
          `display:flex;flex-direction:column;align-items:center;gap:4px;padding:10px;border-radius:6px;cursor:pointer;background:#fff;border:2px solid ${preset.id === chosen.id ? "#185abd" : "#d5dce6"}`;
        btn.innerHTML = `<span style="font-size:26px;line-height:1.15;${preset.previewStyle}">A</span><small style="font-size:11px;color:#445">${escapeHtml(preset.label)}</small>`;
        btn.addEventListener("click", () => {
          chosen = preset;
          renderCards();
        }, { signal });
        grid.append(btn);
      }
    };
    renderCards();

    const textLabel = document.createElement("label");
    textLabel.textContent = "艺术字文本";
    const textInput = document.createElement("input");
    textInput.value = sel.text.trim() || "艺术字标题";
    textInput.setAttribute("aria-label", "艺术字文本");
    textLabel.append(textInput);

    stack.append(grid, textLabel);
    addButton("取消", close);
    addButton(
      "确定",
      () => {
        const txt = textInput.value.trim();
        if (!txt) throw new DocumentError("InvalidOperation", "请输入艺术字文本。");
        const styleCss = Object.entries(chosen.properties)
          .map(([k, v]) => `${k}:${v}`)
          .join(";");
        if (sel.to > sel.from && sel.text.trim() === txt) {
          commit([
            {
              op: "formatWordText",
              nodeId: sel.paragraphId!,
              from: sel.from,
              to: sel.to,
              expectedText: sel.text,
              properties: { ...chosen.properties },
            },
          ]);
        } else {
          commit([
            {
              op: "insertFragment",
              nodeId: sel.paragraphId!,
              position: "after",
              html: `<p><span data-word-run="true" style="${escapeHtml(styleCss)}">${escapeHtml(txt)}</span></p>`,
            },
          ]);
        }
        close();
      },
      true,
    );
    ui.show();
  }

  return {
    openCoverPageDialog,
    removeCoverPage,
    openQuickPartsDialog,
    openTemplateDialog,
    saveCurrentAsDefaultTemplate,
    openWordArtDialog,
    dispose: close,
  };
}

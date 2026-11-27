import {
  DocumentError,
  findWordText,
  formatWordBibliography,
  formatWordCitation,
  type DocumentEditingSession,
  type DocumentOperation,
  type WordBibliography,
  type WordBibliographyStyle,
  type WordCitationSource,
  type WordCitationSourceType,
} from "@doffi/core";
import { createOfficeDialog } from "./office-dialog.js";
import type { WordParitySelection } from "./word-parity-controls.js";

const SOURCE_TYPES: readonly (readonly [WordCitationSourceType, string])[] = [
  ["Book", "书籍"],
  ["JournalArticle", "期刊文章"],
  ["ConferenceProceedings", "会议论文集"],
  ["Report", "报告"],
  ["InternetSite", "网站"],
  ["DocumentFromInternetSite", "网站文档"],
  ["BookSection", "书籍章节"],
  ["ArticleInAPeriodical", "定期刊物文章"],
  ["ElectronicSource", "电子资料来源"],
  ["Misc", "其他"],
];

const BIB_STYLES: readonly (readonly [WordBibliographyStyle, string])[] = [
  ["APA", "APA (第 7 版)"],
  ["MLA", "MLA (第 9 版)"],
  ["GB7714", "GB/T 7714-2015"],
  ["Chicago", "Chicago (第 17 版)"],
  ["IEEE", "IEEE"],
];

const TOA_CATEGORY_OPTIONS: readonly (readonly [string, string])[] = [
  ["1", "1 - 案例"],
  ["2", "2 - 法规"],
  ["3", "3 - 其他引文"],
  ["4", "4 - 规则"],
  ["5", "5 - 条约"],
  ["6", "6 - 条例"],
  ["7", "7 - 宪法条文"],
  ["8", "8 - 类别 8"],
  ["9", "9 - 类别 9"],
  ["10", "10 - 类别 10"],
  ["11", "11 - 类别 11"],
  ["12", "12 - 类别 12"],
  ["13", "13 - 类别 13"],
  ["14", "14 - 类别 14"],
  ["15", "15 - 类别 15"],
  ["16", "16 - 类别 16"],
];

const FIELD_CATALOG: readonly {
  readonly name: string;
  readonly category: string;
  readonly template: string;
  readonly desc: string;
}[] = [
  { name: "PAGE", category: "编号", template: " PAGE ", desc: "插入当前页的页码。" },
  { name: "NUMPAGES", category: "编号", template: " NUMPAGES ", desc: "插入文档总页数。" },
  { name: "SECTION", category: "编号", template: " SECTION ", desc: "插入当前节编号。" },
  { name: "SECTIONPAGES", category: "编号", template: " SECTIONPAGES ", desc: "插入当前节的总页数。" },
  { name: "SEQ", category: "编号", template: ' SEQ 图 \\* ARABIC ', desc: "按顺序为章节、图表或条目编号（支持 \\r、\\c、\\n、\\h）。" },
  { name: "DATE", category: "日期和时间", template: ' DATE \\@ "yyyy年M月d日" ', desc: "插入当前日期。" },
  { name: "TIME", category: "日期和时间", template: ' TIME \\@ "HH:mm:ss" ', desc: "插入当前时间。" },
  { name: "CREATEDATE", category: "日期和时间", template: ' CREATEDATE \\@ "yyyy-MM-dd" ', desc: "插入文档创建日期。" },
  { name: "SAVEDATE", category: "日期和时间", template: ' SAVEDATE \\@ "yyyy-MM-dd" ', desc: "插入文档最后保存日期。" },
  { name: "PRINTDATE", category: "日期和时间", template: ' PRINTDATE \\@ "yyyy-MM-dd" ', desc: "插入文档打印日期。" },
  { name: "AUTHOR", category: "文档信息", template: " AUTHOR ", desc: "插入文档属性中的作者名称。" },
  { name: "TITLE", category: "文档信息", template: " TITLE ", desc: "插入文档标题。" },
  { name: "SUBJECT", category: "文档信息", template: " SUBJECT ", desc: "插入文档主题。" },
  { name: "KEYWORDS", category: "文档信息", template: " KEYWORDS ", desc: "插入文档关键字。" },
  { name: "COMMENTS", category: "文档信息", template: " COMMENTS ", desc: "插入文档备注/摘要。" },
  { name: "FILENAME", category: "文档信息", template: " FILENAME ", desc: "插入当前文档文件名。" },
  { name: "TEMPLATE", category: "文档信息", template: " TEMPLATE ", desc: "插入附加模板名称。" },
  { name: "NUMWORDS", category: "文档信息", template: " NUMWORDS ", desc: "插入文档总字数。" },
  { name: "NUMCHARS", category: "文档信息", template: " NUMCHARS ", desc: "插入文档总字符数。" },
  { name: "LASTSAVEDBY", category: "文档信息", template: " LASTSAVEDBY ", desc: "插入最后修改者姓名。" },
  { name: "REVNUM", category: "文档信息", template: " REVNUM ", desc: "插入文档修订次数。" },
  { name: "DOCPROPERTY", category: "文档信息", template: ' DOCPROPERTY "Author" ', desc: "插入指定文档属性值（Author/Title/Subject/Category/Words/Pages 等）。" },
  { name: "STYLEREF", category: "链接和引用", template: ' STYLEREF "Heading 1" ', desc: "插入页面中最近指定样式（如标题 1）的文本，常用于页眉。" },
  { name: "REF", category: "链接和引用", template: ' REF Bookmark1 \\h ', desc: "插入书签内容（支持 \\p 相对位置、\\n/\\r/\\w 编号、\\h 超链接）。" },
  { name: "PAGEREF", category: "链接和引用", template: ' PAGEREF Bookmark1 \\h ', desc: "插入书签所在页码（支持 \\p、\\h）。" },
  { name: "NOTEREF", category: "链接和引用", template: ' NOTEREF NoteBookmark \\h ', desc: "插入脚注或尾注引用标记。" },
  { name: "CITATION", category: "索引和目录", template: " CITATION Source1 ", desc: "插入引文源的文内引用。" },
  { name: "BIBLIOGRAPHY", category: "索引和目录", template: " BIBLIOGRAPHY ", desc: "生成当前文档的书目/参考文献列表。" },
  { name: "INDEX", category: "索引和目录", template: ' INDEX \\h "A" \\c "2" ', desc: "汇总所有 XE 索引标记生成多栏索引。" },
  { name: "TOA", category: "索引和目录", template: ' TOA \\h \\c "1" \\p ', desc: "汇总所有 TA 标记生成引文目录。" },
  { name: "TOC", category: "索引和目录", template: ' TOC \\o "1-3" \\h \\z \\u ', desc: "按标题大纲级别生成文档目录。" },
  { name: "IF", category: "等式和公式", template: ' IF { PAGE } = 1 "首页" "续页" ', desc: "根据条件比较结果输出真值或假值（支持嵌套 { ... } 域）。" },
  { name: "COMPARE", category: "等式和公式", template: " COMPARE 10 > 5 ", desc: "比较两个值，真返回 1，假返回 0。" },
  { name: "=", category: "等式和公式", template: " =SUM(ABOVE) ", desc: "计算算术或表格公式表达式。" },
  { name: "SET", category: "链接和引用", template: ' SET DocVer "v1.0" ', desc: "定义可在后续 REF 域中引用的文档变量。" },
  { name: "QUOTE", category: "文档信息", template: ' QUOTE "doffi" \\* Upper ', desc: "输出指定文字并可结合 \\* / \\# / \\@ 开关格式化。" },
  { name: "SYMBOL", category: "等式和公式", template: " SYMBOL 0x2605 ", desc: "按 Unicode 码点插入特殊字符。" },
];

const SWITCH_SNIPPETS: readonly (readonly [string, string])[] = [
  ["", "无附加开关"],
  [" \\* Upper", "全部大写 (\\* Upper)"],
  [" \\* Lower", "全部小写 (\\* Lower)"],
  [" \\* FirstCap", "句首大写 (\\* FirstCap)"],
  [" \\* Caps", "词首大写 (\\* Caps)"],
  [" \\* ROMAN", "大写罗马数字 (\\* ROMAN)"],
  [" \\* roman", "小写罗马数字 (\\* roman)"],
  [" \\* ALPHABETIC", "大写字母 (\\* ALPHABETIC)"],
  [" \\* CHINESENUM3", "中文小写数字 (\\* CHINESENUM3)"],
  [" \\* CHINESENUM2", "中文大写数字 (\\* CHINESENUM2)"],
  [" \\* CardText", "英文基数词 (\\* CardText)"],
  [" \\* Ordinal", "英文序数 (\\* Ordinal)"],
  [' \\# "#,##0.00"', "千分位小数 (\\# \"#,##0.00\")"],
  [' \\@ "yyyy年M月d日"', "中文日期 (\\@ \"yyyy年M月d日\")"],
];

function escapeFieldQuotes(value: string): string {
  return value.replaceAll("\\", "\\\\").replaceAll('"', '\\"');
}

function parseAuthorsInput(raw: string): WordCitationSource["authors"] {
  const parts = raw
    .split(/[;；]/)
    .map((s) => s.trim())
    .filter(Boolean);
  if (!parts.length) return undefined;
  return parts.map((name) => {
    if (name.includes(",")) {
      const [last, rest] = name.split(",").map((s) => s.trim());
      return { ...(last ? { last } : {}), ...(rest ? { first: rest } : {}) };
    }
    if (/[\u3400-\u9fff]/.test(name) || !name.includes(" ")) return { last: name };
    const tokens = name.split(/\s+/);
    const last = tokens.pop()!;
    return { first: tokens.join(" "), last };
  });
}

function formatAuthorsForInput(authors: WordCitationSource["authors"]): string {
  if (!authors?.length) return "";
  return authors
    .map((a) => {
      if (a.corporate) return a.corporate;
      if (a.last && a.first)
        return /[\u3400-\u9fff]/.test(a.last + a.first) ? a.last + a.first : `${a.last}, ${a.first}`;
      return a.last ?? a.first ?? "";
    })
    .filter(Boolean)
    .join("; ");
}

export function createWordAcademicControls(
  parent: ShadowRoot,
  before: HTMLElement,
  session: DocumentEditingSession,
  readSelection: () => WordParitySelection,
  callbacks: {
    onUpdateFields?(): void;
    onError(error: unknown): void;
  },
): {
  openCitationsManager(): void;
  openInsertCitationDialog(): void;
  openBibliographyDialog(): void;
  setBibliographyStyle(style: WordBibliographyStyle): void;
  openMarkIndexEntryDialog(): void;
  openInsertIndexDialog(): void;
  openMarkToaCitationDialog(): void;
  openInsertToaDialog(): void;
  openInsertFieldDialog(): void;
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

  const documentRootId = (): string =>
    new DOMParser()
      .parseFromString(session.snapshot().html, "text/html")
      .querySelector<HTMLElement>('[data-office-kind="document"]')!.dataset.nodeId!;

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

  const refreshFieldsSoon = (): void => {
    queueMicrotask(() => {
      try {
        callbacks.onUpdateFields?.();
      } catch {
        // Ignore transient layout refresh errors.
      }
    });
  };

  const currentBibliography = (): WordBibliography => {
    const bib = session.snapshot().metadata.word?.bibliography;
    return {
      style: bib?.style ?? "APA",
      sources: bib?.sources ? JSON.parse(JSON.stringify(bib.sources)) as WordCitationSource[] : [],
    };
  };

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

  const makeInput = (labelText: string, value: string, ariaLabel = labelText, type = "text") => {
    const label = document.createElement("label");
    label.textContent = labelText;
    const input = document.createElement("input");
    input.type = type;
    input.value = value;
    input.setAttribute("aria-label", ariaLabel);
    label.append(input);
    return { label, input };
  };

  const makeSelect = (
    labelText: string,
    options: readonly (readonly [string, string])[],
    initial: string,
    ariaLabel = labelText,
  ) => {
    const label = document.createElement("label");
    label.textContent = labelText;
    const select = document.createElement("select");
    select.setAttribute("aria-label", ariaLabel);
    for (const [val, text] of options) {
      const opt = document.createElement("option");
      opt.value = val;
      opt.textContent = text;
      select.append(opt);
    }
    select.value = initial;
    label.append(select);
    return { label, select };
  };

  const makeCheckbox = (labelText: string, checked: boolean, ariaLabel = labelText) => {
    const label = document.createElement("label");
    const input = document.createElement("input");
    input.type = "checkbox";
    input.checked = checked;
    input.setAttribute("aria-label", ariaLabel);
    label.append(input, labelText);
    return { label, input };
  };

  function setBibliographyStyle(style: WordBibliographyStyle): void {
    const bib = currentBibliography();
    commit([
      {
        op: "setWordOptions",
        nodeId: documentRootId(),
        bibliography: { ...bib, style },
      },
    ]);
    refreshFieldsSoon();
  }

  function openCitationsManager(): void {
    const bib = currentBibliography();
    const { ui, signal, addButton } = showDialog("管理源（引文管理器）", "620px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    let sources: WordCitationSource[] = [...bib.sources];
    let editingIndex = -1;

    const styleSelect = makeSelect("引文样式", BIB_STYLES, bib.style ?? "APA", "引文样式");
    const listBox = document.createElement("div");
    listBox.setAttribute("role", "listbox");
    listBox.setAttribute("aria-label", "当前文档引文源");
    listBox.style.cssText =
      "border:1px solid #d0d7e2;border-radius:4px;max-height:130px;overflow:auto;background:#fff;padding:4px;display:flex;flex-direction:column;gap:4px";

    const previewBox = document.createElement("div");
    previewBox.style.cssText =
      "font-size:12px;color:#334;background:#f7f9fc;border:1px solid #dde4ef;border-radius:4px;padding:8px;white-space:pre-wrap";

    const formGrid = document.createElement("div");
    formGrid.className = "dialog-options";
    const typeSel = makeSelect("文献类型", SOURCE_TYPES, "Book", "文献类型");
    const tagInp = makeInput("标记 (Tag)", "", "引文标记");
    const titleInp = makeInput("标题", "", "文献标题");
    const authorInp = makeInput("作者（多作者以 ; 分隔）", "", "文献作者");
    const yearInp = makeInput("年份", String(new Date().getFullYear()), "出版年份");
    const cityInp = makeInput("城市", "", "出版城市");
    const pubInp = makeInput("出版者", "", "出版者");
    const journalInp = makeInput("期刊名称", "", "期刊名称");
    const volInp = makeInput("卷 / 期", "", "卷期");
    const pagesInp = makeInput("页码", "", "文献页码");
    const urlInp = makeInput("URL", "", "文献链接");
    formGrid.append(
      typeSel.label,
      tagInp.label,
      titleInp.label,
      authorInp.label,
      yearInp.label,
      cityInp.label,
      pubInp.label,
      journalInp.label,
      volInp.label,
      pagesInp.label,
      urlInp.label,
    );

    const resetForm = () => {
      editingIndex = -1;
      typeSel.select.value = "Book";
      tagInp.input.value = `Src${sources.length + 1}`;
      titleInp.input.value = "";
      authorInp.input.value = "";
      yearInp.input.value = String(new Date().getFullYear());
      cityInp.input.value = "";
      pubInp.input.value = "";
      journalInp.input.value = "";
      volInp.input.value = "";
      pagesInp.input.value = "";
      urlInp.input.value = "";
    };

    const loadForm = (index: number) => {
      const s = sources[index];
      if (!s) return;
      editingIndex = index;
      typeSel.select.value = s.sourceType;
      tagInp.input.value = s.tag;
      titleInp.input.value = s.title;
      authorInp.input.value = formatAuthorsForInput(s.authors);
      yearInp.input.value = s.year ?? "";
      cityInp.input.value = s.city ?? "";
      pubInp.input.value = s.publisher ?? "";
      journalInp.input.value = s.journalName ?? "";
      volInp.input.value = [s.volume, s.issue].filter(Boolean).join("/");
      pagesInp.input.value = s.pages ?? "";
      urlInp.input.value = s.url ?? "";
    };

    const renderList = () => {
      listBox.replaceChildren();
      const style = styleSelect.select.value as WordBibliographyStyle;
      if (!sources.length) {
        const empty = document.createElement("div");
        empty.style.cssText = "padding:8px;font-size:12px;color:#666";
        empty.textContent = "暂无引文源，请在下方填写后点击“保存引文源”。";
        listBox.append(empty);
        previewBox.textContent = "引文预览：（选择或添加一个引文源）";
        return;
      }
      sources.forEach((src, index) => {
        const row = document.createElement("div");
        row.style.cssText =
          "display:flex;align-items:center;gap:6px;padding:4px 6px;border-radius:4px;background:" +
          (index === editingIndex ? "#e8f1ff" : "#fafbfd");
        const btn = document.createElement("button");
        btn.type = "button";
        btn.style.cssText =
          "flex:1;text-align:left;border:0;background:transparent;padding:2px;cursor:pointer;font-size:12px";
        btn.textContent = `[${src.tag}] ${src.title}${src.year ? ` (${src.year})` : ""}`;
        btn.addEventListener("click", () => {
          loadForm(index);
          renderList();
        }, { signal });
        const del = document.createElement("button");
        del.type = "button";
        del.textContent = "删除";
        del.style.cssText = "font-size:11px;padding:2px 6px";
        del.addEventListener("click", () => {
          sources = sources.filter((_, i) => i !== index);
          if (editingIndex === index) resetForm();
          else if (editingIndex > index) editingIndex--;
          renderList();
        }, { signal });
        row.append(btn, del);
        listBox.append(row);
      });
      const active = sources[editingIndex >= 0 ? editingIndex : 0]!;
      const bibLine = formatWordBibliography([active], style)[0]?.text ?? "";
      previewBox.textContent = `文内引文：${formatWordCitation(active, style, { ordinal: 1 })}\n书目条目：${bibLine}`;
    };

    const actionRow = document.createElement("div");
    actionRow.style.cssText = "display:flex;gap:8px;flex-wrap:wrap";
    const saveSrcBtn = document.createElement("button");
    saveSrcBtn.type = "button";
    saveSrcBtn.textContent = "保存引文源";
    saveSrcBtn.addEventListener(
      "click",
      () => {
        const title = titleInp.input.value.trim();
        if (!title) {
          ui.error("请输入文献标题。");
          return;
        }
        const rawTag = tagInp.input.value.trim() || `Src${sources.length + 1}`;
        const tag = rawTag.replace(/[^A-Za-z0-9_.-]/g, "_") || `Src${sources.length + 1}`;
        const [volume, issue] = volInp.input.value
          .split("/")
          .map((s) => s.trim())
          .filter(Boolean);
        const authors = parseAuthorsInput(authorInp.input.value);
        const entry: WordCitationSource = {
          tag,
          sourceType: typeSel.select.value as WordCitationSourceType,
          title,
          ...(authors ? { authors } : {}),
          ...(yearInp.input.value.trim() ? { year: yearInp.input.value.trim() } : {}),
          ...(cityInp.input.value.trim() ? { city: cityInp.input.value.trim() } : {}),
          ...(pubInp.input.value.trim() ? { publisher: pubInp.input.value.trim() } : {}),
          ...(journalInp.input.value.trim() ? { journalName: journalInp.input.value.trim() } : {}),
          ...(volume ? { volume } : {}),
          ...(issue ? { issue } : {}),
          ...(pagesInp.input.value.trim() ? { pages: pagesInp.input.value.trim() } : {}),
          ...(urlInp.input.value.trim() ? { url: urlInp.input.value.trim() } : {}),
        };
        if (editingIndex >= 0) {
          sources[editingIndex] = entry;
        } else {
          if (sources.some((s) => s.tag === entry.tag)) {
            ui.error(`引文标记 "${entry.tag}" 已存在，请更换标记。`);
            return;
          }
          sources.push(entry);
          editingIndex = sources.length - 1;
        }
        ui.error("");
        renderList();
      },
      { signal },
    );
    const newSrcBtn = document.createElement("button");
    newSrcBtn.type = "button";
    newSrcBtn.textContent = "新建空白源";
    newSrcBtn.addEventListener("click", () => {
      resetForm();
      renderList();
    }, { signal });
    actionRow.append(saveSrcBtn, newSrcBtn);

    styleSelect.select.addEventListener("change", renderList, { signal });
    if (sources.length) loadForm(0);
    else resetForm();
    renderList();

    stack.append(styleSelect.label, listBox, previewBox, formGrid, actionRow);

    addButton("取消", close);
    addButton(
      "应用到文档",
      () => {
        commit([
          {
            op: "setWordOptions",
            nodeId: documentRootId(),
            bibliography: {
              style: styleSelect.select.value as WordBibliographyStyle,
              sources,
            },
          },
        ]);
        refreshFieldsSoon();
        close();
      },
      true,
    );
    ui.show();
  }

  function openInsertCitationDialog(): void {
    const sel = readSelection();
    if (!sel.paragraphId) throw new DocumentError("InvalidOperation", "请先在正文段落中放置光标。");
    const bib = currentBibliography();
    if (!bib.sources.length) {
      openCitationsManager();
      return;
    }
    const { ui, addButton } = showDialog("插入引文", "480px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const sourceSel = makeSelect(
      "选择引文源",
      bib.sources.map((s) => [s.tag, `[${s.tag}] ${s.title}${s.year ? ` (${s.year})` : ""}`] as const),
      bib.sources[0]!.tag,
      "选择引文源",
    );
    const pages = makeInput("页码（可选）", "", "引文页码");
    const suppressAuthor = makeCheckbox("隐藏作者 (\\n)", false, "隐藏作者");
    const suppressYear = makeCheckbox("隐藏年份 (\\y)", false, "隐藏年份");
    stack.append(sourceSel.label, pages.label, suppressAuthor.label, suppressYear.label);

    addButton("管理源…", () => openCitationsManager());
    addButton("取消", close);
    addButton(
      "插入",
      () => {
        const tag = sourceSel.select.value;
        const src = bib.sources.find((s) => s.tag === tag);
        if (!src) return;
        const switches = [
          pages.input.value.trim() ? `\\p "${escapeFieldQuotes(pages.input.value.trim())}"` : "",
          suppressAuthor.input.checked ? "\\n" : "",
          suppressYear.input.checked ? "\\y" : "",
        ]
          .filter(Boolean)
          .join(" ");
        const instruction = ` CITATION ${tag}${switches ? " " + switches : ""} `;
        const cached = formatWordCitation(src, bib.style ?? "APA", {
          ...(pages.input.value.trim() ? { pages: pages.input.value.trim() } : {}),
          suppressAuthor: suppressAuthor.input.checked,
          suppressYear: suppressYear.input.checked,
          ordinal: 1,
        });
        commit([
          {
            op: "insertWordField",
            nodeId: sel.paragraphId!,
            offset: sel.to,
            instruction,
            cachedText: cached,
          },
        ]);
        refreshFieldsSoon();
        close();
      },
      true,
    );
    ui.show();
  }

  function openBibliographyDialog(): void {
    const sel = readSelection();
    if (!sel.paragraphId) throw new DocumentError("InvalidOperation", "请先在正文段落中放置光标。");
    const bib = currentBibliography();
    const { ui, addButton } = showDialog("插入书目 / 参考文献", "500px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const styleSel = makeSelect("参考文献格式", BIB_STYLES, bib.style ?? "APA", "参考文献格式");
    const titleSel = makeSelect(
      "书目标题",
      [
        ["参考文献", "参考文献"],
        ["书目", "书目"],
        ["引用的作品", "引用的作品"],
        ["", "无标题"],
      ],
      "参考文献",
      "书目标题",
    );
    const info = document.createElement("p");
    info.style.cssText = "margin:0;font-size:12px;color:#555";
    info.textContent = bib.sources.length
      ? `当前文档包含 ${bib.sources.length} 条引文源，插入后可随“样式”切换自动更新。`
      : "当前文档尚无引文源，可先在“管理源”中添加文献。";
    stack.append(styleSel.label, titleSel.label, info);

    addButton("管理源…", () => openCitationsManager());
    addButton("取消", close);
    addButton(
      "插入书目",
      () => {
        const chosenStyle = styleSel.select.value as WordBibliographyStyle;
        const ops: DocumentOperation[] = [];
        if (chosenStyle !== bib.style) {
          ops.push({
            op: "setWordOptions",
            nodeId: documentRootId(),
            bibliography: { ...bib, style: chosenStyle },
          });
        }
        const headingText = titleSel.select.value;
        const cached = bib.sources.length
          ? formatWordBibliography(bib.sources, chosenStyle)
              .map((item) => item.text)
              .join("\n")
          : "此文档中没有引文源。";
        if (headingText) {
          ops.push({
            op: "replaceText",
            nodeId: sel.paragraphId!,
            from: sel.to,
            to: sel.to,
            expectedText: "",
            text: headingText + "\n",
          });
        }
        ops.push({
          op: "insertWordField",
          nodeId: sel.paragraphId!,
          instruction: " BIBLIOGRAPHY ",
          cachedText: cached,
        });
        commit(ops);
        refreshFieldsSoon();
        close();
      },
      true,
    );
    ui.show();
  }

  function openMarkIndexEntryDialog(): void {
    const sel = readSelection();
    if (!sel.paragraphId) throw new DocumentError("InvalidOperation", "请先在正文段落中放置光标。");
    const { ui, addButton } = showDialog("标记索引项 (XE)", "500px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const mainEntry = makeInput("主索引项", sel.text.trim(), "主索引项");
    const subEntry = makeInput("次索引项（可选）", "", "次索引项");
    const seeAlso = makeInput("交叉引用（如 参见 ...，可选）", "", "交叉引用");
    const bold = makeCheckbox("页码加粗 (\\b)", false, "页码加粗");
    const italic = makeCheckbox("页码斜体 (\\i)", false, "页码斜体");
    stack.append(mainEntry.label, subEntry.label, seeAlso.label, bold.label, italic.label);

    const buildInstruction = (): string => {
      const main = mainEntry.input.value.trim();
      if (!main) throw new DocumentError("InvalidOperation", "请输入主索引项文本。");
      const entry = subEntry.input.value.trim() ? `${main}:${subEntry.input.value.trim()}` : main;
      const switches = [
        bold.input.checked ? "\\b" : "",
        italic.input.checked ? "\\i" : "",
        seeAlso.input.value.trim() ? `\\t "${escapeFieldQuotes(seeAlso.input.value.trim())}"` : "",
      ]
        .filter(Boolean)
        .join(" ");
      return ` XE "${escapeFieldQuotes(entry)}"${switches ? " " + switches : ""} `;
    };

    addButton("关闭", close);
    addButton("标记全部", () => {
      const term = mainEntry.input.value.trim();
      const instruction = buildInstruction();
      const matches = findWordText(session, term, { matchCase: false });
      if (!matches.length) throw new DocumentError("InvalidOperation", `未在文档中找到“${term}”。`);
      // Insert right-to-left per paragraph so offsets stay exact in one atomic transaction.
      const sorted = [...matches].sort((a, b) =>
        a.nodeId === b.nodeId ? b.to - a.to : 0,
      );
      commit(
        sorted.map((m) => ({
          op: "insertWordField" as const,
          nodeId: m.nodeId,
          offset: m.to,
          instruction,
          cachedText: "",
        })),
      );
      refreshFieldsSoon();
      close();
    });
    addButton(
      "标记",
      () => {
        const instruction = buildInstruction();
        commit([
          {
            op: "insertWordField",
            nodeId: sel.paragraphId!,
            offset: sel.to,
            instruction,
            cachedText: "",
          },
        ]);
        refreshFieldsSoon();
        close();
      },
      true,
    );
    ui.show();
  }

  function openInsertIndexDialog(): void {
    const sel = readSelection();
    if (!sel.paragraphId) throw new DocumentError("InvalidOperation", "请先在正文段落中放置光标。");
    const { ui, addButton } = showDialog("索引", "460px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const columns = makeSelect(
      "栏数",
      [
        ["1", "1 栏"],
        ["2", "2 栏（默认）"],
        ["3", "3 栏"],
      ],
      "2",
      "索引栏数",
    );
    const alphaHeader = makeCheckbox("显示字母标题 (\\h \"A\")", true, "显示字母标题");
    const sep = makeInput("页码分隔符 (\\e)", ", ", "页码分隔符");
    const xeCount = (session.snapshot().metadata.word?.fields ?? []).filter((f) =>
      /^\s*XE\s+/i.test(f.instruction),
    ).length;
    const info = document.createElement("p");
    info.style.cssText = "margin:0;font-size:12px;color:#555";
    info.textContent = xeCount
      ? `当前文档已标记 ${xeCount} 个索引项，插入后将按页码汇总并压缩连续页码。`
      : "当前文档尚未标记任何索引项，可先使用“标记条目”添加 XE 标记。";
    stack.append(columns.label, alphaHeader.label, sep.label, info);

    addButton("取消", close);
    addButton(
      "确定",
      () => {
        const parts = [
          "INDEX",
          alphaHeader.input.checked ? '\\h "A"' : "",
          `\\c "${columns.select.value}"`,
          sep.input.value !== ", " ? `\\e "${escapeFieldQuotes(sep.input.value)}"` : "",
        ].filter(Boolean);
        commit([
          {
            op: "insertWordField",
            nodeId: sel.paragraphId!,
            instruction: ` ${parts.join(" ")} `,
            cachedText: "索引",
          },
        ]);
        refreshFieldsSoon();
        close();
      },
      true,
    );
    ui.show();
  }

  function openMarkToaCitationDialog(): void {
    const sel = readSelection();
    if (!sel.paragraphId) throw new DocumentError("InvalidOperation", "请先在正文段落中放置光标。");
    const { ui, addButton } = showDialog("标记引文 (TA)", "500px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const selectedText = sel.text.trim();
    const longCit = makeInput("所选引文（完整引用 \\l）", selectedText, "完整引文");
    const shortCit = makeInput("简短引文 (\\s)", selectedText, "简短引文");
    const category = makeSelect("类别 (\\c)", TOA_CATEGORY_OPTIONS, "1", "引文类别");
    stack.append(longCit.label, shortCit.label, category.label);

    const buildInstruction = (): string => {
      const l = longCit.input.value.trim();
      const s = shortCit.input.value.trim() || l;
      if (!l) throw new DocumentError("InvalidOperation", "请输入完整引文内容。");
      return ` TA \\l "${escapeFieldQuotes(l)}" \\s "${escapeFieldQuotes(s)}" \\c ${category.select.value} `;
    };

    addButton("关闭", close);
    addButton("标记全部", () => {
      const term = shortCit.input.value.trim() || longCit.input.value.trim();
      const instruction = buildInstruction();
      const matches = findWordText(session, term, { matchCase: false });
      if (!matches.length) throw new DocumentError("InvalidOperation", `未在文档中找到“${term}”。`);
      const sorted = [...matches].sort((a, b) =>
        a.nodeId === b.nodeId ? b.to - a.to : 0,
      );
      commit(
        sorted.map((m) => ({
          op: "insertWordField" as const,
          nodeId: m.nodeId,
          offset: m.to,
          instruction,
          cachedText: "",
        })),
      );
      refreshFieldsSoon();
      close();
    });
    addButton(
      "标记",
      () => {
        const instruction = buildInstruction();
        commit([
          {
            op: "insertWordField",
            nodeId: sel.paragraphId!,
            offset: sel.to,
            instruction,
            cachedText: "",
          },
        ]);
        refreshFieldsSoon();
        close();
      },
      true,
    );
    ui.show();
  }

  function openInsertToaDialog(): void {
    const sel = readSelection();
    if (!sel.paragraphId) throw new DocumentError("InvalidOperation", "请先在正文段落中放置光标。");
    const { ui, addButton } = showDialog("引文目录 (TOA)", "460px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const category = makeSelect(
      "类别",
      [["all", "全部类别"], ...TOA_CATEGORY_OPTIONS],
      "all",
      "引文目录类别",
    );
    const usePassim = makeCheckbox("使用“同上”(passim，≥5 页自动合并 \\p)", true, "使用 passim");
    const showHeader = makeCheckbox("显示类别标题 (\\h)", true, "显示类别标题");
    const taCount = (session.snapshot().metadata.word?.fields ?? []).filter((f) =>
      /^\s*TA\s+/i.test(f.instruction),
    ).length;
    const info = document.createElement("p");
    info.style.cssText = "margin:0;font-size:12px;color:#555";
    info.textContent = taCount
      ? `当前文档已标记 ${taCount} 条引文目录项 (TA)。`
      : "当前文档尚未标记任何引文目录项，可先使用“标记引文”添加 TA 标记。";
    stack.append(category.label, usePassim.label, showHeader.label, info);

    addButton("取消", close);
    addButton(
      "确定",
      () => {
        const parts = [
          "TOA",
          showHeader.input.checked ? "\\h" : "",
          category.select.value !== "all" ? `\\c "${category.select.value}"` : "",
          usePassim.input.checked ? "\\p" : "",
        ].filter(Boolean);
        commit([
          {
            op: "insertWordField",
            nodeId: sel.paragraphId!,
            instruction: ` ${parts.join(" ")} `,
            cachedText: "引文目录",
          },
        ]);
        refreshFieldsSoon();
        close();
      },
      true,
    );
    ui.show();
  }

  function openInsertFieldDialog(): void {
    const sel = readSelection();
    if (!sel.paragraphId) throw new DocumentError("InvalidOperation", "请先在正文段落中放置光标。");
    const { ui, signal, addButton } = showDialog("域", "560px");
    const stack = document.createElement("div");
    stack.className = "dialog-stack";
    ui.content.append(stack);

    const categories = ["全部", ...new Set(FIELD_CATALOG.map((item) => item.category))];
    const catSel = makeSelect(
      "类别",
      categories.map((c) => [c, c] as const),
      "全部",
      "域类别",
    );
    const fieldSel = document.createElement("select");
    fieldSel.setAttribute("aria-label", "域名");
    fieldSel.size = 7;
    fieldSel.style.cssText = "width:100%;font-size:12px";

    const desc = document.createElement("p");
    desc.style.cssText = "margin:0;font-size:12px;color:#445";

    const instr = makeInput("域代码", FIELD_CATALOG[0]!.template, "域代码");
    instr.input.style.fontFamily = "monospace";
    const sw = makeSelect("格式开关", SWITCH_SNIPPETS, "", "域格式开关");

    const populateFields = () => {
      const cat = catSel.select.value;
      const items = cat === "全部" ? FIELD_CATALOG : FIELD_CATALOG.filter((f) => f.category === cat);
      fieldSel.replaceChildren();
      for (const item of items) {
        const opt = document.createElement("option");
        opt.value = item.name;
        opt.textContent = item.name;
        fieldSel.append(opt);
      }
      if (items[0]) {
        fieldSel.value = items[0].name;
        instr.input.value = items[0].template;
        desc.textContent = items[0].desc;
      }
    };
    populateFields();

    catSel.select.addEventListener("change", populateFields, { signal });
    fieldSel.addEventListener(
      "change",
      () => {
        const item = FIELD_CATALOG.find((f) => f.name === fieldSel.value);
        if (item) {
          instr.input.value = item.template;
          desc.textContent = item.desc;
        }
      },
      { signal },
    );
    sw.select.addEventListener(
      "change",
      () => {
        if (sw.select.value && !instr.input.value.includes(sw.select.value.trim())) {
          instr.input.value = instr.input.value.trimEnd() + sw.select.value + " ";
        }
      },
      { signal },
    );

    stack.append(catSel.label, fieldSel, desc, instr.label, sw.label);
    addButton("取消", close);
    addButton(
      "确定",
      () => {
        const code = instr.input.value.trim();
        if (!code) throw new DocumentError("InvalidOperation", "域代码不能为空。");
        commit([
          {
            op: "insertWordField",
            nodeId: sel.paragraphId!,
            offset: sel.to,
            instruction: ` ${code} `,
            cachedText: code.split(/\s+/)[0] ?? "域",
          },
        ]);
        refreshFieldsSoon();
        close();
      },
      true,
    );
    ui.show();
  }

  return {
    openCitationsManager,
    openInsertCitationDialog,
    openBibliographyDialog,
    setBibliographyStyle,
    openMarkIndexEntryDialog,
    openInsertIndexDialog,
    openMarkToaCitationDialog,
    openInsertToaDialog,
    openInsertFieldDialog,
    dispose: close,
  };
}

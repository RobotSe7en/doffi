import { DocumentError } from "@doffi/core";
import type { DocumentEditingSession, DocumentOperation } from "@doffi/core";
import { createOfficeDialog } from "./office-dialog.js";

export type WordTableQuickAction =
  | "insertRowAbove"
  | "insertRowBelow"
  | "insertColLeft"
  | "insertColRight"
  | "deleteRow"
  | "deleteCol"
  | "deleteTable";

export interface WordToolsSelection {
  readonly paragraphIds: readonly string[];
  readonly tableId?: string;
  readonly cellId?: string;
  readonly caret?: { readonly nodeId: string; readonly offset: number };
}

const escapeHtml = (value: string): string =>
  value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");

function parseSortableValue(raw: string, type: "text" | "number" | "date"): number | string {
  const trimmed = raw.replace(/\uFFFC/g, "").trim();
  if (type === "number") {
    const cleaned = trimmed.replace(/[,，\s¥$€£%]/g, "");
    const num = Number(cleaned);
    return Number.isFinite(num) ? num : Number.POSITIVE_INFINITY;
  }
  if (type === "date") {
    const normalized = trimmed
      .replace(/(\d{4})年(\d{1,2})月(\d{1,2})日?/, "$1-$2-$3")
      .replaceAll("/", "-");
    const ts = Date.parse(normalized);
    return Number.isFinite(ts) ? ts : Number.POSITIVE_INFINITY;
  }
  return trimmed;
}

function compareValues(
  a: number | string,
  b: number | string,
  direction: "asc" | "desc",
): number {
  const factor = direction === "asc" ? 1 : -1;
  if (typeof a === "number" && typeof b === "number") return (a - b) * factor;
  return String(a).localeCompare(String(b), "zh-CN", { numeric: true }) * factor;
}

function tableGeometry(session: DocumentEditingSession, tableId: string) {
  const table = session.getNode(tableId);
  if (!table) throw new DocumentError("InvalidOperation", "当前表格不存在。");
  const rowContainerId =
    table.childNodeIds.length === 1 && session.getNode(table.childNodeIds[0]!)?.tagName === "tbody"
      ? table.childNodeIds[0]!
      : tableId;
  const rowIds = table.childNodeIds.flatMap((id) => {
    const node = session.getNode(id)!;
    return node.tagName === "tr" ? [id] : node.childNodeIds;
  });
  const cells: { id: string; row: number; column: number; rows: number; columns: number }[] = [];
  const occupied: string[][] = Array.from({ length: rowIds.length }, () => []);
  let columnCount = 0;
  for (const [r, rowId] of rowIds.entries()) {
    for (const cellId of session.getNode(rowId)!.childNodeIds) {
      const cell = session.getNode(cellId)!;
      let c = 0;
      while (occupied[r]![c]) c++;
      const rows = Number(cell.attributes.rowspan ?? 1) || rowIds.length - r;
      const columns = Number(cell.attributes.colspan ?? 1);
      cells.push({ id: cellId, row: r, column: c, rows, columns });
      for (let y = r; y < r + rows; y++)
        for (let x = c; x < c + columns; x++) occupied[y]![x] = cellId;
      columnCount = Math.max(columnCount, c + columns);
    }
  }
  return { rowContainerId, rowIds, cells, columnCount };
}

/** Word general-purpose tools: Sort, Text<->Table conversion, Blank Page, Date/Time, and quick table ops. */
export function createWordToolsControls(
  parent: ShadowRoot,
  before: HTMLElement,
  session: DocumentEditingSession,
  readSelection: () => WordToolsSelection,
  onError: (error: unknown) => void,
): {
  openSort(): void;
  openTextToTable(): void;
  openTableToText(): void;
  openDateTime(): void;
  insertBlankPage(): void;
  runTableAction(action: WordTableQuickAction): void;
  setTableCellShading(color: string): void;
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

  const showDialog = (title: string) => {
    close();
    controller = new AbortController();
    const ui = createOfficeDialog(document, title);
    ui.dialog.style.width = "420px";
    ui.content.style.minHeight = "0";
    activeDialog = ui;
    parent.insertBefore(ui.dialog, before);
    ui.dialog.addEventListener("cancel", (event) => {
      event.preventDefault();
      close();
    }, { signal: controller.signal });
    const addButton = (label: string, run: () => void, primary = false) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = label;
      if (primary) btn.className = "primary";
      btn.addEventListener("click", () => {
        try {
          run();
        } catch (error) {
          ui.error(error);
          onError(error);
        }
      }, { signal: controller!.signal });
      ui.footer.append(btn);
      return btn;
    };
    return { ui, signal: controller.signal, addButton };
  };

  function openSort(): void {
    const sel = readSelection();
    const inTable = Boolean(sel.tableId);
    let tableInfo: ReturnType<typeof tableGeometry> | undefined;
    let paragraphParentId: string | undefined;
    let paragraphRange: string[] = [];

    if (inTable) {
      tableInfo = tableGeometry(session, sel.tableId!);
      if (tableInfo.rowIds.length < 2)
        throw new DocumentError("InvalidOperation", "表格至少需要两行才能排序。");
    } else {
      if (!sel.paragraphIds.length)
        throw new DocumentError("InvalidOperation", "请先选择要排序的段落或表格。");
      const firstNode = session.getNode(sel.paragraphIds[0]!);
      if (!firstNode?.parentNodeId)
        throw new DocumentError("InvalidOperation", "当前段落无法排序。");
      paragraphParentId = firstNode.parentNodeId;
      const parentNode = session.getNode(paragraphParentId)!;
      const siblings = parentNode.childNodeIds;
      if (sel.paragraphIds.length === 1) {
        // Like Microsoft Word: sorting with a collapsed caret sorts contiguous
        // sibling paragraphs around the caret.
        const idx = siblings.indexOf(sel.paragraphIds[0]!);
        let start = idx,
          end = idx;
        while (start > 0 && /^p$|^h[1-6]$/.test(session.getNode(siblings[start - 1]!)?.tagName ?? ""))
          start--;
        while (
          end + 1 < siblings.length &&
          /^p$|^h[1-6]$/.test(session.getNode(siblings[end + 1]!)?.tagName ?? "")
        )
          end++;
        paragraphRange = siblings.slice(start, end + 1);
      } else {
        const sameParent = sel.paragraphIds.filter(
          (id) => session.getNode(id)?.parentNodeId === paragraphParentId,
        );
        const indices = sameParent.map((id) => siblings.indexOf(id)).filter((i) => i >= 0);
        const min = Math.min(...indices),
          max = Math.max(...indices);
        paragraphRange = siblings
          .slice(min, max + 1)
          .filter((id) => /^p$|^h[1-6]$/.test(session.getNode(id)?.tagName ?? ""));
      }
      if (paragraphRange.length < 2)
        throw new DocumentError("InvalidOperation", "至少需要两个相邻段落才能排序。");
    }

    const { ui, addButton } = showDialog("排序");
    const body = document.createElement("div");
    body.style.cssText = "display:flex;flex-direction:column;gap:12px;padding-top:10px";
    ui.content.append(body);

    let columnSelect: HTMLSelectElement | undefined;
    let hasHeaderBox: HTMLInputElement | undefined;
    if (inTable && tableInfo) {
      const colLabel = document.createElement("label");
      colLabel.textContent = "主要关键字:";
      columnSelect = document.createElement("select");
      columnSelect.setAttribute("aria-label", "排序列");
      for (let c = 0; c < tableInfo.columnCount; c++) {
        const opt = document.createElement("option");
        opt.value = String(c);
        const headerCell = tableInfo.cells.find((cell) => cell.row === 0 && cell.column === c);
        const preview = headerCell ? session.getNode(headerCell.id)?.text.trim().slice(0, 16) : "";
        opt.textContent = preview ? `列 ${c + 1} (${preview})` : `列 ${c + 1}`;
        columnSelect.append(opt);
      }
      colLabel.append(columnSelect);
      body.append(colLabel);

      const headerLabel = document.createElement("label");
      hasHeaderBox = document.createElement("input");
      hasHeaderBox.type = "checkbox";
      hasHeaderBox.setAttribute("aria-label", "有标题行");
      const firstRowNode = session.getNode(tableInfo.rowIds[0]!);
      hasHeaderBox.checked = Boolean(
        firstRowNode?.childNodeIds.some((id) => session.getNode(id)?.tagName === "th") ||
          firstRowNode?.attributes.style?.includes("--word-row-repeat-header:1"),
      );
      headerLabel.append(hasHeaderBox, "有标题行（保留首行不参与排序）");
      body.append(headerLabel);
    }

    const typeLabel = document.createElement("label");
    typeLabel.textContent = "类型:";
    const typeSelect = document.createElement("select");
    typeSelect.setAttribute("aria-label", "排序类型");
    for (const [val, label] of [
      ["text", "文本"],
      ["number", "数字"],
      ["date", "日期"],
    ] as const) {
      const opt = document.createElement("option");
      opt.value = val;
      opt.textContent = label;
      typeSelect.append(opt);
    }
    typeLabel.append(typeSelect);
    body.append(typeLabel);

    const dirLabel = document.createElement("label");
    dirLabel.textContent = "顺序:";
    const dirSelect = document.createElement("select");
    dirSelect.setAttribute("aria-label", "排序方向");
    for (const [val, label] of [
      ["asc", "升序"],
      ["desc", "降序"],
    ] as const) {
      const opt = document.createElement("option");
      opt.value = val;
      opt.textContent = label;
      dirSelect.append(opt);
    }
    dirLabel.append(dirSelect);
    body.append(dirLabel);

    addButton("取消", close);
    addButton(
      "确定",
      () => {
        const sortType = typeSelect.value as "text" | "number" | "date";
        const dir = dirSelect.value as "asc" | "desc";
        if (inTable && tableInfo) {
          const skipFirst = Boolean(hasHeaderBox?.checked);
          const colIndex = Number(columnSelect?.value ?? "0");
          const sortableRows = skipFirst ? tableInfo.rowIds.slice(1) : [...tableInfo.rowIds];
          if (sortableRows.length < 2) {
            close();
            return;
          }
          const keyed = sortableRows.map((rowId, order) => {
            const rowIndex = tableInfo!.rowIds.indexOf(rowId);
            const cell = tableInfo!.cells.find(
              (item) => item.row === rowIndex && item.column <= colIndex && colIndex < item.column + item.columns,
            );
            const raw = cell ? (session.getNode(cell.id)?.text ?? "") : "";
            return { rowId, order, key: parseSortableValue(raw, sortType) };
          });
          keyed.sort((a, b) => compareValues(a.key, b.key, dir) || a.order - b.order);
          const nextOrder = keyed.map((item) => item.rowId);
          if (nextOrder.some((id, idx) => id !== sortableRows[idx])) {
            commit([
              { op: "reorderWordNodes", nodeId: tableInfo.rowContainerId, childNodeIds: nextOrder },
            ]);
          }
        } else if (paragraphParentId) {
          const keyed = paragraphRange.map((nodeId, order) => ({
            nodeId,
            order,
            key: parseSortableValue(session.getNode(nodeId)?.text ?? "", sortType),
          }));
          keyed.sort((a, b) => compareValues(a.key, b.key, dir) || a.order - b.order);
          const nextOrder = keyed.map((item) => item.nodeId);
          if (nextOrder.some((id, idx) => id !== paragraphRange[idx])) {
            commit([{ op: "reorderWordNodes", nodeId: paragraphParentId, childNodeIds: nextOrder }]);
          }
        }
        close();
      },
      true,
    );
    ui.show();
  }

  function openTextToTable(): void {
    const sel = readSelection();
    if (!sel.paragraphIds.length)
      throw new DocumentError("InvalidOperation", "请先选择要转换为表格的段落。");
    const paragraphs = sel.paragraphIds
      .map((id) => session.getNode(id))
      .filter((node): node is NonNullable<typeof node> => Boolean(node && /^p$|^h[1-6]$/.test(node.tagName)));
    if (!paragraphs.length)
      throw new DocumentError("InvalidOperation", "请选择正文段落以转换为表格。");

    const { ui, signal, addButton } = showDialog("将文字转换成表格");
    const body = document.createElement("div");
    body.style.cssText = "display:flex;flex-direction:column;gap:12px;padding-top:10px";
    ui.content.append(body);

    const sepLabel = document.createElement("label");
    sepLabel.textContent = "文字分隔位置:";
    const sepSelect = document.createElement("select");
    sepSelect.setAttribute("aria-label", "文字分隔位置");
    const hasTab = paragraphs.some((p) => p.text.includes("\t"));
    const hasComma = paragraphs.some((p) => /[,，]/.test(p.text));
    for (const [val, label] of [
      ["tab", "制表符"],
      ["comma", "逗号"],
      ["paragraph", "段落标记"],
      ["other", "其他字符"],
    ] as const) {
      const opt = document.createElement("option");
      opt.value = val;
      opt.textContent = label;
      sepSelect.append(opt);
    }
    sepSelect.value = hasTab ? "tab" : hasComma ? "comma" : "paragraph";
    sepLabel.append(sepSelect);
    body.append(sepLabel);

    const customLabel = document.createElement("label");
    customLabel.textContent = "自定义分隔符:";
    const customInput = document.createElement("input");
    customInput.type = "text";
    customInput.maxLength = 4;
    customInput.value = "|";
    customInput.setAttribute("aria-label", "自定义分隔符");
    customLabel.append(customInput);
    customLabel.hidden = sepSelect.value !== "other";
    sepSelect.addEventListener("change", () => {
      customLabel.hidden = sepSelect.value !== "other";
    }, { signal });
    body.append(customLabel);

    addButton("取消", close);
    addButton(
      "确定",
      () => {
        const mode = sepSelect.value;
        const rows: string[][] = [];
        if (mode === "paragraph") {
          for (const p of paragraphs) rows.push([p.text.replace(/\uFFFC/g, "")]);
        } else {
          const delim =
            mode === "tab"
              ? /\t/
              : mode === "comma"
                ? /[,，]/
                : new RegExp(customInput.value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") || ",");
          for (const p of paragraphs)
            rows.push(p.text.replace(/\uFFFC/g, "").split(delim).map((cell) => cell.trim()));
        }
        const columns = Math.max(1, ...rows.map((row) => row.length));
        const normalized = rows.map((row) =>
          Array.from({ length: columns }, (_, c) => row[c] ?? ""),
        );
        const anchorId = paragraphs[0]!.nodeId;
        commit([
          {
            op: "insertWordTable",
            nodeId: anchorId,
            position: "before",
            rows: normalized.length,
            columns,
            cells: normalized,
          },
          ...paragraphs.map((p) => ({ op: "removeNode" as const, nodeId: p.nodeId })),
        ]);
        close();
      },
      true,
    );
    ui.show();
  }

  function openTableToText(): void {
    const sel = readSelection();
    if (!sel.tableId)
      throw new DocumentError("InvalidOperation", "请先将光标放在要转换的表格中。");
    const info = tableGeometry(session, sel.tableId);
    const { ui, signal, addButton } = showDialog("表格转换成文本");
    const body = document.createElement("div");
    body.style.cssText = "display:flex;flex-direction:column;gap:12px;padding-top:10px";
    ui.content.append(body);

    const sepLabel = document.createElement("label");
    sepLabel.textContent = "文字分隔符:";
    const sepSelect = document.createElement("select");
    sepSelect.setAttribute("aria-label", "表格转换分隔符");
    for (const [val, label] of [
      ["tab", "制表符"],
      ["comma", "逗号"],
      ["paragraph", "段落标记"],
      ["other", "其他字符"],
    ] as const) {
      const opt = document.createElement("option");
      opt.value = val;
      opt.textContent = label;
      sepSelect.append(opt);
    }
    sepLabel.append(sepSelect);
    body.append(sepLabel);

    const customLabel = document.createElement("label");
    customLabel.textContent = "自定义分隔符:";
    const customInput = document.createElement("input");
    customInput.type = "text";
    customInput.maxLength = 4;
    customInput.value = "|";
    customInput.setAttribute("aria-label", "自定义转换分隔符");
    customLabel.append(customInput);
    customLabel.hidden = true;
    sepSelect.addEventListener("change", () => {
      customLabel.hidden = sepSelect.value !== "other";
    }, { signal });
    body.append(customLabel);

    addButton("取消", close);
    addButton(
      "确定",
      () => {
        const mode = sepSelect.value;
        const lines: string[] = [];
        for (const rowId of info.rowIds) {
          const rowTexts = session
            .getNode(rowId)!
            .childNodeIds.map((cellId) => session.getNode(cellId)?.text ?? "");
          if (mode === "paragraph") {
            lines.push(...rowTexts);
          } else {
            const sep = mode === "tab" ? "\t" : mode === "comma" ? "," : customInput.value || ",";
            lines.push(rowTexts.join(sep));
          }
        }
        const html = lines.map((line) => "<p>" + escapeHtml(line) + "</p>").join("");
        commit([
          { op: "insertFragment", nodeId: sel.tableId!, position: "after", html: html || "<p></p>" },
          { op: "deleteWordTable", nodeId: sel.tableId! },
        ]);
        close();
      },
      true,
    );
    ui.show();
  }

  function openDateTime(): void {
    const sel = readSelection();
    const targetNodeId = sel.caret?.nodeId ?? sel.paragraphIds[0];
    if (!targetNodeId)
      throw new DocumentError("InvalidOperation", "请先在正文段落中放置光标。");
    const now = new Date();
    const yyyy = String(now.getFullYear());
    const m = String(now.getMonth() + 1);
    const mm = m.padStart(2, "0");
    const d = String(now.getDate());
    const dd = d.padStart(2, "0");
    const hh = String(now.getHours()).padStart(2, "0");
    const min = String(now.getMinutes()).padStart(2, "0");
    const sec = String(now.getSeconds()).padStart(2, "0");
    const formats: readonly [string, string][] = [
      [`${yyyy}年${m}月${d}日`, 'DATE \\@ "yyyy年M月d日"'],
      [`${yyyy}/${m}/${d}`, 'DATE \\@ "yyyy/M/d"'],
      [`${yyyy}-${mm}-${dd}`, 'DATE \\@ "yyyy-MM-dd"'],
      [`${yyyy}年${m}月${d}日 ${hh}:${min}`, 'DATE \\@ "yyyy年M月d日 HH:mm"'],
      [`${hh}:${min}:${sec}`, 'TIME \\@ "HH:mm:ss"'],
    ];

    const { ui, addButton } = showDialog("日期和时间");
    const body = document.createElement("div");
    body.style.cssText = "display:flex;flex-direction:column;gap:12px;padding-top:10px";
    ui.content.append(body);

    const fmtLabel = document.createElement("label");
    fmtLabel.textContent = "可用格式:";
    const fmtSelect = document.createElement("select");
    fmtSelect.setAttribute("aria-label", "日期和时间格式");
    fmtSelect.size = formats.length;
    for (const [index, [text]] of formats.entries()) {
      const opt = document.createElement("option");
      opt.value = String(index);
      opt.textContent = text;
      fmtSelect.append(opt);
    }
    fmtSelect.value = "0";
    fmtLabel.append(fmtSelect);
    body.append(fmtLabel);

    const autoLabel = document.createElement("label");
    const autoBox = document.createElement("input");
    autoBox.type = "checkbox";
    autoBox.setAttribute("aria-label", "自动更新");
    autoLabel.append(autoBox, "自动更新（作为 Word 域插入）");
    body.append(autoLabel);

    addButton("取消", close);
    addButton(
      "确定",
      () => {
        const [text, instruction] = formats[Number(fmtSelect.value)] ?? formats[0]!;
        const offset = sel.caret?.offset ?? (session.getNode(targetNodeId)?.text.length ?? 0);
        if (autoBox.checked) {
          commit([
            {
              op: "insertWordField",
              nodeId: targetNodeId,
              instruction: " " + instruction + " ",
              cachedText: text,
              offset,
            },
          ]);
        } else {
          commit([
            {
              op: "replaceText",
              nodeId: targetNodeId,
              from: offset,
              to: offset,
              expectedText: "",
              text,
            },
          ]);
        }
        close();
      },
      true,
    );
    ui.show();
  }

  function insertBlankPage(): void {
    const sel = readSelection();
    const paragraphId = sel.paragraphIds[0];
    if (!paragraphId)
      throw new DocumentError("InvalidOperation", "请先在正文段落中放置光标。");
    commit([
      {
        op: "insertFragment",
        nodeId: paragraphId,
        position: "after",
        html: '<p><br data-word-break="page"></p><p><br data-word-break="page"></p>',
      },
    ]);
  }

  function runTableAction(action: WordTableQuickAction): void {
    const sel = readSelection();
    if (!sel.tableId)
      throw new DocumentError("InvalidOperation", "请先将光标放在表格单元格内。");
    if (action === "deleteTable") {
      commit([{ op: "deleteWordTable", nodeId: sel.tableId }]);
      return;
    }
    const info = tableGeometry(session, sel.tableId);
    const active = info.cells.find((cell) => cell.id === sel.cellId) ?? info.cells[0]!;
    const change: Extract<DocumentOperation, { op: "editWordTable" }>["change"] =
      action === "insertRowAbove"
        ? { kind: "insertRows", at: active.row, count: 1 }
        : action === "insertRowBelow"
          ? { kind: "insertRows", at: active.row + active.rows, count: 1 }
          : action === "deleteRow"
            ? { kind: "deleteRows", at: active.row, count: active.rows }
            : action === "insertColLeft"
              ? { kind: "insertColumns", at: active.column, count: 1 }
              : action === "insertColRight"
                ? { kind: "insertColumns", at: active.column + active.columns, count: 1 }
                : { kind: "deleteColumns", at: active.column, count: active.columns };
    commit([{ op: "editWordTable", nodeId: sel.tableId, change }]);
  }

  function setTableCellShading(color: string): void {
    const sel = readSelection();
    if (!sel.cellId)
      throw new DocumentError("InvalidOperation", "请先选择表格单元格。");
    commit([
      {
        op: "formatWordCell",
        nodeId: sel.cellId,
        properties: { "background-color": color },
      },
    ]);
  }

  return {
    openSort,
    openTextToTable,
    openTableToText,
    openDateTime,
    insertBlankPage,
    runTableAction,
    setTableCellShading,
    dispose: close,
  };
}

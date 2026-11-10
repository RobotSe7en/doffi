import { DocumentError } from "@doffi/core";

/** Transient command availability belongs to the mounted view, not the source. */
export interface EditorCommandState {
  readonly id: string;
  readonly label: string;
  readonly enabled: boolean;
  readonly active?: boolean;
}
export interface EditorCommandTarget {
  commands(): readonly EditorCommandState[];
  executeCommand(id: string): Promise<void>;
}
export interface EditorCommandRegistry extends EditorCommandTarget {
  register(id: string, label: string, run: () => unknown, enabled: () => boolean, active?: () => boolean): void;
  notify(): void;
}
export function createEditorCommands(alive: () => boolean, changed?: (commands: readonly EditorCommandState[]) => void, available: () => boolean = () => true): EditorCommandRegistry {
  const entries = new Map<string, { label: string; enabled(): boolean; active?: () => boolean; run(): unknown }>();
  let published = "";
  const commands = (): readonly EditorCommandState[] => [...entries].map(([id, entry]) => ({ id, label: entry.label, enabled: alive() && available() && entry.enabled(), ...(entry.active ? { active: entry.active() } : {}) }));
  return {
    register(id: string, label: string, run: () => unknown, enabled: () => boolean, active?: () => boolean): void {
      if (entries.has(id)) throw new DocumentError("InvalidOperation", "Duplicate editor command: " + id);
      entries.set(id, { label, run, enabled, ...(active ? { active } : {}) });
    },
    commands,
    async executeCommand(id: string): Promise<void> {
      if (!alive()) throw new DocumentError("InvalidOperation", "The editor has been disposed.");
      const entry = entries.get(id);
      if (!entry || !available() || !entry.enabled()) throw new DocumentError("InvalidOperation", "当前不可执行此命令：" + id);
      await entry.run();
    },
    notify(): void {
      if (!alive()) return;
      const current = commands(), key = JSON.stringify(current); if (key === published) return; published = key;
      try { changed?.(current); } catch { /* The host owns its UI state. */ }
    },
  };
}

/** IDs remain stable when the native view's visible labels change. */
export const workbookCommandIds: Readonly<Record<string, string>> = {
  "确认输入": "acceptInput", "取消输入": "cancelInput", "撤销": "undo", "重做": "redo", "重新计算": "recalculate",
  "向下填充": "fillDown", "向右填充": "fillRight", "数列填充": "fillSeries", "合并单元格": "mergeCells", "合并并清除覆盖内容": "mergeClear", "取消合并": "unmergeCells",
  "冻结选区之前": "freezePanes", "取消冻结": "unfreezePanes", "插入选区行": "insertRows", "删除选区行": "deleteRows", "插入选区列": "insertColumns", "删除选区列": "deleteColumns",
  "升序排序": "sortAscending", "降序排序": "sortDescending", "复制选区": "copy", "剪切选区": "cut", "粘贴选区": "paste", "取消操作": "cancelOperation",
};
export const presentationCommandIds: Readonly<Record<string, string>> = {
  "编辑文字": "editText", "编辑备注": "editNotes", "完成输入": "finishInput", "加粗": "bold", "常规字重": "normalWeight", "斜体": "italic", "下划线": "underline", "段落左对齐": "textAlignLeft", "段落居中": "textAlignCenter", "段落右对齐": "textAlignRight", "项目符号": "bullets", "编号列表": "numbering", "无列表": "noList", "插入图表": "insertChart",
  "撤销": "undo", "重做": "redo", "新增幻灯片": "newSlide", "复制幻灯片": "duplicateSlide", "删除幻灯片": "deleteSlide", "幻灯片上移": "slideUp", "幻灯片下移": "slideDown",
  "文本框": "insertTextBox", "形状": "insertShape", "表格": "insertTable", "连接线": "insertConnector", "插入图片": "insertImage", "复制对象": "duplicateObject", "删除对象": "deleteObject",
  "组合": "groupObjects", "取消组合": "ungroupObjects", "选择上层组合": "selectParentGroup", "置顶": "bringToFront", "置底": "sendToBack", "上移一层": "bringForward", "下移一层": "sendBackward",
  "左对齐": "alignLeft", "水平居中": "alignCenter", "右对齐": "alignRight", "顶端对齐": "alignTop", "垂直居中": "alignMiddle", "底端对齐": "alignBottom",
  "水平等分布": "distributeHorizontal", "垂直等分布": "distributeVertical", "水平翻转": "flipHorizontal", "垂直翻转": "flipVertical",
};

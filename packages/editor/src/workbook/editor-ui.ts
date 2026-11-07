interface WorkbookEditorUi {
  readonly element: HTMLDivElement; readonly tools: HTMLDivElement; readonly input: HTMLTextAreaElement; readonly address: HTMLInputElement;
  readonly mode: HTMLSelectElement; readonly preview: HTMLDivElement; readonly status: HTMLParagraphElement; readonly error: HTMLParagraphElement;
  readonly find: HTMLInputElement; readonly replacement: HTMLInputElement; readonly findField: HTMLSelectElement; readonly findScope: HTMLSelectElement;
  readonly matchCase: HTMLInputElement; readonly entireCell: HTMLInputElement; readonly replacementMode: HTMLSelectElement; readonly searchTools: HTMLDivElement;
}
export function createWorkbookEditorUi(container: HTMLElement, title: string): WorkbookEditorUi {
  const document = container.ownerDocument, element = document.createElement("div"), shadow = element.attachShadow({ mode: "closed" });
  const style = document.createElement("style"), heading = document.createElement("h2"), tools = document.createElement("div"), bar = document.createElement("div");
  const input = document.createElement("textarea"), address = document.createElement("input"), mode = document.createElement("select"), preview = document.createElement("div");
  const status = document.createElement("p"), error = document.createElement("p");
  const search = document.createElement("div"), searchTools = document.createElement("div"), find = document.createElement("input"), replacement = document.createElement("input");
  const findField = document.createElement("select"), findScope = document.createElement("select"), replacementMode = document.createElement("select"), matchCase = document.createElement("input"), entireCell = document.createElement("input");
  style.textContent = ":host{display:block;color:#172638;font:14px/1.5 system-ui,sans-serif}*{box-sizing:border-box}h2{margin:0 0 8px;font-size:17px}.tools,.bar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:8px}button,input,textarea,select{font:inherit;border:1px solid #bdc9d5;border-radius:4px;background:#fff;color:inherit;padding:5px 8px}button{cursor:pointer}button:disabled{opacity:.5;cursor:default}textarea{flex:1;min-width:160px;resize:vertical;min-height:35px;max-height:160px}input{width:120px}p{margin:4px 0;white-space:pre-wrap}.error{color:#a12626}.preview{position:relative}";
  heading.textContent = title; tools.className = "tools"; tools.setAttribute("role", "toolbar"); tools.setAttribute("aria-label", "工作簿编辑"); bar.className = "bar";
  address.setAttribute("aria-label", "选择单元格或区域"); address.placeholder = "A1 或 A1:B10";
  input.setAttribute("aria-label", "单元格或公式输入"); input.rows = 1; input.spellcheck = false; input.disabled = true;
  mode.setAttribute("aria-label", "输入类型");
  search.className = "bar"; search.setAttribute("role", "search"); search.setAttribute("aria-label", "工作簿查找替换"); searchTools.className = "tools";
  find.setAttribute("aria-label", "查找内容"); find.placeholder = "查找内容"; find.maxLength = 32767;
  replacement.setAttribute("aria-label", "替换为"); replacement.placeholder = "替换为"; replacement.maxLength = 32767;
  const choices = (select: HTMLSelectElement, label: string, items: readonly (readonly [string, string])[]) => {
    select.setAttribute("aria-label", label); for (const [value, label] of items) { const option = document.createElement("option"); option.value = value; option.textContent = label; select.append(option); }
  };
  choices(findField, "查找字段", [["source", "源内容"], ["formula", "公式"], ["display", "当前显示值"]]);
  choices(findScope, "查找范围", [["sheet", "当前工作表"], ["selection", "选中区域"], ["workbook", "整个工作簿"]]);
  choices(replacementMode, "替换类型", [["preserve", "保留原类型"], ["auto", "自动识别新类型"], ["text", "常量转为文本"]]);
  for (const [input, text] of [[matchCase, "区分大小写"], [entireCell, "匹配整个单元格"]] as const) {
    const label = document.createElement("label"); input.type = "checkbox"; input.style.width = "auto"; label.append(input, text); search.append(label);
  }
  search.prepend(find, replacement, findField, findScope, replacementMode); search.append(searchTools);
  for (const [value, label] of [["auto", "自动识别"], ["text", "文本"], ["number", "数字"]]) {
    const option = document.createElement("option"); option.value = value!; option.textContent = label!; mode.append(option);
  }
  status.setAttribute("role", "status"); status.setAttribute("aria-label", "工作簿编辑状态"); status.setAttribute("aria-live", "polite"); status.textContent = "正在打开工作簿…";
  error.setAttribute("role", "alert"); error.className = "error"; error.hidden = true; preview.className = "preview";
  bar.append(address, input, mode); shadow.append(style, heading, tools, bar, search, status, error, preview); container.append(element);
  return { element, tools, input, address, mode, preview, status, error, find, replacement, findField, findScope, matchCase, entireCell, replacementMode, searchTools };
}

/** Shared editable Office lists for the ordinary DOM font dialogs. */
export function attachOfficeFieldList(input: HTMLInputElement, values: readonly string[], signal: AbortSignal, fontPreview = false): void {
  const document = input.ownerDocument, view = document.defaultView!, holder = document.createElement("span"), caret = document.createElement("button"), popup = document.createElement("div");
  const label = input.getAttribute("aria-label") || "列表";
  holder.className = "dialog-field-combo"; popup.className = "dialog-field-list"; popup.hidden = true;
  popup.id = "office-list-" + crypto.randomUUID(); popup.setAttribute("role", "listbox"); popup.setAttribute("aria-label", label + "选项");
  input.removeAttribute("list"); input.setAttribute("role", "combobox"); input.setAttribute("aria-autocomplete", "list"); input.setAttribute("aria-expanded", "false"); input.setAttribute("aria-controls", popup.id);
  caret.type = "button"; caret.textContent = "⌄"; caret.tabIndex = -1; caret.setAttribute("aria-label", "展开" + label); caret.setAttribute("aria-expanded", "false");
  input.replaceWith(holder); holder.append(input, caret, popup);
  let active = -1, choices: string[] = [];
  const close = () => { popup.hidden = true; input.setAttribute("aria-expanded", "false"); caret.setAttribute("aria-expanded", "false"); input.removeAttribute("aria-activedescendant"); };
  const select = (value: string) => { input.value = value; close(); input.dispatchEvent(new view.Event("input", { bubbles: true })); input.dispatchEvent(new view.Event("change", { bubbles: true })); close(); input.focus(); };
  const mark = () => { [...popup.children].forEach((button, index) => button.setAttribute("aria-selected", String(index === active))); const button = popup.children[active] as HTMLElement | undefined; if (button) { input.setAttribute("aria-activedescendant", button.id); button.scrollIntoView({ block: "nearest" }); } };
  const open = (filter = "") => {
    choices = values.filter(value => value.toLowerCase().includes(filter.toLowerCase())); active = choices.indexOf(input.value); popup.replaceChildren();
    for (const [index, value] of choices.entries()) { const button = document.createElement("button"); button.type = "button"; button.tabIndex = -1; button.id = popup.id + "-" + index; button.setAttribute("role", "option"); button.textContent = value;
      if (fontPreview) button.style.fontFamily = JSON.stringify(value) + ",sans-serif";
      button.addEventListener("mousedown", event => event.preventDefault()); button.addEventListener("click", () => select(value)); popup.append(button); }
    const rect = holder.getBoundingClientRect(), width = Math.min(view.innerWidth - 24, fontPreview ? Math.max(rect.width, 250) : Math.max(rect.width, 90));
    popup.style.width = width + "px"; popup.style.left = Math.max(12, Math.min(rect.left, view.innerWidth - width - 12)) + "px"; popup.style.top = rect.bottom + "px"; popup.style.maxHeight = Math.min(260, view.innerHeight - rect.bottom - 12) + "px";
    popup.hidden = false; input.setAttribute("aria-expanded", "true"); caret.setAttribute("aria-expanded", "true"); mark();
  };
  caret.addEventListener("mousedown", event => event.preventDefault(), { signal }); caret.addEventListener("click", () => { if (popup.hidden) open(); else close(); input.focus(); }, { signal });
  input.addEventListener("keydown", event => {
    if (event.isComposing) return;
    if (event.key === "Escape" && !popup.hidden) { event.preventDefault(); event.stopPropagation(); close(); }
    else if (event.key === "Tab") close();
    else if (["ArrowDown", "ArrowUp"].includes(event.key)) { event.preventDefault(); if (popup.hidden) open(); active = Math.max(0, Math.min(choices.length - 1, active + (event.key === "ArrowDown" ? 1 : -1))); mark(); }
    else if (event.key === "Enter" && !popup.hidden && choices[active]) { event.preventDefault(); event.stopPropagation(); select(choices[active]!); }
  }, { signal });
  input.addEventListener("input", event => { if (event.isTrusted) open(input.value); }, { signal });
  input.getRootNode().addEventListener("pointerdown", event => { if (!holder.contains(event.target as Node)) close(); }, { signal });
  view.addEventListener("resize", close, { signal });
}

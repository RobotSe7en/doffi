/** Native modal semantics and shared Office chrome, independent of React. */
export const officeDialogStyles = `
.office-dialog{--office-accent:#185abd;box-sizing:border-box;width:660px;max-width:calc(100vw - 32px);max-height:calc(100dvh - 32px);border:1px solid #c7c7c7;border-radius:10px;padding:0;color:#252525;background:#fff;box-shadow:0 16px 60px #0004;font:13px/1.45 -apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif;overflow:auto}
.office-dialog::backdrop{background:#0002}.office-dialog *{box-sizing:border-box}.office-dialog [hidden]{display:none!important}
.office-dialog .dialog-title{padding:6px 20px;margin:0;font-size:13px;font-weight:600;border-bottom:1px solid #eee}
.office-dialog .dialog-tabs{display:flex;justify-content:center;padding:10px 14px 0;margin-bottom:-10px;gap:0;background:#fff;position:sticky;top:0;z-index:1}
.office-dialog .dialog-tabs button{padding:3px 15px;border:1px solid #dedede;background:#ededed;color:#333;border-radius:0;font:inherit;min-height:24px}
.office-dialog .dialog-tabs button:first-child{border-radius:6px 0 0 6px}.office-dialog .dialog-tabs button:last-child{border-radius:0 6px 6px 0}
.office-dialog .dialog-tabs button[aria-selected=true]{background:white;box-shadow:0 1px 3px #0002}
.office-dialog .dialog-content{margin:0 20px 12px;padding:25px 18px 18px;border:1px solid #e6e6e6;border-radius:4px;background:#f8f8f8;min-height:420px}
.office-dialog label{display:flex;flex-direction:column;gap:5px;color:#333;font:inherit}.office-dialog input,.office-dialog select,.office-dialog textarea{font:inherit;border:1px solid #d5d5d5;border-radius:5px;background:white;color:#222;padding:3px 7px;min-height:25px;max-width:100%}
.office-dialog select{width:100%}.office-dialog input:not([type=number]):not([type=checkbox]):not([type=color]){width:100%}
.office-dialog input[type=number]{width:86px}.office-dialog input[type=checkbox]{min-height:0;width:14px;accent-color:var(--office-accent);margin:0 6px 0 0}
.office-dialog label:has(input[type=checkbox]){flex-direction:row;align-items:center}.office-dialog input:focus,.office-dialog select:focus{outline:2px solid #8daed9;outline-offset:1px}
.office-dialog fieldset{margin:0 0 18px;padding:15px 0 0;border:0;border-top:1px solid #dedede;min-width:0;display:grid;grid-template-columns:1fr 1fr;gap:12px 20px}
.office-dialog legend{padding:0 8px 0 0;font-weight:600}.office-dialog .dialog-grid{display:grid;grid-template-columns:2fr 1.1fr .8fr;gap:12px 16px}
.office-dialog .dialog-grid .wide{grid-column:1/-1}.office-dialog .dialog-footer{display:flex;align-items:center;justify-content:flex-end;gap:12px;padding:3px 20px 18px}
.office-dialog .dialog-footer button{padding:4px 17px;border:1px solid transparent;border-radius:6px;background:#efefef;color:#333;font:inherit;min-width:62px}
.office-dialog .dialog-footer button.primary{color:#fff;background:var(--office-accent)}.office-dialog button:disabled{opacity:.4}.office-dialog .dialog-preview{display:grid;place-items:center;min-height:82px;background:white;border:1px solid #888;margin:10px 35px 0;font-size:22px;color:#111}
.office-dialog .dialog-options{display:grid;grid-template-columns:1fr 1fr;gap:8px 30px}.office-dialog .dialog-stack{display:flex;flex-direction:column;gap:12px}.office-dialog .dialog-list{height:155px;overflow:auto;border:1px solid #ccc;background:#fff}.office-dialog .dialog-list button{display:block;width:100%;border:0;border-radius:0;text-align:left;background:white;color:#222;font:inherit;padding:2px 8px}.office-dialog .dialog-list button:hover,.office-dialog .dialog-list button[aria-selected=true]{background:#dce8f7}
.office-dialog .dialog-field-combo{display:flex;position:relative;min-width:0;border:1px solid #d5d5d5;border-radius:5px;background:white;height:25px}.office-dialog .dialog-field-combo>input{border:0;min-width:0;width:100%;min-height:0}.office-dialog .dialog-field-combo>button{border:0;border-radius:0 4px 4px 0;background:#eee;color:#555;padding:0 5px;font:inherit;cursor:pointer}.office-dialog .dialog-field-list{position:fixed;z-index:100;overflow:auto;border:1px solid #bbb;border-radius:5px;padding:4px 0;background:white;box-shadow:0 7px 22px #0003}.office-dialog .dialog-field-list>button{display:block;width:100%;border:0;border-radius:0;background:white;color:#222;padding:3px 12px;text-align:left;font:inherit;font-size:13px;line-height:1.4}.office-dialog .dialog-field-list>button[aria-selected=true],.office-dialog .dialog-field-list>button:hover{background:#dce8f8}
@media(max-width:600px){.office-dialog .dialog-content{margin-inline:10px;padding:18px 12px}.office-dialog .dialog-grid{grid-template-columns:1fr 1fr}.office-dialog .dialog-options{gap:10px}.office-dialog .dialog-preview{margin-inline:0}}
`;

export function createOfficeDialog(document: Document, title: string): {
  dialog: HTMLDialogElement; content: HTMLDivElement; footer: HTMLDivElement;
  tab(label: string): HTMLDivElement; show(): void; close(): void; error(reason: unknown): void;
} {
  const dialog = document.createElement("dialog"), style = document.createElement("style"), heading = document.createElement("h2"), tabs = document.createElement("div"), content = document.createElement("div"), footer = document.createElement("div");
  dialog.className = "office-dialog"; dialog.setAttribute("aria-label", title); heading.textContent = title; heading.className = "dialog-title";
  style.textContent = officeDialogStyles; tabs.className = "dialog-tabs"; tabs.setAttribute("role", "tablist"); tabs.setAttribute("aria-label", title + "选项");
  content.className = "dialog-content"; footer.className = "dialog-footer"; dialog.append(style, heading, tabs, content, footer);
  const error = document.createElement("p"); error.setAttribute("role", "alert"); error.hidden = true; error.style.cssText = "color:#a12626;margin:0 20px 10px;white-space:pre-wrap"; dialog.insertBefore(error, footer);
  dialog.addEventListener("keydown", event => {
    if (event.key !== "Enter" || event.isComposing || event.defaultPrevented || (event.target as Element).matches("textarea,select,button")) return;
    const accept = footer.querySelector<HTMLButtonElement>("button.primary:not(:disabled)"); if (accept) { event.preventDefault(); accept.click(); }
  });
  const choices: { button: HTMLButtonElement; panel: HTMLDivElement }[] = [];
  const select = (index: number) => { for (const [i, item] of choices.entries()) { item.button.setAttribute("aria-selected", String(i === index)); item.button.tabIndex = i === index ? 0 : -1; item.panel.hidden = i !== index; } };
  return { dialog, content, footer,
    tab(label) {
      const button = document.createElement("button"), panel = document.createElement("div"); button.type = "button"; button.textContent = label; button.setAttribute("role", "tab"); panel.setAttribute("role", "tabpanel"); panel.setAttribute("aria-label", label);
      const index = choices.length; choices.push({ button, panel }); tabs.append(button); content.append(panel); select(0);
      button.addEventListener("click", () => select(index));
      button.addEventListener("keydown", event => { if (!["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) return; event.preventDefault(); const next = event.key === "Home" ? 0 : event.key === "End" ? choices.length - 1 : (index + (event.key === "ArrowRight" ? 1 : -1) + choices.length) % choices.length; select(next); choices[next]!.button.focus(); });
      return panel;
    },
    show() { error.hidden = true; dialog.showModal(); }, close() { dialog.close(); },
    error(reason) { error.textContent = reason instanceof Error ? reason.message : String(reason); error.hidden = false; },
  };
}

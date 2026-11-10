import { DocumentError } from "@doffi/core";

/** The native copy/cut/paste listeners still serialize and validate source data. */
export async function executeClipboard(root: HTMLElement, focus: () => void, action: "copy" | "cut" | "paste", current: () => boolean): Promise<void> {
  if (!current() || !root.isConnected) throw new DocumentError("InvalidOperation", "当前输入已失效。");
  const document = root.ownerDocument, view = document.defaultView!;
  focus();
  if (action !== "paste") {
    if (!document.execCommand(action)) throw new DocumentError("InvalidOperation", "浏览器未允许剪贴板操作，请使用键盘快捷键。");
    return;
  }
  if (!view.navigator.clipboard?.read) throw new DocumentError("UnsupportedFeature", "浏览器不允许读取系统剪贴板，请使用 Ctrl/⌘+V。");
  const items = await view.navigator.clipboard.read(), data = new view.DataTransfer();
  let bytes = 0;
  for (const type of ["text/html", "text/plain"]) {
    const item = items.find(item => item.types.includes(type)); if (!item) continue;
    const blob = await item.getType(type); bytes += blob.size;
    if (bytes > 1024 * 1024) throw new DocumentError("ResourceLimitExceeded", "剪贴板内容超过 1 MiB。");
    data.setData(type, await blob.text());
  }
  if (!current() || !root.isConnected) throw new DocumentError("RevisionConflict", "读取剪贴板期间文档或选区已变化。");
  focus(); root.dispatchEvent(new view.ClipboardEvent("paste", { bubbles: true, cancelable: true, clipboardData: data }));
}

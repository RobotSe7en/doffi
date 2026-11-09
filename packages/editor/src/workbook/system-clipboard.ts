import { DocumentError } from "@doffi/core";
import type { DocumentSelection } from "@doffi/core";
import type { WorkbookClipboardData } from "./clipboard-data.js";

/** Keep only the current bounded selection, so native ClipboardEvent handlers
 * can set both formats synchronously even when the source belongs to a Worker. */
export class WorkbookSystemClipboard {
  private current: { key: string; abort: AbortController; work: Promise<void>; data?: WorkbookClipboardData; error?: unknown } | undefined;
  constructor(private readonly read: (selection: DocumentSelection, signal: AbortSignal) => Promise<WorkbookClipboardData>) {}
  prepare(selection: DocumentSelection): void {
    const key = JSON.stringify(selection); if (this.current?.key === key) return;
    this.current?.abort.abort(); const abort = new AbortController();
    const entry: NonNullable<WorkbookSystemClipboard["current"]> = { key, abort, work: Promise.resolve() };
    this.current = entry;
    entry.work = Promise.resolve().then(() => this.read(selection, abort.signal)).then(data => { if (!abort.signal.aborted) entry.data = data; }, error => { entry.error = error; });
  }
  async settled(): Promise<void> { for (;;) { const entry = this.current; await entry?.work; if (entry === this.current) return; } }
  data(selection: DocumentSelection): WorkbookClipboardData {
    const entry = this.current;
    if (!entry || entry.key !== JSON.stringify(selection) || entry.abort.signal.aborted) throw new DocumentError("RevisionConflict", "Clipboard selection belongs to a previous source.");
    if (entry.error) throw entry.error;
    if (!entry.data) throw new DocumentError("InvalidOperation", "剪贴板内容正在准备，请稍后重试复制。");
    return entry.data;
  }
  dispose(): void { this.current?.abort.abort(); this.current = undefined; }
}

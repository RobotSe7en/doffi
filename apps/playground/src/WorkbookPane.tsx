import { useEffect, useRef } from "react";
import { DocumentError } from "@doffi/core";
import { mountWorkbookEditor } from "@doffi/editor";
import type { WorkbookRemoteEditorView } from "@doffi/editor";
import type { HostedWorkbook } from "./open-workbook.js";

interface Mounted { readonly document: HostedWorkbook; readonly slot: HTMLDivElement; readonly editor: WorkbookRemoteEditorView; dispose(): void }
export function WorkbookPane({ document: hosted, preparing, activation, onReady, onUpdate, onError }: {
  document: HostedWorkbook; preparing: boolean; activation?: object | undefined;
  onReady(document: HostedWorkbook): void; onUpdate(document: HostedWorkbook): void; onError(document: HostedWorkbook, error: Error): void;
}) {
  const container = useRef<HTMLDivElement>(null), active = useRef<Mounted>(undefined), pending = useRef<Mounted>(undefined), changing = useRef(false);
  const callbacks = useRef({ onReady, onUpdate, onError }); callbacks.current = { onReady, onUpdate, onError };
  const hidden = useRef(preparing); hidden.current = preparing;
  const resize = useRef<() => Promise<void>>(async () => {});
  useEffect(() => {
    const element = container.current!;
    resize.current = async () => {
      const current = active.current; if (!current || changing.current || hidden.current || !element.clientWidth) return;
      try {
        const widthPt = Math.max(180, element.clientWidth * .75 - 49.5), viewport = current.editor.layout().viewport;
        if (Math.abs(viewport.widthPt - widthPt) > .5) await current.editor.setViewport({ ...viewport, widthPt });
      } catch (error) { if (active.current === current && !changing.current) callbacks.current.onError(current.document, error instanceof Error ? error : new Error("工作簿尺寸更新失败。")); }
    };
    const observer = new ResizeObserver(() => { void resize.current(); });
    observer.observe(element);
    return () => { observer.disconnect(); pending.current?.dispose(); active.current?.dispose(); pending.current = active.current = undefined; };
  }, []);
  useEffect(() => { if (!preparing) void resize.current(); }, [preparing]);
  useEffect(() => {
    let cancelled = false; changing.current = true;
    if (active.current?.document === hosted) {
      // An accepted current-revision canvas can be restored immediately.
      // A viewport job started before hiding must not gate its reactivation.
      const editor = active.current.editor;
      const accepted = editor.layout().revision === hosted.session.state().revision ? Promise.resolve() : editor.settled();
      void accepted.then(() => { if (!cancelled) { changing.current = false; callbacks.current.onReady(hosted); void resize.current(); } }).catch(error => { if (!cancelled) { changing.current = false; callbacks.current.onError(hosted, error instanceof Error ? error : new Error("工作簿显示失败。")); } });
      return () => { cancelled = true; };
    }
    const slot = container.current!.ownerDocument.createElement("div");
    slot.style.cssText = "position:absolute;inset:0;transform:translateX(-200vw);opacity:0;visibility:hidden;pointer-events:none"; slot.inert = true; slot.setAttribute("aria-hidden", "true"); container.current!.append(slot);
    let mounted: Mounted | undefined;
    try {
      const editor = mountWorkbookEditor(slot, hosted.session, { widthPt: Math.max(180, container.current!.clientWidth * .75 - 49.5), heightPt: 450,
        prepareOptions: () => ({ calculationContext: { now: { unixMilliseconds: Date.now(), utcOffsetMinutes: -new Date().getTimezoneOffset() } } }),
        confirmValidationWarnings(confirmation) {
          return window.confirm(confirmation.warnings.map(warning => `${warning.message}（${warning.sheetId}，行 ${warning.row + 1}，列 ${warning.column + 1}）`).join("\n") + "\n是否仍应用本次更改？");
        },
        onViewChange(snapshot) { if (mounted && (active.current === mounted || pending.current === mounted)) { hosted.acceptView(snapshot); if (active.current === mounted) callbacks.current.onUpdate(hosted); } },
        onPendingChange(value) { if (mounted && (active.current === mounted || pending.current === mounted)) hosted.setPendingInput(value); },
        onError(error) { if (mounted && active.current === mounted && !changing.current) callbacks.current.onError(hosted, error); },
      });
      let released = false;
      mounted = { document: hosted, slot, editor, dispose() { if (released) return; released = true; editor.dispose(); hosted.setPendingInput(false); slot.remove(); } };
      const current = mounted; pending.current?.dispose(); pending.current = current;
      void editor.ready.then(() => {
        if (cancelled || pending.current !== current) return;
        if (active.current?.editor.hasPendingInput()) throw new DocumentError("InvalidOperation", "请先确认或取消当前单元格输入或操作，再切换文档。");
        const previous = active.current; active.current = current; pending.current = undefined;
        slot.style.cssText = "position:relative"; slot.inert = false; slot.removeAttribute("aria-hidden"); previous?.dispose();
        changing.current = false; callbacks.current.onReady(hosted); void resize.current();
      }).catch(error => {
        if (cancelled) return; current.dispose(); if (pending.current === current) pending.current = undefined;
        changing.current = false; callbacks.current.onError(hosted, error instanceof Error ? error : new Error("工作簿显示失败。"));
      });
    } catch (error) { slot.remove(); changing.current = false; callbacks.current.onError(hosted, error instanceof Error ? error : new Error("工作簿显示失败。")); }
    return () => { cancelled = true; if (mounted && active.current !== mounted) { mounted.dispose(); if (pending.current === mounted) pending.current = undefined; } };
  }, [hosted, activation]);
  return <div className={"workbook-pane" + (preparing ? " preparing" : "")} aria-label="Excel 工作簿编辑" aria-hidden={preparing || undefined} inert={preparing} ref={container} />;
}

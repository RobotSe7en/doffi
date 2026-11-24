import { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import type { CSSProperties, KeyboardEvent } from "react";
import type { DocumentSession } from "@doffi/core";
import { mountDocumentPreview } from "@doffi/editor";
import { wordWebFonts } from "./word-web-fonts.js";

export function SlideShow({ session, startId, onClose }: {
  session: DocumentSession;
  startId: string;
  onClose(): void;
}) {
  const presentation = useMemo(() => session.snapshot().metadata.presentation!, [session]);
  const slides = presentation.slides.filter(slide => !slide.hidden || slide.slideId === startId);
  const [index, setIndex] = useState(() => Math.max(0, slides.findIndex(slide => slide.slideId === startId)));
  const [frame, setFrame] = useState<HTMLIFrameElement>();
  const [error, setError] = useState("");
  const dialog = useRef<HTMLDialogElement>(null), available = useRef<HTMLDivElement>(null);
  const surface = useRef<HTMLDivElement>(null);
  const width = presentation.size.width * 4 / 3, height = presentation.size.height * 4 / 3;
  const slide = slides[index]!;

  useLayoutEffect(() => {
    const modal = dialog.current!, viewport = available.current!, slot = surface.current!;
    modal.showModal();
    viewport.focus({ preventScroll: true });
    const fit = () => {
      const scale = Math.min(viewport.clientWidth / width, viewport.clientHeight / height);
      slot.style.width = width * scale + "px";
      slot.style.height = height * scale + "px";
      slot.style.setProperty("--slide-scale", String(scale));
    };
    fit();
    const observer = new ResizeObserver(fit);
    observer.observe(viewport);
    return () => {
      observer.disconnect();
      modal.close();
    };
  }, [width, height]);

  useEffect(() => {
    const failed = (reason: unknown) => setError(reason instanceof Error ? reason.message : "无法加载幻灯片。");
    let preview: ReturnType<typeof mountDocumentPreview>;
    try {
      preview = mountDocumentPreview(surface.current!, session, { title: "放映画布", wordWebFonts });
    } catch (reason) {
      failed(reason);
      return;
    }
    const iframe = preview.element;
    // Playback is view-only. Keys remain with the dialog, and hyperlinks
    // cannot navigate away from the current slide.
    iframe.tabIndex = -1;
    iframe.style.setProperty("width", width + "px", "important");
    let active = true;
    void preview.ready.then(() => {
      if (active) setFrame(iframe);
    }).catch((reason) => {
      if (active) failed(reason);
    });
    return () => {
      active = false;
      preview.dispose();
    };
  }, [session, width]);

  useLayoutEffect(() => {
    const document = frame?.contentDocument;
    if (!frame || !document) return;
    const style = document.createElement("style");
    style.textContent = `
      html,body{padding:0!important;margin:0!important;overflow:hidden!important}
      article[data-office-kind="presentation"]{padding:0!important;margin:0!important}
      article>[data-office-kind="slide"]:not([data-node-id="${CSS.escape(slide.rootNodeId)}"]),
      .sdk-ppt-notes,.sdk-ppt-hidden-label{display:none!important}
      article>[data-office-kind="slide"]{margin:0!important}
    `;
    document.head.append(style);
    return () => style.remove();
  }, [frame, slide.rootNodeId]);

  const move = (delta: number) => setIndex(current => Math.min(slides.length - 1, Math.max(0, current + delta)));
  function key(event: KeyboardEvent<HTMLDialogElement>) {
    if (event.metaKey || event.ctrlKey || event.altKey) return;
    if (event.key === "Home") setIndex(0);
    else if (event.key === "End") setIndex(slides.length - 1);
    else if (["ArrowLeft", "ArrowUp", "PageUp"].includes(event.key)) move(-1);
    else if (["ArrowRight", "ArrowDown", "PageDown"].includes(event.key)
      || (event.key === " " && !(event.target instanceof HTMLButtonElement))) move(1);
    else return;
    event.preventDefault();
  }
  return <dialog ref={dialog} className="slide-show" aria-label="幻灯片放映"
    onKeyDown={key} onCancel={event => { event.preventDefault(); onClose(); }}>
    <header className="slide-show-header">
      <span aria-live="polite">{index + 1} / {slides.length} · {slide.name ?? "幻灯片"}</span>
      <button type="button" onClick={onClose}>退出放映</button>
    </header>
    <div className="slide-show-picture" ref={available} tabIndex={0} aria-label="幻灯片画面" aria-busy={!frame && !error}>
      <div ref={surface} className="slide-show-surface" data-ready={Boolean(frame)} aria-hidden={!frame}
        style={{ "--slide-height": height + "px" } as CSSProperties} />
      {!frame && <p role={error ? "alert" : "status"}>{error || "正在加载幻灯片…"}</p>}
    </div>
    <footer className="slide-show-actions">
      <button type="button" disabled={index === 0} onClick={() => move(-1)}>上一张</button>
      <button type="button" disabled={index + 1 === slides.length} onClick={() => move(1)}>下一张</button>
    </footer>
  </dialog>;
}

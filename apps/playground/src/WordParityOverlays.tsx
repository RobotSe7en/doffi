import { useEffect, useRef, useState } from "react";
import type { ReactElement } from "react";
import { compareWordDocuments, openDocument } from "@doffi/core";
import type { ContentDocumentModel, DocumentSnapshot } from "@doffi/core";
import { importDocx } from "@doffi/office";
import { mountDocumentPreview } from "@doffi/editor";
import { OfficeDialog } from "@doffi/editor/react";
import { wordWebFonts } from "./word-web-fonts.js";

/** Print a measured Word/PPT print projection through the native browser print dialog. */
export function printHtmlDocument(printDoc: Document): void {
  const iframe = document.createElement("iframe");
  iframe.setAttribute("aria-hidden", "true");
  iframe.style.cssText =
    "position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden;pointer-events:none";
  document.body.append(iframe);
  const frameDoc = iframe.contentDocument;
  const frameWin = iframe.contentWindow;
  if (!frameDoc || !frameWin) {
    iframe.remove();
    throw new Error("无法创建打印视图。");
  }
  frameDoc.open();
  frameDoc.write("<!doctype html>" + printDoc.documentElement.outerHTML);
  frameDoc.close();
  const cleanup = () => {
    window.setTimeout(() => iframe.remove(), 250);
  };
  frameWin.addEventListener("afterprint", cleanup, { once: true });
  window.setTimeout(() => {
    try {
      frameWin.focus();
      frameWin.print();
    } finally {
      window.setTimeout(() => {
        if (iframe.isConnected) iframe.remove();
      }, 60_000);
    }
  }, 80);
}

/** Immersive Reader overlay matching Word's Column Width, Page Color, Text Spacing, and Line Focus. */
export function ImmersiveReaderOverlay({
  snapshot,
  title,
  onClose,
}: {
  snapshot: DocumentSnapshot;
  title: string;
  onClose(): void;
}): ReactElement {
  const [width, setWidth] = useState<"narrow" | "moderate" | "wide">("moderate");
  const [theme, setTheme] = useState<"light" | "sepia" | "dark">("sepia");
  const [spacing, setSpacing] = useState<number>(1.85);
  const [focusLines, setFocusLines] = useState<0 | 1 | 3 | 5>(0);
  const [focusIndex, setFocusIndex] = useState(0);

  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") onClose();
      else if (focusLines > 0 && (event.key === "ArrowDown" || event.key === "j")) {
        event.preventDefault();
        setFocusIndex((i) => i + 1);
      } else if (focusLines > 0 && (event.key === "ArrowUp" || event.key === "k")) {
        event.preventDefault();
        setFocusIndex((i) => Math.max(0, i - 1));
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose, focusLines]);

  const html = "html" in snapshot ? snapshot.html : "";
  const parsed = new DOMParser().parseFromString(html, "text/html");
  const main =
    parsed.querySelector('[data-office-story="main"]') ??
    parsed.querySelector('[data-office-kind="document"]');
  const blocks = [...(main?.querySelectorAll("p,h1,h2,h3,h4,h5,h6,li") ?? [])]
    .map((el) => ({ tag: el.tagName.toLowerCase(), text: el.textContent?.trim() ?? "" }))
    .filter((b) => b.text.length > 0);

  const clampedFocus = blocks.length ? Math.min(focusIndex, blocks.length - 1) : 0;
  const maxWidth = width === "narrow" ? 540 : width === "moderate" ? 720 : 940;
  const palette =
    theme === "dark"
      ? { bg: "#1c1f24", fg: "#e6ebf2", bar: "#252a32", border: "#343b46" }
      : theme === "sepia"
        ? { bg: "#f6efe0", fg: "#3a2f22", bar: "#ece2cc", border: "#d8cbb0" }
        : { bg: "#ffffff", fg: "#1e242b", bar: "#f5f7fa", border: "#dce2eb" };

  return (
    <div
      role="dialog"
      aria-label="沉浸式阅读器"
      style={{
        position: "fixed",
        inset: 0,
        zIndex: 120,
        background: palette.bg,
        color: palette.fg,
        display: "flex",
        flexDirection: "column",
        fontFamily: '"PingFang SC","Microsoft YaHei",system-ui,sans-serif',
      }}
    >
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 12,
          flexWrap: "wrap",
          padding: "10px 20px",
          background: palette.bar,
          borderBottom: `1px solid ${palette.border}`,
          fontSize: 12,
        }}
      >
        <strong style={{ marginRight: "auto" }}>沉浸式阅读器 · {title}</strong>
        <label>
          列宽{" "}
          <select aria-label="阅读列宽" value={width} onChange={(e) => setWidth(e.currentTarget.value as "narrow" | "moderate" | "wide")}>
            <option value="narrow">窄</option>
            <option value="moderate">适中</option>
            <option value="wide">宽</option>
          </select>
        </label>
        <label>
          页面颜色{" "}
          <select aria-label="阅读页面颜色" value={theme} onChange={(e) => setTheme(e.currentTarget.value as "light" | "sepia" | "dark")}>
            <option value="light">纯白</option>
            <option value="sepia">护眼纸</option>
            <option value="dark">暗夜</option>
          </select>
        </label>
        <label>
          行距{" "}
          <select aria-label="阅读行距" value={String(spacing)} onChange={(e) => setSpacing(Number(e.currentTarget.value))}>
            <option value="1.5">标准 (1.5)</option>
            <option value="1.85">舒适 (1.85)</option>
            <option value="2.3">宽松 (2.3)</option>
          </select>
        </label>
        <label>
          行聚焦{" "}
          <select aria-label="行聚焦" value={String(focusLines)} onChange={(e) => setFocusLines(Number(e.currentTarget.value) as 0 | 1 | 3 | 5)}>
            <option value="0">无</option>
            <option value="1">1 段聚焦</option>
            <option value="3">3 段聚焦</option>
            <option value="5">5 段聚焦</option>
          </select>
        </label>
        <button type="button" onClick={onClose} style={{ padding: "4px 12px", borderRadius: 14, border: `1px solid ${palette.border}`, cursor: "pointer" }}>
          关闭沉浸式阅读器 (Esc)
        </button>
      </div>
      <div style={{ flex: 1, overflow: "auto", padding: "36px 24px" }}>
        <div style={{ maxWidth, margin: "0 auto", fontSize: 18, lineHeight: spacing }}>
          {blocks.length === 0 ? (
            <p style={{ opacity: 0.6 }}>文档暂无正文内容。</p>
          ) : (
            blocks.map((block, idx) => {
              const half = Math.floor(focusLines / 2);
              const inFocus =
                focusLines === 0 || (idx >= clampedFocus - half && idx <= clampedFocus + half);
              const isHeading = /^h[1-6]$/.test(block.tag);
              return (
                <p
                  key={idx}
                  onClick={() => setFocusIndex(idx)}
                  style={{
                    margin: isHeading ? "1.4em 0 0.6em" : "0.7em 0",
                    fontSize: isHeading ? "1.35em" : "1em",
                    fontWeight: isHeading ? 700 : 400,
                    opacity: inFocus ? 1 : 0.18,
                    transition: "opacity 120ms ease",
                    cursor: focusLines > 0 ? "pointer" : "default",
                  }}
                >
                  {block.text}
                </p>
              );
            })
          )}
        </div>
      </div>
    </div>
  );
}

/** Split Window second view pane showing a live independently scrollable preview of the document. */
export function SplitDocumentPane({
  model,
  revision,
  onClose,
}: {
  model: ContentDocumentModel;
  revision: number;
  onClose(): void;
}): ReactElement {
  const hostRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    const el = hostRef.current;
    if (!el) return;
    const preview = mountDocumentPreview(el, model, {
      title: "拆分窗口对照视图",
      layout: "pages",
      wordWebFonts,
    });
    return () => preview.dispose();
  }, [model, revision]);

  return (
    <section
      aria-label="拆分窗口"
      style={{
        display: "flex",
        flexDirection: "column",
        borderTop: "3px solid #185abd",
        height: "42vh",
        minHeight: 200,
        background: "#eceef1",
      }}
    >
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          padding: "4px 12px",
          background: "#eef4fc",
          borderBottom: "1px solid #d0d7e2",
          fontSize: 11,
          color: "#2b579a",
        }}
      >
        <span>拆分窗口（独立滚动对照视图）</span>
        <button type="button" onClick={onClose} style={{ fontSize: 11, padding: "2px 8px", cursor: "pointer" }}>
          取消拆分
        </button>
      </div>
      <div ref={hostRef} style={{ flex: 1, overflow: "auto" }} />
    </section>
  );
}

/** Compare Documents dialog: compares the active Word document with another DOCX file and opens a tracked-revisions diff package. */
export function WordCompareDialog({
  currentModel,
  currentName,
  onResult,
  onClose,
}: {
  currentModel: ContentDocumentModel;
  currentName: string;
  onResult(compared: ContentDocumentModel, label: string): void;
  onClose(): void;
}): ReactElement {
  const [file, setFile] = useState<File>();
  const [author, setAuthor] = useState("文档比较");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  async function runCompare(): Promise<void> {
    if (!file) {
      setError("请选择要比较的修订版 DOCX 文件。");
      return;
    }
    setBusy(true);
    setError("");
    try {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const revised = openDocument(importDocx(bytes, { documentId: "compare-revised" }), {
        kind: "document",
      });
      const diffPkg = compareWordDocuments(currentModel, revised, {
        title: `${currentName} vs ${file.name}`,
        author: author.trim() || "文档比较",
      });
      const resultModel = openDocument(diffPkg, { kind: "document" });
      onResult(resultModel, `比较结果 · ${currentName} ↔ ${file.name}`);
      onClose();
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "文档比较失败。");
    } finally {
      setBusy(false);
    }
  }

  return (
    <OfficeDialog
      title="比较文档"
      close={onClose}
      footer={
        <>
          <button type="button" onClick={onClose}>
            取消
          </button>
          <button
            type="button"
            className="primary"
            disabled={busy || !file}
            onClick={() => {
              void runCompare();
            }}
          >
            {busy ? "正在比较…" : "开始比较"}
          </button>
        </>
      }
    >
      <div className="dialog-stack">
        {error && <p role="alert" style={{ color: "#a00000", margin: 0 }}>{error}</p>}
        <label>
          原文档（当前打开）
          <input type="text" readOnly value={currentName} />
        </label>
        <label>
          修订的文档（选择本地 DOCX）
          <input
            type="file"
            accept=".docx"
            aria-label="选择修订的 DOCX 文档"
            onChange={(event) => setFile(event.currentTarget.files?.[0])}
          />
        </label>
        <label>
          标记更改的作者
          <input
            type="text"
            aria-label="比较修订作者"
            value={author}
            onChange={(event) => setAuthor(event.currentTarget.value)}
          />
        </label>
        <p style={{ margin: 0, fontSize: 12, color: "#555" }}>
          将在新文档中以 Word 原生修订（插入/删除标记）标出两份文档的全部差异，可直接接受/拒绝或导出为 DOCX。
        </p>
      </div>
    </OfficeDialog>
  );
}

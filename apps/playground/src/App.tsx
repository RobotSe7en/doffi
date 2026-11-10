import { useEffect, useRef, useState } from "react";
import { createPresentation, createDocumentSession, createOpaqueId, DocumentError, openDocument } from "@doffi/core";
import type { ContentDocumentModel as DocumentModel, DocumentSession, DocumentSnapshot } from "@doffi/core";
import {
  importDocx,
  importPptx,
  DEFAULT_MAX_ZIP_BYTES,
  writeDocumentZip,
} from "@doffi/office";
import type { DocxExportReport, OfficeImportProgress, PptxExportReport } from "@doffi/office";
import type { DocumentWorkbenchView, PresentationWorkbenchView, WorkbookRemoteEditorView, EditorCommandState, TextFormatState, DocumentTextCommand } from "@doffi/editor";
import { DocumentWorkbench, PresentationWorkbench } from "@doffi/editor/react";
import { reportFiles, sampleProposal, presentationFiles, presentationProposal } from "./samples.js";
import { exportCurrentPptx } from "./export-pptx.js";
import { exportCurrentDocx } from "./export-docx.js";
import { downloadBytes } from "./download.js";
import { Icon } from "./icons.js";
import { openWorkbookFile } from "./open-workbook.js";
import type { HostedWorkbook, OpenedWorkspaceDocument } from "./open-workbook.js";
import { WorkbookPane } from "./WorkbookPane.js";
import { WelcomeScreen } from "./WelcomeScreen.js";
import { OfficeRibbon } from "@doffi/editor/react";
import type { OfficeActions } from "@doffi/editor/react";
import { brandLogo } from "./brand.js";
type WorkspaceFile = { id: string; name: string; kind: "word" | "presentation" | "workbook"; detail: string };
interface WorkbookCandidate { document: HostedWorkbook; accept(): void; reject(error: Error): void }
const examples: WorkspaceFile[] = [
  { id: "report", name: "季度项目报告", kind: "word", detail: "批注与智能体协作样例" },
  { id: "docx", name: "文字与编号.docx", kind: "word", detail: "Word 原生文件样例" },
  { id: "presentation", name: "项目进展演示", kind: "presentation", detail: "PPT 编辑与协作样例" },
  { id: "pptx", name: "多页演示文稿.pptx", kind: "presentation", detail: "PowerPoint 原生文件样例" },
  { id: "xlsx", name: "一万行工作簿.xlsx", kind: "workbook", detail: "Excel 工作簿 · 20 万单元格" },
];
const sample = import.meta.env.BASE_URL + "01-text-numbering.docx";
function base64(bytes: Uint8Array): string {
  let value = ""; for (let offset = 0; offset < bytes.length; offset += 32768) value += String.fromCharCode(...bytes.subarray(offset, offset + 32768)); return btoa(value);
}

export function App() {
  const [home, setHome] = useState(true);
  const [libraryVisible, setLibraryVisible] = useState(() => { try { return localStorage.getItem("doffi:library") === "open"; } catch { return false; } });
  const [officeTab, setOfficeTab] = useState("开始");
  const preferredEditing = (): boolean => { try { return localStorage.getItem("doffi:editing") !== "preview"; } catch { return true; } };
  const [editing, setEditing] = useState(preferredEditing);
  const [nativeCommands, setNativeCommands] = useState<readonly EditorCommandState[]>([]);
  const [hasDocumentSelection, setHasDocumentSelection] = useState(false);
  const [textFormat, setTextFormat] = useState<TextFormatState>({ enabled: false });
  const workbookEditor = useRef<WorkbookRemoteEditorView>(undefined);
  const [changingMode, setChangingMode] = useState(false);
  const [zoom, setZoom] = useState(1);
  const [rulers, setRulers] = useState(() => { try { return localStorage.getItem("doffi:rulers") !== "hidden"; } catch { return true; } });
  const [commentsVisible, setCommentsVisible] = useState(true);
  const [pageCount, setPageCount] = useState(0);
  const [files, setFiles] = useState(examples);
  const [activeId, setActiveId] = useState("");
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState<"all" | WorkspaceFile["kind"]>("all");
  const opened = useRef(new Map<string, OpenedWorkspaceDocument>());
  const sessions = useRef(new Map<DocumentModel, DocumentSession>());
  const fileInput = useRef<HTMLInputElement>(null);
  const [model, setModel] = useState<DocumentModel>();
  const [workbook, setWorkbook] = useState<HostedWorkbook>();
  const [, refreshWorkbookMetadata] = useState(0);
  const [candidate, setCandidate] = useState<WorkbookCandidate>();
  const [progress, setProgress] = useState<OfficeImportProgress>();
  const loadRef = useRef<AbortController>(undefined);
  const activeDocument = useRef<OpenedWorkspaceDocument>(undefined);
  // The workspace owns one inactive preview alongside its cached Worker.
  // Keeping this view retains its local-font row index across format switches.
  const lastWorkbookPane = useRef<HostedWorkbook>(undefined);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [layout, setLayout] = useState<"pages" | "continuous">(() => {
    try { return localStorage.getItem("doffi:layout") === "continuous" ? "continuous" : "pages"; } catch { return "pages"; }
  });
  const [snapshot, setSnapshot] = useState<DocumentSnapshot>();
  const sessionRef = useRef<DocumentSession>(undefined);
  const presentationRef=useRef<PresentationWorkbenchView>(undefined);
  const workbenchRef = useRef<DocumentWorkbenchView>(undefined);
  const [session, setSession] = useState<DocumentSession>();
  const [submitted, setSubmitted] = useState(false);
  const generation = useRef(0);
  const [exporting, setExporting] = useState<"pdf" | "png" | "docx" | "pptx">();
  const [pptxReport,setPptxReport]=useState<PptxExportReport>();
  const isPresentation=Boolean(snapshot?.metadata.presentation);
  const [docxReport, setDocxReport] = useState<DocxExportReport>();
  const exportRef = useRef<{ controller: AbortController; session: DocumentSession; revision: number }>(undefined);
  useEffect(() => { try { localStorage.setItem("doffi:library", libraryVisible ? "open" : "closed"); } catch { /* Optional host preference. */ } }, [libraryVisible]);
  useEffect(() => {
    setSubmitted(false);
    setDocxReport(undefined);
    setPptxReport(undefined);
    if (!model) { setSession(undefined); sessionRef.current = undefined; return; }
    const session = sessions.current.get(model) ?? createDocumentSession(model);
    sessions.current.set(model, session);
    sessionRef.current = session;
    setSession(session);
    setSnapshot(session.snapshot());
    const unsubscribe = session.subscribe((event) => {
      if (event.type !== "error") {
        setSnapshot(session.snapshot());
        setPptxReport(report=>report?.revision===session.snapshot().manifest.revision?report:undefined);
        setDocxReport((report) => report?.revision === session.snapshot().manifest.revision ? report : undefined);
        if (exportRef.current?.session === session && exportRef.current.revision !== session.snapshot().manifest.revision) exportRef.current.controller.abort();
      }
    });
    return () => {
      unsubscribe();
      if (exportRef.current?.session === session) { exportRef.current.controller.abort(); exportRef.current = undefined; setExporting(undefined); }
      if (sessionRef.current === session) sessionRef.current = undefined;
    };
  }, [model]);
  useEffect(
    () => () => {
      generation.current++;
      loadRef.current?.abort(); loadRef.current = undefined;
      exportRef.current?.controller.abort(); exportRef.current = undefined;
      for (const session of sessions.current.values()) session.dispose();
      sessions.current.clear();
      for (const document of opened.current.values()) if (document.kind === "workbook") document.dispose();
      opened.current.clear();
      lastWorkbookPane.current = undefined;
    },
    [],
  );
  function activate(next: DocumentModel, file: WorkspaceFile): void {
    setOfficeTab("开始"); setEditing(next.kind === "document" && preferredEditing()); setNativeCommands([]); setTextFormat({ enabled: false }); setHasDocumentSelection(false); setZoom(1); setPageCount(0);
    if (next === model) {
      const view = next.snapshot().metadata.presentation ? presentationRef.current : workbenchRef.current;
      if (view) void (async () => {
        if ('setEditing' in view) await view.setEditing(preferredEditing()); else await view.ready;
        if (activeDocument.current !== next) return;
        view.setZoom(1); setPageCount(view.layout()?.pages.length ?? 0); setHasDocumentSelection(Boolean(view.selection()));
        if ('editor' in view) { setNativeCommands(view.editor.commands()); setTextFormat(view.editor.formatState()); } else setTextFormat(view.formatState());
      })().catch(reason => setError(reason.message));
    }
    setHome(false);
    window.scrollTo(0, 0);
    activeDocument.current = next;
    setWorkbook(undefined); setCandidate(undefined);
    opened.current.set(file.id, next);
    setFiles(current => current.some(item => item.id === file.id) ? current : [...current, file]);
    setActiveId(file.id);
    setModel(next);
  }
  function canLeaveWorkbook(): boolean {
    const current = activeDocument.current;
    if (current?.kind === "workbook" && current.pendingInput) { setError("请先确认或取消当前单元格输入或操作，再切换文档。"); return false; }
    return true;
  }
  function beginLoad(): { current: number; controller: AbortController } | undefined {
    if (!canLeaveWorkbook()) return;
    const current = ++generation.current; loadRef.current?.abort();
    const controller = new AbortController(); loadRef.current = controller;
    exportRef.current?.controller.abort(); exportRef.current = undefined; setExporting(undefined);
    setLoading(true); setProgress(undefined); setError("");
    return { current, controller };
  }
  function stageWorkbook(document: HostedWorkbook, file: WorkspaceFile, current: number, signal: AbortSignal): Promise<void> {
    if (signal.aborted || current !== generation.current) return Promise.reject(new DocumentError("OperationCancelled", "已取消读取。"));
    return new Promise((resolve, reject) => {
      let settled = false;
      const finish = (error?: Error): void => {
        if (settled) return; settled = true;
        signal.removeEventListener("abort", cancel);
        setCandidate(value => value === next ? undefined : value);
        if (error) { reject(error); return; }
        opened.current.set(file.id, document); setFiles(files => files.some(item => item.id === file.id) ? files : [...files, { ...file, kind: "workbook" }]);
        activeDocument.current = document; lastWorkbookPane.current = document;
        setOfficeTab("开始"); setEditing(false); setNativeCommands([]); setTextFormat({ enabled: false }); setHome(false); window.scrollTo(0, 0); setActiveId(file.id); setWorkbook(document); setModel(undefined); setSnapshot(undefined); setSession(undefined); sessionRef.current = undefined;
        setDocxReport(undefined); setPptxReport(undefined); setSubmitted(false); resolve();
      };
      const cancel = (): void => finish(new DocumentError("OperationCancelled", "已取消读取。"));
      const next: WorkbookCandidate = { document, accept() { if (current === generation.current && !signal.aborted) finish(); }, reject: finish };
      signal.addEventListener("abort", cancel, { once: true }); setCandidate(next);
    });
  }
  function cancelLoad(): void {
    generation.current++; loadRef.current?.abort(); loadRef.current = undefined;
    setLoading(false); setProgress(undefined); setError("已取消读取。");
  }
  async function load(
    bytes: Promise<ArrayBuffer>,
    kind: "docx" | "pptx" | "xlsx" | "package",
    file: WorkspaceFile,
  ): Promise<void> {
    const started = beginLoad(); if (!started) { void bytes.catch(() => {}); return; } const { current, controller } = started;
    let next: OpenedWorkspaceDocument | undefined;
    try {
      const value = new Uint8Array(await bytes);
      if (controller.signal.aborted || current !== generation.current) return;
      next = kind === "xlsx" || kind === "package" ? await openWorkbookFile(value, kind, createOpaqueId(), createOpaqueId(), {
        signal: controller.signal, onProgress(value) { if (current === generation.current) setProgress(value); },
        onFailure(error) { if (activeDocument.current === next) setError(error.message); },
      }) : kind === "docx" ? openDocument(importDocx(value, { documentId: createOpaqueId() }), { kind: "document" })
        : openDocument(importPptx(value, { documentId: createOpaqueId() }), { kind: "presentation" });
      if (current === generation.current && !controller.signal.aborted) {
        if (next.kind === "workbook") { setProgress(undefined); await stageWorkbook(next, file, current, controller.signal); }
        else activate(next, { ...file, kind: next.snapshot().metadata.presentation ? "presentation" : "word" });
      }
    } catch (reason) {
      if (current === generation.current)
        setError(
          reason instanceof DocumentError
            ? [
                reason.code,
                reason.message,
                ...reason.details.map((item) => item.path.join(" / ")),
              ].join(" · ")
            : "无法读取文件。",
        );
    } finally {
      if (next?.kind === "workbook" && opened.current.get(file.id) !== next) next.dispose();
      if (current === generation.current) { setLoading(false); setProgress(undefined); if (loadRef.current === controller) loadRef.current = undefined; }
    }
  }
  function download(): void {
    const session = sessionRef.current;
    if (!session) return;
    try { downloadBytes(writeDocumentZip(session.serialize()), "document.zip", "application/zip"); }
    catch (reason) { setError(reason instanceof Error ? reason.message : "文档包导出失败。"); }
  }
  async function openReport(presentation=false): Promise<void> {
    const started = beginLoad(); if (!started) return; const { current, controller } = started;
    try { const files = await (presentation?presentationFiles():reportFiles()); if (current === generation.current) activate(openDocument(files, { kind: presentation ? "presentation" : "document" }), examples[presentation ? 2 : 0]!); }
    catch (reason) { if (current === generation.current) setError(reason instanceof Error ? reason.message : "报告读取失败"); }
    finally { if (current === generation.current) { setLoading(false); if (loadRef.current === controller) loadRef.current = undefined; } }
  }
  function openFile(file: WorkspaceFile): void {
    const cached = opened.current.get(file.id);
    if (cached !== activeDocument.current && !canLeaveWorkbook()) return;
    if (cached) {
      if (cached.kind === "workbook") {
        if (cached.disposed) { setError("工作簿读取进程已停止，请重新选择原文件。"); return; }
        if (cached === workbook) { setError(""); setHome(false); return; }
        const started = beginLoad(); if (!started) return; const { current, controller } = started;
        void stageWorkbook(cached, file, current, controller.signal).catch(error => { if (current === generation.current) setError(error.message); })
          .finally(() => { if (current === generation.current) { setLoading(false); loadRef.current = undefined; } });
        return;
      }
      generation.current++;
      loadRef.current?.abort(); loadRef.current = undefined;
      exportRef.current?.controller.abort(); exportRef.current = undefined; setExporting(undefined);
      setError(""); setLoading(false); activate(cached, file);
    } else if (file.id === "report" || file.id === "presentation") {
      void openReport(file.id === "presentation");
    } else {
      const url = file.id === "xlsx" ? import.meta.env.BASE_URL + "xlsx/15-ten-thousand-rows.xlsx" : file.id === "pptx" ? import.meta.env.BASE_URL + "pptx/13-multi-slide.pptx" : sample;
      void load(fetch(url).then(response => { if (!response.ok) throw new Error("样例读取失败"); return response.arrayBuffer(); }), file.id === "xlsx" ? "xlsx" : file.id === "pptx" ? "pptx" : "docx", file);
    }
  }
  async function downloadDocx(): Promise<void> {
    const active = sessionRef.current; if (!active || exportRef.current) return;
    const current = generation.current, identity = active.snapshot().manifest;
    const job = { controller: new AbortController(), session: active, revision: identity.revision };
    exportRef.current = job; setExporting("docx"); setDocxReport(undefined); setError("");
    try {
      const result = await exportCurrentDocx(active.serialize(), createOpaqueId(), job.controller.signal);
      if (current !== generation.current || exportRef.current !== job) return;
      if (result.report.documentId !== identity.documentId || result.report.revision !== identity.revision) throw new Error("DOCX 导出结果与当前请求不匹配。");
      if (active.snapshot().manifest.revision !== identity.revision) throw new Error("导出期间文档已变化，请重新导出。");
      downloadBytes(result.bytes, "document.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
      setDocxReport(result.report);
    } catch (reason) {
      if (current === generation.current && exportRef.current === job) setError(active.snapshot().manifest.revision !== identity.revision ? "导出期间文档已变化，请重新导出。"
        : job.controller.signal.aborted ? "已取消 DOCX 导出。" : reason instanceof Error ? reason.message : "DOCX 导出失败。");
    } finally { if (exportRef.current === job) { exportRef.current = undefined; setExporting(undefined); } }
  }
  async function downloadPptx():Promise<void> {
    const active=sessionRef.current;if(!active||exportRef.current)return;
    const current=generation.current,identity=active.snapshot().manifest,job={controller:new AbortController(),session:active,revision:identity.revision};
    exportRef.current=job;setExporting("pptx");setPptxReport(undefined);setError("");
    try {
      const result=await exportCurrentPptx(active.serialize(),createOpaqueId(),job.controller.signal);
      if(current!==generation.current||exportRef.current!==job)return;
      if(result.report.documentId!==identity.documentId||result.report.revision!==identity.revision)throw new Error("PPTX 导出结果与当前请求不匹配。");
      if(active.snapshot().manifest.revision!==identity.revision)throw new Error("导出期间文稿已变化，请重新导出。");
      downloadBytes(result.bytes,"presentation.pptx","application/vnd.openxmlformats-officedocument.presentationml.presentation");setPptxReport(result.report);
    } catch(reason){if(current===generation.current&&exportRef.current===job)setError(active.snapshot().manifest.revision!==identity.revision?"导出期间文稿已变化，请重新导出。":job.controller.signal.aborted?"已取消 PPTX 导出。":reason instanceof Error?reason.message:"PPTX 导出失败。");}
    finally{if(exportRef.current===job){exportRef.current=undefined;setExporting(undefined);}}
  }
  async function loadSampleProposal(): Promise<void> {
    const current = generation.current;
    if(isPresentation){const view=presentationRef.current;if(!view)return;try{const proposal=await presentationProposal();if(current!==generation.current||presentationRef.current!==view)return;const request=view.request()??view.buildRequest();view.loadProposal(request,{...proposal,requestId:request.requestId,documentId:request.documentId,baseRevision:request.baseRevision,annotationIds:request.annotations.map(annotation=>annotation.annotationId)});}catch(reason){if(current===generation.current)setError(reason instanceof Error?reason.message:"固定提案导入失败");}return;}
    const view = workbenchRef.current;
    if (!view) return;
    try {
      const proposal = await sampleProposal();
      if (current !== generation.current || workbenchRef.current !== view) return;
      const request = view.request() ?? view.buildRequest();
      view.loadProposal(request, { ...proposal, requestId: request.requestId, documentId: request.documentId,
        baseRevision: request.baseRevision, annotationIds: request.annotations.map((annotation) => annotation.annotationId) });
    } catch (reason) { if (current === generation.current) setError(reason instanceof Error ? reason.message : "固定提案导入失败"); }
  }
  async function downloadRendered(format:"pdf"|"png"="pdf"): Promise<void> {
    const label=format.toUpperCase();
    const active = sessionRef.current; if (!active || exportRef.current) return;
    const current = generation.current, identity = active.snapshot().manifest;
    const job = { controller: new AbortController(), session: active, revision: identity.revision };
    exportRef.current = job; setExporting(format); setError("");
    try {
      const taskId = createOpaqueId(), bytes = writeDocumentZip(active.serialize());
      const response = await fetch("/__doffi/convert", { method: "POST", signal: job.controller.signal,
        headers: { "Content-Type": "application/json", "X-Doffi-Dev": "1" },
        body: JSON.stringify({ protocolVersion: 1, taskId, format, input: { format: "document-package-zip", documentId: identity.documentId, bytesBase64: base64(bytes) } }),
      });
      const result = await response.json();
      if (current !== generation.current || exportRef.current !== job) return;
      if (!response.ok || !result.ok) throw new Error(typeof result.error?.message === "string" ? result.error.message.slice(0, 256) : label+" 导出失败。");
      if (result.protocolVersion !== 1 || result.taskId !== taskId || result.output?.mimeType !== (format==="pdf"?"application/pdf":"image/png")
        || result.output.report?.documentId !== identity.documentId || result.output.report?.revision !== identity.revision) throw new Error(label+" 导出结果与当前请求不匹配。");
      if (active.snapshot().manifest.revision !== identity.revision) throw new Error("导出期间文档已变化，请重新导出。");
      const content = Uint8Array.from(atob(result.output.bytesBase64), (character) => character.charCodeAt(0));
      downloadBytes(content, (isPresentation?"presentation":"document")+"."+format, format==="pdf"?"application/pdf":"image/png");
    } catch (reason) {
      if (current === generation.current && exportRef.current === job) setError(active.snapshot().manifest.revision !== identity.revision ? "导出期间文档已变化，请重新导出。"
        : job.controller.signal.aborted ? "已取消 "+label+" 导出。" : reason instanceof Error ? reason.message : label+" 导出失败。");
    } finally { if (exportRef.current === job) { exportRef.current = undefined; setExporting(undefined); } }
  }
  const activeFile = files.find(file => file.id === activeId);
  const visibleFiles = files.filter(file => (filter === "all" || file.kind === filter) && (file.name + file.detail).toLowerCase().includes(search.toLowerCase()));
  const exampleLabels: Record<string, string> = { report: "打开报告协作样例", docx: "打开 DOCX 样例", presentation: "打开 PPT 协作样例", pptx: "打开 PPTX 样例", xlsx: "打开一万行 XLSX 样例" };
  const progressLabels = { zip: "解包", metadata: "读取工作簿", worksheet: "读取工作表", content: "读取对象与批注", validate: "校验工作簿" };
  useEffect(() => {
    const view = isPresentation ? presentationRef.current : workbenchRef.current;
    if (!view || workbook) return;
    let current = true;
    void view.ready.then(() => { if (current) setPageCount(view.layout()?.pages.length ?? 0); }).catch(() => {});
    return () => { current = false; };
  }, [snapshot, editing, layout, workbook, isPresentation]);
  useEffect(() => { workbenchRef.current?.setRulers(rulers); }, [rulers, model]);
  useEffect(() => { presentationRef.current?.setReviewVisible(officeTab === "审阅"); }, [officeTab, model]);
  function changeRulers(visible: boolean): void { setRulers(visible); try { localStorage.setItem("doffi:rulers", visible ? "visible" : "hidden"); } catch { /* Optional preference. */ } }
  function changeLayout(value: "pages" | "continuous"): void {
    setLayout(value); setZoom(1); try { localStorage.setItem("doffi:layout", value); } catch { /* Optional host preference. */ }
  }
  function changeZoom(value: number): void {
    try { (isPresentation ? presentationRef.current : workbenchRef.current)?.setZoom(value); setZoom(value); } catch (reason) { setError(reason instanceof Error ? reason.message : "缩放失败。"); }
  }
  async function changeEditing(value: boolean): Promise<void> {
    const view = workbenchRef.current; if (!view || changingMode) return;
    setChangingMode(true);
    try { await view.setEditing(value); if (workbenchRef.current === view) { setEditing(value); try { localStorage.setItem("doffi:editing", value ? "edit" : "preview"); } catch { /* Optional preference. */ } } }
    catch (reason) { if (workbenchRef.current === view) setError(reason instanceof Error ? reason.message : "切换编辑模式失败。"); }
    finally { setChangingMode(false); }
  }
  function nativeEditor() { return workbook ? workbookEditor.current : isPresentation ? presentationRef.current?.editor : undefined; }
  function runNative(id: string): void {
    const editor = nativeEditor(); if (!editor) return;
    void editor.executeCommand(id).catch(reason => setError(reason instanceof Error ? reason.message : "命令执行失败。"));
  }
  function runTextCommand(command: DocumentTextCommand): void | Promise<void> {
    try {
      if (workbook) {
        const editor = workbookEditor.current; if (!editor) return;
        const state = editor.formatState(), properties: Record<string, string> = {};
        if (command.type === "toggle") {
          const active = state[command.property] === true;
          properties[command.property === "bold" ? "font-weight" : command.property === "italic" ? "font-style" : "--office-font-underline"] = command.property === "bold" ? active ? "normal" : "bold" : command.property === "italic" ? active ? "normal" : "italic" : active ? "none" : "single";
        } else if (command.type === "align") properties["--office-cell-horizontal"] = command.value;
        else { properties[command.property] = command.value; if (command.property === "font-family") for (const role of ["ascii", "h-ansi", "east-asia"]) properties["--office-font-" + role] = command.value; }
        return editor.formatCells(properties).catch(reason => setError(reason.message));
      } else if (isPresentation) presentationRef.current?.editor.executeTextCommand(command);
      else workbenchRef.current?.executeTextCommand(command);
    } catch (reason) { setError(reason instanceof Error ? reason.message : "格式命令失败。"); }
  }
  function clipboard(action: "copy" | "cut" | "paste"): void {
    if (workbook) { runNative(action); return; }
    const view = isPresentation ? presentationRef.current?.editor : workbenchRef.current;
    void view?.clipboard(action).catch(reason => setError(reason.message));
  }
  const history = workbook ? workbook.session.state().history : session?.history();
  const actions: OfficeActions = {
    copy: { run: () => clipboard("copy"), disabled: !textFormat.enabled || !textFormat.hasSelection },
    cut: { run: () => clipboard("cut"), disabled: !textFormat.enabled || !textFormat.hasSelection },
    paste: { run: () => clipboard("paste"), disabled: !textFormat.enabled || isPresentation && !textFormat.editingText || !navigator.clipboard?.read, reason: !navigator.clipboard?.read ? "浏览器限制读取剪贴板，请使用 Ctrl/⌘+V" : "请先进入文字编辑" },
    newComment: { run: () => { setCommentsVisible(true); if (workbook) runNative("comments"); else if (isPresentation) presentationRef.current?.startAnnotation(); else {
      const view = workbenchRef.current, chosen = view?.selection();
      void (async () => { if (!view || !chosen) return; await view.ready; view.startAnnotation(chosen); })().catch(reason => setError(reason.message));
    } }, disabled: !workbook && !(editing && !isPresentation ? textFormat.hasSelection : hasDocumentSelection) },
    open: { run: () => fileInput.current?.click() },
    ...(!workbook ? {
      package: { run: download, disabled: !session },
      native: { run: () => { void (isPresentation ? downloadPptx() : downloadDocx()); }, disabled: !session || Boolean(exporting) },
      pdf: { run: () => { void downloadRendered(); }, disabled: !session || Boolean(exporting) },
      png: { run: () => { void downloadRendered("png"); }, disabled: !session || Boolean(exporting) },
      proposal: { run: () => { setOfficeTab("审阅"); void loadSampleProposal(); }, disabled: !session },
    } : {}),
    undo: { run: () => { if (workbook || isPresentation) runNative("undo"); else session?.undo(); }, disabled: !history?.undoSteps || Boolean(workbook?.pendingInput) },
    redo: { run: () => { if (workbook || isPresentation) runNative("redo"); else session?.redo(); }, disabled: !history?.redoSteps || Boolean(workbook?.pendingInput) },
    ...(!workbook && !editing ? {
      textSelection: { run: () => { (isPresentation ? presentationRef.current : workbenchRef.current)?.setMode("text"); } },
      objectSelection: { run: () => { (isPresentation ? presentationRef.current : workbenchRef.current)?.setMode("node"); } },
      regionSelection: { run: () => { (isPresentation ? presentationRef.current : workbenchRef.current)?.setMode("region"); } },
    } : {}),
    fields: { run: () => { try { workbenchRef.current?.updateFields(); } catch (reason) { setError(reason instanceof Error ? reason.message : "更新字段失败。"); } }, disabled: isPresentation || Boolean(workbook) },
    comments: { run: () => { if (workbook) runNative("comments"); else setCommentsVisible(value => !value); } },
    library: { run: () => setLibraryVisible(value => !value) },
    preview: { run: () => { void changeEditing(false); }, disabled: isPresentation || Boolean(workbook) },
  };
  function importFile(file: File): void {
    if (file.size > DEFAULT_MAX_ZIP_BYTES) { setError("文件超过 64 MiB，无法打开。"); return; }
    const name = file.name.toLowerCase();
    if (!/\.(docx|pptx|xlsx|zip)$/.test(name)) { setError("请选择 DOCX、XLSX、PPTX 或 ZIP 文档包。"); return; }
    void load(file.arrayBuffer(), name.endsWith(".docx") ? "docx" : name.endsWith(".pptx") ? "pptx" : name.endsWith(".xlsx") ? "xlsx" : "package",
      { id: createOpaqueId(), name: file.name, kind: name.endsWith(".xlsx") ? "workbook" : name.endsWith(".pptx") ? "presentation" : "word", detail: "本地导入 · " + (file.size < 1024 * 1024 ? Math.max(1, Math.round(file.size / 1024)) + " KB" : (file.size / 1024 / 1024).toFixed(1) + " MB") });
  }
  function newPresentation(): void {
    if (!canLeaveWorkbook()) return;
    generation.current++; exportRef.current?.controller.abort(); exportRef.current = undefined; setExporting(undefined);
    loadRef.current?.abort(); loadRef.current = undefined;
    activate(openDocument(createPresentation({ documentId: createOpaqueId() }), { kind: "presentation" }), { id: createOpaqueId(), name: "未命名演示文稿", kind: "presentation", detail: "新建 PowerPoint 文稿" }); setError("");
  }
  return (
    <main data-office-tab={officeTab} data-revision={workbook?.snapshot().manifest.revision ?? snapshot?.manifest.revision} data-annotation-count={workbook?.snapshot().annotations.length ?? snapshot?.annotations.length} data-sheet-count={workbook?.snapshot().metadata.workbook.sheets.length} data-document-id={workbook?.snapshot().manifest.documentId ?? snapshot?.manifest.documentId} data-loading={loading} className={"office-app" + (home ? " is-home" : "") + (!libraryVisible ? " library-hidden" : "") + (!workbook && !isPresentation ? " word-office" : isPresentation ? " presentation-office" : " workbook-office") + (!commentsVisible ? " hide-comments" : "")}>
      {home && <WelcomeScreen loading={loading} onOpen={id => { const file = files.find(item => item.id === id); if (file) openFile(file); }} onChooseFile={() => fileInput.current?.click()} onFile={importFile} onNewPresentation={newPresentation} />}
      {home && (loading || error) && <div className="welcome-feedback" aria-label="文件读取状态">{loading && <p role="status">{progress ? `${progressLabels[progress.phase]} · ${progress.completed} / ${progress.total}` : candidate ? "正在准备工作簿预览…" : "正在读取…"} <button onClick={cancelLoad}>取消读取</button></p>}{error && <p role="alert">{error}</p>}</div>}
      <aside className="document-sidebar" aria-label="文档列表" hidden={home || !libraryVisible}>
        <a className="brand" href="./" aria-label="doffi Playground">
          <img className="brand-logo" src={brandLogo} alt="" width="1150" height="572" />
          <small>Playground</small>
        </a>
        <div className="library-heading"><span>工作空间</span><span className="workspace-label">PLAYGROUND</span></div>
        <label className="document-search"><Icon name="search" size={16} /><input aria-label="搜索文档" placeholder="搜索文档…" value={search} onChange={event => setSearch(event.currentTarget.value)} /><kbd>⌕</kbd></label>
        <div className="library-title"><h2>我的文档</h2><span>{files.length}</span></div>
        <div className="document-filters" role="group" aria-label="筛选文档">
          {([["all", "全部"], ["word", "Word"], ["presentation", "PPT"], ["workbook", "Excel"]] as const).map(([value, label]) => <button key={value} aria-pressed={filter === value} onClick={() => setFilter(value)}>{label}</button>)}
        </div>
        <nav className="document-list" aria-label="打开文档">
          {visibleFiles.map(file => <button key={file.id} className={"document-item" + (activeId === file.id ? " active" : "")} aria-label={exampleLabels[file.id] ?? "打开 " + file.name} aria-current={activeId === file.id ? "page" : undefined} disabled={loading} onClick={() => openFile(file)}>
            <span className={"file-icon " + file.kind}><Icon name={file.kind === "word" ? "file" : file.kind === "workbook" ? "grid" : "slides"} size={21} /></span>
            <span className="file-description"><strong>{file.name}</strong><small>{file.detail}</small></span>
            {activeId === file.id && <span className="active-indicator" />}
          </button>)}
          {!visibleFiles.length && <p className="no-results">没有找到匹配的文档</p>}
        </nav>
        <div className="sidebar-note"><Icon name="comment" size={20} /><strong>一起完善这份文档</strong><p>选中内容添加批注，<br />审阅并应用修改提案。</p><span>DOCX <i /> XLSX <i /> PPTX</span></div>
        <div className="sidebar-footer"><span className="avatar">D</span><div><strong>本地工作空间</strong><small>文档仅保留在当前页面</small></div><span className="connection-dot" /></div>
      </aside>
      <section className={"workspace" + (home && candidate ? " home-preparing" : "")} aria-label="文档工作区" hidden={home && !candidate} aria-hidden={home || undefined} inert={home}>
        <header className="document-header">
          <div className="document-navigation"><button aria-label="返回首页" title="返回首页" onClick={() => { if (canLeaveWorkbook()) { setHome(true); window.scrollTo(0, 0); } }}><Icon name="home" size={17} /></button><button aria-label="显示文档列表" title="显示或收起文档列表" aria-expanded={libraryVisible} onClick={() => setLibraryVisible(value => !value)}><Icon name="panel" size={17} /></button></div>
          <div className="quick-access" role="toolbar" aria-label="快速访问">{!workbook && <button title="下载文档包" aria-label="下载文档包" disabled={!session || loading} onClick={download}><Icon name="save" size={17} /></button>}<button title="撤销" aria-label="撤销" disabled={loading || actions.undo?.disabled} onMouseDown={event => event.preventDefault()} onClick={actions.undo?.run}><Icon name="undo" size={17} /></button><button title="重做" aria-label="重做" disabled={loading || actions.redo?.disabled} onMouseDown={event => event.preventDefault()} onClick={actions.redo?.run}><Icon name="redo" size={17} /></button></div>
          <div className="document-title"><span className={"file-icon " + (workbook ? "workbook" : isPresentation ? "presentation" : "word")}><Icon name={workbook ? "grid" : isPresentation ? "slides" : "file"} size={18} /></span><div><h1>{activeFile?.name ?? "文档工作台"}</h1><p>{workbook ? "Excel" : isPresentation ? "PowerPoint" : "Word"} · doffi Playground</p></div></div>
          <div className="document-header-actions">{!workbook && !isPresentation && <button className="editing-mode-button" aria-label={editing ? "返回预览与批注" : "编辑正文"} disabled={loading || changingMode || !session} onClick={() => { void changeEditing(!editing); }}><Icon name={editing ? "edit" : "book"} size={15} />{editing ? "返回预览与批注" : "编辑正文"}<span className="dropdown-arrow" aria-hidden="true">⌄</span></button>}<button className="subtle-button" disabled={loading} onClick={newPresentation}><Icon name="plus" size={16} />新建 PPT</button><button className="primary-button" disabled={loading} onClick={() => fileInput.current?.click()}><Icon name="upload" size={16} />打开文件</button></div>
          <input
            ref={fileInput}
            className="file-input"
            hidden
            aria-label="打开文件"
            type="file"
            disabled={loading}
            accept=".docx,.pptx,.xlsx,.zip"
            onChange={(event) => {
              const file = event.currentTarget.files?.[0];
              if (file) importFile(file);
              event.currentTarget.value = "";
            }}
          />
        </header>
      <div role="toolbar" aria-label="文件与导出"><OfficeRibbon kind={workbook ? "workbook" : isPresentation ? "presentation" : "document"} tab={officeTab} onTab={setOfficeTab} mode={!workbook && !isPresentation && !editing ? "preview" : "edit"} busy={loading || changingMode} actions={actions} commands={nativeCommands} onCommand={runNative} formatState={textFormat}
        onTextCommand={runTextCommand} rulers={rulers} onRulers={changeRulers} layout={layout} onLayout={changeLayout} zoom={zoom} onZoom={changeZoom} /></div>
      <div className="workspace-feedback">
      {exporting && <p role="status">正在导出 {exporting.toUpperCase()}… <button onClick={() => exportRef.current?.controller.abort()}>取消导出</button></p>}
      {loading && <p role="status">{progress ? `${progressLabels[progress.phase]} · ${progress.completed} / ${progress.total}` : candidate ? "正在准备工作簿预览…" : "正在读取…"} <button onClick={cancelLoad}>取消读取</button></p>}
      {error && <p className="error-message" role="alert">{error}</p>}
      {submitted && <p className="submission-message"><Icon name="check" size={16} />Playground 已接收批注请求；使用固定提案验证返回与应用。</p>}
      {docxReport && <section className="export-report" aria-label="DOCX 导出报告">
        <p>已生成 DOCX · 保留原生批注 {docxReport.exportedAnnotationIds.length} 条。</p>
        {docxReport.unexportedAnnotations.length > 0 && <>
          <p>{docxReport.unexportedAnnotations.length} 条批注无法精确绑定到 Word 文字，已列入报告。下载文档包可完整保留这些批注。</p>
          <ul>{docxReport.unexportedAnnotations.slice(0, 16).map((item) => <li key={item.annotationId}>{item.annotationId}：{{ region: "矩形批注", needs_reanchor: "需要重新定位", disjoint: "不连续的文字范围", cross_story: "跨正文或辅助内容的范围", unrepresentable: "没有可精确绑定的文字范围" }[item.reason]}</li>)}</ul>
          {docxReport.unexportedAnnotations.length > 16 && <p>其余批注详见完整报告。</p>}
        </>}
        <button onClick={() => downloadBytes(new TextEncoder().encode(JSON.stringify(docxReport, null, 2)), "document.export-report.json", "application/json")}>下载导出报告</button>
      </section>}
      {pptxReport&&<section className="export-report" aria-label="PPTX 导出报告">
        <p>已生成 PPTX · 原生批注 {pptxReport.exportedAnnotationIds.length} 条 · 协作信息损失 {pptxReport.annotationLosses.length} 项。</p>
        {pptxReport.unexportedAnnotations.length>0&&<p>{pptxReport.unexportedAnnotations.length} 条批注没有可导出的幻灯片，详见完整报告。</p>}
        {pptxReport.annotationLosses.length>0&&<p>部分批注锚点、格式或来源信息无法原样表示，下载文档包可完整保留。</p>}
        <button onClick={()=>downloadBytes(new TextEncoder().encode(JSON.stringify(pptxReport,null,2)),"presentation.export-report.json","application/json")}>下载导出报告</button>
      </section>}
      </div>
      {!model && !workbook && <div className="empty-workspace"><div className="empty-paper"><div className="paper-lines"><i /><i /><i /><i /></div><span className="empty-comment"><Icon name="comment" size={24} /></span></div><span className="empty-eyebrow">YOUR NEXT IDEA STARTS HERE</span><h2>让文档成为协作的起点</h2><p>从左侧选择一份样例，或打开自己的 Office 文档。<br />选中内容、写下批注，让每一个修改想法都有迹可循。</p><div className="format-tags"><span><i className="word-dot" />Word · DOCX</span><span><i className="ppt-dot" />PowerPoint · PPTX</span><span><i className="html-dot" />HTML 文档包</span></div></div>}
      {(workbook || candidate || lastWorkbookPane.current) && <div style={{ display: "contents" }}>
        <WorkbookPane document={(candidate?.document ?? workbook ?? lastWorkbookPane.current)!} preparing={!workbook} activation={candidate}
        onEditor={(document, view) => { if (activeDocument.current === document) { workbookEditor.current = view; setNativeCommands(view.commands()); setTextFormat(view.formatState()); } }}
        onCommands={(document, commands) => { if (activeDocument.current === document) setNativeCommands(commands); }}
        onFormat={(document, state) => { if (activeDocument.current === document) setTextFormat(state); }}
        onReady={document => { if (candidate?.document === document) candidate.accept(); }}
        onUpdate={document => { if (workbook === document) refreshWorkbookMetadata(value => value + 1); }}
        onError={(document, error) => { if (candidate?.document === document) candidate.reject(error); else if (workbook === document) setError(error.message); }} /></div>}
      {model&&session&&isPresentation&&<PresentationWorkbench session={session} appearance="workspace" annotationTrigger="command" className="preview" onSelectionChange={selection => setHasDocumentSelection(Boolean(selection))} onCommandStateChange={setNativeCommands} onFormatStateChange={setTextFormat} onViewChange={view=>{presentationRef.current=view; if(view) {view.setReviewVisible(officeTab === "审阅"); setNativeCommands(view.editor.commands()); setTextFormat(view.editor.formatState()); void view.ready.then(() => { if (presentationRef.current === view) { view.setZoom(1); setPageCount(view.layout()?.pages.length ?? 0); } }).catch(() => {});}}} onSubmitAnnotations={()=>{setSubmitted(true);}} onError={reason=>setError(reason.message)}/>}
      {model && session && !isPresentation && <DocumentWorkbench session={session} appearance="workspace" annotationTrigger="command" layout={layout} className="preview" onSelectionChange={selection => setHasDocumentSelection(Boolean(selection))} onFormatStateChange={setTextFormat}
        onViewChange={(view) => { workbenchRef.current = view; if (view) { view.setRulers(rulers); if (editing) void view.setEditing(true).catch(reason => setError(reason.message)); void view.ready.then(() => { if (workbenchRef.current === view) setPageCount(view.layout()?.pages.length ?? 0); }).catch(() => {}); } }}
        onSubmitAnnotations={() => { setSubmitted(true); }}
        onError={(reason) => setError(reason.message)} />}

        <footer className="workspace-footer"><span>{pageCount > 0 && !workbook && <>{pageCount} {isPresentation ? "张幻灯片" : "页"} <i /></>}简体中文 <i /><span className="connection-dot" />本地工作空间</span><div className="office-status-controls">{!workbook && !isPresentation && <><button aria-label="页面视图" title="页面视图" onClick={() => changeLayout("pages")}><Icon name="file" size={14} /></button><button aria-label="连续视图" title="连续视图" onClick={() => changeLayout("continuous")}><Icon name="globe" size={14} /></button></>}<span className="zoom-controls" title={workbook ? "工作簿缩放尚未实现" : undefined}><button aria-label="缩小" disabled={Boolean(workbook) || zoom <= .5} onClick={() => changeZoom(Math.max(.5, zoom - .25))}>−</button><input aria-label="文档缩放" disabled={Boolean(workbook)} type="range" min="50" max="200" step="25" value={zoom * 100} onChange={event => changeZoom(Number(event.currentTarget.value) / 100)} /><button aria-label="放大" disabled={Boolean(workbook) || zoom >= 2} onClick={() => changeZoom(Math.min(2, zoom + .25))}>＋</button><span>{Math.round(zoom * 100)}%</span></span></div></footer>
      </section>
    </main>
  );
}

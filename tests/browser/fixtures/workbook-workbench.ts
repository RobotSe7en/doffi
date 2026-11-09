import { createDocumentSession, openDocument } from "@doffi/core";
import type { WorkbookEditProposal, WorkbookInteractionRequest, WorkbookSession } from "@doffi/core";
import { readDocumentZip } from "@doffi/office";
import { connectWorkbookSession, mountDocumentWorkbench } from "@doffi/editor";
import type { WorkbookSessionConnection, WorkbookWorkbenchOptions, WorkbookWorkbenchView } from "@doffi/editor";
import { WorkbookWorkbench } from "@doffi/editor/react";
import { createElement } from "react";
import { createRoot } from "react-dom/client";
import { decodeTestBytes } from "./package-transport.js";
import { measureWorkbookDrawingBounds, observeWorkbookViews } from "./workbook-view-observer.js";

let cleanup = () => {}, workbench: WorkbookWorkbenchView, session: WorkbookSession | WorkbookSessionConnection;
let releaseHost: ((value: WorkbookEditProposal) => void) | undefined, submitted: WorkbookInteractionRequest | undefined, hostMode = "return", worker: Worker | undefined;
let errors: string[] = [], disposed = false, unmount: (() => void) | undefined;
let observed: ReturnType<typeof observeWorkbookViews> | undefined;
function edit(request: WorkbookInteractionRequest, value: number, column = 0): WorkbookEditProposal { return { protocolVersion: 1, requestId: request.requestId, documentId: request.documentId, baseRevision: request.baseRevision, proposalId: "reviewed-proposal", annotationIds: request.annotations.map(annotation => annotation.annotationId), summary: "将 A1 改为 " + value + "，核对只读公式当前结果。", operations: [{ op: "setWorkbookCells", sheetId: "sheet-1", cells: [{ row: 0, column, source: { kind: "value", value: { type: "number", value } } }] }] }; }
export async function mount(encoded: string, workerCode?: string, react = false, confirmWarnings = false, observeDrawings = false) {
  cleanup(); errors = []; hostMode = "return"; releaseHost = undefined; submitted = undefined; disposed = false;
  const host = document.createElement("div"); document.body.append(host); let url: string | undefined;
  if (workerCode) { url = URL.createObjectURL(new Blob([workerCode], { type: "text/javascript" })); worker = new Worker(url); const channel = new MessageChannel(), bytes = decodeTestBytes(encoded); worker.postMessage({ kind: "open", format: "package", bytes, port: channel.port1 }, [bytes.buffer, channel.port1]); session = await connectWorkbookSession(channel.port2); }
  else { worker = undefined; session = createDocumentSession(openDocument(readDocumentZip(decodeTestBytes(encoded)), { kind: "workbook" }), { sessionLimits: { maxCandidates: 1 } }); }
  const config: WorkbookWorkbenchOptions = { widthPt: 420, heightPt: 240, ...(confirmWarnings ? { confirmValidationWarnings: () => window.confirm("应用工作簿验证警告？") } : {}), allowGlobalScope: true, onError: error => errors.push((error as { code?: string }).code ?? error.message), onSubmitAnnotations: async request => {
    submitted = request as WorkbookInteractionRequest; if (hostMode === "fail") throw Error("private host error");
    if (hostMode === "hold") return new Promise<WorkbookEditProposal>(resolve => { releaseHost = resolve; });
    return hostMode === "void" ? undefined : edit(submitted, 10);
  } };
  const attach = Element.prototype.attachShadow;
  Element.prototype.attachShadow = function(options) { return attach.call(this, { ...options, mode: "open" }); };
  try {
    if (react) { const root = createRoot(host); unmount = () => root.unmount(); await new Promise<void>(resolve => { root.render(createElement(WorkbookWorkbench, { ...config, session, onViewChange: value => { if (value) { workbench = value; resolve(); } } })); }); }
    else { unmount = undefined; workbench = mountDocumentWorkbench(host, session, { ...config, format: "workbook" }); }
    if (observeDrawings) { Element.prototype.attachShadow = attach; observed = observeWorkbookViews(workbench.editor.element, { openHostShadows: true }); }
  } finally { if (!observed) Element.prototype.attachShadow = attach; }
  await workbench.ready; await workbench.settled();
  cleanup = () => { observed?.dispose(); observed = undefined; if (!disposed) { unmount?.(); workbench.dispose(); disposed = true; } session.dispose(); worker?.terminate(); if (url) URL.revokeObjectURL(url); host.remove(); };
}
export async function inspect() {
  const lease = "kind" in session ? undefined : await session.view(), snapshot = "kind" in session ? session.snapshot() : lease!.cells.snapshot;
  try {
    const cells = "kind" in session ? [session.getCell("sheet-1", 0, 0), session.getCell("sheet-1", 0, 1)] : (await lease!.cells.read({ documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, maxCells: 2, maxCodeUnits: 65536, windows: [{ sheetId: "sheet-1", row: 0, column: 0, rows: 1, columns: 2 }] }, new AbortController().signal)).cells;
    return { revision: snapshot.manifest.revision, annotations: snapshot.annotations, history: "kind" in session ? session.history() : session.state().history, cells, drawings: snapshot.metadata.workbook.objects ?? [], content: snapshot.content?.html, request: workbench.request(), errors: [...errors], submitted: submitted?.requestId, held: !!releaseHost, disposed };
  } finally { lease?.dispose(); }
}
export function build() { return workbench.buildRequest(); }
export function load(value = 10, column = 0) { const request = workbench.request(); if (!request) throw Error("No request"); return workbench.loadProposal(request, edit(request, value, column)).then(() => {}); }
export function loadNoop() { const request = workbench.request(); if (!request) throw Error("No request"); return workbench.loadProposal(request, { ...edit(request, 10), operations: [{ op: "setWorkbookCellFormats", sheetId: "sheet-1", cells: [{ row: 0, column: 1 }] }] }).then(() => {}); }
export async function selectRows(rows: number) { await workbench.editor.selectRange("sheet-1", { from: { row: 0, column: 0 }, to: { row: rows - 1, column: 0 } }); }
export function formatRows(rows: number) { const request = workbench.request(); if (!request) throw Error("No request"); return workbench.loadProposal(request, { ...edit(request, 10), operations: [{ op: "formatWorkbookRanges", sheetId: "sheet-1", ranges: [{ from: { row: 0, column: 0 }, to: { row: rows - 1, column: 0 } }], properties: { color: "red" } }] }).then(() => {}); }
export function formatFont() { const request = workbench.request(); if (!request) throw Error("No request"); return workbench.loadProposal(request, { ...edit(request, 10), operations: [{ op: "formatWorkbookRanges", sheetId: "sheet-1", ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }], properties: { "font-size": "40pt" } }] }).then(() => {}); }
export function drawingBounds() { return measureWorkbookDrawingBounds(observed?.current()); }
export async function addFixedObject() {
  await workbench.editor.perform([{ op: "updateWorkbookSheet", sheetId: "sheet-1", changes: { defaultRowHeightCustom: false, rowDimensions: [] } },
    { op: "insertWorkbookObject", sheetId: "sheet-1", anchor: { kind: "two-cell", from: { row: 2, column: 1, offsetXPt: 2, offsetYPt: 3 }, to: { row: 5, column: 3, offsetXPt: 4, offsetYPt: 5 }, moveWithCells: false, sizeWithCells: false }, object: { kind: "geometry", preset: "rect", html: "<p>外部固定对象</p>" } }]);
  await workbench.settled(); return (await inspect()).drawings[0]!.nodeId;
}
export async function addObject() {
  const current = "kind" in session ? session.snapshot().manifest : session.state(), transaction = { protocolVersion: 1 as const, documentId: current.documentId, baseRevision: current.revision,
    operations: [{ op: "insertWorkbookObject" as const, sheetId: "sheet-1", anchor: { kind: "absolute" as const, xPt: 10, yPt: 30, widthPt: 100, heightPt: 40 }, object: { kind: "text" as const, html: "<p>原对象</p>" } }] };
  if ("kind" in session) session.commit(session.prepare(transaction)); else await session.commit(await session.prepare({ kind: "transaction", transaction }));
  await workbench.settled(); const lease = "kind" in session ? undefined : await session.view();
  try { return ("kind" in session ? session.snapshot() : lease!.cells.snapshot).metadata.workbook.objects![0]!.nodeId; } finally { lease?.dispose(); }
}
export function replaceObject(nodeId: string) { const request = workbench.request(); if (!request) throw Error("No request"); return workbench.loadProposal(request, { ...edit(request, 10), operations: [{ op: "setWorkbookObjectText", nodeId, html: "<p>新对象😀</p>" }] }).then(() => {}); }
export function mode(value: string) { hostMode = value; }
export function deliver(value = 10) { if (!releaseHost || !submitted) throw Error("No host result held"); const resolve = releaseHost; releaseHost = undefined; resolve(edit(submitted, value)); }
export async function external(value: number) {
  const identity = "kind" in session ? session.snapshot().manifest : session.state(), transaction = { protocolVersion: 1 as const, documentId: identity.documentId, baseRevision: identity.revision, operations: [{ op: "setWorkbookCells" as const, sheetId: "sheet-1", cells: [{ row: 0, column: 0, source: { kind: "value" as const, value: { type: "number" as const, value } } }] }] };
  if ("kind" in session) session.commit(session.prepare(transaction)); else await session.commit(await session.prepare({ kind: "transaction", transaction }));
  if (!disposed) await workbench.settled();
}
export function focusInput() { workbench.editor.focusInput(true); }
export function cancelInput() { return workbench.editor.cancelInput(); }
export function settled() { return workbench.settled(); }
export function revision() { return "kind" in session ? session.snapshot().manifest.revision : session.state().revision; }
export async function disposeView() { unmount?.(); workbench.dispose(); disposed = true; await external(42); return inspect(); }
export async function applyAndDispose() { const receipt = workbench.applyProposal(); unmount?.(); workbench.dispose(); disposed = true; return receipt; }
export async function finish() {
  if (!disposed) { unmount?.(); workbench.dispose(); disposed = true; }
  if (!worker) { cleanup(); return { hostUsable: true }; }
  session.dispose(); const active = worker;
  const result = await new Promise<unknown>(resolve => { const timer = setTimeout(() => resolve({ error: "audit timeout" }), 10000); active.onmessage = event => { clearTimeout(timer); resolve(event.data); }; active.postMessage({ kind: "closed-audit" }); }); cleanup(); return result;
}
Object.assign(window, { workbookWorkbench: { mount, inspect, build, load, loadNoop, selectRows, formatRows, formatFont, drawingBounds, addFixedObject, addObject, replaceObject, mode, deliver, external, focusInput, cancelInput, settled, revision, disposeView, applyAndDispose, finish } });

import { DocumentError } from "@doffi/core";
import type { DeepReadonly, WorkbookRange, WorkbookSnapshot } from "@doffi/core";
import type { FontResolution } from "../contracts.js";
import { createWorkbookAnnotationView } from "./annotation-view.js";
import type { WorkbookAnnotationView } from "./annotation-view.js";
import { createWorkbookCommentView } from "./comment-view.js";
import type { WorkbookCommentView } from "./comment-view.js";
import type { WorkbookContentSources } from "./content-source.js";
import type { WorkbookSurface } from "./surface.js";
import type { WorkbookRenderedWindow, WorkbookWindowDecoration } from "./window-controller.js";
import { intersectRect, paneRect } from "./projection.js";

export interface WorkbookPreviewDecoration extends WorkbookWindowDecoration {
  readonly fonts: readonly FontResolution[];
  openComment(annotationId: string): void;
  showComments(annotationIds: readonly string[]): void;
  closeComment(): void;
  selectCell(row: number, column: number): void;
  selectRange(range: DeepReadonly<WorkbookRange>): void;
}
export function decorateWorkbookPreview(surface: WorkbookSurface, snapshot: WorkbookSnapshot, window: WorkbookRenderedWindow,
  content: WorkbookContentSources, maxNodes: number, attempt: (action: () => void) => void): WorkbookPreviewDecoration {
  if (maxNodes < 8) throw new DocumentError("ResourceLimitExceeded", "This window has no remaining annotation-control budget.");
  const document = surface.document, sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === window.frame.sheetId)!;
  const root = document.createElement("div"); root.dataset.officeView = "workbook-preview-controls";
  root.style.cssText = "position:absolute;inset:0;pointer-events:none";
  const selection = document.createElement("div"); selection.dataset.officeView = "workbook-cell-selection"; selection.setAttribute("aria-hidden", "true");
  selection.style.cssText = "position:absolute;inset:0;pointer-events:none";
  const annotations = new Map(snapshot.annotations.map(annotation => [annotation.annotationId, annotation]));
  let markers: WorkbookAnnotationView | undefined, panel: WorkbookCommentView | undefined, panelNodes = 0, panelEvents: AbortController | undefined, disposed = false;
  const check = (): void => { if (disposed) throw new DocumentError("InvalidOperation", "These workbook preview controls have been disposed."); };
  const width = Math.min(240, window.frame.layout.widthPt), top = Math.min(24, window.frame.layout.heightPt / 10);
  const show = (ids: readonly string[], selected: string | undefined): void => {
    check();
    if (ids.length > 5000 || new Set(ids).size !== ids.length || ids.some(id => !annotations.has(id))) throw new DocumentError("InvalidOperation", "Comment controls require unique annotations from this workbook.");
    const events = new document.defaultView!.AbortController();
    let nextMarkers: WorkbookAnnotationView | undefined, nextPanel: WorkbookCommentView | undefined, count = 0, units = 0;
    // Two roots plus up to four fragments of the selected merged cell. Reserve
    // this space so selecting a cell never evicts an open comment panel.
    const account = (amount: number): void => { count += amount; if (6 + (nextMarkers?.nodes ?? 0) + count > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Comment controls exceed their assigned window-node budget."); };
    const textUnits = (text: string): void => { units += text.length; if (units > 8 * 1024 * 1024) throw new DocumentError("ResourceLimitExceeded", "Comment controls exceed their text budget."); };
    try {
      nextMarkers = createWorkbookAnnotationView(document, window.annotations, window.frame.layout, sheet.rightToLeft,
        ids => attempt(() => ids.length === 1 ? show(ids, ids[0]) : show(ids, undefined)), maxNodes - 6);
      nextMarkers.select(selected);
      const children: HTMLElement[] = [selection, nextMarkers.element];
      if (ids.length) {
        account(1); const box = document.createElement("aside"); box.setAttribute("aria-label", "批注面板"); box.dataset.officeView = "workbook-preview-comments";
        box.style.cssText = `position:absolute;box-sizing:border-box;right:0;top:${top}pt;width:${width}pt;max-height:${window.frame.layout.heightPt - top}pt;overflow:auto;background:white;box-shadow:0 2px 8px #0003;pointer-events:auto;font:11pt sans-serif;color:#20242a`;
        const button = (label: string, action: () => void): void => {
          account(2); textUnits(label); const button = document.createElement("button"); button.type = "button"; button.textContent = label;
          button.style.cssText = "display:block;box-sizing:border-box;width:100%;padding:8px;text-align:start;font:inherit;color:inherit;background:#f4f6fa;border:1px solid #dce0e6;cursor:pointer";
          button.addEventListener("click", () => attempt(action), { signal: events.signal }); box.append(button);
        };
        button("关闭批注", closeComment);
        if (selected !== undefined) {
          const remaining = maxNodes - 6 - nextMarkers.nodes - count;
          if (remaining < 1) throw new DocumentError("ResourceLimitExceeded", "This window has no remaining comment-body budget.");
          nextPanel = createWorkbookCommentView(surface, snapshot, content, selected, { maxNodes: remaining }); account(nextPanel.nodes); box.append(nextPanel.element);
        } else {
          for (const id of ids) {
            const annotation = annotations.get(id)!;
            button((annotation.author?.name ?? "未署名") + " · " + ({ open: "待处理", resolved: "已处理", closed: "已关闭", needs_reanchor: "需要重新定位" }[annotation.status]), () => openComment(id));
          }
        }
        children.push(box);
      }
      root.replaceChildren(...children); markers?.dispose(); panel?.dispose(); panelEvents?.abort();
      markers = nextMarkers; panel = nextPanel; panelNodes = count; panelEvents = events; nextMarkers = undefined; nextPanel = undefined;
    } catch (error) { nextMarkers?.dispose(); nextPanel?.dispose(); events.abort(); throw error; }
  };
  const openComment = (annotationId: string): void => show([annotationId], annotationId);
  const closeComment = (): void => show([], undefined);
  const selectRange = (range: DeepReadonly<WorkbookRange>): void => {
    check(); const dimensions = window.dimensions, x = dimensions.columns.position(range.from.column), y = dimensions.rows.position(range.from.row);
    const width = dimensions.columns.position(range.to.column + 1) - x, height = dimensions.rows.position(range.to.row + 1) - y;
    const next = document.createDocumentFragment();
    for (const pane of window.frame.layout.panes) {
      const rect = { x: x - pane.sourceLeftPt + pane.inlineStartPt, y: y - pane.sourceTopPt + pane.blockStartPt, width, height }, clip = intersectRect(rect, paneRect(pane));
      if (!clip) continue;
      const node = document.createElement("div"), left = sheet.rightToLeft ? window.frame.layout.widthPt - clip.x - clip.width : clip.x;
      node.style.cssText = `position:absolute;box-sizing:border-box;left:${left}pt;top:${clip.y}pt;width:${clip.width}pt;height:${clip.height}pt;background:rgba(52,100,189,.06);pointer-events:none`;
      const edge = (name: string, visible: boolean): void => { if (visible) node.style.setProperty("border-" + name, "2px solid #3464bd"); };
      edge("top", Math.abs(rect.y - clip.y) < 1e-6); edge("bottom", Math.abs(rect.y + rect.height - clip.y - clip.height) < 1e-6);
      edge(sheet.rightToLeft ? "right" : "left", Math.abs(rect.x - clip.x) < 1e-6); edge(sheet.rightToLeft ? "left" : "right", Math.abs(rect.x + rect.width - clip.x - clip.width) < 1e-6);
      next.append(node);
    }
    selection.replaceChildren(next);
  };
  const selectCell = (row: number, column: number): void => {
    const projected = window.frame.projections.find(piece => piece.range.from.row === row && piece.range.from.column === column);
    if (projected) selectRange(projected.range); else { check(); selection.replaceChildren(); }
  };
  try { closeComment(); window.element.append(root); }
  catch (error) { root.remove(); throw error; }
  return Object.freeze({ get nodes() { return 2 + selection.childNodes.length + (markers?.nodes ?? 0) + panelNodes; }, get fonts() { return panel?.fonts ?? []; }, openComment, selectCell, selectRange,
    showComments: (ids: readonly string[]) => show(ids, ids.length === 1 ? ids[0] : undefined), closeComment,
    dispose() { if (disposed) return; disposed = true; markers?.dispose(); panel?.dispose(); panelEvents?.abort(); root.remove(); root.replaceChildren(); markers = undefined; panel = undefined; panelNodes = 0; } });
}

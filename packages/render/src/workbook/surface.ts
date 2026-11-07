import { DocumentError } from "@doffi/core";
import type { WorkbookSnapshot } from "@doffi/core";
import { FontProbeCache } from "../font-probe-cache.js";

export interface WorkbookSurface {
  readonly frame: HTMLIFrameElement;
  readonly document: Document;
  /** Source projections retain the document's :root/html/body cascade. */
  readonly source: HTMLElement;
  /** Only materialized display nodes belong in this isolated view. */
  readonly view: HTMLElement;
  /** Internal leases share the unchanged scriptless document and source CSS,
   * while each window owns a separate shadow tree and publication lifetime. */
  fork?(): WorkbookSurface;
  activate?(): void;
  fontState?(): { readonly ready: boolean; readonly epoch: number; readonly identity: object };
  dispose(): void;
}
export interface WorkbookSurfaceOptions {
  readonly widthPt: number;
  readonly heightPt: number;
  readonly maxCssCodeUnits?: number;
  readonly signal?: AbortSignal;
}

/** Prepare a hidden candidate, independent of the caller's current surface.
 * Source CSS remains in its own scriptless document. A separate shadow view
 * prevents even generic !important rules from changing generated layout boxes.
 * The compositor reveals/replaces the candidate only after its whole window
 * succeeds; neither this frame nor its measurement DOM is persisted source. */
export function createWorkbookSurface(container: HTMLElement, snapshot: WorkbookSnapshot, options: WorkbookSurfaceOptions): Promise<WorkbookSurface> {
  const { widthPt, heightPt, signal } = options, maxCss = options.maxCssCodeUnits ?? 4 * 1024 * 1024;
  if (!container.isConnected || !container.ownerDocument.defaultView || ![widthPt, heightPt].every(n => Number.isFinite(n) && n > 0 && n <= 1_000_000)
    || !Number.isSafeInteger(maxCss) || maxCss < 1) throw new DocumentError("InvalidOperation", "A workbook surface requires a connected container, finite viewport and positive CSS budget.");
  let codeUnits = 0;
  for (const [index, sheet] of snapshot.styles.entries()) { codeUnits += sheet.css.length + (index ? 1 : 0); if (codeUnits > maxCss) throw new DocumentError("ResourceLimitExceeded", "Workbook surface styles exceed their source budget."); }
  if (signal?.aborted) return Promise.reject(new DocumentError("OperationCancelled", "Workbook surface preparation was cancelled."));
  let styleText: string | undefined = snapshot.styles.map(style => style.css).join("\n");
  const frame = container.ownerDocument.createElement("iframe");
  // WebKit also disables trusted parent-installed event callbacks without
  // allow-scripts. Source scripts/handlers stay forbidden by the CSP below;
  // only SDK listeners installed from the owning document can handle input.
  frame.title = "工作簿只读预览"; frame.setAttribute("sandbox", "allow-same-origin allow-scripts");
  frame.style.cssText = `all:initial!important;display:block!important;position:absolute!important;left:0!important;top:0!important;border:0!important;width:${widthPt}pt!important;height:${heightPt}pt!important;visibility:hidden!important`;
  return new Promise((resolve, reject) => {
    let disposed = false, pending = true;
    let ownedDocument: Document | undefined, sourceRoot: HTMLElement | undefined, viewRoot: HTMLElement | undefined;
    const leases = new Map<HTMLElement, HTMLElement>(), fonts = new FontProbeCache();
    const detach = (): void => { frame.removeEventListener("load", load); signal?.removeEventListener("abort", abort); clearTimeout(timer); };
    const destroy = (): void => {
      if (disposed) return; disposed = true; styleText = undefined; detach(); frame.remove(); frame.removeAttribute("srcdoc");
      sourceRoot?.replaceChildren(); viewRoot?.replaceChildren(); ownedDocument?.head.replaceChildren(); ownedDocument?.body.replaceChildren();
      fonts.clear();
    };
    const fail = (error: unknown): void => { if (!pending) return; pending = false; destroy(); reject(error); };
    const abort = (): void => fail(new DocumentError("OperationCancelled", "Workbook surface preparation was cancelled."));
    function lease(document: Document, source: HTMLElement, host: HTMLElement, view: HTMLElement): WorkbookSurface {
      let released = false; leases.set(view, host);
      const check = (): void => { if (released || disposed) throw new DocumentError("InvalidOperation", "This workbook surface lease has been released."); };
      return Object.freeze({ frame, document, source, view,
        fork() {
          check(); const nextHost = document.createElement("office-workbook-view-surface"), nextView = document.createElement("div");
          nextHost.style.cssText = host.style.cssText; nextHost.style.setProperty("visibility", "hidden", "important"); nextHost.style.setProperty("pointer-events", "none", "important");
          nextView.style.cssText = view.style.cssText;
          nextHost.attachShadow({ mode: "closed" }).append(nextView); document.body.append(nextHost);
          return lease(document, source, nextHost, nextView);
        },
        activate() {
          check();
          // New window contents remain detached until the owner commits. No
          // observer can see two complete published grids in this document.
          for (const other of leases.values()) if (other !== host) other.remove();
          if (!host.isConnected) document.body.append(host);
          host.style.removeProperty("visibility"); host.style.removeProperty("pointer-events");
        },
        fontState() { check(); return { ...fonts.watch(document.fonts), identity: document.fonts }; },
        dispose() {
          if (released) return; released = true; leases.delete(view); host.remove(); view.replaceChildren();
          if (!leases.size) destroy();
        },
      });
    }
    const load = (): void => {
      if (disposed || !pending) return;
      try {
        const document = frame.contentDocument;
        if (!document?.defaultView) throw new DocumentError("UnsupportedFeature", "This browser cannot access the scriptless workbook surface.");
        ownedDocument = document;
        // These boxes belong to the viewport, not to an Office content node.
        // Keep typography/custom properties for source inheritance while
        // preventing source page-box rules from resizing/offsetting the view.
        for (const root of [document.documentElement, document.body]) for (const [name, value] of Object.entries({ margin: "0", padding: "0", border: "0", width: "100%", height: "100%", "min-width": "0", "min-height": "0", "max-width": "none", "max-height": "none", "box-sizing": "border-box", overflow: "hidden", "overscroll-behavior": "none" })) root.style.setProperty(name, value, "important");
        const styles = document.createElement("style"); styles.dataset.officeSourceStyles = "true";
        styles.textContent = styleText!; styleText = undefined; document.head.append(styles);
        const source = document.createElement("office-workbook-source-surface"), host = document.createElement("office-workbook-view-surface"), view = document.createElement("div");
        sourceRoot = source; viewRoot = view;
        source.setAttribute("aria-hidden", "true");
        source.style.cssText = "position:absolute;visibility:hidden;pointer-events:none;contain:layout;left:0;top:0;width:0;height:0;overflow:hidden";
        // Inline important wins over source selectors targeting the host;
        // the shadow boundary then protects all generated descendants.
        host.style.cssText = "all:initial!important;display:block!important;position:absolute!important;left:0!important;top:0!important;width:100%!important;height:100%!important;direction:ltr!important;unicode-bidi:isolate!important";
        view.style.cssText = "position:relative;width:100%;height:100%;overflow:hidden;direction:ltr";
        host.attachShadow({ mode: "closed" }).append(view); document.body.append(source, host);
        void document.fonts.ready.then(() => {
          if (disposed || !pending) return;
          pending = false; detach(); resolve(lease(document, source, host, view));
        }).catch(fail);
      } catch (error) { fail(error); }
    };
    const timer = setTimeout(() => fail(new DocumentError("UnsupportedFeature", "The workbook surface did not become ready within its load deadline.")), 30_000);
    signal?.addEventListener("abort", abort, { once: true }); frame.addEventListener("load", load, { once: true });
    // No source HTML or CSS is interpolated into srcdoc. Policy precedes all
    // later DOM adoption; authorized image leases are the only resource URLs.
    try {
      frame.srcdoc = '<!doctype html><html><head><meta http-equiv="Content-Security-Policy" content="default-src \'none\'; script-src \'none\'; img-src blob:; style-src \'unsafe-inline\'; font-src \'none\'; base-uri \'none\'; form-action \'none\'"></head><body></body></html>';
      container.append(frame);
    } catch (error) { fail(error); }
  });
}

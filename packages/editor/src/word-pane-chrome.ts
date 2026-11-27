/** Shared macOS Microsoft Word 16.114.1 side task-pane & inspector chrome for Shadow DOM surfaces. */
export const wordPaneStyles = `
:host([data-word-dock-host="true"]){display:grid!important;grid-template-columns:auto minmax(0,1fr) auto;grid-template-rows:auto minmax(0,1fr) auto;align-items:stretch}
:host([data-word-dock-host="true"])>.toolbar{grid-column:1/-1;grid-row:1}
:host([data-word-dock-host="true"])>.stage,:host([data-word-dock-host="true"])>.workspace-body{grid-column:2;grid-row:2;min-width:0}
:host([data-word-dock-host="true"])>.error{grid-column:1/-1;grid-row:3}
.word-dock-pane{box-sizing:border-box;display:flex;flex-direction:column;width:304px;max-width:min(350px,44vw);height:var(--doffi-canvas-height,calc(100dvh - 190px));max-height:var(--doffi-canvas-height,calc(100dvh - 190px));min-height:320px;background:#fafafa;color:#242424;font:12px/1.45 -apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",sans-serif;overflow:hidden;position:relative;z-index:4}
.word-dock-pane[data-word-dock="left"]{grid-column:1;grid-row:2;border-right:1px solid #dfdfdf;width:288px}
.word-dock-pane[data-word-dock="right"]{grid-column:3;grid-row:2;border-left:1px solid #dfdfdf;width:316px}
.word-dock-pane[data-word-dock="floating"]{position:fixed;top:156px;right:24px;z-index:85;width:310px;height:auto;min-height:0;max-height:calc(100dvh - 190px);border:1px solid #d0d0d0;border-radius:10px;background:#fff;box-shadow:0 12px 38px rgba(0,0,0,.18)}
.word-dock-pane *{box-sizing:border-box}
.word-dock-pane [hidden]{display:none!important}
.word-pane-header{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:10px 14px;background:#f5f5f5;border-bottom:1px solid #e4e4e4;flex-shrink:0}
.word-pane-title{margin:0;font-size:13px;font-weight:600;color:#222;letter-spacing:.01em}
.word-pane-close{display:inline-flex;align-items:center;justify-content:center;width:24px;height:24px;padding:0;border:1px solid transparent;border-radius:5px;background:transparent;color:#555;font-size:16px;line-height:1;cursor:pointer}
.word-pane-close:hover{background:#e6e6e6;color:#111;border-color:transparent}
.word-pane-body{flex:1;min-height:0;overflow-y:auto;overflow-x:hidden;padding:12px 14px;display:flex;flex-direction:column;gap:10px}
.word-pane-section{background:#fff;border:1px solid #e5e5e5;border-radius:7px;padding:10px 11px;display:flex;flex-direction:column;gap:8px}
.word-pane-section-title{margin:0;font-size:11px;font-weight:600;color:#5c5c5c;text-transform:none}
.word-pane-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px 10px;align-items:end}
.word-pane-grid .wide{grid-column:1/-1}
.word-pane-row{display:flex;align-items:center;gap:6px;flex-wrap:wrap}
.word-pane-tabs{display:flex;gap:3px;padding:8px 12px 0;background:#f5f5f5;border-bottom:1px solid #e4e4e4;flex-wrap:wrap;flex-shrink:0}
.word-pane-tabs button{padding:5px 9px;border:1px solid transparent;border-bottom:2px solid transparent;border-radius:5px 5px 0 0;background:transparent;color:#555;font-size:11.5px;cursor:pointer}
.word-pane-tabs button:hover{background:#eaeaea;color:#111}
.word-pane-tabs button[aria-selected="true"]{background:#fff;color:#185abd;border-color:#e4e4e4;border-bottom-color:#185abd;font-weight:600}
.word-dock-pane label{display:flex;flex-direction:column;gap:4px;font-size:11.5px;color:#3b3b3b;margin:0;min-width:0}
.word-dock-pane label:has(>input[type="checkbox"]){flex-direction:row;align-items:center;gap:6px;cursor:pointer}
.word-dock-pane input:not([type="checkbox"]):not([type="color"]):not([type="file"]),.word-dock-pane select,.word-dock-pane textarea{width:100%!important;min-width:0;min-height:26px;padding:4px 8px;border:1px solid #d4d4d4;border-radius:5px;background:#fff;color:#222;font:inherit;font-size:12px}
.word-dock-pane input[type="color"]{width:100%!important;height:26px;padding:2px;border:1px solid #d4d4d4;border-radius:5px;background:#fff;cursor:pointer}
.word-dock-pane input[type="checkbox"]{width:14px;height:14px;margin:0;accent-color:#185abd;flex-shrink:0}
.word-dock-pane input:focus,.word-dock-pane select:focus,.word-dock-pane textarea:focus{outline:2px solid #8daed9;outline-offset:1px;border-color:#185abd}
.word-dock-pane button{border:1px solid #d2d2d2;border-radius:5px;background:#fff;color:#292929;padding:5px 10px;font-size:12px;cursor:pointer;line-height:1.3}
.word-dock-pane button:hover:not(:disabled){background:#f0f4fa;border-color:#b4cbe8}
.word-dock-pane button:disabled{opacity:.42;cursor:default}
.word-dock-pane button.word-btn-primary{background:#185abd;border-color:#185abd;color:#fff;font-weight:500}
.word-dock-pane button.word-btn-primary:hover:not(:disabled){background:#144ca3;border-color:#144ca3}
.word-dock-pane button.word-btn-danger{color:#b42318;border-color:#f0c7c4;background:#fff}
.word-dock-pane button.word-btn-danger:hover:not(:disabled){background:#fef3f2;border-color:#e39b95}
.word-dock-pane details{border:1px solid #e6e6e6;border-radius:6px;background:#fff;padding:6px 9px}
.word-dock-pane details summary{cursor:pointer;font-size:11.5px;font-weight:500;color:#444;user-select:none}
.word-dock-pane details[open] summary{margin-bottom:8px;padding-bottom:6px;border-bottom:1px solid #efefef}
.word-pane-list{display:flex;flex-direction:column;gap:3px;max-height:220px;overflow-y:auto;padding:4px;border:1px solid #e5e5e5;border-radius:6px;background:#fff}
.word-style-item{display:flex!important;align-items:center;justify-content:space-between;gap:8px;width:100%;padding:6px 8px!important;border:1px solid transparent!important;border-radius:5px!important;background:transparent!important;text-align:left!important;min-height:30px}
.word-style-item:hover{background:#eef4fc!important;border-color:#d2e2f8!important}
.word-style-item[aria-selected="true"]{background:#e3eefc!important;border-color:#9fc0eb!important}
.word-style-badge{font-size:10px;color:#666;background:#f0f0f0;border-radius:3px;padding:1px 5px;flex-shrink:0}
.word-stat-table{width:100%;border-collapse:collapse;font-size:12px}
.word-stat-table td{padding:5px 2px;border-bottom:1px solid #f0f0f0}
.word-stat-table td:last-child{text-align:right;font-weight:600;font-variant-numeric:tabular-nums;color:#185abd}
.word-pane-muted{margin:0;font-size:11px;color:#6b7280;line-height:1.45}
@media(max-width:860px){.word-dock-pane[data-word-dock="left"],.word-dock-pane[data-word-dock="right"]{position:fixed;top:148px;bottom:28px;height:auto;max-height:none;z-index:70;box-shadow:0 8px 28px rgba(0,0,0,.18)}.word-dock-pane[data-word-dock="left"]{left:0}.word-dock-pane[data-word-dock="right"]{right:0}}
`;

export function createWordDockedPane(
  document: Document,
  options: {
    readonly ariaLabel: string;
    readonly title: string;
    readonly dock: "left" | "right" | "floating";
    readonly closeLabel: string;
    readonly onClose: () => void;
    readonly signal: AbortSignal;
    readonly width?: string;
  },
): {
  readonly panel: HTMLElement;
  readonly header: HTMLElement;
  readonly body: HTMLElement;
  readonly closeButton: HTMLButtonElement;
} {
  const panel = document.createElement("section");
  panel.className = "word-dock-pane";
  panel.dataset.wordDock = options.dock;
  panel.setAttribute("role", "region");
  panel.setAttribute("aria-label", options.ariaLabel);
  if (options.width) panel.style.width = options.width;

  const header = document.createElement("div");
  header.className = "word-pane-header";
  const heading = document.createElement("h3");
  heading.className = "word-pane-title";
  heading.textContent = options.title;
  const closeButton = document.createElement("button");
  closeButton.type = "button";
  closeButton.className = "word-pane-close";
  closeButton.setAttribute("aria-label", options.closeLabel);
  closeButton.title = options.closeLabel;
  closeButton.textContent = "×";
  closeButton.addEventListener("click", options.onClose, { signal: options.signal });
  header.append(heading, closeButton);

  const body = document.createElement("div");
  body.className = "word-pane-body";
  panel.append(header, body);
  return { panel, header, body, closeButton };
}

/** Close any sibling docked pane on the same side so sidebars never stack awkwardly. */
export function replaceSameSideDock(parent: ShadowRoot, dock: "left" | "right" | "floating", keep?: HTMLElement): void {
  for (const existing of parent.querySelectorAll<HTMLElement>(`.word-dock-pane[data-word-dock="${dock}"]`)) {
    if (existing !== keep) existing.remove();
  }
}

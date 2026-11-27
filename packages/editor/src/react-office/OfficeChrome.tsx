import { useEffect } from "react";
import type { CSSProperties, ReactElement, ReactNode } from "react";
import { OfficeIcon as Icon } from "./icons.js";
import { OfficeMenu } from "./OfficeMenu.js";
import { officeRibbonStyles } from "./styles.js";
import { OfficeRibbon } from "./OfficeRibbon.js";
import type { OfficeActions, OfficeRibbonProps } from "./OfficeRibbon.js";

const chromeStyles = `
.doffi-office-titlebar{height:40px;min-height:40px;padding:3px 15px;display:flex;align-items:center;gap:7px;position:relative;box-sizing:border-box;background:#f5f5f5;color:#444;font:12px -apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif}
.doffi-office-titlebar *{box-sizing:border-box}.doffi-office-titlebar button{display:inline-flex;align-items:center;justify-content:center;gap:5px;border:1px solid transparent;border-radius:5px;background:transparent;padding:4px;color:inherit;font:inherit;cursor:pointer;white-space:nowrap}.doffi-office-titlebar button:disabled{opacity:.35;cursor:default}.doffi-office-titlebar button:hover:not(:disabled){background:#e8e8e8}
.doffi-office-titlebar .title-quick{display:flex;align-items:center;gap:4px}.doffi-office-titlebar .title-heading{display:flex;align-items:center;justify-content:center;gap:4px;position:absolute;left:50%;transform:translateX(-50%);max-width:45%;pointer-events:none;min-width:0}.doffi-office-titlebar h1{font:inherit;margin:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.doffi-office-titlebar .title-heading>svg{color:#185abd;flex-shrink:0}.doffi-office-titlebar[data-kind=presentation] .title-heading>svg{color:#b7472a}.doffi-office-titlebar[data-kind=workbook] .title-heading>svg{color:#107c41}
.doffi-office-titlebar .title-actions{display:flex;align-items:center;gap:8px;margin-left:auto}.doffi-office-titlebar .title-mode{border:1px solid #d5d5d5;border-radius:15px;padding:4px 11px;background:#fafafa}.doffi-office-titlebar .title-primary{color:white;background:#185abd;border-radius:15px;padding:4px 12px}.doffi-office-titlebar[data-kind=presentation] .title-primary{background:#b7472a}.doffi-office-titlebar[data-kind=workbook] .title-primary{background:#107c41}.doffi-office-titlebar .title-primary:hover:not(:disabled){filter:brightness(.95);background:#185abd}
.doffi-office-titlebar .office-menu>.ribbon-tool>span:not(:first-child),.doffi-office-titlebar .menu-arrow{display:none}.doffi-office-titlebar .office-menu>.ribbon-tool{min-width:28px}.doffi-office-titlebar .autosave{display:flex;align-items:center;gap:7px;margin-right:4px;font-size:11px}.doffi-office-titlebar .autosave button{width:34px;height:17px;padding:1px;border-radius:12px;background:#d9d9d9;justify-content:flex-start}.doffi-office-titlebar .autosave button[aria-checked=true]{background:#185abd;justify-content:flex-end}.doffi-office-titlebar .autosave i{width:14px;height:14px;border-radius:50%;background:white;box-shadow:0 1px 3px #0003}
.doffi-office-statusbar{min-height:25px;height:25px;display:flex;align-items:center;justify-content:space-between;gap:12px;padding:2px 16px;border-top:1px solid #ddd;background:#f5f5f5;color:#666;box-sizing:border-box;font:11px -apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif;white-space:nowrap}.doffi-office-statusbar .status-summary{display:flex;align-items:center;gap:18px;min-width:0;overflow:hidden}.doffi-office-statusbar .status-controls{display:flex;align-items:center;gap:2px}.doffi-office-statusbar button{display:inline-flex;align-items:center;justify-content:center;padding:3px 6px;border:0;border-radius:0;background:transparent;font:inherit;color:inherit;cursor:pointer}.doffi-office-statusbar button[aria-pressed=true]{background:#dedede}.doffi-office-statusbar button:disabled{opacity:.35}.doffi-office-statusbar input[type=range]{appearance:none;width:120px;height:2px;background:#bbb;margin:0 4px}.doffi-office-statusbar input[type=range]::-webkit-slider-thumb{appearance:none;width:13px;height:13px;border:1px solid #ccc;background:#fff;border-radius:50%;box-shadow:0 1px 3px #0002}.doffi-office-statusbar input[type=range]::-moz-range-thumb{width:12px;height:12px;border:1px solid #ccc;background:white;border-radius:50%}.doffi-office-statusbar .zoom-number{min-width:42px;text-align:right}.doffi-office-statusbar .status-page-input{display:flex;align-items:center;gap:2px}.doffi-office-statusbar .status-page-input input{font:inherit;width:35px;min-height:0;border:1px solid #ccc;border-radius:3px;padding:1px 3px}
.doffi-office-workspace{display:flex;flex-direction:column;min-width:0;min-height:100dvh;height:100%;background:#f5f5f5;position:relative}.doffi-office-workspace>.office-workspace-content{display:flex;flex-direction:column;flex:1;min-height:0;overflow:auto}.doffi-office-focus-exit{position:fixed;top:12px;right:18px;z-index:20;padding:6px 12px;border:1px solid #d0d7e2;border-radius:16px;background:rgba(255,255,255,.94);color:#243041;font:12px -apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif;cursor:pointer;box-shadow:0 2px 8px #0002}.doffi-office-focus-exit:hover{background:#eef3fa}
@media(max-width:680px){.doffi-office-titlebar{padding-inline:7px;gap:3px}.doffi-office-titlebar .title-heading{position:static;transform:none;max-width:none;flex:1;justify-content:flex-start}.doffi-office-titlebar .title-quick{gap:0}.doffi-office-titlebar .title-actions{gap:3px}.doffi-office-titlebar .title-primary{font-size:0;padding:4px}.doffi-office-titlebar .title-primary svg{width:15px}.doffi-office-titlebar .autosave{display:none}.doffi-office-statusbar{padding-inline:8px;gap:3px;font-size:10px}.doffi-office-statusbar .status-summary{gap:7px}.doffi-office-statusbar input[type=range]{width:65px}.doffi-office-statusbar .status-page-input{display:none}}
`;

export interface OfficeTitleBarProps {
  title: string; kind: OfficeRibbonProps["kind"]; actions: OfficeActions; busy?: boolean;
  editing?: boolean; onEditingChange?(): void; onHome?(): void; onToggleDocuments?(): void; documentsVisible?: boolean;
  onNewWord?(): void; onNewPresentation?(): void; onFile?(): void;
  commentsVisible?: boolean; autosave?: { enabled: boolean; disabled?: boolean; change(enabled: boolean): void };
}
export function OfficeTitleBar({ title, kind, actions, busy, editing, onEditingChange, onHome, onToggleDocuments, documentsVisible, onNewWord, onNewPresentation, onFile, commentsVisible, autosave }: OfficeTitleBarProps): ReactElement {
  const action = (id: string, label: string, icon: Parameters<typeof Icon>[0]["name"]) => actions[id] && <button type="button" aria-label={label} title={label} disabled={busy || actions[id]?.disabled}
    onMouseDown={event => event.preventDefault()} onClick={actions[id]?.run}><Icon name={icon} size={16} /></button>;
  return <header className="doffi-office-titlebar" data-kind={kind}><style>{chromeStyles + officeRibbonStyles}</style>
    <span className="autosave">自动保存<button role="switch" type="button" aria-label="自动保存" aria-checked={autosave?.enabled ?? false} disabled={!autosave || autosave.disabled || busy} onClick={() => autosave?.change(!autosave.enabled)}><i /></button></span>
    <div className="title-quick" role="toolbar" aria-label="快速访问">
      {onHome && <button type="button" aria-label="返回首页" title="返回首页" onClick={onHome}><Icon name="home" size={16} /></button>}
      {action("package", "下载文档包", "save")}{action("undo", "撤销", "undo")}{action("redo", "重做", "redo")}
      <button type="button" aria-label="打印" title="打印" disabled={busy || !actions.print || actions.print.disabled} onClick={actions.print?.run}><Icon name="print" size={16} /></button>
      {action("native", "另存为", "saveAs")}
      <div className="doffi-office-ribbon"><OfficeMenu label="文件菜单" icon={<span>•••</span>} disabled={Boolean(busy)}>{close => <>
        {onNewWord && <button role="menuitem" type="button" onClick={() => { onNewWord(); close(); }}>新建 Word</button>}
        {onNewPresentation && <button role="menuitem" type="button" onClick={() => { onNewPresentation(); close(); }}>新建 PPT</button>}
        {actions.open && <button role="menuitem" type="button" onClick={() => { actions.open!.run(); close(); }}>打开…</button>}
        {onFile && <button role="menuitem" type="button" onClick={() => { onFile(); close(); }}>文件和导出…</button>}
        {onToggleDocuments && <button role="menuitem" type="button" aria-expanded={documentsVisible} onClick={() => { onToggleDocuments(); close(); }}>显示文档列表</button>}
      </>}</OfficeMenu></div>
    </div>
    <div className="title-heading"><Icon name={kind === "presentation" ? "powerPoint" : kind === "workbook" ? "excel" : "word"} size={16} /><h1>{title}</h1></div>
    <div className="title-actions">{actions.comments && <button type="button" aria-label="批注" aria-pressed={commentsVisible} title="批注" disabled={busy || actions.comments.disabled} onClick={actions.comments.run}><Icon name="comment" size={17} /></button>}
      {onEditingChange && <button type="button" className="title-mode" aria-label={editing ? "返回预览与批注" : "编辑正文"} disabled={busy} onClick={onEditingChange}><Icon name={editing ? "edit" : "book"} size={14} />{editing ? "编辑" : "查看"}<span>⌄</span></button>}
      {actions.open && <button type="button" className="title-primary" aria-label="选择本地文件" disabled={busy || actions.open.disabled} onClick={actions.open.run}><Icon name="folder" size={16} />打开文件</button>}
      <button type="button" aria-label="查找" title="查找" disabled={busy || !actions.find || actions.find.disabled} onClick={actions.find?.run}><Icon name="search" size={16} /></button>
    </div>
  </header>;
}

export interface OfficeStatusBarProps {
  kind: OfficeRibbonProps["kind"]; pageCount?: number; currentPage?: number; wordCount?: number; language?: string;
  zoom: number; onZoom(value: number): void; disabled?: boolean; layout?: "pages" | "continuous"; onLayout?(layout: "pages" | "continuous"): void;
  onPage?(page: number): void; onFit?(): void; onSlideShow?(): void; summary?: ReactNode;
}
export function OfficeStatusBar({ kind, pageCount = 0, currentPage, wordCount, language = "简体中文(中国大陆)", zoom, onZoom, disabled, layout, onLayout, onPage, onFit, onSlideShow, summary }: OfficeStatusBarProps): ReactElement {
  return <footer className="doffi-office-statusbar"><style>{chromeStyles}</style><span className="status-summary">{summary ?? <>
    {kind === "workbook" ? <span>就绪</span> : pageCount > 0 && <span>{kind === "presentation" ? currentPage ? `幻灯片 ${currentPage} / ${pageCount}` : `${pageCount} 张幻灯片` : currentPage ? `第 ${currentPage} 页，共 ${pageCount} 页` : `共 ${pageCount} 页`}</span>}
    {wordCount !== undefined && <span>{wordCount} 个字</span>}{kind !== "workbook" && <span>{language}</span>}
  </>}</span><div className="status-controls">
    {kind === "document" && onLayout && <><button type="button" title="页面视图" aria-label="页面视图" aria-pressed={layout === "pages"} onClick={() => onLayout("pages")}><Icon name="file" size={14} /></button><button type="button" title="连续视图" aria-label="连续视图" aria-pressed={layout === "continuous"} onClick={() => onLayout("continuous")}><Icon name="globe" size={14} /></button></>}
    {kind === "workbook" && <button type="button" aria-label="普通视图" title="普通视图" aria-pressed="true"><Icon name="grid" size={14} /></button>}
    {onPage && pageCount > 1 && <form className="status-page-input" aria-label="转到页面" onSubmit={event => { event.preventDefault(); onPage(Number(new FormData(event.currentTarget).get("page"))); }}><input name="page" type="number" aria-label="目标页码" required min={1} max={pageCount} defaultValue={1} /><button type="submit">转到</button></form>}
    {onSlideShow && <button type="button" aria-label="从当前幻灯片放映" title="从当前幻灯片放映" disabled={disabled} onClick={onSlideShow}><Icon name="slides" size={15} /></button>}
    <button type="button" aria-label="缩小" disabled={disabled || zoom <= .25} onClick={() => onZoom(Math.max(.25, Math.round((zoom - .25) * 100) / 100))}>−</button>
    <input type="range" aria-label="文档缩放" disabled={disabled} min="25" max="400" step="1" value={Math.min(400, Math.max(25, Math.round(zoom * 100)))} onChange={event => onZoom(Number(event.currentTarget.value) / 100)} />
    <button type="button" aria-label="放大" disabled={disabled || zoom >= 4} onClick={() => onZoom(Math.min(4, Math.round((zoom + .25) * 100) / 100))}>＋</button><span className="zoom-number">{Math.round(zoom * 100)}%</span>
    {onFit && <button type="button" aria-label="幻灯片适应窗口" title="幻灯片适应窗口" disabled={disabled} onClick={onFit}><Icon name="slides" size={14} /></button>}
  </div></footer>;
}

/** Hosts provide actions and the native editor; all shared Office chrome ships in the SDK. */
export interface OfficeWorkspaceProps {
  titleBar: OfficeTitleBarProps; ribbon: OfficeRibbonProps; statusBar: OfficeStatusBarProps;
  focusMode?: boolean; onExitFocusMode?(): void; children: ReactNode;
}
export function OfficeWorkspace({ titleBar, ribbon, statusBar, focusMode, onExitFocusMode, children }: OfficeWorkspaceProps): ReactElement {
  useEffect(() => {
    if (!focusMode || !onExitFocusMode) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") onExitFocusMode();
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [focusMode, onExitFocusMode]);
  const style = {
    "--doffi-text-controls": "none", "--doffi-editing-controls": "none", "--doffi-interaction-controls": "none", "--doffi-collaboration-controls": !focusMode && ribbon.tab === "审阅" ? "flex" : "none",
    "--doffi-presentation-controls": "none", "--doffi-workbook-tools": "none", "--doffi-workbook-search": "none", "--doffi-workbook-panels": "none", "--doffi-workbook-mode-label": "none", "--doffi-workbook-preview-toolbar": "none",
    "--doffi-canvas-height": focusMode ? "100dvh" : "calc(100dvh - 190px)", "--doffi-presentation-canvas-height": focusMode ? "100dvh" : "calc(100dvh - 220px)", "--doffi-annotations-display": !focusMode && titleBar.commentsVisible ? "block" : "none", "--doffi-annotations-width": !focusMode && titleBar.commentsVisible ? "280px" : "0px",
    "--doffi-presentation-comments-display": !focusMode && titleBar.commentsVisible ? "block" : "none", "--doffi-presentation-comments-width": !focusMode && titleBar.commentsVisible ? "260px" : "0px",
  } as CSSProperties;
  return <div className="doffi-office-workspace" data-focus-mode={focusMode ? "true" : undefined} style={style}><style>{chromeStyles}</style>
    {focusMode && onExitFocusMode && <button type="button" className="doffi-office-focus-exit" onClick={onExitFocusMode}>退出专注模式 (Esc)</button>}
    {!focusMode && <OfficeTitleBar {...titleBar} />}
    {!focusMode && <OfficeRibbon {...ribbon} />}
    <div className="office-workspace-content">{children}</div>
    {!focusMode && <OfficeStatusBar {...statusBar} />}
  </div>;
}

import { OfficeIcon as Icon } from "./icons.js";
import { useLayoutEffect, useId, useRef, useState } from "react";
import type { ReactElement, ReactNode } from "react";

/** Office menus preserve the document selection until a command is chosen. */
export function OfficeMenu({ label, icon, children, disabled = false, large = false, split, active, className = "" }: {
  label: string; icon: ReactNode; children(close: () => void): ReactNode;
  disabled?: boolean; large?: boolean; active?: boolean | undefined; split?: () => void; className?: string;
}): ReactElement {
  const [open, setOpen] = useState(false), [position, setPosition] = useState({ left: 0, top: 0 });
  const root = useRef<HTMLDivElement>(null), button = useRef<HTMLButtonElement>(null), id = useId();
  const close = () => setOpen(false);
  useLayoutEffect(() => {
    if (!open) return;
    const rect = root.current!.getBoundingClientRect();
    setPosition({ left: Math.max(8, Math.min(rect.left, window.innerWidth - 300)), top: rect.bottom + 3 });
    const dismiss = (event: Event) => { if (!root.current?.contains(event.target as Node)) close(); };
    const key = (event: KeyboardEvent) => { if (event.key === "Escape") { event.preventDefault(); close(); button.current?.focus(); } };
    document.addEventListener("pointerdown", dismiss); document.addEventListener("keydown", key);
    window.addEventListener("resize", close);
    return () => { document.removeEventListener("pointerdown", dismiss); document.removeEventListener("keydown", key); window.removeEventListener("resize", close); };
  }, [open]);
  return <div ref={root} className={"office-menu " + (large ? "large " : "") + className}>
    <button ref={button} className={"ribbon-tool " + (large ? "large" : "")} type="button" aria-label={label}
      disabled={disabled} title={label} aria-pressed={active} aria-haspopup={split ? undefined : "menu"} aria-expanded={split ? undefined : open}
      onMouseDown={event => event.preventDefault()} onClick={() => split ? split() : setOpen(!open)}>
      {icon}<span>{label}</span>{!split && <small className="menu-arrow" aria-hidden="true">⌄</small>}
    </button>
    {split && <button type="button" className="menu-split-arrow" aria-label={label + "选项"} aria-haspopup="menu" aria-expanded={open}
      aria-controls={id} disabled={disabled} onMouseDown={event => event.preventDefault()} onClick={() => setOpen(!open)}>⌄</button>}
    {open && <div id={id} role="menu" aria-label={label + "选项"} className="office-menu-popup" style={{ left: position.left, top: position.top }}
      onKeyDown={event => {
        if (!["ArrowDown", "ArrowUp", "Home", "End"].includes(event.key)) return;
        const items = [...event.currentTarget.querySelectorAll<HTMLButtonElement>('button:not(:disabled)')];
        const current = items.indexOf(document.activeElement as HTMLButtonElement);
        const index = event.key === "Home" ? 0 : event.key === "End" ? items.length - 1 : (current + (event.key === "ArrowDown" ? 1 : -1) + items.length) % items.length;
        event.preventDefault(); items[index]?.focus();
      }}>{children(close)}</div>}
  </div>;
}

const themeColors = ["#ffffff", "#000000", "#e7e6e6", "#44546a", "#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47"];
const standardColors = ["#c00000", "#ff0000", "#ffc000", "#ffff00", "#92d050", "#00b050", "#00b0f0", "#0070c0", "#002060", "#7030a0"];
function tint(color: string, factor: number): string {
  const values = [1, 3, 5].map(index => parseInt(color.slice(index, index + 2), 16));
  return "#" + values.map(value => Math.round(factor > 0 ? value + (255 - value) * factor : value * (1 + factor)).toString(16).padStart(2, "0")).join("");
}
export function OfficeColorMenu({ label, value, disabled, apply, highlight = false, fill = false }: {
  label: string; value: string; disabled: boolean; apply(color: string): void; highlight?: boolean; fill?: boolean;
}): ReactElement {
  const [color, setColor] = useState(value);
  const icon = <span className={"office-color-symbol " + (highlight ? "highlight-symbol" : fill ? "fill-symbol" : "")}
    style={{ borderBottomColor: color }}>{highlight ? <Icon name="highlight" size={18} /> : fill ? <Icon name="fill" size={18} /> : "A"}</span>;
  const select = (next: string, close: () => void) => { setColor(next); apply(next); close(); };
  return <OfficeMenu label={label} icon={icon} disabled={disabled} split={() => apply(color)} className="color-menu">
    {close => <>
      {!highlight && <button type="button" role="menuitem" onClick={() => select("#000000", close)}>自动</button>}
      <div className="office-menu-heading">{highlight ? "突出显示颜色" : "主题颜色"}</div>
      <div className="office-color-grid">
        {(highlight ? standardColors : themeColors).map(color => <button key={color} role="menuitem" type="button" aria-label={color} title={color}
          style={{ backgroundColor: color }} onClick={() => select(color, close)} />)}
        {!highlight && [.8, .6, .4, -.25, -.5].flatMap(factor => themeColors.map(color => {
          const next = tint(color, factor); return <button key={factor + color} role="menuitem" type="button" aria-label={next} title={next}
            style={{ backgroundColor: next }} onClick={() => select(next, close)} />;
        }))}
      </div>
      {!highlight && <><div className="office-menu-heading">标准色</div><div className="office-color-grid">{standardColors.map(color =>
        <button key={color} role="menuitem" type="button" aria-label={color} title={color} style={{ backgroundColor: color }} onClick={() => select(color, close)} />)}</div></>}
      {highlight && <button type="button" role="menuitem" onClick={() => select("transparent", close)}>无颜色</button>}
      <label className="office-custom-color">其他颜色…<input type="color" aria-label={label} value={/^#[0-9a-f]{6}$/i.test(color) ? color : "#000000"}
        onChange={event => select(event.currentTarget.value, close)} /></label>
    </>}
  </OfficeMenu>;
}

import { wordColorLuminance } from "@doffi/core";
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

const defaultThemeColors = ["#ffffff", "#000000", "#e7e6e6", "#44546a", "#4472c4", "#ed7d31", "#a5a5a5", "#ffc000", "#5b9bd5", "#70ad47"];
const standardColors = ["#c00000", "#ff0000", "#ffc000", "#ffff00", "#92d050", "#00b050", "#00b0f0", "#0070c0", "#002060", "#7030a0"];
function themeVariation(color: string, row: number): string {
  const hex = color.replace(/^#/, "");
  if (!/^[0-9a-f]{6}$/i.test(hex)) return color;
  const r = parseInt(hex.slice(0, 2), 16),
    g = parseInt(hex.slice(2, 4), 16),
    b = parseInt(hex.slice(4, 6), 16);
  const lum = (Math.max(r, g, b) + Math.min(r, g, b)) / 510;
  // Microsoft Office theme tint/shade rows: white darkens progressively, black
  // lightens progressively, light backgrounds darken, accents tint then shade.
  const spec: readonly [number | undefined, number | undefined][] =
    lum >= 0.99
      ? [[undefined, 0xf2], [undefined, 0xd9], [undefined, 0xbf], [undefined, 0xa6], [undefined, 0x80]]
      : lum <= 0.01
        ? [[0x80, undefined], [0xa6, undefined], [0xbf, undefined], [0xd9, undefined], [0xf2, undefined]]
        : lum >= 0.8
          ? [[undefined, 0xe6], [undefined, 0xbf], [undefined, 0x80], [undefined, 0x40], [undefined, 0x1a]]
          : [[0x33, undefined], [0x66, undefined], [0x99, undefined], [undefined, 0xbf], [undefined, 0x80]];
  const [tintHex, shadeHex] = spec[row]!;
  return (
    "#" +
    wordColorLuminance(
      hex,
      tintHex !== undefined ? tintHex.toString(16).padStart(2, "0") : undefined,
      shadeHex !== undefined ? shadeHex.toString(16).padStart(2, "0") : undefined,
    ).toLowerCase()
  );
}
export function OfficeColorMenu({ label, value, disabled, apply, highlight = false, fill = false, themeColors }: {
  label: string; value: string; disabled: boolean; apply(color: string): void; highlight?: boolean; fill?: boolean; themeColors?: readonly string[] | undefined;
}): ReactElement {
  const palette = themeColors && themeColors.length === 10 ? themeColors : defaultThemeColors;
  const [color, setColor] = useState(value);
  const icon = <span className={"office-color-symbol " + (highlight ? "highlight-symbol" : fill ? "fill-symbol" : "")}
    style={{ borderBottomColor: color }}>{highlight ? <Icon name="highlight" size={18} /> : fill ? <Icon name="fill" size={18} /> : "A"}</span>;
  const select = (next: string, close: () => void) => { setColor(next); apply(next); close(); };
  return <OfficeMenu label={label} icon={icon} disabled={disabled} split={() => apply(color)} className="color-menu">
    {close => <>
      {!highlight && <button type="button" role="menuitem" onClick={() => select("#000000", close)}>自动</button>}
      <div className="office-menu-heading">{highlight ? "突出显示颜色" : "主题颜色"}</div>
      <div className="office-color-grid">
        {(highlight ? standardColors : palette).map((swatch, index) => <button key={index + swatch} role="menuitem" type="button" aria-label={swatch} title={swatch}
          style={{ backgroundColor: swatch }} onClick={() => select(swatch, close)} />)}
        {!highlight && [0, 1, 2, 3, 4].flatMap(row => palette.map((swatch, col) => {
          const next = themeVariation(swatch, row); return <button key={`${row}-${col}-${next}`} role="menuitem" type="button" aria-label={next} title={next}
            style={{ backgroundColor: next }} onClick={() => select(next, close)} />;
        }))}
      </div>
      {!highlight && <><div className="office-menu-heading">标准色</div><div className="office-color-grid">{standardColors.map(swatch =>
        <button key={swatch} role="menuitem" type="button" aria-label={swatch} title={swatch} style={{ backgroundColor: swatch }} onClick={() => select(swatch, close)} />)}</div></>}
      {highlight && <button type="button" role="menuitem" onClick={() => select("transparent", close)}>无颜色</button>}
      <label className="office-custom-color">其他颜色…<input type="color" aria-label={label} value={/^#[0-9a-f]{6}$/i.test(color) ? color : "#000000"}
        onChange={event => select(event.currentTarget.value, close)} /></label>
    </>}
  </OfficeMenu>;
}

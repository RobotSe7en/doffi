import { useLayoutEffect, useRef, useState } from "react";
import type { ReactElement, ReactNode } from "react";
import { officeDialogStyles } from "../office-dialog.js";

export function OfficeDialog({ title, tabs, children, footer, close, className = "" }: {
  title: string; tabs?: readonly string[]; children: ReactNode | ((tab: number) => ReactNode);
  footer: ReactNode; close(): void; className?: string;
}): ReactElement {
  const ref = useRef<HTMLDialogElement>(null), [selected, setSelected] = useState(0);
  useLayoutEffect(() => { const dialog = ref.current!; dialog.showModal(); return () => { dialog.close(); }; }, []);
  return <dialog ref={ref} className={"office-dialog " + className} aria-label={title} onKeyDown={event => {
    if (event.key !== "Enter" || event.nativeEvent.isComposing || event.defaultPrevented || (event.target as Element).matches("textarea,select,button")) return;
    const primary = ref.current?.querySelector<HTMLButtonElement>("button.primary:not(:disabled)"); if (primary) { event.preventDefault(); primary.click(); }
  }} onCancel={event => { event.preventDefault(); close(); }}>
    <style>{officeDialogStyles}</style><h2 className="dialog-title">{title}</h2>
    {tabs && <div className="dialog-tabs" role="tablist" aria-label={title + "选项"}>{tabs.map((name, index) => <button key={name} type="button" role="tab" aria-selected={selected === index} tabIndex={selected === index ? 0 : -1}
      onClick={() => setSelected(index)} onKeyDown={event => { if (event.key !== "ArrowLeft" && event.key !== "ArrowRight") return; event.preventDefault(); const next = (index + (event.key === "ArrowRight" ? 1 : -1) + tabs.length) % tabs.length; setSelected(next); (event.currentTarget.parentElement!.children[next] as HTMLElement).focus(); }}>{name}</button>)}</div>}
    <div className="dialog-content">{typeof children === "function" ? children(selected) : children}</div><div className="dialog-footer">{footer}</div>
  </dialog>;
}

import { useEffect, useId, useRef, useState } from 'react';
import type { ReactElement } from 'react';

/** Editable Office fields keep the document selection while committing one value. */
export function OfficeComboBox({ label, value, options, disabled, commit, hint, className = '' }: {
  label: string; value: string; options: readonly string[]; disabled: boolean; commit(value: string): string | void | Promise<void>; className?: string; hint?: string | undefined;
}): ReactElement {
  const [draft, setDraft] = useState(value), [open, setOpen] = useState(false), [active, setActive] = useState(-1), [error, setError] = useState(''), [position, setPosition] = useState({ left: 0, top: 0, width: 0 });
  const [pendingFocus, setPendingFocus] = useState<HTMLElement | null>(null);
  const optionsId = useId();
  const dirty = useRef(false), input = useRef<HTMLInputElement>(null), holder = useRef<HTMLDivElement>(null);
  useEffect(() => { if (!dirty.current) setDraft(value); }, [value]);
  useEffect(() => { if (pendingFocus && !disabled && !pendingFocus.matches(':disabled')) { if (pendingFocus.isConnected) { window.focus(); pendingFocus.focus({ preventScroll: true }); } setPendingFocus(null); } }, [pendingFocus, disabled]);
  useEffect(() => {
    if (!open) return;
    const rect = holder.current!.getBoundingClientRect(), width = Math.min(280, Math.max(rect.width, label === '字体' ? 210 : 90), window.innerWidth - 16);
    setPosition({ left: Math.max(8, Math.min(rect.left, window.innerWidth - width - 8)), top: rect.bottom, width });
    const dismiss = (event: Event) => { if (event.target instanceof Element && event.target.closest('[role=listbox]')) return; setOpen(false); };
    document.addEventListener('scroll', dismiss, true); window.addEventListener('resize', dismiss);
    return () => { document.removeEventListener('scroll', dismiss, true); window.removeEventListener('resize', dismiss); };
  }, [open, label]);
  useEffect(() => { if (open && active >= 0) holder.current?.querySelector('[role=listbox] [aria-selected=true]')?.scrollIntoView({ block: 'nearest' }); }, [active, open]);
  const choices = options.filter(option => !dirty.current || option.toLowerCase().includes(draft.toLowerCase()));
  function apply(next = draft): boolean | Promise<void> {
    if (!dirty.current && next === value) { setOpen(false); return true; }
    dirty.current = false;
    const result = commit(next.trim()); if (typeof result === 'string') { dirty.current = true; setError(result); return false; }
    setDraft(next.trim()); setError(''); setOpen(false); setActive(-1);
    return result ?? true;
  }
  return <div ref={holder} className={'office-combo ' + className} onBlur={event => {
    if (holder.current?.contains(event.relatedTarget as Node | null)) return;
    if (dirty.current) apply(); setOpen(false);
  }}>
    <input ref={input} role="combobox" spellCheck={false} autoComplete="off" aria-label={label} aria-expanded={open} aria-autocomplete="list" aria-controls={optionsId} aria-activedescendant={open && active >= 0 && choices[active] ? optionsId + '-' + active : undefined}
      aria-invalid={Boolean(error)} disabled={disabled} value={draft} placeholder={label} title={error || hint || label}
      onChange={event => { dirty.current = true; setDraft(event.currentTarget.value); setError(''); setOpen(true); setActive(-1); }}
      onKeyDown={event => {
        if (event.nativeEvent.isComposing) return;
        if (event.key === 'Escape') { event.preventDefault(); dirty.current = false; setDraft(value); setError(''); setOpen(false); setActive(-1); }
        else if (event.key === 'Enter') { event.preventDefault(); apply(open && active >= 0 ? choices[active]! : draft); }
        else if (event.key === 'Tab') {
          const fields = [...document.querySelectorAll<HTMLElement>('button,input,select,textarea,a[href],[tabindex]')].filter(node => node.tabIndex >= 0 && !node.matches(':disabled') && node.getClientRects().length);
          const next = fields[fields.indexOf(event.currentTarget) + (event.shiftKey ? -1 : 1)]; if (!next) return;
          event.preventDefault(); const result = dirty.current ? apply() : true; if (result === false) return;
          const focusNext = () => setPendingFocus(next);
          if (result === true) focusNext(); else void result.then(focusNext);
          setOpen(false);
        }
        else if (event.key === 'ArrowDown' || event.key === 'ArrowUp') { event.preventDefault(); setOpen(true); setActive(index => Math.max(0, Math.min(choices.length - 1, index + (event.key === 'ArrowDown' ? 1 : -1)))); }
      }} />
    {hint && <span role="status" aria-label={hint} title={hint} className="office-font-hint">!</span>}
    <button type="button" disabled={disabled} aria-label={'展开' + label} aria-expanded={open} tabIndex={-1}
      onMouseDown={event => event.preventDefault()} onClick={() => { dirty.current = false; setDraft(value); setOpen(!open); setActive(-1); input.current?.focus(); }}>▾</button>
    {open && <div role="listbox" aria-label={label + '选项'} id={optionsId} className="office-combo-options" style={{ position: "fixed", left: position.left, top: position.top, width: position.width }}>
      {choices.map((option, index) => <button type="button" role="option" tabIndex={-1} id={optionsId + '-' + index} aria-selected={index === active} key={option}
        onMouseDown={event => event.preventDefault()} onClick={() => apply(option)}>{option}</button>)}
      {!choices.length && <span>按 Enter 应用输入的{label}</span>}
    </div>}
    {error && <span role="alert" className="office-combo-error">{error}</span>}
  </div>;
}

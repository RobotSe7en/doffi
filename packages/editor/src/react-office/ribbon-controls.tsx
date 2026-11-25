import type { ReactElement, ReactNode } from "react";
import { OfficeIcon as Icon } from "./icons.js";
export type Glyph = Parameters<typeof Icon>[0]["name"];
export function Group({
  title,
  children,
  className = "",
}: {
  title: string;
  children: ReactNode;
  className?: string;
}): ReactElement {
  return (
    <div className={"ribbon-group " + className} role="group" aria-label={title}>
      <div className="ribbon-group-content">{children}</div>
      <span className="ribbon-group-caption">{title}</span>
    </div>
  );
}
export function Tool({
  label,
  icon = "file",
  glyph,
  large,
  run,
  disabled = !run,
  active,
  reason,
  className = "",
}: {
  label: string;
  icon?: Glyph;
  glyph?: ReactNode;
  large?: boolean;
  run?: (() => void) | undefined;
  disabled?: boolean;
  active?: boolean | undefined;
  reason?: string | undefined;
  className?: string;
}): ReactElement {
  return (
    <button
      type="button"
      className={"ribbon-tool" + (large ? " large" : "") + " " + className}
      data-icon={icon}
      aria-label={label}
      aria-pressed={active}
      title={
        label + (disabled && reason ? " · " + reason : "")
      }
      disabled={disabled}
      onMouseDown={(event) => {
        const focused = document.activeElement;
        if (focused instanceof HTMLInputElement && focused.closest(".office-combo")) focused.blur();
        event.preventDefault();
      }}
      onClick={run}
    >
      {glyph ? (
        <span className="ribbon-glyph" aria-hidden="true">
          {glyph}
        </span>
      ) : (
        <Icon name={icon} size={large ? 32 : 17} />
      )}
      <span>{label}</span>
    </button>
  );
}

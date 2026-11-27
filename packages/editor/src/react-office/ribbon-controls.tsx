import type { ReactElement, ReactNode } from "react";
import { OfficeIcon as Icon } from "./icons.js";
export type Glyph = Parameters<typeof Icon>[0]["name"];
export function Group({
  title,
  children,
  className = "",
  launcher,
}: {
  title: string;
  children: ReactNode;
  className?: string;
  launcher?:
    | {
        readonly label: string;
        readonly run: () => void | Promise<void>;
        readonly disabled?: boolean;
      }
    | undefined;
}): ReactElement {
  return (
    <div className={"ribbon-group " + className} role="group" aria-label={title}>
      <div className="ribbon-group-content">{children}</div>
      <span className="ribbon-group-caption">{title}</span>
      {launcher && (
        <button
          type="button"
          className="ribbon-dialog-launcher"
          aria-label={launcher.label}
          title={launcher.label}
          disabled={launcher.disabled}
          onMouseDown={(event) => event.preventDefault()}
          onClick={launcher.run}
        >
          <svg
            width="9"
            height="9"
            viewBox="0 0 10 10"
            fill="none"
            stroke="currentColor"
            strokeWidth="1.15"
            aria-hidden="true"
          >
            <path d="M1.5 6V1.5H6M4.2 4.2L8.5 8.5M8.5 5.6V8.5H5.6" />
          </svg>
          <span>{launcher.label}</span>
        </button>
      )}
    </div>
  );
}
export function Tool({
  label,
  icon = "file",
  glyph,
  large,
  run,
  onDoubleClick,
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
  onDoubleClick?: (() => void) | undefined;
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
      onDoubleClick={onDoubleClick}
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

export interface WordContextMenuItem {
  readonly label?: string;
  readonly separator?: boolean;
  readonly disabled?: boolean;
  readonly run?: () => void;
}

/** Accessible floating right-click context menu hosted inside the editor shadow DOM. */
export function createWordContextMenu(
  shadow: ShadowRoot,
  onError: (error: unknown) => void,
): {
  open(clientX: number, clientY: number, items: readonly WordContextMenuItem[]): void;
  close(): void;
  dispose(): void;
} {
  const document = shadow.ownerDocument;
  let menu: HTMLElement | undefined;
  let controller: AbortController | undefined;

  const close = (): void => {
    controller?.abort();
    controller = undefined;
    menu?.remove();
    menu = undefined;
  };

  const open = (clientX: number, clientY: number, items: readonly WordContextMenuItem[]): void => {
    close();
    const visible = items.filter(
      (item, index) =>
        !item.separator ||
        (index > 0 && index < items.length - 1 && !items[index - 1]?.separator),
    );
    if (!visible.length) return;
    controller = new AbortController();
    const signal = controller.signal;
    menu = document.createElement("div");
    menu.setAttribute("role", "menu");
    menu.setAttribute("aria-label", "上下文菜单");
    menu.style.cssText =
      "position:fixed;z-index:1000;min-width:176px;max-width:260px;padding:4px 0;background:#ffffff;border:1px solid #d0d7e2;border-radius:6px;box-shadow:0 6px 20px rgba(15,23,42,.16);font:12px -apple-system,BlinkMacSystemFont,'PingFang SC',sans-serif;color:#20252b";
    for (const item of visible) {
      if (item.separator) {
        const hr = document.createElement("div");
        hr.setAttribute("role", "separator");
        hr.style.cssText = "height:1px;margin:4px 0;background:#e5e9f0";
        menu.append(hr);
        continue;
      }
      const btn = document.createElement("button");
      btn.type = "button";
      btn.setAttribute("role", "menuitem");
      btn.textContent = item.label ?? "";
      btn.disabled = Boolean(item.disabled || !item.run);
      btn.style.cssText =
        "display:block;width:100%;text-align:left;padding:6px 14px;border:0;background:transparent;color:inherit;font:inherit;cursor:" +
        (btn.disabled ? "default" : "pointer") +
        ";opacity:" +
        (btn.disabled ? ".42" : "1");
      btn.addEventListener("mouseenter", () => {
        if (!btn.disabled) btn.style.background = "#eef4fc";
      }, { signal });
      btn.addEventListener("mouseleave", () => {
        btn.style.background = "transparent";
      }, { signal });
      btn.addEventListener("mousedown", (event) => event.preventDefault(), { signal });
      btn.addEventListener(
        "click",
        () => {
          close();
          try {
            item.run?.();
          } catch (error) {
            onError(error);
          }
        },
        { signal },
      );
      menu.append(btn);
    }
    shadow.append(menu);
    const rect = menu.getBoundingClientRect();
    const view = document.defaultView!;
    const left = Math.max(8, Math.min(clientX, view.innerWidth - rect.width - 8));
    const top = Math.max(8, Math.min(clientY, view.innerHeight - rect.height - 8));
    menu.style.left = left + "px";
    menu.style.top = top + "px";
    menu.addEventListener(
      "keydown",
      (event) => {
        if (event.key === "Escape") {
          event.preventDefault();
          close();
          return;
        }
        if (!["ArrowDown", "ArrowUp", "Home", "End"].includes(event.key)) return;
        const buttons = [...menu!.querySelectorAll<HTMLButtonElement>("button:not(:disabled)")];
        if (!buttons.length) return;
        const current = buttons.indexOf(shadow.activeElement as HTMLButtonElement);
        const next =
          event.key === "Home"
            ? 0
            : event.key === "End"
              ? buttons.length - 1
              : (current + (event.key === "ArrowDown" ? 1 : -1) + buttons.length) % buttons.length;
        event.preventDefault();
        buttons[next]?.focus();
      },
      { signal },
    );
    const dismiss = (event: Event) => {
      if (menu && !event.composedPath().includes(menu)) close();
    };
    document.addEventListener("pointerdown", dismiss, { signal, capture: true });
    view.addEventListener("blur", close, { signal });
    view.addEventListener("resize", close, { signal });
    menu.querySelector<HTMLButtonElement>("button:not(:disabled)")?.focus();
  };

  return { open, close, dispose: close };
}

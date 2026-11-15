import type { PageSetup } from "@doffi/core";
import type { DocumentRenderReport } from "@doffi/render";

const svgNS = "http://www.w3.org/2000/svg",
  px = 4 / 3,
  centimeter = 96 / 2.54;
export interface WordRulers {
  refresh(): void;
  setVisible(visible: boolean): void;
  dispose(): void;
}
export interface WordRulerEditing {
  pageId?(): string | undefined;
  read(): { left: number; right: number; first: number; tabs: readonly number[] } | undefined;
  commit(kind: "left" | "right" | "first" | "tab", value: number, previous?: number): void;
  commitMargin?(
    side: "left" | "right" | "top" | "bottom",
    value: number,
    sectionId?: string,
    pageIndex?: number,
  ): void;
  onError(error: unknown): void;
}
/** Build chrome before inserting an iframe: moving a live iframe resets its document. */
export function createWordViewport(stage: HTMLElement): HTMLElement {
  const document = stage.ownerDocument,
    viewport = document.createElement("div"),
    across = document.createElement("div"),
    down = document.createElement("div"),
    corner = document.createElement("div");
  viewport.dataset.doffiViewport = "";
  across.dataset.doffiRulerRail = "horizontal";
  down.dataset.doffiRulerRail = "vertical";
  corner.setAttribute("aria-hidden", "true");
  viewport.style.cssText = "position:relative;min-width:0;min-height:0;background:#eceef1";
  across.style.cssText =
    "position:relative;grid-area:1/2;overflow:hidden;border-bottom:1px solid #d3d5d8";
  down.style.cssText =
    "position:relative;grid-area:2/1;overflow:hidden;border-right:1px solid #d3d5d8";
  corner.style.cssText =
    "grid-area:1/1;border-right:1px solid #d3d5d8;border-bottom:1px solid #d3d5d8";
  viewport.append(corner, across, down, stage);
  for (const rail of [across, down, corner]) rail.style.display = "none";
  return viewport;
}

/** Two derived rulers in viewport rails; no document nodes or per-page ruler trees. */
export function mountWordRulers(
  stage: HTMLElement,
  frame: HTMLIFrameElement,
  root: () => HTMLElement | undefined,
  layout: () => DocumentRenderReport | undefined,
  defaultSetup: () => Readonly<PageSetup>,
  editing?: WordRulerEditing,
): WordRulers {
  const document = stage.ownerDocument,
    viewport = stage.parentElement!,
    across = viewport.querySelector<HTMLElement>("[data-doffi-ruler-rail=horizontal]")!,
    down = viewport.querySelector<HTMLElement>("[data-doffi-ruler-rail=vertical]")!,
    corner = viewport.firstElementChild as HTMLElement;
  const original = {
    height: stage.style.height,
    minHeight: stage.style.minHeight,
    maxHeight: stage.style.maxHeight,
  };
  stage.style.gridArea = "2/2";
  stage.style.minWidth = "0";
  const horizontal = document.createElementNS(svgNS, "svg"),
    vertical = document.createElementNS(svgNS, "svg");
  for (const [ruler, axis, label] of [
    [horizontal, "horizontal", "水平标尺（厘米）"],
    [vertical, "vertical", "垂直标尺（厘米）"],
  ] as const) {
    ruler.dataset.doffiRuler = axis;
    ruler.setAttribute("role", "img");
    ruler.setAttribute("aria-label", label);
    ruler.style.cssText =
      "position:absolute;z-index:4;pointer-events:none;overflow:hidden;background:#f5f5f5;color:#525965";
    ruler.style.display = "none";
    (axis === "horizontal" ? across : down).append(ruler);
  }
  let visible = false,
    disposed = false,
    queued: number | undefined,
    painted = "";
  const controls = new AbortController();
  let handleControls = new AbortController(),
    handleKey = "",
    dragging = false;
  let horizontalGeometry: { start: number; end: number; width: number; zoom: number } | undefined,
    verticalGeometry: typeof horizontalGeometry,
    activeSectionId: string | undefined,
    activePageIndex: number | undefined;
  function handles(): void {
    if (dragging) return;
    const nextKey = JSON.stringify([
      editing?.read(),
      horizontalGeometry,
      verticalGeometry,
      activeSectionId,
      activePageIndex,
    ]);
    if (
      nextKey === handleKey &&
      horizontal.querySelector("[data-ruler-handle],[data-ruler-margin]")
    )
      return;
    handleKey = nextKey;
    handleControls.abort();
    handleControls = new AbortController();
    for (const ruler of [horizontal, vertical])
      ruler
        .querySelectorAll("[data-ruler-handle],[data-ruler-margin]")
        .forEach((node) => node.remove());
    const state = editing?.read(),
      geometry = horizontalGeometry;
    if (!geometry) return;
    if (editing?.commitMargin) {
      const addMargin = (side: "left" | "right" | "top" | "bottom", label: string): void => {
        const isVertical = side === "top" || side === "bottom",
          ruler = isVertical ? vertical : horizontal,
          box = isVertical ? verticalGeometry : geometry;
        if (!box) return;
        const end = side === "right" || side === "bottom",
          value = end ? box.end : box.start,
          maximum = box.width - (end ? box.start : box.end) - 1,
          scale = px * box.zoom,
          sectionId = activeSectionId,
          pageIndex = activePageIndex;
        const handle = document.createElementNS(svgNS, "path");
        handle.dataset.rulerMargin = side;
        handle.setAttribute("d", isVertical ? "M8 -4H20V4H8Z" : "M-4 8H4V12H-4Z");
        handle.setAttribute("fill", "#9a7755");
        handle.style.cssText =
          "pointer-events:all;outline:none;cursor:" + (isVertical ? "ns" : "ew") + "-resize";
        handle.setAttribute("role", "slider");
        handle.setAttribute("tabindex", "0");
        handle.setAttribute("aria-label", label);
        handle.setAttribute("aria-orientation", isVertical ? "vertical" : "horizontal");
        handle.setAttribute("aria-valuenow", String(value));
        handle.setAttribute("aria-valuemin", "0");
        handle.setAttribute("aria-valuemax", String(maximum));
        const place = (next: number): void => {
          const at = (end ? box.width - next : next) * scale;
          handle.setAttribute(
            "transform",
            isVertical ? "translate(0 " + at + ")" : "translate(" + at + " 0)",
          );
          handle.setAttribute("aria-valuenow", String(next));
        };
        place(value);
        const clamp = (value: number): number =>
          Math.round(Math.max(0, Math.min(maximum, value)) * 2) / 2;
        const finish = (next: number): void => {
          const focus = (handle.getRootNode() as Document | ShadowRoot).activeElement === handle;
          try {
            editing!.commitMargin!(side, next, sectionId, pageIndex);
          } catch (error) {
            editing!.onError(error);
          }
          refresh();
          if (focus)
            queueMicrotask(() =>
              ruler.querySelector<SVGElement>("[data-ruler-margin=" + side + "]")?.focus(),
            );
        };
        handle.addEventListener(
          "keydown",
          (event) => {
            const decrease = isVertical ? "ArrowUp" : "ArrowLeft",
              increase = isVertical ? "ArrowDown" : "ArrowRight";
            if (![decrease, increase, "Home", "End"].includes(event.key)) return;
            event.preventDefault();
            finish(
              event.key === "Home"
                ? 0
                : event.key === "End"
                  ? maximum
                  : clamp(
                      value +
                        (event.key === increase ? 1 : -1) *
                          (end ? -1 : 1) *
                          (event.shiftKey ? 10 : 1),
                    ),
            );
          },
          { signal: handleControls.signal },
        );
        handle.addEventListener(
          "pointerdown",
          (event) => {
            event.preventDefault();
            event.stopPropagation();
            dragging = true;
            handle.setPointerCapture(event.pointerId);
            const start = isVertical ? event.clientY : event.clientX;
            let next = value;
            const move = (event: PointerEvent): void => {
              next = clamp(
                value +
                  (((isVertical ? event.clientY : event.clientX) - start) / scale) * (end ? -1 : 1),
              );
              place(next);
            };
            const cleanup = (): void => {
              dragging = false;
              handle.removeEventListener("pointermove", move);
              handle.removeEventListener("pointerup", up);
              handle.removeEventListener("pointercancel", cancel);
            };
            const up = (): void => {
              cleanup();
              finish(next);
            };
            const cancel = (): void => {
              cleanup();
              refresh();
            };
            handle.addEventListener("pointermove", move);
            handle.addEventListener("pointerup", up);
            handle.addEventListener("pointercancel", cancel);
          },
          { signal: handleControls.signal },
        );
        ruler.append(handle);
      };
      addMargin("left", "页面左边距");
      addMargin("right", "页面右边距");
      addMargin("top", "页面上边距");
      addMargin("bottom", "页面下边距");
    }
    if (!state) return;
    const scale = px * geometry.zoom,
      origin = geometry.start * scale,
      available = geometry.width - geometry.start - geometry.end;
    const add = (kind: "left" | "right" | "first" | "tab", value: number, label: string): void => {
      const handle = document.createElementNS(svgNS, "path");
      handle.dataset.rulerHandle = kind;
      handle.setAttribute(
        "d",
        kind === "first" ? "M-5 0H5L0 7Z" : kind === "tab" ? "M-3 11V18H4" : "M-5 19H5L0 12Z",
      );
      handle.setAttribute("fill", kind === "tab" ? "#2271cf" : "#657b96");
      const at = () =>
        origin +
        (kind === "right" ? available - value : kind === "first" ? state.left + value : value) *
          scale;
      handle.setAttribute("transform", `translate(${at()} 0)`);
      handle.style.cssText = "pointer-events:all;cursor:ew-resize;outline:none";
      handle.setAttribute("role", "slider");
      handle.setAttribute("aria-label", label);
      handle.setAttribute("tabindex", "0");
      handle.setAttribute("aria-valuenow", String(value));
      const finish = (next: number): void => {
        const keepFocus = (handle.getRootNode() as ShadowRoot).activeElement === handle;
        try {
          editing!.commit(kind, next, kind === "tab" ? value : undefined);
        } catch (error) {
          editing!.onError(error);
        }
        refresh();
        if (keepFocus)
          queueMicrotask(() =>
            horizontal.querySelector<SVGElement>("[data-ruler-handle=" + kind + "]")?.focus(),
          );
      };
      handle.addEventListener(
        "keydown",
        (event) => {
          if (!["ArrowLeft", "ArrowRight", "Delete", "Backspace"].includes(event.key)) return;
          event.preventDefault();
          if ((event.key === "Delete" || event.key === "Backspace") && kind === "tab") finish(-1);
          else if (event.key.startsWith("Arrow"))
            finish(
              value +
                (event.key === "ArrowRight" ? 1 : -1) *
                  (kind === "right" ? -1 : 1) *
                  (event.shiftKey ? 10 : 1),
            );
        },
        { signal: handleControls.signal },
      );
      handle.addEventListener(
        "pointerdown",
        (event) => {
          event.preventDefault();
          event.stopPropagation();
          dragging = true;
          handle.setPointerCapture(event.pointerId);
          const start = event.clientX;
          let next = value;
          const move = (move: PointerEvent): void => {
            next = value + ((move.clientX - start) / scale) * (kind === "right" ? -1 : 1);
            next =
              Math.round(
                Math.max(kind === "first" ? -state.left : 0, Math.min(available, next)) * 2,
              ) / 2;
            handle.setAttribute(
              "transform",
              `translate(${origin + (kind === "right" ? available - next : kind === "first" ? state.left + next : next) * scale} 0)`,
            );
          };
          const end = (up: PointerEvent): void => {
            dragging = false;
            handle.removeEventListener("pointermove", move);
            handle.removeEventListener("pointerup", end);
            handle.removeEventListener("pointercancel", cancel);
            finish(
              kind === "tab" && Math.abs(up.clientY - horizontal.getBoundingClientRect().top) > 40
                ? -1
                : next,
            );
          };
          const cancel = (): void => {
            dragging = false;
            handle.removeEventListener("pointermove", move);
            handle.removeEventListener("pointerup", end);
            handle.removeEventListener("pointercancel", cancel);
            refresh();
          };
          handle.addEventListener("pointermove", move);
          handle.addEventListener("pointerup", end);
          handle.addEventListener("pointercancel", cancel);
        },
        { signal: handleControls.signal },
      );
      horizontal.append(handle);
    };
    add("left", state.left, "段落左缩进");
    add("right", state.right, "段落右缩进");
    add("first", state.first, "首行缩进");
    state.tabs.forEach((value) => add("tab", value, "制表位 " + value + " 磅"));
  }
  if (editing) {
    horizontal.style.pointerEvents = "auto";
    horizontal.addEventListener(
      "dblclick",
      (event) => {
        if ((event.target as Element).closest("[data-ruler-handle]") || !horizontalGeometry) return;
        const geometry = horizontalGeometry,
          value =
            (event.clientX - horizontal.getBoundingClientRect().left) / (px * geometry.zoom) -
            geometry.start;
        if (value < 0 || value > geometry.width - geometry.start - geometry.end) return;
        try {
          editing.commit("tab", Math.round(value * 2) / 2);
        } catch (error) {
          editing.onError(error);
        }
        refresh();
      },
      { signal: controls.signal },
    );
  }
  function paint(
    ruler: SVGSVGElement,
    length: number,
    start: number,
    end: number,
    zoom: number,
    vertical: boolean,
  ): void {
    const size = length * zoom,
      origin = start * zoom,
      unit = (centimeter * zoom) / 10;
    ruler.setAttribute("width", String(vertical ? 20 : size));
    ruler.setAttribute("height", String(vertical ? size : 20));
    const rectangle = document.createElementNS(svgNS, "rect");
    rectangle.setAttribute("fill", "#dfe1e4");
    rectangle.setAttribute("width", String(vertical ? 20 : size));
    rectangle.setAttribute("height", String(vertical ? size : 20));
    const printable = document.createElementNS(svgNS, "rect");
    printable.setAttribute("fill", "#fff");
    printable.setAttribute(vertical ? "y" : "x", String(origin));
    printable.setAttribute(
      "width",
      String(vertical ? 20 : Math.max(0, size - origin - end * zoom)),
    );
    printable.setAttribute(
      "height",
      String(vertical ? Math.max(0, size - origin - end * zoom) : 20),
    );
    const marks = document.createElementNS(svgNS, "path");
    marks.setAttribute("stroke", "#71767f");
    marks.setAttribute("stroke-width", "0.6");
    const children: SVGElement[] = [rectangle, printable, marks],
      paths: string[] = [];
    const stride = Math.max(1, Math.ceil(size / unit / 1500));
    for (
      let tick = Math.ceil(-origin / unit / stride) * stride;
      tick <= (size - origin) / unit;
      tick += stride
    ) {
      const at = origin + tick * unit,
        major = tick % 10 === 0,
        height = major ? 8 : tick % 5 === 0 ? 6 : 3;
      paths.push(vertical ? `M20 ${at}h-${height}` : `M${at} 20v-${height}`);
      if (major) {
        const text = document.createElementNS(svgNS, "text");
        text.textContent = String(Math.abs(tick / 10));
        text.setAttribute("x", String(vertical ? 7 : at));
        text.setAttribute("y", String(vertical ? at + 3 : 9));
        text.setAttribute("text-anchor", "middle");
        text.setAttribute("fill", "#525965");
        text.setAttribute("font-family", "system-ui,sans-serif");
        text.setAttribute("font-size", "8");
        children.push(text);
      }
    }
    marks.setAttribute("d", paths.join(""));
    ruler.replaceChildren(...children);
  }
  function refresh(): void {
    if (disposed) return;
    const show = visible;
    viewport.style.display = show
      ? "var(--doffi-preview-stage,grid)"
      : "var(--doffi-preview-stage,block)";
    viewport.style.gridTemplateColumns = "24px minmax(0,1fr)";
    viewport.style.gridTemplateRows = "24px minmax(0,1fr)";
    viewport.style.height = show ? "var(--doffi-canvas-height,70vh)" : "";
    for (const rail of [across, down, corner]) rail.style.display = show ? "block" : "none";
    stage.style.height = show ? "100%" : original.height;
    stage.style.minHeight = show ? "0" : original.minHeight;
    stage.style.maxHeight = show ? "100%" : original.maxHeight;
    const source = root(),
      report = source ? layout() : undefined,
      hidden = !show || !source || !report || !stage.clientWidth;
    horizontal.style.display = hidden ? "none" : "block";
    vertical.style.display = hidden || report?.layout !== "pages" ? "none" : "block";
    if (hidden || !source || !report) return;
    const sourceBox = source.getBoundingClientRect(),
      zoom = Number(frame.style.transform.match(/scale\(([\d.]+)\)/)?.[1] ?? 1);
    const frameBox = frame.getBoundingClientRect(),
      stageBox = stage.getBoundingClientRect();
    const top = (stageBox.top - frameBox.top) / zoom - sourceBox.top,
      page =
        report.pages.find((page) => page.pageId === editing?.pageId?.()) ??
        report.pages.find((page) => (page.rect.y + page.rect.height) * px > top + 20 / zoom) ??
        report.pages.at(-1) ??
        (report.layout === "continuous"
          ? { rect: { x: 0, y: 0 }, setup: { ...defaultSetup(), width: sourceBox.width / px } }
          : undefined);
    if (!page) {
      horizontal.style.display = vertical.style.display = "none";
      return;
    }
    const setup: Readonly<PageSetup> = page.setup,
      left = frameBox.left - stageBox.left + (sourceBox.left + page.rect.x * px) * zoom,
      y = frameBox.top - stageBox.top + (sourceBox.top + page.rect.y * px) * zoom;
    horizontal.style.left = left + "px";
    horizontal.style.top = "2px";
    vertical.style.left = "1px";
    vertical.style.top = y + "px";
    horizontalGeometry = {
      start: setup.margins.left,
      end: setup.margins.right,
      width: setup.width,
      zoom,
    };
    verticalGeometry = {
      start: setup.margins.top,
      end: setup.margins.bottom,
      width: setup.height,
      zoom,
    };
    activeSectionId = "sectionId" in page ? page.sectionId : undefined;
    activePageIndex = "index" in page ? page.index : undefined;
    const key = JSON.stringify([setup, zoom]);
    if (key === painted) {
      handles();
      return;
    }
    painted = key;
    paint(
      horizontal,
      setup.width * px,
      setup.margins.left * px,
      setup.margins.right * px,
      zoom,
      false,
    );
    paint(
      vertical,
      setup.height * px,
      setup.margins.top * px,
      setup.margins.bottom * px,
      zoom,
      true,
    );
    handles();
  }
  const schedule = () => {
    if (queued === undefined && !disposed)
      queued = document.defaultView!.requestAnimationFrame(() => {
        queued = undefined;
        refresh();
      });
  };
  stage.addEventListener("scroll", schedule, { passive: true });
  return {
    refresh,
    setVisible(value: boolean) {
      visible = value;
      refresh();
    },
    dispose() {
      disposed = true;
      controls.abort();
      handleControls.abort();
      stage.removeEventListener("scroll", schedule);
      if (queued !== undefined) document.defaultView!.cancelAnimationFrame(queued);
      horizontal.remove();
      vertical.remove();
      for (const [name, value] of Object.entries(original))
        stage.style[name as keyof typeof original] = value;
    },
  };
}

/** The iframe keeps one complete source flow. Its outer scrolling viewport owns
 * edge scrolling, while the browser selection remains a range in that flow. */
export function bindWordSelectionScroll(
  stage: HTMLElement,
  iframe: HTMLIFrameElement,
  root: HTMLElement,
  options: {
    signal: AbortSignal;
    enabled(): boolean;
    ignored(node: Node): boolean;
    changed(): void;
  },
): void {
  const frame = root.ownerDocument,
    host = stage.ownerDocument,
    view = host.defaultView!;
  let dragging = false,
    anchor: { node: Node; offset: number } | undefined,
    point: { x: number; y: number } | undefined,
    scheduled: number | undefined;
  const stop = (): void => {
    dragging = false;
    anchor = undefined;
    point = undefined;
    if (scheduled !== undefined) view.cancelAnimationFrame(scheduled);
    scheduled = undefined;
  };
  const zoom = (): number => Number(iframe.style.transform.match(/scale\(([\d.]+)\)/)?.[1] ?? 1);
  function tick(): void {
    scheduled = undefined;
    if (!dragging || !point || !options.enabled()) return;
    const bounds = stage.getBoundingClientRect(),
      edge = 32;
    const velocity = (value: number, from: number, to: number): number =>
      value < from + edge
        ? -Math.min(24, (from + edge - value) * 0.65)
        : value > to - edge
          ? Math.min(24, (value - to + edge) * 0.65)
          : 0;
    const dx = velocity(point.x, bounds.left, bounds.right),
      dy = velocity(point.y, bounds.top, bounds.bottom);
    if (!dx && !dy) return;
    const left = stage.scrollLeft,
      top = stage.scrollTop;
    stage.scrollBy(dx, dy);
    if (left === stage.scrollLeft && top === stage.scrollTop) return;
    const box = iframe.getBoundingClientRect(),
      scale = zoom(),
      x = (Math.max(bounds.left + 2, Math.min(bounds.right - 3, point.x)) - box.left) / scale,
      y = (Math.max(bounds.top + 2, Math.min(bounds.bottom - 3, point.y)) - box.top) / scale;
    const caret = frame.caretPositionFromPoint?.(x, y),
      range = caret ? undefined : frame.caretRangeFromPoint?.(x, y),
      node = caret?.offsetNode ?? range?.startContainer,
      offset = caret?.offset ?? range?.startOffset;
    const selection = frame.getSelection();
    if (!anchor && selection?.anchorNode && root.contains(selection.anchorNode))
      anchor = { node: selection.anchorNode, offset: selection.anchorOffset };
    if (
      anchor &&
      node &&
      offset !== undefined &&
      root.contains(node) &&
      !options.ignored(node) &&
      anchor.node.isConnected
    ) {
      selection?.setBaseAndExtent(anchor.node, anchor.offset, node, offset);
      options.changed();
    }
    scheduled = view.requestAnimationFrame(tick);
  }
  const schedule = (): void => {
    if (scheduled === undefined) scheduled = view.requestAnimationFrame(tick);
  };
  frame.addEventListener(
    "mousedown",
    (event) => {
      const node = event.target as Node;
      if (
        event.button !== 0 ||
        !options.enabled() ||
        !root.contains(node) ||
        options.ignored(node) ||
        frame.documentElement.hasAttribute("data-doffi-line-cursor")
      )
        return;
      dragging = true;
      anchor = undefined;
      const box = iframe.getBoundingClientRect(),
        scale = zoom();
      point = { x: box.left + event.clientX * scale, y: box.top + event.clientY * scale };
    },
    { signal: options.signal },
  );
  frame.addEventListener(
    "mousemove",
    (event) => {
      if (!dragging) return;
      if (!(event.buttons & 1)) {
        stop();
        return;
      }
      const box = iframe.getBoundingClientRect(),
        scale = zoom();
      point = { x: box.left + event.clientX * scale, y: box.top + event.clientY * scale };
      schedule();
    },
    { signal: options.signal },
  );
  host.addEventListener(
    "mousemove",
    (event) => {
      if (!dragging) return;
      if (!(event.buttons & 1)) {
        stop();
        return;
      }
      point = { x: event.clientX, y: event.clientY };
      schedule();
    },
    { signal: options.signal },
  );
  for (const document of [frame, host])
    document.addEventListener("mouseup", stop, { signal: options.signal });
  options.signal.addEventListener("abort", stop, { once: true });
}

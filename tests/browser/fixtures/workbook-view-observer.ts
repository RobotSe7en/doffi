/** Observe only this fixture's private workbook surfaces, including the
 * asynchronously created Worker preview. Source surfaces remain closed; tests
 * may open parent UI shadows to drive worksheet tabs through accessibility. */
export function observeWorkbookViews(editor: HTMLElement, options: { openHostShadows?: boolean } = {}) {
  const preview = editor.shadowRoot!.querySelector('.preview')!, captured = new WeakMap<Element, ShadowRoot>(), views = new WeakMap<HTMLIFrameElement, ShadowRoot>();
  const watched = new Set<HTMLIFrameElement>(), observed = new WeakSet<Node>(), restores: (() => void)[] = [];
  const observer = new MutationObserver(watch);
  const observe = (root: Node) => { if (!observed.has(root)) { observed.add(root); observer.observe(root, { childList: true, subtree: true }); } };
  function watch(): void {
    const scan = (root: Element | ShadowRoot) => {
      observe(root);
      for (const element of root.querySelectorAll('*')) {
        if (element.localName === 'iframe') {
          const frame = element as HTMLIFrameElement;
          if (!watched.has(frame)) {
            watched.add(frame); frame.addEventListener('load', () => {
              const prototype = (frame.contentWindow as Window & typeof globalThis).Element.prototype, attach = prototype.attachShadow;
              prototype.attachShadow = function(init) { const value = attach.call(this, init); if (this.localName === 'office-workbook-view-surface') views.set(frame, value); return value; };
              restores.push(() => { prototype.attachShadow = attach; });
            }, { capture: true, once: true });
          }
        }
        const shadow = captured.get(element) ?? element.shadowRoot; if (shadow) scan(shadow);
      }
    };
    scan(preview);
  }
  const prototype = editor.ownerDocument.defaultView!.Element.prototype, attach = prototype.attachShadow;
  // Preview hosts attach their shadow before joining the connected container.
  // Retain only weak handles here; scan/observe content solely under this UI.
  prototype.attachShadow = function(init) { const shadow = attach.call(this, options.openHostShadows ? { ...init, mode: 'open' } : init); captured.set(this, shadow); return shadow; };
  watch();
  return { root: (frame: HTMLIFrameElement) => views.get(frame), current: () => { const frame = [...watched].find(frame => frame.isConnected); return frame ? views.get(frame) : undefined; },
    dispose() { observer.disconnect(); for (const restore of restores.splice(0).reverse()) restore(); prototype.attachShadow = attach; watched.clear(); } };
}
export function measureWorkbookChart(view: ShadowRoot | undefined) {
  const svg = view?.querySelector<SVGSVGElement>('svg[data-sdk-chart=line]'); if (!svg) throw Error('The observed current chart is unavailable.');
  const category = svg.querySelector<SVGLineElement>('[data-sdk-chart-axis=category]')!, value = svg.querySelector<SVGLineElement>('[data-sdk-chart-axis=value]')!;
  return { categories: [...svg.querySelectorAll('[data-sdk-chart-label=category-axis]')].map(label => label.textContent), labels: [...svg.querySelectorAll('[data-sdk-chart-label=data-label]')].map(label => label.textContent),
    leaders: [...svg.querySelectorAll('[data-sdk-chart-leader-line]')].map(line => ({ color: line.getAttribute('stroke'), width: Number(line.getAttribute('stroke-width')) })),
    frames: [...svg.querySelectorAll('[data-sdk-chart-label-frame]')].map(frame => ({ fill: frame.getAttribute('fill'), color: frame.getAttribute('stroke'), width: Number(frame.getAttribute('stroke-width')) })),
    axes: { y: category.y1.baseVal.value, top: value.y1.baseVal.value, bottom: value.y2.baseVal.value } };
}
export function measureWorkbookDrawingText(view: ShadowRoot | undefined, nodeId: string) {
  return [...view?.querySelectorAll<HTMLElement>('[data-office-text-style][data-source-node-id]') ?? []]
    .filter(element => element.dataset.sourceNodeId === nodeId).map(element => {
      const style = element.ownerDocument.defaultView!.getComputedStyle(element);
      return { text: element.textContent, color: style.color, size: parseFloat(style.fontSize), weight: style.fontWeight };
    });
}

export function measureWorkbookDrawingBounds(view: ShadowRoot | undefined) {
  if (!view) throw Error("The observed workbook view is unavailable.");
  const origin = view.host.getBoundingClientRect();
  const objects = [...view.querySelectorAll<HTMLElement>('[data-office-view="workbook-drawing"],[data-office-view="workbook-note"]')].map(node => {
    const body = node.querySelector<HTMLElement>('[data-office-view="workbook-drawing-body"]')!, rect = body.getBoundingClientRect();
    return { nodeId: node.dataset.sourceNodeId!, kind: node.dataset.officeView!, xPt: (rect.left - origin.left) * .75, yPt: (rect.top - origin.top) * .75, widthPt: rect.width * .75, heightPt: rect.height * .75 };
  });
  const first = view.querySelector<HTMLElement>('[data-office-view="workbook-cell"][data-row="0"][data-column="0"]'), second = view.querySelector<HTMLElement>('[data-office-view="workbook-cell"][data-row="0"][data-column="1"]');
  return { objects, columnWidthPt: first && second ? (second.getBoundingClientRect().left - first.getBoundingClientRect().left) * .75 : undefined };
}

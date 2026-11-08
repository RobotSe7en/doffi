/** Observe only this fixture's private workbook surfaces, including the
 * asynchronously created Worker preview. No shadow is opened or source changed. */
export function observeWorkbookViews(editor: HTMLElement) {
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
  prototype.attachShadow = function(init) { const shadow = attach.call(this, init); captured.set(this, shadow); return shadow; };
  watch();
  return { root: (frame: HTMLIFrameElement) => views.get(frame), current: () => { const frame = [...watched].find(frame => frame.isConnected); return frame ? views.get(frame) : undefined; },
    dispose() { observer.disconnect(); for (const restore of restores.splice(0).reverse()) restore(); prototype.attachShadow = attach; watched.clear(); } };
}
export function measureWorkbookChart(view: ShadowRoot | undefined) {
  const svg = view?.querySelector<SVGSVGElement>('svg[data-sdk-chart=line]'); if (!svg) throw Error('The observed current chart is unavailable.');
  const category = svg.querySelector<SVGLineElement>('[data-sdk-chart-axis=category]')!, value = svg.querySelector<SVGLineElement>('[data-sdk-chart-axis=value]')!;
  return { categories: [...svg.querySelectorAll('[data-sdk-chart-label=category-axis]')].map(label => label.textContent), labels: [...svg.querySelectorAll('[data-sdk-chart-label=data-label]')].map(label => label.textContent),
    leaders: [...svg.querySelectorAll('[data-sdk-chart-leader-line]')].map(line => ({ color: line.getAttribute('stroke'), width: Number(line.getAttribute('stroke-width')) })),
    axes: { y: category.y1.baseVal.value, top: value.y1.baseVal.value, bottom: value.y2.baseVal.value } };
}

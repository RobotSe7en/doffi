import type { PageSetup } from '@doffi/core';
import type { DocumentRenderReport } from '@doffi/render';

const svgNS = 'http://www.w3.org/2000/svg', px = 4 / 3, centimeter = 96 / 2.54;
export interface WordRulers { refresh(): void; setVisible(visible: boolean): void; dispose(): void }
/** Build chrome before inserting an iframe: moving a live iframe resets its document. */
export function createWordViewport(stage: HTMLElement): HTMLElement {
  const document = stage.ownerDocument, viewport = document.createElement('div'), across = document.createElement('div'), down = document.createElement('div'), corner = document.createElement('div');
  viewport.dataset.doffiViewport = ''; across.dataset.doffiRulerRail = 'horizontal'; down.dataset.doffiRulerRail = 'vertical'; corner.setAttribute('aria-hidden', 'true');
  viewport.style.cssText = 'position:relative;min-width:0;min-height:0;background:#eceef1';
  across.style.cssText = 'position:relative;grid-area:1/2;overflow:hidden;border-bottom:1px solid #d3d5d8';
  down.style.cssText = 'position:relative;grid-area:2/1;overflow:hidden;border-right:1px solid #d3d5d8';
  corner.style.cssText = 'grid-area:1/1;border-right:1px solid #d3d5d8;border-bottom:1px solid #d3d5d8';
  viewport.append(corner, across, down, stage);
  for (const rail of [across, down, corner]) rail.style.display = "none";
  return viewport;
}

/** Two derived rulers in viewport rails; no document nodes or per-page ruler trees. */
export function mountWordRulers(stage: HTMLElement, frame: HTMLIFrameElement, root: () => HTMLElement | undefined, layout: () => DocumentRenderReport | undefined, defaultSetup: () => Readonly<PageSetup>): WordRulers {
  const document = stage.ownerDocument, viewport = stage.parentElement!, across = viewport.querySelector<HTMLElement>('[data-doffi-ruler-rail=horizontal]')!, down = viewport.querySelector<HTMLElement>('[data-doffi-ruler-rail=vertical]')!, corner = viewport.firstElementChild as HTMLElement;
  const original = { height: stage.style.height, minHeight: stage.style.minHeight, maxHeight: stage.style.maxHeight };
  stage.style.gridArea = '2/2'; stage.style.minWidth = '0';
  const horizontal = document.createElementNS(svgNS, 'svg'), vertical = document.createElementNS(svgNS, 'svg');
  for (const [ruler, axis, label] of [[horizontal, 'horizontal', '水平标尺（厘米）'], [vertical, 'vertical', '垂直标尺（厘米）']] as const) {
    ruler.dataset.doffiRuler = axis; ruler.setAttribute('role', 'img'); ruler.setAttribute('aria-label', label);
    ruler.style.cssText = 'position:absolute;z-index:4;pointer-events:none;overflow:hidden;background:#f5f5f5;color:#525965'; ruler.style.display = 'none'; (axis === 'horizontal' ? across : down).append(ruler);
  }
  let visible = false, disposed = false, queued: number | undefined, painted = '';
  function paint(ruler: SVGSVGElement, length: number, start: number, end: number, zoom: number, vertical: boolean): void {
    const size = length * zoom, origin = start * zoom, unit = centimeter * zoom / 10;
    ruler.setAttribute('width', String(vertical ? 20 : size)); ruler.setAttribute('height', String(vertical ? size : 20));
    const rectangle = document.createElementNS(svgNS, 'rect'); rectangle.setAttribute('fill', '#dfe1e4'); rectangle.setAttribute('width', String(vertical ? 20 : size)); rectangle.setAttribute('height', String(vertical ? size : 20));
    const printable = document.createElementNS(svgNS, 'rect'); printable.setAttribute('fill', '#fff'); printable.setAttribute(vertical ? 'y' : 'x', String(origin)); printable.setAttribute('width', String(vertical ? 20 : Math.max(0, size - origin - end * zoom))); printable.setAttribute('height', String(vertical ? Math.max(0, size - origin - end * zoom) : 20));
    const marks = document.createElementNS(svgNS, 'path'); marks.setAttribute('stroke', '#71767f'); marks.setAttribute('stroke-width', '0.6');
    const children: SVGElement[] = [rectangle, printable, marks], paths: string[] = [];
    const stride = Math.max(1, Math.ceil(size / unit / 1500));
    for (let tick = Math.ceil(-origin / unit / stride) * stride; tick <= (size - origin) / unit; tick += stride) {
      const at = origin + tick * unit, major = tick % 10 === 0, height = major ? 8 : tick % 5 === 0 ? 6 : 3;
      paths.push(vertical ? `M20 ${at}h-${height}` : `M${at} 20v-${height}`);
      if (major) { const text = document.createElementNS(svgNS, 'text'); text.textContent = String(Math.abs(tick / 10)); text.setAttribute('x', String(vertical ? 7 : at)); text.setAttribute('y', String(vertical ? at + 3 : 9)); text.setAttribute('text-anchor', 'middle'); text.setAttribute('fill', '#525965'); text.setAttribute('font-family', 'system-ui,sans-serif'); text.setAttribute('font-size', '8'); children.push(text); }
    }
    marks.setAttribute('d', paths.join('')); ruler.replaceChildren(...children);
  }
  function refresh(): void {
    if (disposed) return;
    const show = visible;
    viewport.style.display = show ? 'var(--doffi-preview-stage,grid)' : 'var(--doffi-preview-stage,block)'; viewport.style.gridTemplateColumns = '24px minmax(0,1fr)'; viewport.style.gridTemplateRows = '24px minmax(0,1fr)';
    viewport.style.height = show ? 'var(--doffi-canvas-height,70vh)' : '';
    for (const rail of [across, down, corner]) rail.style.display = show ? 'block' : 'none';
    stage.style.height = show ? '100%' : original.height; stage.style.minHeight = show ? '0' : original.minHeight; stage.style.maxHeight = show ? '100%' : original.maxHeight;
    const source = root(), report = source ? layout() : undefined, hidden = !show || !source || !report || !stage.clientWidth;
    horizontal.style.display = hidden ? 'none' : 'block'; vertical.style.display = hidden || report?.layout !== 'pages' ? 'none' : 'block';
    if (hidden || !source || !report) return;
    const sourceBox = source.getBoundingClientRect(), zoom = Number(frame.style.transform.match(/scale\(([\d.]+)\)/)?.[1] ?? 1);
    const frameBox = frame.getBoundingClientRect(), stageBox = stage.getBoundingClientRect();
    const top = (stageBox.top - frameBox.top) / zoom - sourceBox.top, page = report.pages.find(page => (page.rect.y + page.rect.height) * px > top + 20 / zoom) ?? report.pages.at(-1)
      ?? (report.layout === 'continuous' ? { rect: { x: 0, y: 0 }, setup: { ...defaultSetup(), width: sourceBox.width / px } } : undefined);
    if (!page) { horizontal.style.display = vertical.style.display = 'none'; return; }
    const setup: Readonly<PageSetup> = page.setup, left = frameBox.left - stageBox.left + (sourceBox.left + page.rect.x * px) * zoom, y = frameBox.top - stageBox.top + (sourceBox.top + page.rect.y * px) * zoom;
    horizontal.style.left = left + 'px'; horizontal.style.top = '2px'; vertical.style.left = '1px'; vertical.style.top = y + 'px';
    const key = JSON.stringify([setup, zoom]); if (key === painted) return; painted = key;
    paint(horizontal, setup.width * px, setup.margins.left * px, setup.margins.right * px, zoom, false);
    paint(vertical, setup.height * px, setup.margins.top * px, setup.margins.bottom * px, zoom, true);
  }
  const schedule = () => { if (queued === undefined && !disposed) queued = document.defaultView!.requestAnimationFrame(() => { queued = undefined; refresh(); }); };
  stage.addEventListener('scroll', schedule, { passive: true });
  return { refresh, setVisible(value: boolean) { visible = value; refresh(); }, dispose() { disposed = true; stage.removeEventListener('scroll', schedule); if (queued !== undefined) document.defaultView!.cancelAnimationFrame(queued); horizontal.remove(); vertical.remove(); for (const [name, value] of Object.entries(original)) stage.style[name as keyof typeof original] = value; } };
}

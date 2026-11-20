import { evaluatePresentationCustomGeometry, evaluatePresentationPreset, readPresentationGeometry } from "@doffi/core";
import type { DeepReadonly, PresentationObject } from "@doffi/core";
import { markViewNode } from "../view-nodes.js";
import { svgElement } from "./svg.js";

type Connector = DeepReadonly<Extract<PresentationObject, { kind: "connector" }>>;
/** Endpoint bindings are resolved after all objects, independently of z-order.
 * Cached connector geometry remains source data; no correction is persisted. */
export function resolveConnector(node: HTMLElement, definition: Connector, objects: ReadonlyMap<string, DeepReadonly<PresentationObject>>): void {
  if (!definition.startNodeId && !definition.endNodeId) return;
  const document = node.ownerDocument, surface = node.closest<HTMLElement>('[data-office-kind="slide"]')!, svg = node.querySelector("svg")!, path = svg.querySelector(":scope > path")!;
  const endpoint = (id: string | undefined, index: number | undefined): DOMPoint | undefined => {
    if (!id) return;
    const element = surface.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(id) + '"]'); if (!element) return;
    const computed = document.defaultView!.getComputedStyle(element), geometry = { width: parseFloat(computed.width) * 0.75, height: parseFloat(computed.height) * 0.75 };
    const object = objects.get(id)!, preset = object.kind === "geometry" ? object.preset : "rect", adjustments: Record<string, number> = {};
    for (const attr of element.attributes) if (attr.name.startsWith("data-office-adjust-")) adjustments[attr.name.slice(19)] = Number(attr.value);
    const custom = object.kind === "geometry" || object.kind === "image" ? object.customGeometry : undefined;
    const site = (custom ? evaluatePresentationCustomGeometry(custom, geometry.width, geometry.height) : evaluatePresentationPreset(preset, geometry.width, geometry.height, adjustments)).connections[index ?? 0]; if (!site) return;
    const probe = markViewNode(svgElement(document, "svg", { width: geometry.width, height: geometry.height, viewBox: `0 0 ${geometry.width} ${geometry.height}` })); probe.style.cssText = "position:absolute;left:0;top:0;visibility:hidden"; probe.style.width = geometry.width + "pt"; probe.style.height = geometry.height + "pt"; element.append(probe);
    const matrix = probe.getScreenCTM(); probe.remove();
    const local = svg.getScreenCTM(); if (!matrix || !local) return;
    return new document.defaultView!.DOMPoint(site.x, site.y).matrixTransform(matrix).matrixTransform(local.inverse());
  };
  const start = endpoint(definition.startNodeId, definition.startSite), end = endpoint(definition.endNodeId, definition.endSite);
  // One unattached end retains the cached preset endpoint.
  const geometry = readPresentationGeometry(Object.fromEntries([...node.attributes].map(a => [a.name, a.value])))!;
  const a = start ?? { x: 0, y: 0 }, b = end ?? { x: geometry.width, y: geometry.height };
  if (Math.hypot(a.x, a.y) < 0.02 && Math.hypot(b.x - geometry.width, b.y - geometry.height) < 0.02) return;
  path.setAttribute("d", definition.routing === "straight" ? `M${a.x} ${a.y}L${b.x} ${b.y}` : `M${a.x} ${a.y}L${a.x} ${b.y}L${b.x} ${b.y}`);
}

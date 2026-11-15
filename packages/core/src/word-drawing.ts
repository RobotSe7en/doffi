import type { DeepReadonly } from "./runtime.js";
import type { WordMetadata } from "./contracts/word.js";
import type { WordDrawing } from "./contracts/word-drawing.js";
import type { DocumentNode } from "./runtime.js";
import { evaluatePresentationPreset } from "./presentation-geometry.js";
import { inverse, local, multiply, point } from "./session/presentation-geometry.js";
import type { Matrix } from "./session/presentation-geometry.js";
import { DocumentError } from "./errors.js";

type Source = { getNode(id: string): Pick<DocumentNode, "parentNodeId" | "tagName"> | undefined };
const identity: Matrix = [1, 0, 0, 1, 0, 0];
/** Resolve one paragraph's drawing coordinate system, including nested groups.
 * This is derived from source ownership and never retained as document state. */
export function wordDrawingTransforms(word: DeepReadonly<WordMetadata>, source: Source) {
  const objects = new Map((word.drawings ?? []).map((item) => [item.nodeId, item])),
    cache = new Map<string, Matrix>();
  const object = (id: string): DeepReadonly<WordDrawing> => {
    const item = objects.get(id);
    if (!item) throw new DocumentError("InvalidOperation", "The drawing target is missing.");
    return item;
  };
  const parentId = (id: string): string | undefined => {
    for (
      let parent = source.getNode(id)?.parentNodeId;
      parent;
      parent = source.getNode(parent)?.parentNodeId
    )
      if (objects.has(parent)) return parent;
    return;
  };
  const root = (id: string): DeepReadonly<WordDrawing> => {
    let current = id;
    for (let next = parentId(current); next; next = parentId(current)) current = next;
    return object(current);
  };
  const paragraph = (id: string): string | undefined => {
    for (
      let current: string | null | undefined = id;
      current;
      current = source.getNode(current)?.parentNodeId
    )
      if (/^p$|^h[1-6]$/.test(source.getNode(current)?.tagName ?? "")) return current;
    return;
  };
  const parent = (id: string): Matrix => {
    const owner = parentId(id);
    if (!owner) return identity;
    const group = object(owner);
    if (group.kind !== "group")
      throw new DocumentError("InvalidDocument", "Only groups can own drawing objects.");
    const frame = group.coordinateSpace,
      sx = group.geometry.width / frame.width,
      sy = group.geometry.height / frame.height;
    return multiply(world(owner), [sx, 0, 0, sy, -frame.x * sx, -frame.y * sy]);
  };
  const world = (id: string): Matrix => {
    let matrix = cache.get(id);
    if (!matrix) {
      matrix = multiply(parent(id), local(object(id).geometry));
      cache.set(id, matrix);
    }
    return matrix;
  };
  const endpoint = (id: string, site: number): { x: number; y: number } => {
    const item = object(id),
      g = item.geometry,
      preset = evaluatePresentationPreset(
        item.kind === "shape" ? item.preset : "rect",
        g.width,
        g.height,
        item.kind === "shape" ? item.adjustments : {},
      ),
      connection = preset.connections[site];
    if (!connection)
      throw new DocumentError(
        "InvalidOperation",
        "The target drawing has no such connection site.",
      );
    return point(world(id), connection);
  };
  const sameCanvas = (a: string, b: string): boolean => {
    const ar = root(a),
      br = root(b);
    return (
      a !== b &&
      objects.get(b)?.kind !== "connector" &&
      paragraph(a) === paragraph(b) &&
      paragraph(a) !== undefined &&
      ((ar.nodeId === br.nodeId && ar.kind === "group") ||
        (ar.layout.mode === "floating" && br.layout.mode === "floating")) &&
      (ar.layout.horizontal ?? "column") === (br.layout.horizontal ?? "column") &&
      (ar.layout.vertical ?? "paragraph") === (br.layout.vertical ?? "paragraph") &&
      !ar.layout.alignH &&
      !ar.layout.alignV &&
      !br.layout.alignH &&
      !br.layout.alignV
    );
  };
  return { world, parent, endpoint, root, paragraph, sameCanvas };
}

/** Update the native connector fallback in the same transaction as its targets. */
export function updateWordDrawingConnectors(word: WordMetadata, source: Source): void {
  const objects = new Set((word.drawings ?? []).map((item) => item.nodeId)),
    transforms = wordDrawingTransforms(word, source);
  for (const item of word.drawings ?? []) {
    if (item.kind !== "connector") continue;
    const binding = item.connector;
    for (const side of ["startNodeId", "endNodeId"] as const)
      if (binding[side] && !objects.has(binding[side]!)) delete binding[side];
    if (!binding.startNodeId && !binding.endNodeId) continue;
    const g = item.geometry,
      matrix = transforms.world(item.nodeId),
      back = inverse(transforms.parent(item.nodeId));
    const a = point(
        back,
        binding.startNodeId
          ? transforms.endpoint(binding.startNodeId, binding.startSite ?? 0)
          : point(matrix, { x: 0, y: 0 }),
      ),
      b = point(
        back,
        binding.endNodeId
          ? transforms.endpoint(binding.endNodeId, binding.endSite ?? 0)
          : point(matrix, { x: g.width, y: g.height }),
      );
    const clean = (value: number): number => {
      if (!Number.isFinite(value))
        throw new DocumentError("InvalidOperation", "Connector geometry must be finite.");
      return Math.abs(value) < 1e-10 ? 0 : value;
    };
    item.geometry = {
      x: clean(Math.min(a.x, b.x)),
      y: clean(Math.min(a.y, b.y)),
      width: clean(Math.abs(a.x - b.x)),
      height: clean(Math.abs(a.y - b.y)),
      rotation: 0,
      flipH: b.x < a.x,
      flipV: b.y < a.y,
    };
  }
}

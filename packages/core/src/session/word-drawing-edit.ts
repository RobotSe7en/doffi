import { ensureWordMainStoryContainer } from "./word-story-container.js";
import { DocumentError } from "../errors.js";
import type { WordDocumentMetadata } from "../contracts/document.js";
import type { WordDrawing, WordDrawingOperation } from "../contracts/word-drawing.js";
import type { DocumentEdit } from "../document/edit.js";
import type { ParsedHtml } from "../document/html.js";
import type { DocumentReadLimits } from "../runtime.js";
import { checkLimit } from "../document/validation.js";
import { copyWordStorySemantics } from "./word-story-copy.js";
import { local, multiply, point } from "./presentation-geometry.js";
function fail(message: string): never {
  throw new DocumentError("InvalidOperation", message);
}
const escape = (text: string): string =>
  text
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replace(/\r\n?/g, "\n")
    .replaceAll("\n", "<br>");
export const isWordDrawingOperation: (operation: {
  readonly op: string;
}) => operation is WordDrawingOperation = (operation): operation is WordDrawingOperation =>
  [
    "setWordChart",
    "insertWordDrawing",
    "setWordDrawingGeometry",
    "setWordDrawingLayout",
    "setWordDrawingAppearance",
    "setWordDrawingCrop",
    "setWordDrawingPreset",
    "setWordDrawingText",
    "setWordDrawingConnector",
    "removeWordDrawing",
    "copyWordDrawing",
    "groupWordDrawings",
    "ungroupWordDrawing",
  ].includes(operation.op);
export function applyWordDrawing(
  metadata: WordDocumentMetadata,
  edit: DocumentEdit,
  parse: () => ParsedHtml,
  operation: WordDrawingOperation,
  limits: DocumentReadLimits,
): void {
  const word = metadata.word!,
    drawings = (word.drawings ??= []),
    source = parse(),
    target = source.peekNode(operation.nodeId);
  if (!target) fail("The drawing target is missing.");
  word.drawings = drawings;
  const article = source.nodeIds().find((id) => source.peekNode(id)!.parentNodeId === null)!;
  const story = (text: string): string => {
    ensureWordMainStoryContainer(metadata, edit, parse);
    const id = edit.insertSource(
      article,
      '<div data-office-story="textBox"><p>' + escape(text) + "</p></div>",
    ).rootNodeIds[0]!;
    word.stories.push({ storyId: id, kind: "textBox", rootNodeId: id });
    const normal = word.styles.find((style) => style.styleId === "Normal");
    if (normal) {
      const paragraph = parse().peekNode(id)!.childNodeIds[0]!;
      edit.attributes(paragraph, { "data-word-style": normal.styleId, class: normal.cssClass });
    }
    return id;
  };
  if (operation.op === "insertWordDrawing") {
    if (!/^p$|^h[1-6]$/.test(target.tagName)) fail("Insert a drawing in a source paragraph.");
    checkLimit(drawings.length + 1, limits.maxNodes);
    const object = operation.object,
      html =
        object.kind === "image"
          ? '<img data-word-drawing="image" src="' +
            escape(edit.resourceReference(object.assetPath)) +
            '" alt="' +
            escape(object.alt) +
            '">'
          : '<span data-word-drawing="' + object.kind + '">\ufffc</span>';
    const id = edit.insertSourceAt(operation.nodeId, operation.offset, html).rootNodeIds[0]!,
      base = { nodeId: id, geometry: operation.geometry, layout: operation.layout };
    let drawing: WordDrawing;
    if (object.kind === "chart") drawing = { ...base, kind: "chart", chart: object.chart };
    else if (object.kind === "image")
      drawing = { ...base, kind: "image", ...(object.crop ? { crop: object.crop } : {}) };
    else if (object.kind === "shape")
      drawing = {
        ...base,
        kind: "shape",
        preset: object.preset,
        ...(object.text !== undefined ? { textStoryId: story(object.text) } : {}),
      };
    else if (object.kind === "textBox")
      drawing = { ...base, kind: "textBox", textStoryId: story(object.text) };
    else drawing = { ...base, kind: "connector", connector: object.connector };
    drawings.push(drawing);
    for (const [property, value] of Object.entries(
      operation.properties ??
        (object.kind === "shape"
          ? {
              "--office-fill-kind": "solid",
              "--office-fill-color": "#4472c4",
              "--office-line-kind": "solid",
              "--office-line-color": "#224466",
              "--office-line-width": "1pt",
            }
          : {}),
    ))
      edit.sourceDrawingStyle(id, property, value);
    return;
  }
  if (operation.op === "groupWordDrawings") {
    if (new Set(operation.nodeIds).size !== operation.nodeIds.length)
      fail("Grouping requires distinct drawing identities.");
    const selected = operation.nodeIds.map(
      (id) => drawings.find((item) => item.nodeId === id) ?? fail("A selected drawing is missing."),
    );
    if (selected.some((item) => item.layout.relativeWidth || item.layout.relativeHeight))
      fail("Resolve relative drawing sizes to fixed extents before grouping.");
    const contains = (id: string): boolean => {
      for (
        let current: string | null = id;
        current;
        current = source.peekNode(current)!.parentNodeId
      )
        if (current === operation.nodeId) return true;
      return false;
    };
    if (
      selected.some(
        (item) =>
          !contains(item.nodeId) ||
          item.layout.mode !== "floating" ||
          !(() => {
            let parent = source.peekNode(item.nodeId)!.parentNodeId;
            while (parent && parent !== target.nodeId) {
              if (drawings.some((drawing) => drawing.nodeId === parent)) return false;
              parent = source.peekNode(parent)!.parentNodeId;
            }
            return parent === target.nodeId;
          })(),
      )
    )
      fail("Group floating drawings directly in the selected paragraph.");
    if (
      selected.some(
        (item) =>
          item.layout.horizontal !== selected[0]!.layout.horizontal ||
          item.layout.vertical !== selected[0]!.layout.vertical,
      )
    )
      fail("Grouped drawings must use the same coordinate reference.");
    const left = Math.min(...selected.map((item) => item.geometry.x)),
      top = Math.min(...selected.map((item) => item.geometry.y)),
      right = Math.max(...selected.map((item) => item.geometry.x + item.geometry.width)),
      bottom = Math.max(...selected.map((item) => item.geometry.y + item.geometry.height));
    const ordinal = new Map(source.nodeIds().map((id, index) => [id, index])),
      ids = selected.map((item) => item.nodeId).sort((a, b) => ordinal.get(a)! - ordinal.get(b)!);
    const offset = source.textOffset(ids[0]!)! - source.textOffset(operation.nodeId)!,
      id = edit.insertSourceAt(operation.nodeId, offset, '<span data-word-drawing="group"></span>')
        .rootNodeIds[0]!;
    edit.moveSource(ids, id);
    const layout = { ...selected[0]!.layout };
    for (const item of selected) {
      item.geometry = { ...item.geometry, x: item.geometry.x - left, y: item.geometry.y - top };
      item.layout = { mode: "inline" };
    }
    drawings.push({
      nodeId: id,
      kind: "group",
      geometry: {
        x: left,
        y: top,
        width: Math.max(0.01, right - left),
        height: Math.max(0.01, bottom - top),
        rotation: 0,
        flipH: false,
        flipV: false,
      },
      layout: { ...layout, mode: "floating" },
      coordinateSpace: {
        x: 0,
        y: 0,
        width: Math.max(0.01, right - left),
        height: Math.max(0.01, bottom - top),
      },
    });
    return;
  }
  const drawing = drawings.find((item) => item.nodeId === operation.nodeId);
  if (!drawing) fail("The Word drawing is missing.");
  if (operation.op === "setWordChart") {
    if (drawing.kind !== "chart") fail("Chart settings require a Word chart.");
    Object.assign(drawing.chart, operation.changes);
  } else if (operation.op === "setWordDrawingGeometry") drawing.geometry = operation.geometry;
  else if (operation.op === "setWordDrawingLayout") drawing.layout = operation.layout;
  else if (operation.op === "setWordDrawingAppearance")
    for (const [property, value] of Object.entries(operation.properties))
      edit.sourceDrawingStyle(drawing.nodeId, property, value);
  else if (operation.op === "setWordDrawingCrop") {
    if (drawing.kind !== "image") fail("Cropping requires an image.");
    drawing.crop = operation.crop;
  } else if (operation.op === "setWordDrawingPreset") {
    if (drawing.kind !== "shape") fail("Preset editing requires a shape.");
    drawing.preset = operation.preset;
  } else if (operation.op === "setWordDrawingText") {
    if (drawing.kind !== "shape" && drawing.kind !== "textBox")
      fail("Text properties require a shape or text box.");
    drawing.text = operation.text;
    if (!drawing.textStoryId) drawing.textStoryId = story("");
  } else if (operation.op === "setWordDrawingConnector") {
    if (drawing.kind !== "connector") fail("Connector editing requires a connector.");
    drawing.connector = operation.connector;
  } else if (operation.op === "removeWordDrawing") {
    const removed = new Set<string>(),
      visit = (id: string): void => {
        removed.add(id);
        for (const child of source.peekNode(id)!.childNodeIds) visit(child);
      };
    visit(drawing.nodeId);
    for (const item of drawings)
      if (removed.has(item.nodeId) && "textStoryId" in item && item.textStoryId) {
        const associated = word.stories.find((story) => story.storyId === item.textStoryId);
        if (associated) edit.removeSource(associated.rootNodeId);
      }
    edit.removeSource(drawing.nodeId);
    word.drawings = drawings.filter((item) => !removed.has(item.nodeId));
    word.stories = word.stories.filter((story) => parse().hasNode(story.rootNodeId));
    return;
  } else if (operation.op === "copyWordDrawing") {
    const parent = target.parentNodeId!,
      siblings = source.peekNode(parent)!.childNodeIds,
      clone = edit.cloneSource(
        drawing.nodeId,
        parent,
        siblings[siblings.indexOf(drawing.nodeId) + 1],
      ),
      id = clone.rootNodeIds[0]!;
    copyWordStorySemantics(word, clone.identities, edit, source, id);
    const copy = drawings.find((item) => item.nodeId === id)!;
    copy.geometry = {
      ...copy.geometry,
      x: copy.geometry.x + (operation.offsetX ?? 12),
      y: copy.geometry.y + (operation.offsetY ?? 12),
    };
    return;
  } else if (operation.op === "ungroupWordDrawing") {
    if (drawing.kind !== "group") fail("Ungrouping requires a group.");
    if (drawing.layout.relativeWidth || drawing.layout.relativeHeight)
      fail("Resolve the relative group size to fixed extents before ungrouping.");
    const parent = target.parentNodeId!,
      siblings = source.peekNode(parent)!.childNodeIds,
      before = siblings[siblings.indexOf(drawing.nodeId) + 1],
      children = drawings.filter(
        (item) => source.peekNode(item.nodeId)?.parentNodeId === drawing.nodeId,
      ),
      g = drawing.geometry,
      frame = drawing.coordinateSpace;
    const sx = g.width / frame.width,
      sy = g.height / frame.height,
      groupTransform = multiply(local(g), [sx, 0, 0, sy, -frame.x * sx, -frame.y * sy]);
    for (const child of children) {
      const cg = child.geometry,
        matrix = multiply(groupTransform, local(cg)),
        scaleX = Math.hypot(matrix[0], matrix[1]),
        scaleY = Math.hypot(matrix[2], matrix[3]),
        dot = matrix[0] * matrix[2] + matrix[1] * matrix[3];
      if (Math.abs(dot) > 1e-9 * Math.max(1, scaleX * scaleY)) {
        // A rotated child of an anisotropically scaled group has shear. DrawingML
        // cannot encode that in a leaf xfrm: retain a single-object carrier instead
        // of approximating the child's paint, text or connector coordinates.
        const carrier = edit.insertSource(parent, '<span data-word-drawing="group"></span>', before)
          .rootNodeIds[0]!;
        edit.moveSource([child.nodeId], carrier);
        drawings.push({
          ...drawing,
          nodeId: carrier,
          geometry: { ...g },
          coordinateSpace: { ...frame },
        });
        continue;
      }
      const width = cg.width * scaleX,
        height = cg.height * scaleY,
        center = point(matrix, { x: cg.width / 2, y: cg.height / 2 }),
        flipH = g.flipH !== cg.flipH,
        det = matrix[0] * matrix[3] - matrix[1] * matrix[2];
      child.geometry = {
        x: center.x - width / 2,
        y: center.y - height / 2,
        width,
        height,
        rotation:
          (Math.atan2(matrix[1] * (flipH ? -1 : 1), matrix[0] * (flipH ? -1 : 1)) * 180) / Math.PI,
        flipH,
        flipV: det < 0 !== flipH,
      };
      child.layout = { ...drawing.layout };
      edit.moveSource([child.nodeId], parent, before);
    }
    edit.removeSource(drawing.nodeId);
    word.drawings = drawings.filter((item) => item !== drawing);
    return;
  }
}

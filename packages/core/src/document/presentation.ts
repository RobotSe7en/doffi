import { parse, walk, ident } from "css-tree";
import type { Anchor,Annotation } from "../contracts/annotations.js";
import type { PresentationGeometry, PresentationMetadata, PresentationObject } from "../contracts/presentation.js";
import { PresentationGeometrySchema } from "../contracts/presentation.js";
import type { ParsedHtml } from "./html.js";
import type { DeepReadonly,DocumentReadLimits } from "../runtime.js";
import { checkLimit, invalid } from "./validation.js";
import { evaluatePresentationCustomGeometry } from "../presentation-custom-geometry.js";

const geometryNames = ["x", "y", "width", "height"] as const;
const numeric = /^-?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i;
export function placeholderReferenceKeys(placeholder:{readonly type:string;readonly index:number}):readonly string[]{return[placeholder.type+":"+placeholder.index,...(placeholder.type==="title"||placeholder.type==="ctrTitle"?[(placeholder.type==="title"?"ctrTitle":"title")+":"+placeholder.index]:[])];}
/** Native classic positions are fixed slide points, independent of font layout. */
export function isPresentationCommentPoint(anchor:DeepReadonly<Anchor>):boolean{return anchor.kind==="region"&&anchor.layoutKey==="pptx-classic-point"&&anchor.rect.width===1&&anchor.rect.height===1&&!anchor.nodeIds.length;}
function number(value: string | undefined): number {
  if (value === undefined || !numeric.test(value) || !Number.isFinite(Number(value))) invalid("Presentation geometry must contain finite numbers.");
  return Number(value);
}
function boolean(value: string | undefined): boolean {
  if (value === undefined || value === "false") return false;
  if (value === "true") return true;
  return invalid("Presentation flips must be true or false.");
}
/** Source-local transform. CSS positioning and screen scale are never inputs. */
export function readPresentationGeometry(attributes: Readonly<Record<string, string>>): PresentationGeometry | undefined {
  if (!geometryNames.some((name) => attributes["data-office-" + name] !== undefined)) {
    if (["rotation", "flip-h", "flip-v"].some((name) => attributes["data-office-" + name] !== undefined)) invalid("A partial presentation transform is invalid.");
    return undefined;
  }
  const value = {
    x: number(attributes["data-office-x"]), y: number(attributes["data-office-y"]),
    width: number(attributes["data-office-width"]), height: number(attributes["data-office-height"]),
    rotation: attributes["data-office-rotation"] === undefined ? 0 : number(attributes["data-office-rotation"]),
    flipH: boolean(attributes["data-office-flip-h"]), flipV: boolean(attributes["data-office-flip-v"]),
  };
  if (!PresentationGeometrySchema.safeParse(value).success) invalid("Presentation dimensions must be nonnegative.");
  return value;
}
function unique<Value>(values: readonly Value[], key: (value: Value) => string): Map<string, Value> {
  const result = new Map<string, Value>();
  for (const value of values) {
    const id = key(value);
    if (result.has(id)) invalid("Presentation identities must be unique within their definition type.");
    result.set(id, value);
  }
  return result;
}

/** Validate references against the one HTML/CSS source, without DOM or Office I/O. */
export function validatePresentation(presentation: PresentationMetadata, annotations: readonly Annotation[], html: ParsedHtml, styles: readonly { css: string }[], limits: DocumentReadLimits): void {
  let work = Object.values(presentation).reduce<number>((count, value) => count + (Array.isArray(value) ? value.length : 0), 0);
  checkLimit(work, limits.maxNodes);
  for (const object of presentation.objects) if (object.kind === "table") { work += object.grid.length; checkLimit(work, limits.maxNodes); }
  for (const dataSet of presentation.dataSets) {
    work += dataSet.categories.length + dataSet.series.length;
    for (const series of dataSet.series) work += series.values.length;
    checkLimit(work, limits.maxNodes);
  }
  for (const annotation of annotations) { work += 1 + (annotation.replies?.length ?? 0); checkLimit(work, limits.maxNodes); }
  for (const chart of presentation.charts) { work += (chart.seriesStyles?.length ?? 0) + (chart.pointStyles?.length ?? 0) + (chart.labelContents?.length ?? 0); checkLimit(work, limits.maxNodes); }
  for (const comment of presentation.comments) { work += comment.replyAuthors?.length ?? 0; checkLimit(work, limits.maxNodes); }
  const classes = new Set<string>();
  for (const style of styles) walk(parse(style.css), (node) => { if (node.type === "ClassSelector") classes.add(ident.decode(node.name)); });
  const themes = unique(presentation.themes, (item) => item.themeId);
  const masters = unique(presentation.masters, (item) => item.masterId);
  const layouts = unique(presentation.layouts, (item) => item.layoutId);
  const slides = unique(presentation.slides, (item) => item.slideId);
  for (const item of unique(presentation.nativePreview ?? [], item => item.feature + ":" + item.ownerId).values()) {
    if (item.feature === "theme-defaults" && !themes.has(item.ownerId) || item.feature === "animation" && !slides.has(item.ownerId) ||
        item.feature === "tags" && !html.hasNode(item.ownerId)) invalid("A preserved PPTX fragment must have a current owner.");
  }
  const objects = unique(presentation.objects, (item) => item.nodeId);
  const charts = unique(presentation.charts, (item) => item.chartId);
  const dataSets = unique(presentation.dataSets, (item) => item.dataSetId);
  const notesMasters = unique(presentation.notesMasters, (item) => item.notesMasterId);
  const commentAuthors = unique(presentation.commentAuthors ?? [], (item) => item.authorId);
  unique(presentation.stories, (item) => item.storyId);
  const roots = new Map<string, string>();
  const node = (id: string) => {
    const value = html.peekNode(id);
    if (!value) invalid("A presentation reference targets a missing source node.");
    return value;
  };
  const presentationRoots = html.nodeIds().filter((id) => node(id).attributes["data-office-kind"] === "presentation");
  if (presentationRoots.length !== 1 || node(presentationRoots[0]!).parentNodeId !== null) invalid("Presentation requires exactly one root.");
  const root = presentationRoots[0]!;
  if (html.directText(root)?.trim()) invalid("Unregistered text cannot appear directly in a presentation root.");
  const addRoot = (id: string, kind: string): void => {
    const value = node(id);
    if (roots.has(id) || !html.isContainer(id) || value.parentNodeId !== root
      || value.attributes["data-office-kind"] !== kind || value.attributes["data-office-story"] !== kind) invalid("Presentation content roots require distinct, matching source containers.");
    roots.set(id, kind);
  };
  for (const theme of themes.values()) if (!classes.has(theme.cssClass)) invalid("A presentation theme refers to an undefined CSS class.");
  for (const master of masters.values()) {
    if (!themes.has(master.themeId)) invalid("A slide master refers to a missing theme.");
    for (const cssClass of [master.titleStyle, master.bodyStyle, master.otherStyle])
      if (cssClass !== undefined && !classes.has(cssClass)) invalid("A master text style refers to an undefined CSS class.");
    addRoot(master.rootNodeId, "master");
  }
  for (const layout of layouts.values()) {
    if (!masters.has(layout.masterId)) invalid("A slide layout refers to a missing master.");
    addRoot(layout.rootNodeId, "layout");
  }
  for (const slide of slides.values()) {
    if (!layouts.has(slide.layoutId)) invalid("A slide refers to a missing layout.");
    addRoot(slide.rootNodeId, "slide");
  }
  for (const master of presentation.notesMasters) {
    if (!themes.has(master.themeId)) invalid("A notes master refers to a missing theme.");
    if (master.textStyle !== undefined && !classes.has(master.textStyle)) invalid("A notes master text style must refer to a defined CSS class.");
    addRoot(master.rootNodeId, "notes-master");
  }
  const notes = new Set<string>();
  for (const story of presentation.stories) {
    if (story.slideId !== undefined && !slides.has(story.slideId)) invalid("A presentation story refers to a missing slide.");
    if (story.kind === "notes") {
      if (!story.slideId || notes.has(story.slideId)) invalid("Each notes story requires its own existing slide.");
      if (story.notesMasterId !== undefined && !notesMasters.has(story.notesMasterId)) invalid("A notes story refers to a missing notes master.");
      notes.add(story.slideId);
    }
    addRoot(story.rootNodeId, story.kind);
  }
  for (const id of node(root).childNodeIds) if (!roots.has(id)) invalid("Presentation root children must be registered source surfaces.");
  const owner = (id: string): string => {
    for (let value: string | null = id; value !== null; value = node(value).parentNodeId) if (roots.has(value)) return value;
    return invalid("Presentation objects must belong to a registered surface.");
  };
  const bySurface = new Map<string, PresentationObject[]>();
  const placeholders = new Map<string, Map<string, PresentationObject>>();
  const transforms = new Map<string, PresentationGeometry | undefined>();
  const slidesByRoot = new Map(presentation.slides.map((item) => [item.rootNodeId, item]));
  const layoutsByRoot = new Map(presentation.layouts.map((item) => [item.rootNodeId, item]));
  const notesByRoot = new Map(presentation.stories.filter(item => item.kind === "notes").map(item => [item.rootNodeId, item]));
  const nativeIds = new Map<string, Set<string>>();
  const placeholderKeys = new Map<string, Set<string>>();
  for (const object of objects.values()) {
    const value = node(object.nodeId), surface = owner(object.nodeId), parent = value.parentNodeId;
    if (!html.isContainer(object.nodeId) || value.attributes["data-office-kind"] !== object.kind
      || parent === null || parent !== surface && objects.get(parent)?.kind !== "group") invalid("A presentation object has an invalid source type or parent.");
    if (roots.get(surface) === "comment") invalid("Rich comments cannot contain positioned slide objects.");
    const transform = readPresentationGeometry(value.attributes);
    const connectorPreset = value.attributes["data-office-connector-preset"];
    if (connectorPreset !== undefined && (object.kind !== "connector" || !(object.routing === "straight" ? ["line", "straightConnector1"] : ["bentConnector2", "bentConnector3", "bentConnector4", "bentConnector5"]).includes(connectorPreset))) invalid("A native connector preset must match its routing kind.");
    for (const [name, raw] of Object.entries(value.attributes)) if (name.startsWith("data-office-adjust-")) {
      if (!["geometry", "connector", "text"].includes(object.kind) || !/^data-office-adjust-[a-z][a-z0-9]*$/.test(name) || !/^-?\d+$/.test(raw) || !Number.isSafeInteger(Number(raw))) invalid("Preset adjustments require safe integer constants on a shape or connector.");
    }
    transforms.set(object.nodeId, transform);
    if (!transform && (!object.placeholder || parent !== surface)) invalid("Positioned objects require a local transform; only surface placeholders may inherit one.");
    if (transform && object.kind !== "connector" && !(object.kind === "geometry" && object.preset === "line") && (!transform.width || !transform.height) &&
        !(value.attributes["data-office-hidden"] === "true" && (object.kind === "text" || object.kind === "geometry"))) invalid("Visible two-dimensional objects require positive dimensions.");
    if ((object.kind === "geometry" || object.kind === "image") && object.customGeometry) {
      if (!transform || (object.kind === "geometry" && object.preset !== "rect")) invalid("Custom geometry requires a local transform and its base geometry slot.");
      if (object.kind === "image" && (object.customGeometry.paths.length !== 1 || object.customGeometry.paths[0]!.commands.at(-1)?.kind !== "close"))
        invalid("A custom picture mask requires one closed path.");
      for (const path of object.customGeometry.paths) { work += path.commands.length; checkLimit(work, limits.maxNodes); }
      evaluatePresentationCustomGeometry(object.customGeometry, transform.width, transform.height);
    }
    const opacity = value.attributes["data-office-image-opacity"];
    if (opacity !== undefined && (object.kind !== "image" || number(opacity) < 0 || number(opacity) > 1)) invalid("Picture opacity must be an image alpha factor in [0, 1].");
    if (object.nativeId !== undefined) {
      const ids = nativeIds.get(surface) ?? new Set<string>();
      if (ids.has(object.nativeId)) invalid("Native object identities must be unique within their surface.");
      ids.add(object.nativeId); nativeIds.set(surface, ids);
    }
    if (object.placeholder) {
      const key = object.placeholder.type + ":" + object.placeholder.index;
      const keys = placeholderKeys.get(surface) ?? new Set<string>();
      if (keys.has(key)) invalid("Placeholder identities must be unique within a surface.");
      keys.add(key); placeholderKeys.set(surface, keys);
      const entries = placeholders.get(surface) ?? new Map<string, PresentationObject>();
      entries.set(key, object); placeholders.set(surface, entries);
    }
    if (object.kind === "group") {
      for (const name of geometryNames) {
        const part = number(value.attributes["data-office-child-" + name]);
        if (["width", "height"].includes(name) && part <= 0) invalid("Group child extents must be positive.");
      }
      if (value.childNodeIds.some((id) => !objects.has(id))) invalid("A group may only contain registered objects.");
      if (html.directText(object.nodeId)?.trim()) invalid("Unregistered text cannot appear directly in a group.");
    }
    if (object.kind === "image") {
      if (value.childNodeIds.length !== 1 || node(value.childNodeIds[0]!).tagName !== "img" || html.directText(object.nodeId)?.trim()) invalid("An image object requires exactly one image source.");
      for (const name of ["left", "top", "right", "bottom"]) {
        const raw = value.attributes["data-office-crop-" + name];
        if (raw !== undefined && (number(raw) < 0 || number(raw) >= 1)) invalid("Image crop fractions must lie in [0, 1).");
      }
      const crop = (name: string): number => Number(value.attributes["data-office-crop-" + name] ?? 0);
      if (crop("left") + crop("right") >= 1 || crop("top") + crop("bottom") >= 1) invalid("An image crop must retain a visible area.");
    }
    if (object.kind === "table") {
      const tables = value.childNodeIds.filter((id) => node(id).tagName === "table");
      if (tables.length !== 1 || value.childNodeIds.length !== 1 || html.directText(object.nodeId)?.trim()) invalid("A table object requires exactly one table source.");
      const rows = node(tables[0]!).childNodeIds.flatMap((id) => ["thead", "tbody", "tfoot"].includes(node(id).tagName) ? node(id).childNodeIds : [id]);
      const occupied = new Array<number>(object.grid.length).fill(0);
      for (const [rowIndex, rowId] of rows.entries()) {
        if (node(rowId).tagName !== "tr") invalid("Presentation table children must be rows.");
        let column = 0;
        for (const cellId of node(rowId).childNodeIds) {
          const cell = node(cellId);
          if (!["td", "th"].includes(cell.tagName)) invalid("Presentation table rows require cells.");
          while (column < occupied.length && occupied[column]! > rowIndex) { checkLimit(++work, limits.maxNodes); column++; }
          const span = Number(cell.attributes.colspan ?? 1);
          const rowSpan = Number(cell.attributes.rowspan ?? 1) || rows.length - rowIndex;
          if (column + span > occupied.length || rowIndex + rowSpan > rows.length) invalid("A presentation table span exceeds its grid.");
          for (let end = column + span; column < end; column++) {
            checkLimit(++work, limits.maxNodes);
            if (occupied[column]! > rowIndex) invalid("Presentation table spans overlap.");
            occupied[column] = rowIndex + rowSpan;
          }
        }
        for (const end of occupied) { checkLimit(++work, limits.maxNodes); if (end <= rowIndex) invalid("Presentation table cells must cover their row grid."); }
      }
      if (!rows.length) invalid("A presentation table requires rows.");
    }
    if (object.kind === "chart" && charts.get(object.chartId)?.nodeId !== object.nodeId) invalid("A chart object requires a matching chart definition.");
    if (object.kind === "connector") for (const target of [object.startNodeId, object.endNodeId])
      if (target !== undefined && (!objects.has(target) || target === object.nodeId || owner(target) !== surface)) invalid("Connector endpoints must be other objects on the same surface.");
    const list = bySurface.get(surface) ?? []; list.push(object); bySurface.set(surface, list);
  }
  // The inherited placeholder transform must resolve through layout/master, not
  // through a nearby object or a duplicate copy of a rendered instance.
  for (const [surface, list] of bySurface) for (const object of list) {
    if (transforms.get(object.nodeId)) continue;
    const slide = slidesByRoot.get(surface);
    const layout = slide ? layouts.get(slide.layoutId) : layoutsByRoot.get(surface);
    const notes = notesByRoot.get(surface);
    if (object.placeholder && (notes || roots.get(surface) === "notes-master")) continue;
    const parents = layout ? [layout.rootNodeId, masters.get(layout.masterId)!.rootNodeId].filter((id) => id !== surface)
      : notes?.notesMasterId ? [notesMasters.get(notes.notesMasterId)!.rootNodeId] : [];
    const keys=placeholderReferenceKeys(object.placeholder!);
    if (!parents.some((id) => {
      const inherited=keys.map(key=>placeholders.get(id)?.get(key)).find(value=>value!==undefined);
      return inherited !== undefined && transforms.get(inherited.nodeId) !== undefined;
    })) invalid("A placeholder transform cannot be resolved from its layout or master.");
  }
  for (const id of html.nodeIds()) {
    const value = node(id), kind = value.attributes["data-office-kind"];
    if (kind !== undefined && id !== root && !roots.has(id) && !objects.has(id)) invalid("An unregistered presentation source role is invalid.");
    if (Object.keys(value.attributes).some((key) => key.startsWith("data-word-"))) invalid("Word semantics cannot be embedded in a presentation source.");
    if (value.attributes["data-office-story"] !== undefined && !roots.has(id)) invalid("Presentation stories must be registered content roots.");
    if (roots.has(id) && !["notes", "comment"].includes(roots.get(id)!) && (value.childNodeIds.some((child) => !objects.has(child)) || html.directText(id)?.trim())) invalid("A slide surface may only contain registered objects.");
  }
  for (const story of presentation.stories) if (story.kind === "notes") {
    const children = node(story.rootNodeId).childNodeIds;
    if (children.some(id => objects.has(id)) && (children.some(id => !objects.has(id)) || html.directText(story.rootNodeId)?.trim())) invalid("Native notes cannot mix positioned objects and unpositioned rich text.");
    if (story.bodyNodeId !== undefined && (!objects.has(story.bodyNodeId) || owner(story.bodyNodeId) !== story.rootNodeId || !["text", "geometry"].includes(objects.get(story.bodyNodeId)!.kind))) invalid("A notes body reference must target text on its own notes surface.");
  }
  for (const dataSet of dataSets.values()) {
    unique(dataSet.series, (item) => item.seriesId);
    if (dataSet.series.some((series) => series.values.length !== dataSet.categories.length)) invalid("Chart categories and series lengths must match.");
  }
  const negativeDataSets = new Set(presentation.dataSets.filter((dataSet) => dataSet.series.some((series) => series.values.some((value) => value !== null && value < 0))).map((item) => item.dataSetId));
  for (const chart of charts.values()) {
    const object = objects.get(chart.nodeId), dataSet = dataSets.get(chart.dataSetId);
    if (!object || object.kind !== "chart" || object.chartId !== chart.chartId || !dataSet) invalid("Chart references must resolve to their object and data set.");
    if (chart.titleNodeId !== undefined) {
      let parent: string | null = chart.titleNodeId;
      while (parent !== null && parent !== chart.nodeId) parent = node(parent).parentNodeId;
      if (parent !== chart.nodeId || chart.titleNodeId === chart.nodeId) invalid("A chart title must be inside its chart source.");
    }
    const stack = [...node(chart.nodeId).childNodeIds];
    if (html.directText(chart.nodeId)?.trim()) invalid("Chart text requires an explicitly referenced title.");
    while (stack.length) {
      checkLimit(++work, limits.maxNodes);
      const id = stack.pop()!;
      if (id === chart.titleNodeId) continue;
      const value = node(id);
      if (html.directText(id)?.trim() || ["img", "br", "table"].includes(value.tagName)) invalid("Chart content outside its explicit title cannot be preserved.");
      stack.push(...value.childNodeIds);
    }
    if (chart.minimum !== undefined && chart.maximum !== undefined && chart.minimum >= chart.maximum) invalid("Chart axis bounds are invalid.");
    for (const cssClass of [chart.plotCssClass, chart.legendCssClass, chart.textCssClass, chart.labelCssClass, chart.categoryAxisCssClass, chart.valueAxisCssClass, chart.gridCssClass, chart.categoryGridCssClass]) if (cssClass !== undefined && !classes.has(cssClass)) invalid("Chart styles must refer to defined CSS classes.");
    const seriesIds = new Set(dataSet.series.map(item => item.seriesId));
    for (const style of unique(chart.seriesStyles ?? [], item => item.seriesId).values()) if (!seriesIds.has(style.seriesId) || !classes.has(style.cssClass) || style.labelCssClass !== undefined && !classes.has(style.labelCssClass)) invalid("Chart series styles must reference data and CSS classes.");
    for (const style of unique(chart.pointStyles ?? [], item => item.seriesId + ":" + item.index).values()) if (!seriesIds.has(style.seriesId) || style.index >= dataSet.categories.length || !classes.has(style.cssClass) || style.labelCssClass !== undefined && !classes.has(style.labelCssClass)) invalid("Chart point styles must reference data and CSS classes.");
    if (chart.labelContents && (new Set(chart.labelContents).size !== chart.labelContents.length || chart.showLabels !== !!chart.labelContents.length)) invalid("Chart label options must be unique and agree with label visibility.");
    if (chart.kind === "pie" && (dataSet.series.length !== 1 || negativeDataSets.has(dataSet.dataSetId))) invalid("A pie chart requires one nonnegative series.");
  }
  for (const link of unique(presentation.hyperlinks, (item) => item.nodeId).values()) {
    node(link.nodeId);
    if ((link.target === undefined) === (link.slideId === undefined) || link.slideId !== undefined && !slides.has(link.slideId)
      || link.target !== undefined && (!/^(https?:\/\/|mailto:)/i.test(link.target) || /[\u0000-\u0020\u007f]/.test(link.target))) invalid("Presentation hyperlinks require one supported target.");
  }
  const annotationMap = unique(annotations, (item) => item.annotationId);
  const authorNames = new Map<string, string>();
  const checkAuthor = (id: string, format: "classic" | "modern", author: Annotation["author"]): void => {
    const definition = commentAuthors.get(id);
    if (!definition || definition.format !== format || !author) invalid("A native comment author reference is invalid.");
    const display = JSON.stringify(author), previous = authorNames.get(id);
    if (previous !== undefined && previous !== display) invalid("One native author identity cannot have conflicting display information.");
    authorNames.set(id, display);
  };
  for (const comment of unique(presentation.comments, (item) => item.annotationId).values()) {
    const annotation = annotationMap.get(comment.annotationId);
    if (!annotation) invalid("Native comment metadata refers to a missing annotation.");
    if (comment.recoveryStatus !== undefined && annotation.status !== "needs_reanchor") invalid("Comment recovery workflow state requires an unresolved anchor.");
    if (comment.authorId) checkAuthor(comment.authorId, comment.format, annotation.author);
    for (const reply of unique(comment.replyAuthors ?? [], item => item.replyId).values()) {
      const source = annotation.replies?.find(item => item.replyId === reply.replyId);
      if (!source || comment.format !== "modern") invalid("Native reply authors require a reply in their modern thread.");
      checkAuthor(reply.authorId, comment.format, source.author);
    }
  }
  unique(presentation.commentAuthors ?? [], item => item.format + ":" + item.nativeId);
  for (const author of commentAuthors.values()) if (author.format === "classic" && (author.userId !== undefined || author.providerId !== undefined)) invalid("Classic authors cannot contain modern provider attributes.");
  const replyIds = new Set(annotationMap.keys());
  const bodies = new Set<string>();
  for (const annotation of annotations) for (const body of [annotation, ...(annotation.replies ?? [])]) {
    if ("replyId" in body) {
      if (replyIds.has(body.replyId)) invalid("Replies require distinct document-wide identities.");
      replyIds.add(body.replyId);
    }
    if (body.bodyNodeId !== undefined) {
      if (roots.get(body.bodyNodeId) !== "comment" || bodies.has(body.bodyNodeId)) invalid("Rich comments require distinct comment source roots.");
      bodies.add(body.bodyNodeId);
    }
  }
}

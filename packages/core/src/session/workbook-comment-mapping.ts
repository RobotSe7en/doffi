import type { Annotation } from "../contracts/annotations.js";
import type { WorkbookCoordinate, WorkbookMetadata, WorkbookRange } from "../contracts/workbook.js";
import type { DeepReadonly, WorkbookStructureMapping } from "../runtime.js";
import { mapWorkbookRanges } from "../workbook/range-mapping.js";

type Step = Extract<WorkbookStructureMapping, { kind: "axis" | "move" | "sort" }>;
interface Origin {
  readonly anchor: DeepReadonly<{ readonly sheetId: string; readonly ranges: readonly WorkbookRange[] }>;
  readonly status: Exclude<Annotation["status"], "needs_reanchor">; readonly start: number;
  readonly replies: ReadonlyMap<string, DeepReadonly<WorkbookCoordinate>>;
}
/** A batch may temporarily split one annotation between worksheets. Retain
 * its exact pre-step identities privately so a later inverse move can compose
 * them back; no intermediate invalidation becomes an extra user decision. */
export class WorkbookCommentMapping {
  private readonly origins = new Map<string, Origin>();
  capture(workbook: WorkbookMetadata, annotations: readonly Annotation[], start: number): void {
    const bindings = new Map(workbook.comments?.map(comment => [comment.annotationId, comment]));
    for (const annotation of annotations) if (!this.origins.has(annotation.annotationId) && annotation.status !== "needs_reanchor" && annotation.anchor.kind === "cells") {
      const binding = bindings.get(annotation.annotationId);
      this.origins.set(annotation.annotationId, { anchor: { sheetId: annotation.anchor.sheetId, ranges: annotation.anchor.ranges.map(range => ({ from: { ...range.from }, to: { ...range.to } })) },
        status: annotation.status, start, replies: new Map(binding?.format === "modern" ? binding.replies.flatMap(reply => reply.nativeCell ? [[reply.replyId, { ...reply.nativeCell }]] : []) : []) });
    }
  }
  private apply(workbook: WorkbookMetadata, annotation: Annotation, steps: readonly WorkbookStructureMapping[], origin: Origin): void {
    const mapped = mapWorkbookRanges(origin.anchor.ranges.map(range => ({ sheetId: origin.anchor.sheetId, range })), steps.slice(origin.start));
    const sheets = new Set(mapped.map(value => value.sheetId)), binding = workbook.comments?.find(comment => comment.annotationId === annotation.annotationId);
    if (!mapped.length || sheets.size !== 1) { annotation.status = "needs_reanchor"; if (binding) binding.recoveryStatus = origin.status; return; }
    const sheetId = mapped[0]!.sheetId;
    annotation.anchor = { kind: "cells", sheetId, ranges: mapped.map(value => ({ from: { ...value.range.from }, to: { ...value.range.to } })) }; annotation.status = origin.status;
    if (binding) { delete binding.recoveryStatus; delete binding.recoveryReasons; }
    for (const message of [annotation, ...annotation.replies ?? []]) if (message.bodyNodeId) {
      const story = workbook.stories?.find(story => story.rootNodeId === message.bodyNodeId); if (story) story.sheetId = sheetId;
    }
    if (binding?.format === "modern") for (const reply of binding.replies) {
      const old = origin.replies.get(reply.replyId); if (!old) continue;
      const points = mapWorkbookRanges([{ sheetId: origin.anchor.sheetId, range: { from: old, to: old } }], steps.slice(origin.start));
      if (points.length === 1 && points[0]!.sheetId === sheetId) reply.nativeCell = { ...points[0]!.range.from }; else delete reply.nativeCell;
    }
  }
  explicit(workbook: WorkbookMetadata, annotations: readonly Annotation[], steps: readonly WorkbookStructureMapping[], id: string): void {
    const origin = this.origins.get(id), annotation = annotations.find(annotation => annotation.annotationId === id);
    if (origin && annotation) this.apply(workbook, annotation, steps, origin); this.origins.delete(id);
  }
  *finish(workbook: WorkbookMetadata, annotations: readonly Annotation[], steps: readonly WorkbookStructureMapping[]): Generator<{ completed: number; total: number }, void> {
    let completed = 0;
    for (const annotation of annotations) { const origin = this.origins.get(annotation.annotationId); if (!origin) continue; this.apply(workbook, annotation, steps, origin); if (++completed % 64 === 0) yield { completed, total: 0 }; }
  }
}
/** Apply exact cell identity mappings to both host and native threads. Bodies
 * remain source-owned stories; a lost/mixed-sheet result retains the old
 * anchor as recovery provenance instead of choosing one fragment. */
export function* mapWorkbookCommentAnchors(workbook: WorkbookMetadata, annotations: Annotation[], step: Step): Generator<{ completed: number; total: number }, void> {
  const bindings = new Map(workbook.comments?.map(comment => [comment.annotationId, comment])), stories = new Map(workbook.stories?.map(story => [story.rootNodeId, story]));
  let completed = 0;
  for (const annotation of annotations) {
    if (annotation.status === "needs_reanchor" || annotation.anchor.kind !== "cells") continue;
    const originalSheet = annotation.anchor.sheetId;
    if (step.kind === "move" ? originalSheet !== step.source.sheetId && originalSheet !== step.target.sheetId : originalSheet !== step.sheetId) continue;
    const mapped = mapWorkbookRanges(annotation.anchor.ranges.map(range => ({ sheetId: originalSheet, range })), [step]), sheets = new Set(mapped.map(value => value.sheetId));
    const binding = bindings.get(annotation.annotationId);
    if (!mapped.length || sheets.size !== 1) {
      if (binding) binding.recoveryStatus = annotation.status;
      annotation.status = "needs_reanchor";
    } else {
      const sheetId = mapped[0]!.sheetId;
      annotation.anchor = { kind: "cells", sheetId, ranges: mapped.map(value => ({ from: { ...value.range.from }, to: { ...value.range.to } })) };
      for (const message of [annotation, ...annotation.replies ?? []]) if (message.bodyNodeId) { const story = stories.get(message.bodyNodeId); if (story) story.sheetId = sheetId; }
      if (binding?.format === "modern") for (const reply of binding.replies) if (reply.nativeCell) {
        const points = mapWorkbookRanges([{ sheetId: originalSheet, range: { from: reply.nativeCell, to: reply.nativeCell } }], [step]);
        if (points.length === 1 && points[0]!.sheetId === sheetId) reply.nativeCell = { ...points[0]!.range.from }; else delete reply.nativeCell;
      }
    }
    if (++completed % 64 === 0) yield { completed, total: 0 };
  }
  yield { completed, total: completed };
}

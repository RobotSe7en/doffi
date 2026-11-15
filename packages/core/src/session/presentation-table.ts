import type { PresentationOperation } from "../contracts/presentation-operations.js";
import type { PresentationEditContext } from "./presentation-edit.js";
import { invalid } from "../document/validation.js";
import { applyPresentationGeometry, presentationTransforms } from "./presentation-geometry.js";
import { editSourceTable } from "./table-edit.js";
/** Presentation geometry stays outside the shared HTML table editor. */
export function editPresentationTable(
  ctx: PresentationEditContext,
  operation: Extract<PresentationOperation, { op: "editPresentationTable" }>,
): void {
  const object = ctx.p.objects.find((object) => object.nodeId === operation.nodeId);
  if (object?.kind !== "table") invalid("Table operations require a current table object.");
  const html = ctx.parse(),
    tableId = html.peekNode(object.nodeId)!.childNodeIds[0]!,
    geometry = presentationTransforms(ctx.p, html).geometry(object.nodeId);
  const result = editSourceTable(ctx, tableId, object.grid, geometry, operation.change);
  object.grid = result.grid;
  if (result.width !== geometry.width || result.height !== geometry.height)
    applyPresentationGeometry(
      {
        op: "setPresentationGeometry",
        nodeId: object.nodeId,
        geometry: { ...geometry, width: result.width, height: result.height },
      },
      ctx.p,
      ctx.edit,
      ctx.parse(),
    );
}

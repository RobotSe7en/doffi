import { z } from "zod";
export const DrawingGeometrySchema = z.strictObject({
  x: z.number(),
  y: z.number(),
  width: z.number().nonnegative(),
  height: z.number().nonnegative(),
  rotation: z.number(),
  flipH: z.boolean(),
  flipV: z.boolean(),
});

/** Shared first-party geometry/text subset for independently anchored objects. */
export const DrawingPresetSchema = z.enum([
  "rect",
  "roundRect",
  "ellipse",
  "triangle",
  "diamond",
  "parallelogram",
  "trapezoid",
  "pentagon",
  "hexagon",
  "rightArrow",
  "leftArrow",
  "upArrow",
  "downArrow",
  "leftRightArrow",
  "chevron",
  "line",
]);
export const DrawingTextPropertiesSchema = z.strictObject({
  autoFit: z.enum(["none", "normal", "shape"]).optional(),
  wrap: z.boolean().optional(),
  fontScale: z.number().positive().max(1).optional(),
  lineSpacingReduction: z.number().min(0).max(1).optional(),
});
export type DrawingTextProperties = z.infer<typeof DrawingTextPropertiesSchema>;

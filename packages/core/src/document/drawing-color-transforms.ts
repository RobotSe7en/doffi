import { DocumentError } from "../errors.js";

export const drawingColorTransformNames = ["lumMod", "lumOff", "satMod", "tint", "shade"] as const;
export const maxDrawingColorTransforms = 64;
type TransformName = typeof drawingColorTransformNames[number];

export function drawingColorTransformMaximum(name: TransformName): number {
  return name === "lumMod" || name === "satMod" ? 2147483647 : 100000;
}

export function parseDrawingColorTransforms(source: string): readonly (readonly [TransformName, number])[] {
  const value = source.trim();
  if (!value || value === "none") return [];
  const tokens = value.split(/\s+/);
  if (tokens.length > maxDrawingColorTransforms * 2)
    throw new DocumentError("ResourceLimitExceeded", "A DrawingML color exceeds its transform budget.");
  const transforms: [TransformName, number][] = [];
  for (let index = 0; index < tokens.length; index += 2) {
    const name = drawingColorTransformNames.find(candidate => candidate === tokens[index]);
    const token = tokens[index + 1], amount = Number(token);
    if (!name || !token || !/^\d+$/.test(token) || amount > drawingColorTransformMaximum(name))
      throw new DocumentError("InvalidDocument", "A DrawingML color transform is invalid.");
    transforms.push([name, amount]);
  }
  return transforms;
}

import { expect, test } from "vitest";
import {
  avoidWordDrawingOverlap,
  wordPolygonIntervals,
  wordWrapIntervals,
  wordWrapPoints,
} from "../src/word-wrap-geometry.js";
const notch = [
  { x: 0, y: 0 },
  { x: 100, y: 0 },
  { x: 100, y: 100 },
  { x: 75, y: 100 },
  { x: 75, y: 25 },
  { x: 25, y: 25 },
  { x: 25, y: 100 },
  { x: 0, y: 100 },
];
test("through permits internal openings while tight closes them over the entire glyph band", () => {
  expect(wordPolygonIntervals(notch, 40, 60, true)).toEqual([
    { left: 0, right: 25 },
    { left: 75, right: 100 },
  ]);
  expect(wordPolygonIntervals(notch, 40, 60, false)).toEqual([{ left: 0, right: 100 }]);
  expect(wordPolygonIntervals(notch, 20, 40, true)).toEqual([{ left: 0, right: 100 }]);
  expect(wordPolygonIntervals(notch, 110, 120, true)).toEqual([]);
  expect(wordPolygonIntervals([...notch, notch[0]!], 40, 60, true)).toEqual(
    wordPolygonIntervals(notch, 40, 60, true),
  );
});
test("diagonal edges project the full band and mirrored, rotated contours use native coordinates", () => {
  expect(
    wordPolygonIntervals(
      [
        { x: 50, y: 0 },
        { x: 100, y: 50 },
        { x: 50, y: 100 },
        { x: 0, y: 50 },
      ],
      10,
      20,
      true,
    ),
  ).toEqual([{ left: 30, right: 70 }]);
  const points = wordWrapPoints(
    [
      { x: 0, y: 0 },
      { x: 21600, y: 0 },
      { x: 0, y: 21600 },
    ],
    { left: 10, top: 20, width: 200, height: 100 },
    { rotation: 90, flipH: true, flipV: false },
  );
  expect(points[0]!.x).toBeCloseTo(160);
  expect(points[0]!.y).toBeCloseTo(170);
  expect(points[1]!.x).toBeCloseTo(160);
  expect(points[1]!.y).toBeCloseTo(-30);
  expect(points[2]!.x).toBeCloseTo(60);
  expect(points[2]!.y).toBeCloseTo(170);
});
test("wrap sides, tie direction, top/bottom and four asymmetric distances govern usable text intervals", () => {
  const bounds = { left: 40, right: 100, top: 10, bottom: 70 },
    line = { left: 0, right: 200, top: 20, bottom: 35 },
    paragraph = { left: 0, right: 200 };
  const layout = {
    mode: "floating" as const,
    wrap: "square" as const,
    distances: { left: 3, right: 6, top: 1.5, bottom: 3 },
  };
  expect(wordWrapIntervals(bounds, undefined, layout, line, paragraph, false)).toEqual([
    { left: 36, right: 108 },
  ]);
  expect(
    wordWrapIntervals(bounds, undefined, { ...layout, wrapSide: "left" }, line, paragraph, false),
  ).toEqual([{ left: 36, right: 200 }]);
  expect(
    wordWrapIntervals(bounds, undefined, { ...layout, wrapSide: "right" }, line, paragraph, false),
  ).toEqual([{ left: 0, right: 108 }]);
  expect(
    wordWrapIntervals(
      bounds,
      undefined,
      { ...layout, wrapSide: "largest" },
      line,
      paragraph,
      false,
    ),
  ).toEqual([{ left: 0, right: 108 }]);
  const centered = { ...bounds, left: 60, right: 140 };
  expect(
    wordWrapIntervals(
      centered,
      undefined,
      { mode: "floating", wrapSide: "largest" },
      line,
      paragraph,
      false,
    ),
  ).toEqual([{ left: 60, right: 200 }]);
  expect(
    wordWrapIntervals(
      centered,
      undefined,
      { mode: "floating", wrapSide: "largest" },
      line,
      paragraph,
      true,
    ),
  ).toEqual([{ left: 0, right: 140 }]);
  expect(
    wordWrapIntervals(
      bounds,
      undefined,
      { ...layout, wrap: "topAndBottom" },
      line,
      paragraph,
      false,
    ),
  ).toEqual([paragraph]);
  expect(
    wordWrapIntervals(bounds, undefined, { ...layout, wrap: "none" }, line, paragraph, false),
  ).toEqual([]);
  expect(
    wordWrapIntervals(bounds, undefined, { ...layout, behindText: true }, line, paragraph, false),
  ).toEqual([]);
  expect(
    wordWrapIntervals(
      bounds,
      undefined,
      layout,
      { ...line, top: 71, bottom: 74 },
      paragraph,
      false,
    ),
  ).toHaveLength(1);
});
test("collision resolution rechecks earlier objects after a later displacement and keeps allowed overlaps", () => {
  const box = { left: 20, right: 50, top: 0, bottom: 20 },
    placed = [
      { left: 0, right: 100, top: 40, bottom: 60, allowOverlap: false },
      { left: 0, right: 100, top: 0, bottom: 50, allowOverlap: false },
    ];
  expect(avoidWordDrawingOverlap(box, true, placed)).toBe(60);
  expect(avoidWordDrawingOverlap(box, true, [{ ...box, allowOverlap: true }])).toBe(0);
  expect(avoidWordDrawingOverlap(box, false, [{ ...box, allowOverlap: true }])).toBe(20);
  let work = 0;
  expect(() =>
    wordPolygonIntervals(notch, 0, 100, true, () => {
      if (++work > 3) throw new Error("budget");
    }),
  ).toThrow("budget");
});

test("nonfinite transformed polygon coordinates fail explicitly", () => {
  expect(() =>
    wordWrapPoints(
      [{ x: 1e308, y: 0 }],
      { left: 0, top: 0, width: 1e308, height: 1 },
      { rotation: 0, flipH: false, flipV: false },
    ),
  ).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
});

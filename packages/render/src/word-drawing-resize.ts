import { DocumentError, evaluatePresentationPreset } from "@doffi/core";
import type { DeepReadonly, WordDrawing } from "@doffi/core";
import type { RenderLimits } from "./contracts.js";
import { createDrawingGeometryView } from "./drawing-geometry.js";
import { renderWordChart } from "./word-chart.js";
import { ownViewAttribute, ownViewStyle } from "./view-nodes.js";

/** Recompose the derived paint at its measured size. Text insets, font sizes
 * and line widths keep their physical units; groups scale their child canvas.
 * This deliberately avoids scaling a bitmap of the complete object. */
export function createWordDrawingResize(
  document: Document,
  limits: Readonly<RenderLimits>,
  locale?: string,
) {
  let serial = 0;
  const clips = new WeakMap<HTMLElement, string>();
  return (
    frame: HTMLElement,
    item: DeepReadonly<WordDrawing>,
    g: WordDrawing["geometry"],
  ): void => {
    ownViewAttribute(frame, "data-sdk-word-width");
    ownViewAttribute(frame, "data-sdk-word-height");
    ownViewStyle(frame, "width", "height", "transform", "clip-path");
    frame.dataset.sdkWordWidth = String(g.width);
    frame.dataset.sdkWordHeight = String(g.height);
    frame.style.transform = `rotate(${g.rotation}deg) scale(${g.flipH ? -1 : 1},${g.flipV ? -1 : 1})`;
    if (
      Math.abs((parseFloat(frame.style.width) || 0) - g.width) < 0.0001 &&
      Math.abs((parseFloat(frame.style.height) || 0) - g.height) < 0.0001
    )
      return;
    frame.style.width = g.width + "pt";
    frame.style.height = g.height + "pt";
    if ((g.width === 0 || g.height === 0) && item.kind !== "connector") {
      if (!clips.has(frame)) clips.set(frame, frame.style.clipPath);
      frame.style.clipPath = "inset(100%)";
      return;
    }
    if (clips.has(frame)) {
      frame.style.clipPath = clips.get(frame)!;
      clips.delete(frame);
    }
    if (item.kind === "group") {
      const c = item.coordinateSpace,
        wrapper = frame.firstElementChild as HTMLElement;
      wrapper.style.transform = `scale(${g.width / c.width},${g.height / c.height}) translate(${-c.x}pt,${-c.y}pt)`;
      return;
    }
    for (const svg of [...frame.children].filter((child) => child.localName === "svg"))
      svg.remove();
    let count = document.querySelectorAll("*").length;
    const account = (amount: number): void => {
      if ((count += amount) > limits.maxNodes)
        throw new DocumentError(
          "ResourceLimitExceeded",
          "Relative drawing paint exceeds the render node budget.",
        );
    };
    const computed = document.defaultView!.getComputedStyle(frame),
      preset =
        item.kind === "connector"
          ? item.connector.routing === "elbow"
            ? "bentConnector3"
            : "straightConnector1"
          : item.kind === "shape"
            ? item.preset
            : "rect",
      geometry = evaluatePresentationPreset(
        preset,
        g.width,
        g.height,
        item.kind === "shape" ? item.adjustments : {},
      ),
      svg = createDrawingGeometryView(document, computed, {
        preset,
        geometry,
        widthPt: g.width,
        heightPt: g.height,
        id: "word-relative-" + ++serial,
        ...(item.kind === "image"
          ? { paint: "stroke" as const }
          : {
              fillFallback:
                item.kind === "chart"
                  ? "#ffffff"
                  : item.kind === "shape" && preset !== "line"
                    ? "#4472c4"
                    : "none",
              strokeFallback: item.kind === "connector" || preset === "line" ? "#224466" : "none",
            }),
        orientation: { rotationDegrees: g.rotation, flipH: g.flipH, flipV: g.flipV },
        maxNodes: limits.maxNodes - count,
      });
    account(svg.querySelectorAll("*").length + 1);
    const text = frame.querySelector<HTMLElement>(":scope > .sdk-word-drawing-text");
    if (text) frame.insertBefore(svg, text);
    else frame.append(svg);
    if (item.kind === "chart")
      frame.append(
        renderWordChart(
          frame,
          item.chart,
          g.width,
          g.height,
          "word-relative-chart-" + ++serial,
          account,
          locale,
        ),
      );
    if (text && (item.kind === "shape" || item.kind === "textBox")) {
      const rect = geometry.textRect ?? { x: 0, y: 0, width: g.width, height: g.height };
      Object.assign(text.style, {
        left: rect.x + "pt",
        top: rect.y + "pt",
        width: rect.width + "pt",
        height: rect.height + "pt",
      });
      if (item.text?.autoFit === "normal") {
        const base = parseFloat(text.dataset.sdkWordBaseFont ?? "16"),
          insets = item.text.insets ?? { left: 7.2, right: 7.2, top: 3.6, bottom: 3.6 },
          child = text.firstElementChild as HTMLElement;
        text.style.fontSize = base + "px";
        const scale =
          item.text.fontScale ??
          Math.min(
            1,
            Math.max(
              0.01,
              (rect.height - insets.top - insets.bottom) / 0.75 / Math.max(1, child.scrollHeight),
            ),
          );
        text.style.fontSize = base * scale + "px";
      }
    }
  };
}

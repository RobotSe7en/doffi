import { DocumentError } from "@doffi/core";
import type { DeepReadonly, PresentationChart, PresentationDataSet } from "@doffi/core";
import { markViewNode } from "../view-nodes.js";
import { createChartStyleProbe } from "../chart-style-probe.js";
import { renderBasicChart } from "../chart.js";
import type { BasicChartScene } from "../chart.js";
import type { DrawingGradientOrientation } from "../drawing-gradient.js";

export function formatChartNumber(value: number, format = "General"): string {
  if (format === "General" || !format) return String(Number(value.toPrecision(12)));
  const sections = format.split(";"), section = sections[value < 0 && sections.length > 1 ? 1 : value === 0 && sections.length > 2 ? 2 : 0]!;
  if (section === "") return "";
  if (!/^\(?[#,0]+(?:\.[0#]{1,20})?%?\)?$/.test(section)) throw new DocumentError("UnsupportedFeature", "This chart number format is outside the basic numeric format scope.");
  const pattern = section.replace(/"([^"]*)"/g, "$1").replace(/\\(.)/g, "$1");
  const match = /^(.*?)([#,0]+)(?:\.([0#]+))?(%?)(.*?)$/.exec(pattern);
  if (!match) throw new DocumentError("UnsupportedFeature", "This chart number format is not supported by the preview.");
  const magnitude = Math.abs(value) * (match[4] ? 100 : 1), decimals = match[3] ?? "";
  const number = magnitude.toLocaleString("en-US", { useGrouping: match[2]!.includes(","), minimumFractionDigits: decimals.replace(/#/g, "").length, maximumFractionDigits: decimals.length });
  return (value < 0 && sections.length === 1 ? "-" : "") + match[1] + number + match[4] + match[5];
}
/** PPT data sets contain bounded constants; workbook references have a separate
 * data adapter. Only the derived chart geometry and appearance are shared. */
export function renderChart(node: HTMLElement, chart: DeepReadonly<PresentationChart>, data: DeepReadonly<PresentationDataSet>, width: number, height: number, serial: string, account: (amount: number) => void, orientation?: DrawingGradientOrientation): BasicChartScene {
  const document = node.ownerDocument, probes: HTMLElement[] = [];
  const title = chart.titleNodeId ? node.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(chart.titleNodeId) + '"]') : undefined;
  try {
    let styleScope = node;
    if (chart.textCssClass) {
      account(1); styleScope = markViewNode(document.createElement("span")); styleScope.className = chart.textCssClass;
      styleScope.style.cssText = "position:absolute;visibility:hidden"; node.append(styleScope); probes.push(styleScope);
    }
    const styleProbe = createChartStyleProbe(styleScope, node, account, probes);
    return renderBasicChart(node, chart, data, width, height, serial, account, { formatNumber: formatChartNumber, ...(orientation ? { orientation } : {}),
      ...(title ? { title } : {}), style: styleProbe,
    });
  } finally { for (const probe of probes) probe.remove(); }
}

import type { DeepReadonly, WorkbookPrint, WorkbookPrintRun, WorkbookSheet } from "@doffi/core";
import type { FontResolution, WorkbookPrintedPage, WorkbookPrintSettings } from "../contracts.js";
import { resolvePreviewFontsInScope } from "../fonts.js";
import type { WorkbookSurface } from "./surface.js";

const textProperties = ["font-family", "font-size", "font-weight", "font-style", "color", "text-decoration-line",
  "text-decoration-color", "text-decoration-style", "vertical-align", "letter-spacing"] as const;
type Field = Extract<WorkbookPrintRun, { kind: "field" }>["field"];

export function workbookPrintFields(sheet: DeepReadonly<WorkbookSheet>, page: WorkbookPrintedPage, pages: number,
  settings: WorkbookPrintSettings, locale: string): Record<Field, string> {
  const clock = settings.context?.now;
  const local = clock ? new Date(clock.unixMilliseconds + clock.utcOffsetMinutes * 60_000) : undefined;
  const date = local && Number.isFinite(local.valueOf()) ? local : undefined;
  return {
    page: String(page.number), pages: String(pages), sheet: sheet.name,
    file: settings.fileName ?? "", path: settings.filePath ?? "",
    date: date ? new Intl.DateTimeFormat(locale, { timeZone: "UTC" }).format(date) : "",
    time: date ? new Intl.DateTimeFormat(locale, { timeZone: "UTC", hour: "numeric", minute: "2-digit" }).format(date) : "",
  };
}

export function paintWorkbookPrintBands(target: HTMLElement, surface: WorkbookSurface, page: WorkbookPrintedPage, sheetPageIndex: number,
  print: DeepReadonly<WorkbookPrint>, themeClass: string | undefined, fields: Record<Field, string>, account: (n: number) => void): readonly FontResolution[] {
  const document = surface.document, fonts: FontResolution[] = [], headerFooter = print.headerFooter;
  const alternate = headerFooter.differentFirst && sheetPageIndex === 0 ? "first"
    : headerFooter.differentOddEven && page.number % 2 === 0 ? "even" : "odd";
  const paper = page.setup;
  for (const edge of ["Header", "Footer"] as const) {
    const band = headerFooter[`${alternate}${edge}` as const];
    if (!band || !(band.left.length || band.center.length || band.right.length)) continue;
    const isHeader = edge === "Header", distance = isHeader ? print.marginsPt?.header ?? 21.6 : print.marginsPt?.footer ?? 21.6;
    const maxHeight = (isHeader ? paper.margins.top : paper.margins.bottom) - distance;
    if (maxHeight <= 0) continue;
    const left = headerFooter.alignWithMargins ? paper.margins.left : 18;
    const right = headerFooter.alignWithMargins ? paper.margins.right : 18;
    const scale = headerFooter.scaleWithDocument ? page.scale : 1;
    const strip = target.ownerDocument.createElement("div");
    strip.dataset.workbookPrintBand = edge.toLowerCase();
    strip.style.cssText = `position:absolute;left:${left}pt;width:${Math.max(0, paper.width - left - right) / scale}pt;max-height:${maxHeight / scale}pt;overflow:hidden;${isHeader ? "top" : "bottom"}:${distance}pt;transform:scale(${scale});transform-origin:${isHeader ? "top" : "bottom"} left;display:grid;grid-template-columns:repeat(3,minmax(0,1fr));align-items:${isHeader ? "start" : "end"};color:black;font:11pt Arial,sans-serif;direction:ltr`;
    const probe = document.createElement("div");
    probe.style.font = "11pt Arial, sans-serif";
    if (themeClass) probe.className = themeClass;
    // The probe resolves only text appearance in the scriptless source cascade.
    surface.source.append(probe);
    try {
      account(1);
      for (const side of ["left", "center", "right"] as const) {
        const section = target.ownerDocument.createElement("div");
        section.style.cssText = `text-align:${side};white-space:pre-wrap;overflow-wrap:anywhere`;
        account(1);
        for (const run of band[side]) {
          const text = run.kind === "text" ? run.text : fields[run.field];
          if (!text) continue;
          account(2);
          const source = document.createElement("span"), span = target.ownerDocument.createElement("span");
          if (run.cssClass) source.className = run.cssClass;
          source.textContent = text;
          probe.append(source);
          fonts.push(...resolvePreviewFontsInScope(document, probe, "span", surface.fontProbes));
          const computed = document.defaultView!.getComputedStyle(source);
          for (const key of textProperties) {
            if (key === "font-size" && !run.cssClass) continue;
            span.style.setProperty(key, computed.getPropertyValue(key));
          }
          span.textContent = text; section.append(span); source.remove();
        }
        strip.append(section);
      }
      target.append(strip);
    } finally { probe.remove(); }
  }
  return fonts;
}

import { needsWordTableLineExclusion } from "./word-center-wrap.js";
import {
  SourceStyleResolver,
  resolveSourceCssLength,
  DEFAULT_DOCUMENT_READ_LIMITS,
} from "@doffi/core";
import type { ContentDocumentModel, SourceStyleNode } from "@doffi/core";
import { DocumentError } from "@doffi/core";
import type { RenderLimits, RenderedPage } from "./contracts.js";
import { markViewNode, ownViewStyle } from "./view-nodes.js";
/** The actual source table remains editable. Only the float carrier and its
 * measured margins are view data; native coordinates stay in the source CSS. */
export function composeWordFloatingTables(
  document: Document,
  model: ContentDocumentModel,
  limits: Readonly<RenderLimits>,
): { readonly changed: boolean; measure(pages: readonly RenderedPage[], root: HTMLElement): void } {
  const snapshot = model.snapshot(),
    word = snapshot.metadata.word,
    entries: {
      table: HTMLTableElement;
      box: HTMLElement;
      anchor: HTMLElement;
      css: Record<string, string>;
      width: number;
    }[] = [],
    resolver = new SourceStyleResolver(
      snapshot.styles.map((sheet) => sheet.css),
      DEFAULT_DOCUMENT_READ_LIMITS,
    ),
    project = (id: string): SourceStyleNode => {
      const node = model.getNode(id)!;
      return {
        tag: node.tagName,
        attrs: node.attributes,
        parent: node.parentNodeId ? project(node.parentNodeId) : undefined,
      };
    };
  let nodeCount = document.querySelectorAll("*").length;
  const amount = (css: Record<string, string>, key: string): number =>
    resolveSourceCssLength(css["--word-table-position-" + key] ?? "0pt", {
      font: parseFloat(css["font-size"] ?? "12"),
      percent: 468,
    });
  for (const table of document.querySelectorAll<HTMLTableElement>("table[data-node-id]")) {
    const id = table.dataset.nodeId!;
    if (!word?.tables.some((info) => info.nodeId === id)) continue;
    const css = resolver.get(project(id));
    if (
      css["--word-table-floating"] !== "1" ||
      table.closest(
        '[data-office-story="textBox"],[data-office-story="comment"],[data-office-story="footnote"],[data-office-story="endnote"]',
      )
    )
      continue;
    const h = css["--word-table-position-h-anchor"] ?? "text",
      v = css["--word-table-position-v-anchor"] ?? "margin";
    if (
      !css["--word-table-position-x-align"] &&
      !css["--word-table-position-y-align"] &&
      amount(css, "x") === 0 &&
      amount(css, "y") === 0 &&
      h === "text" &&
      v !== "text"
    )
      continue;
    if ((nodeCount += 2) > limits.maxNodes)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "Floating tables exceed the render node budget.",
      );
    const width =
        table.getBoundingClientRect().width * 0.75 ||
        word.tables.find((info) => info.nodeId === id)!.grid.reduce((sum, width) => sum + width, 0),
      box = markViewNode(document.createElement("div"), "unwrap"),
      anchor = markViewNode(document.createElement("div"));
    box.className = "sdk-word-floating-table";
    anchor.className = "sdk-word-table-anchor";
    anchor.dataset.sdkWordTableAnchor = id;
    anchor.style.cssText = "height:0;padding:0;margin:0;border:0;pointer-events:none";
    table.before(anchor, box);
    box.append(table);
    box.style.cssText = "display:flow-root;box-sizing:content-box;float:left;margin:0;";
    if (
      needsWordTableLineExclusion(css) &&
      table.closest("[data-office-story]")?.getAttribute("data-office-story") === "main"
    ) {
      box.style.position = "absolute";
      box.style.cssFloat = "none";
    }
    const left = amount(css, "left-distance"),
      right = amount(css, "right-distance"),
      top = amount(css, "top-distance"),
      bottom = amount(css, "bottom-distance");
    box.style.width = width + "pt";
    box.style.padding = `${top}pt ${right}pt ${bottom}pt ${left}pt`;
    ownViewStyle(table, "width", "max-width", "margin", "float");
    table.style.width = width + "pt";
    table.style.maxWidth = "none";
    table.style.margin = "0";
    table.style.float = "none";
    entries.push({ table, box, anchor, css, width });
  }
  let changed = false;
  return {
    get changed(): boolean {
      return changed;
    },
    measure(pages: readonly RenderedPage[], root: HTMLElement): void {
      changed = false;
      const origin = root.getBoundingClientRect(),
        placed: {
          page: string;
          left: number;
          right: number;
          top: number;
          bottom: number;
          never: boolean;
        }[] = [];
      const align = (
        value: string | undefined,
        size: number,
        object: number,
        offset: number,
        odd: boolean,
      ): number =>
        value === "center"
          ? (size - object) / 2
          : value === "right" ||
              value === "bottom" ||
              (value === "outside" && odd) ||
              (value === "inside" && !odd)
            ? size - object
            : value && value !== "none" && value !== "inline"
              ? 0
              : offset;
      const instances = entries.flatMap((entry) => [
        entry,
        ...[
          ...document.querySelectorAll<HTMLTableElement>(
            '.sdk-paper .sdk-word-floating-table > table[data-node-id="' +
              CSS.escape(entry.table.dataset.nodeId!) +
              '"]',
          ),
        ].map((table) => ({
          ...entry,
          table,
          box: table.parentElement!,
          anchor: table
            .closest(".sdk-paper")!
            .querySelector<HTMLElement>(
              '[data-sdk-word-table-anchor="' + CSS.escape(entry.table.dataset.nodeId!) + '"]',
            )!,
        })),
      ]);
      for (const { table, box, anchor, css, width } of instances) {
        if (!anchor || !table.getClientRects().length) continue;
        const a = anchor.getBoundingClientRect(),
          y = (a.top - origin.top) * 0.75,
          paper = box.closest<HTMLElement>(".sdk-paper"),
          page = paper
            ? pages.find((page) => page.pageId === paper.dataset.pageId)
            : pages.find(
                (page) =>
                  !page.blank && y >= page.rect.y - 0.1 && y < page.rect.y + page.rect.height,
              );
        if (!page) continue;
        const setup = page.setup,
          m = setup.margins,
          odd = page.index % 2 === 0,
          section = snapshot.metadata.sections?.find(
            (section) => section.sectionId === page.sectionId,
          ),
          h = css["--word-table-position-h-anchor"] ?? "text",
          v = css["--word-table-position-v-anchor"] ?? "margin",
          x = amount(css, "x") > 0 ? amount(css, "x") - 0.05 : amount(css, "x"),
          offsetY = amount(css, "y") > 0 ? amount(css, "y") - 0.05 : amount(css, "y"),
          height = table.getBoundingClientRect().height * 0.75;
        let startX = page.rect.x + m.left,
          sizeX = setup.width - m.left - m.right;
        if (h === "page") {
          startX = page.rect.x;
          sizeX = setup.width;
        } else if (h === "text" && (section?.columns?.count ?? 1) > 1) {
          const count = section!.columns!.count,
            gap = section!.columns!.gap ?? 0,
            span = (sizeX - gap * (count - 1)) / count,
            index = Math.max(
              0,
              Math.min(
                count - 1,
                Math.floor(((a.left - origin.left) * 0.75 - startX) / (span + gap)),
              ),
            );
          startX += index * (span + gap);
          sizeX = section!.columns!.widths?.[index] ?? span;
        }
        let startY = v === "text" ? y : page.rect.y + (v === "page" ? 0 : m.top),
          sizeY =
            v === "page" ? setup.height : v === "margin" ? setup.height - m.top - m.bottom : 0;
        const cell = table.parentElement?.closest("td,th");
        if (cell) {
          const bounds = cell.getBoundingClientRect();
          startX = (bounds.left - origin.left) * 0.75;
          sizeX = bounds.width * 0.75;
          if (v !== "text") {
            startY = (bounds.top - origin.top) * 0.75;
            sizeY = bounds.height * 0.75;
          }
        }
        const left = startX + align(css["--word-table-position-x-align"], sizeX, width, x, odd);
        let top =
          startY +
          (v === "text"
            ? offsetY
            : align(css["--word-table-position-y-align"], sizeY, height, offsetY, odd));
        for (const before of placed)
          if (
            (css["--word-table-overlap"] === "never" || before.never) &&
            before.page === page.pageId &&
            left < before.right &&
            left + width > before.left &&
            top < before.bottom &&
            top + height > before.top
          )
            top = before.bottom + amount(css, "top-distance");
        placed.push({
          never: css["--word-table-overlap"] === "never",
          page: page.pageId,
          left,
          right: left + width,
          top,
          bottom: top + height + amount(css, "bottom-distance"),
        });
        if (paper) {
          const bounds = paper.getBoundingClientRect();
          box.style.position = "absolute";
          box.style.cssFloat = "none";
          box.style.margin = "0";
          box.style.left =
            origin.left + (left - amount(css, "left-distance")) / 0.75 - bounds.left + "px";
          box.style.top =
            origin.top + (top - amount(css, "top-distance")) / 0.75 - bounds.top + "px";
          paper.append(box);
          continue;
        }
        if (
          needsWordTableLineExclusion(css) &&
          table.closest("[data-office-story]")?.getAttribute("data-office-story") === "main"
        ) {
          const parent = (box.offsetParent as HTMLElement | null) ?? root,
            bounds = parent.getBoundingClientRect();
          const x =
              origin.left +
              (left - amount(css, "left-distance")) / 0.75 -
              bounds.left -
              parent.clientLeft,
            y =
              origin.top +
              (top - amount(css, "top-distance")) / 0.75 -
              bounds.top -
              parent.clientTop;
          for (const [key, value] of [
            ["left", x],
            ["top", y],
          ] as const)
            if (Math.abs((parseFloat(box.style.getPropertyValue(key)) || 0) - value) > 0.05) {
              box.style.setProperty(key, value + "px");
              changed = true;
            }
          continue;
        }
        const float = left + width / 2 > page.rect.x + setup.width / 2 ? "right" : "left",
          actual = table.getBoundingClientRect(),
          dx =
            float === "left"
              ? left - (actual.left - origin.left) * 0.75
              : (actual.right - origin.left) * 0.75 - left - width,
          dy = top - (actual.top - origin.top) * 0.75;
        if (box.style.cssFloat !== float) {
          box.style.cssFloat = float;
          box.style.marginLeft = box.style.marginRight = "0pt";
          changed = true;
          continue;
        }
        const nudge = (key: "margin-left" | "margin-right" | "margin-top", delta: number): void => {
          if (Math.abs(delta) < 0.05) return;
          const value =
            Math.round(((parseFloat(box.style.getPropertyValue(key)) || 0) + delta) * 1000) / 1000;
          box.style.setProperty(key, value + "pt");
          changed = true;
        };
        nudge(float === "left" ? "margin-left" : "margin-right", dx);
        nudge("margin-top", dy);
      }
    },
  };
}

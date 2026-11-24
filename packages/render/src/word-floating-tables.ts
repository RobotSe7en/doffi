import { floatingStory } from "./word-floating-placement.js";
import {
  columnRegion,
  floatingTableColumnOffset,
  floatingTableFragments,
} from "./column-regions.js";
import type { WordFloatingPlacement } from "./word-floating-placement.js";
import { needsWordTableLineExclusion } from "./word-center-wrap.js";
import {
  SourceStyleResolver,
  resolveSourceCssLength,
  DEFAULT_DOCUMENT_READ_LIMITS,
} from "@doffi/core";
import type { ContentDocumentModel, SourceStyleNode } from "@doffi/core";
import { DocumentError } from "@doffi/core";
import type { LayoutDiagnostic, RenderLimits, RenderedPage } from "./contracts.js";
import { markViewNode, ownViewStyle } from "./view-nodes.js";
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

/** The actual source table remains editable. Only the float carrier and its
 * measured margins are view data; native coordinates stay in the source CSS. */
export function composeWordFloatingTables(
  document: Document,
  model: ContentDocumentModel,
  limits: Readonly<RenderLimits>,
): {
  readonly active: boolean;
  readonly changed: boolean;
  readonly diagnostics: readonly LayoutDiagnostic[];
  measure(
    pages: readonly RenderedPage[],
    root: HTMLElement,
    placement: WordFloatingPlacement,
    storiesOnly?: boolean,
  ): void;
} {
  const snapshot = model.snapshot(),
    word = snapshot.metadata.word,
    continued = new WeakSet<HTMLElement>(),
    entries: {
      table: HTMLTableElement;
      box: HTMLElement;
      anchor: HTMLElement;
      css: Record<string, string>;
      width: number;
    }[] = [],
    project = (id: string): SourceStyleNode => {
      const node = model.getNode(id)!;
      return {
        tag: node.tagName,
        attrs: node.attributes,
        parent: node.parentNodeId ? project(node.parentNodeId) : undefined,
      };
    };
  let resolver: SourceStyleResolver | undefined;
  let nodeCount = document.querySelectorAll("*").length;
  const amount = (css: Record<string, string>, key: string): number =>
    resolveSourceCssLength(css["--word-table-position-" + key] ?? "0pt", {
      font: parseFloat(css["font-size"] ?? "12"),
      percent: 468,
    });
  for (const table of document.querySelectorAll<HTMLTableElement>("table[data-node-id]")) {
    const id = table.dataset.nodeId!;
    if (!word?.tables.some((info) => info.nodeId === id)) continue;
    resolver ??= new SourceStyleResolver(snapshot.styles.map((sheet) => sheet.css), DEFAULT_DOCUMENT_READ_LIMITS);
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
    if (h === "text" && !anchor.closest("td,th"))
      floatingTableColumnOffset.set(box, (span, page) =>
        align(
          css["--word-table-position-x-align"],
          span,
          width / 0.75,
          (amount(css, "x") - (amount(css, "x") > 0 ? 0.05 : 0)) / 0.75,
          page % 2 === 0,
        ),
      );
    entries.push({ table, box, anchor, css, width });
  }
  let changed = false;
  let diagnostics: LayoutDiagnostic[] = [],
    storyDiagnostics: LayoutDiagnostic[] = [];
  return {
    get active() {
      return entries.length > 0;
    },
    get diagnostics() {
      return [...diagnostics, ...storyDiagnostics];
    },
    get changed(): boolean {
      return changed;
    },
    measure(pages, root, placement, storiesOnly = false): void {
      if (!storiesOnly) {
        changed = false;
        diagnostics = [];
      }
      storyDiagnostics = [];
      const origin = root.getBoundingClientRect(),
        byPage = new Map(pages.map((page) => [page.pageId, page])),
        pageAt = (y: number) =>
          pages.find(
            (page) => !page.blank && y >= page.rect.y - 0.1 && y < page.rect.y + page.rect.height,
          );
      const entriesById = new Map(entries.map((entry) => [entry.table.dataset.nodeId!, entry])),
        instances = storiesOnly ? [] : [...entries];
      for (const anchor of document.querySelectorAll<HTMLElement>(
        ".sdk-paper [data-sdk-word-table-anchor], .sdk-repeat-header [data-sdk-word-table-anchor]",
      )) {
        const entry = entriesById.get(anchor.dataset.sdkWordTableAnchor!);
        if (!entry) continue;
        const paper = anchor.closest(".sdk-paper");
        if (storiesOnly && !paper) continue;
        const candidate = anchor.nextElementSibling,
          table = (
            candidate?.classList.contains("sdk-word-floating-table") ? candidate : paper
          )?.querySelector<HTMLTableElement>(
            '.sdk-word-floating-table > table[data-node-id="' +
              CSS.escape(entry.table.dataset.nodeId!) +
              '"]',
          );
        if (table) instances.push({ ...entry, table, box: table.parentElement!, anchor });
      }
      for (const { table, box, anchor, css, width } of instances) {
        if (!anchor || !table.getClientRects().length) continue;
        const a = anchor.getBoundingClientRect(),
          y = (a.top - origin.top) * 0.75,
          paper = box.closest<HTMLElement>(".sdk-paper"),
          page = paper ? byPage.get(paper.dataset.pageId!) : pageAt(y);
        if (!page) continue;
        const setup = page.setup,
          m = setup.margins,
          odd = page.index % 2 === 0,
          h = css["--word-table-position-h-anchor"] ?? "text",
          v = css["--word-table-position-v-anchor"] ?? "margin",
          x = amount(css, "x") > 0 ? amount(css, "x") - 0.05 : amount(css, "x"),
          offsetY = amount(css, "y") > 0 ? amount(css, "y") - 0.05 : amount(css, "y"),
          height = table.getBoundingClientRect().height * 0.75;
        let startX = page.rect.x + m.left,
          sizeX = setup.width - m.left - m.right;
        const cell = anchor.closest<HTMLElement>("td,th"),
          story = floatingStory(anchor),
          column = !story ? columnRegion(anchor, a, cell ?? undefined) : undefined;
        if (h === "page") {
          startX = page.rect.x;
          sizeX = setup.width;
        } else if (h === "text") {
          const region = story?.getBoundingClientRect() ?? column;
          if (region) {
            startX = (region.left - origin.left) * 0.75;
            sizeX = (region.right - region.left) * 0.75;
          }
        }
        let startY = v === "text" ? y : page.rect.y + (v === "page" ? 0 : m.top),
          sizeY =
            v === "page" ? setup.height : v === "margin" ? setup.height - m.top - m.bottom : 0;
        if (cell) {
          const bounds = column?.fragment ?? cell.getBoundingClientRect(),
            style = document.defaultView!.getComputedStyle(cell),
            inset = (side: string) =>
              (parseFloat(style.getPropertyValue("padding-" + side)) || 0) +
              (parseFloat(style.getPropertyValue("border-" + side + "-width")) || 0),
            left = bounds.left + inset("left"),
            right = bounds.right - inset("right");
          startX = (left - origin.left) * 0.75;
          sizeX = (right - left) * 0.75;
          if (v !== "text") {
            startY = (bounds.top + inset("top") - origin.top) * 0.75;
            sizeY = (bounds.bottom - inset("bottom") - origin.top) * 0.75 - startY;
          }
        }
        const left = startX + align(css["--word-table-position-x-align"], sizeX, width, x, odd);
        const desiredTop =
          startY +
          (v === "text"
            ? offsetY
            : align(css["--word-table-position-y-align"], sizeY, height, offsetY, odd));
        const physicalBottom = origin.top + (page.rect.y + page.rect.height) / 0.75,
          columnBottom = column?.bottom ?? physicalBottom - m.bottom / 0.75;
        if (
          !paper &&
          box.style.position === "absolute" &&
          box.closest(".sdk-page-flow") &&
          height > 0 &&
          origin.top + (desiredTop + height + amount(css, "bottom-distance")) / 0.75 >
            columnBottom + 0.5 &&
          (v === "text" || height > setup.height - m.top - m.bottom)
        ) {
          // Keep the real rows in the measured flow when the positioned table
          // needs the table paginator's continuation handling.
          box.style.position = "";
          box.style.cssFloat = "none";
          continued.add(box);
          box.style.left = box.style.top = "";
          changed = true;
          continue;
        }
        const scope = cell ?? story ?? root,
          allowOverlap = css["--word-table-overlap"] !== "never",
          fragments =
            column && box.style.position !== "absolute"
              ? [...floatingTableFragments(table).values()]
              : [],
          [first, ...continuations] = fragments;
        if (
          !paper &&
          !continued.has(box) &&
          box.style.position !== "absolute" &&
          (continuations.some(({ top }) => pageAt((top - origin.top) * 0.75) !== page) ||
            (!column && height > setup.height - m.top - m.bottom))
        ) {
          // The default native behavior puts following content after the
          // last row when a floating table breaks over a physical page.
          continued.add(box);
          box.style.cssFloat = "none";
          box.style.marginLeft = box.style.marginRight = "0pt";
          changed = true;
          continue;
        }
        for (const bounds of continuations) {
          const continuation = pageAt((bounds.top - origin.top) * 0.75);
          if (continuation)
            placement.add({ anchor, page: continuation, scope, allowOverlap, bounds });
        }
        const collisionHeight = first ? (first.bottom - first.top) * 0.75 : height;
        placement.add({
          anchor,
          page,
          scope,
          bounds: {
            left: origin.left + (left - amount(css, "left-distance")) / 0.75,
            right: origin.left + (left + width + amount(css, "right-distance")) / 0.75,
            top: origin.top + (desiredTop - amount(css, "top-distance")) / 0.75,
            bottom:
              origin.top + (desiredTop + collisionHeight + amount(css, "bottom-distance")) / 0.75,
          },
          allowOverlap,
          apply(delta) {
            const top = desiredTop + delta * 0.75;
            if (delta > 0.1 && origin.top + (top + collisionHeight) / 0.75 > physicalBottom + 0.1)
              (paper ? storyDiagnostics : diagnostics).push({
                code: "ContentOverflow",
                nodeId: table.dataset.nodeId!,
                message: "Collision avoidance moves this table beyond its anchor page.",
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
              return;
            }
            if (box.style.position === "absolute") {
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
              return;
            }
            const columnAligned = column && floatingTableColumnOffset.has(box);
            const center = column
              ? ((column.left + column.right) / 2 - origin.left) * 0.75
              : page.rect.x + setup.width / 2;
            const float = continued.has(box)
                ? "none"
                : !columnAligned && left + width / 2 > center
                  ? "right"
                  : "left",
              actual = table.rows[0]?.getBoundingClientRect() ?? table.getBoundingClientRect(),
              dx =
                float !== "right"
                  ? left - (actual.left - origin.left) * 0.75
                  : (actual.right - origin.left) * 0.75 - left - width,
              dy =
                column && (actual.top < column.top - 0.5 || actual.top >= column.bottom)
                  ? 0
                  : top - (actual.top - origin.top) * 0.75;
            if (box.style.cssFloat !== float) {
              box.style.cssFloat = float;
              box.style.marginLeft = box.style.marginRight = "0pt";
              changed = true;
              return;
            }
            const nudge = (
              key: "margin-left" | "margin-right" | "margin-top",
              delta: number,
            ): void => {
              if (Math.abs(delta) < 0.05) return;
              const value =
                Math.round(((parseFloat(box.style.getPropertyValue(key)) || 0) + delta) * 1000) /
                1000;
              box.style.setProperty(key, value + "pt");
              changed = true;
            };
            if (!columnAligned) nudge(float === "right" ? "margin-right" : "margin-left", dx);
            nudge("margin-top", dy);
          },
        });
      }
    },
  };
}

import { DocumentError } from "@doffi/core";
import type { DocumentEditingSession, DocumentOperation } from "@doffi/core";
import { addWordPageBorderControls } from "./word-page-border-controls.js";
import { addWordSectionNoteControls } from "./word-page-note-controls.js";
/** Page forms edit native section definitions; auxiliary stories use the same
 * rich-text adapter over a temporary view projection of the original session. */
export function createWordPageControls(
  parent: ShadowRoot,
  before: HTMLElement,
  session: DocumentEditingSession,
  sectionRoot: () => string,
  editStory: (storyId?: string) => void,
  onError: (error: unknown) => void,
): { open(): void; dispose(): void } {
  const document = before.ownerDocument;
  let panel: HTMLElement | undefined, controller: AbortController | undefined;
  const close = (): void => {
    controller?.abort();
    panel?.remove();
    controller = undefined;
    panel = undefined;
  };
  function open(): void {
    close();
    controller = new AbortController();
    const signal = controller.signal;
    panel = document.createElement("section");
    panel.setAttribute("role", "region");
    panel.setAttribute("aria-label", "页面设置");
    panel.style.cssText =
      "padding:12px;background:white;border:1px solid #dbe3ef;max-height:300px;overflow:auto;display:flex;flex-wrap:wrap;align-items:center;gap:8px";
    parent.insertBefore(panel, before);
    const metadata = session.snapshot().metadata;
    if (metadata.presentation)
      throw new DocumentError("InvalidOperation", "页面设置需要 Word 文档。");
    const button = (name: string, run: () => void): void => {
      const node = document.createElement("button");
      node.type = "button";
      node.textContent = name;
      node.addEventListener(
        "click",
        () => {
          try {
            run();
          } catch (error) {
            onError(error);
          }
        },
        { signal },
      );
      panel!.append(node);
    };
    const field = (name: string, value: number | string): HTMLInputElement => {
      const label = document.createElement("label"),
        input = document.createElement("input");
      label.textContent = name;
      input.type = "number";
      input.step = "0.1";
      input.min = "0";
      input.style.width = "80px";
      input.value = String(value);
      input.setAttribute("aria-label", name);
      label.append(input);
      panel!.append(label);
      return input;
    };
    const choose = (
      name: string,
      options: readonly (readonly [string, string])[],
      value: string,
    ): HTMLSelectElement => {
      const label = document.createElement("label"),
        input = document.createElement("select");
      label.textContent = name;
      input.setAttribute("aria-label", name);
      for (const [key, text] of options) {
        const node = document.createElement("option");
        node.value = key;
        node.textContent = text;
        input.append(node);
      }
      input.value = value;
      label.append(input);
      panel!.append(label);
      return input;
    };
    const checkbox = (name: string, value: boolean): HTMLInputElement => {
      const label = document.createElement("label"),
        input = document.createElement("input");
      label.textContent = name;
      input.type = "checkbox";
      input.checked = value;
      input.setAttribute("aria-label", name);
      label.append(input);
      panel!.append(label);
      return input;
    };
    const mirror = checkbox("镜像页边距", metadata.word?.mirrorMargins ?? false);
    button("应用镜像页边距", () => {
      const root = new DOMParser()
        .parseFromString(session.snapshot().html, "text/html")
        .querySelector<HTMLElement>('[data-office-kind="document"]')!;
      apply([
        { op: "setWordOptions", nodeId: root.dataset.nodeId!, mirrorMargins: mirror.checked },
      ]);
    });
    button("关闭页面设置", close);
    button("返回正文", () => editStory());
    const section = choose(
      "当前节",
      metadata.sections.map(
        (section, index) => [section.rootNodeId, "第 " + (index + 1) + " 节"] as const,
      ),
      sectionRoot(),
    );
    const selected = () => {
      const metadata = session.snapshot().metadata;
      if (metadata.presentation) throw new DocumentError("InvalidOperation", "需要 Word 文档。");
      const target = metadata.sections.find((target) => target.rootNodeId === section.value);
      if (!target) throw new DocumentError("InvalidOperation", "当前节已不存在。");
      return { metadata, target };
    };
    const sectionStart = choose(
      "节起始位置",
      [
        ["continuous", "连续"],
        ["nextPage", "下一页"],
        ["nextColumn", "下一栏"],
        ["evenPage", "偶数页"],
        ["oddPage", "奇数页"],
      ],
      selected().target.breakType ?? "nextPage",
    );
    button("应用节起始位置", () =>
      apply([
        {
          op: "setWordSection",
          nodeId: selected().target.rootNodeId,
          changes: {
            breakType: sectionStart.value as
              | "continuous"
              | "nextPage"
              | "nextColumn"
              | "evenPage"
              | "oddPage",
          },
        },
      ]),
    );
    const setup = selected().target.pageSetup ?? metadata.pageSetup,
      mm = (pt: number) => Math.round(((pt * 25.4) / 72) * 100) / 100,
      pt = (mm: string) => (Number(mm) * 72) / 25.4;
    const width = field("纸张宽度（毫米）", mm(setup.width)),
      height = field("纸张高度（毫米）", mm(setup.height));
    const marginTop = field("上边距（毫米）", mm(setup.margins.top)),
      marginBottom = field("下边距（毫米）", mm(setup.margins.bottom)),
      marginLeft = field("左边距（毫米）", mm(setup.margins.left)),
      marginRight = field("右边距（毫米）", mm(setup.margins.right));
    const paper = choose(
      "纸张",
      [
        ["custom", "自定义"],
        ["a4", "A4"],
        ["a3", "A3"],
        ["letter", "Letter"],
      ],
      "custom",
    );
    paper.addEventListener(
      "change",
      () => {
        const sizes: Record<string, [number, number]> = {
          a4: [210, 297],
          a3: [297, 420],
          letter: [215.9, 279.4],
        };
        const size = sizes[paper.value];
        if (size) {
          width.value = String(size[0]);
          height.value = String(size[1]);
        }
      },
      { signal },
    );
    button("切换横纵方向", () => {
      const previous = width.value;
      width.value = height.value;
      height.value = previous;
    });
    const scope = choose(
      "应用范围",
      [
        ["section", "本节"],
        ["all", "全文"],
      ],
      "section",
    );
    const apply = (operations: DocumentOperation[]): void => {
      const snapshot = session.snapshot();
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: snapshot.manifest.documentId,
          baseRevision: snapshot.manifest.revision,
          operations,
        }),
      );
    };
    const loadBorders = addWordPageBorderControls(
      document,
      panel,
      session,
      selected,
      () => scope.value === "all",
      apply,
      onError,
      signal,
    );
    section.addEventListener("change", loadBorders, { signal });
    const loadNotes = addWordSectionNoteControls(
      document,
      panel,
      session,
      selected,
      () => scope.value === "all",
      apply,
      onError,
      signal,
    );
    section.addEventListener("change", loadNotes, { signal });
    button("应用页面设置", () => {
      const latest = selected(),
        targets = scope.value === "all" ? latest.metadata.sections : [latest.target];
      apply(
        targets.map((target) => ({
          op: "setWordSection",
          nodeId: target.rootNodeId,
          changes: {
            pageSetup: {
              unit: "pt",
              width: pt(width.value),
              height: pt(height.value),
              margins: {
                top: pt(marginTop.value),
                bottom: pt(marginBottom.value),
                left: pt(marginLeft.value),
                right: pt(marginRight.value),
              },
            },
          },
        })),
      );
    });
    const gutter = field("装订线（毫米）", mm(selected().target.gutter ?? 0)),
      gutterPosition = choose(
        "装订线位置",
        [
          ["left", "左"],
          ["right", "右"],
          ["top", "上"],
        ],
        selected().target.gutterPosition ?? "left",
      );
    button("应用装订线", () => {
      const latest = selected(),
        targets =
          gutterPosition.value === "top" || scope.value === "all"
            ? latest.metadata.sections
            : [latest.target];
      apply(
        targets.map((target) => ({
          op: "setWordSection",
          nodeId: target.rootNodeId,
          changes: {
            gutter: pt(gutter.value),
            gutterPosition: gutterPosition.value as "left" | "right" | "top",
          },
        })),
      );
    });
    const gridType = choose(
        "文档网格",
        [
          ["default", "无网格"],
          ["lines", "只指定行网格"],
          ["linesAndChars", "行和字符网格"],
          ["snapToChars", "字符网格"],
        ],
        selected().target.grid?.type ?? "default",
      ),
      linePitch = field("网格行距（磅）", selected().target.grid?.linePitch ?? 24),
      characterSpacing = field(
        "字符间距增量（磅）",
        (selected().target.grid?.characterSpacing ?? 0) / 4096,
      );
    characterSpacing.min = "-100";
    characterSpacing.step = "0.25";
    button("应用文档网格", () => {
      const latest = selected(),
        targets = scope.value === "all" ? latest.metadata.sections : [latest.target];
      apply(
        targets.map((target) => ({
          op: "setWordSection",
          nodeId: target.rootNodeId,
          changes: {
            grid: {
              type: gridType.value as "default" | "lines" | "linesAndChars" | "snapToChars",
              linePitch: Number(linePitch.value),
              characterSpacing: Math.round(Number(characterSpacing.value) * 4096),
            },
          },
        })),
      );
    });
    const tableGrid = checkbox(
      "表格文字对齐行网格",
      metadata.word?.adjustLineHeightInTable ?? false,
    );
    button("应用表格行网格", () => {
      const root = new DOMParser()
        .parseFromString(session.snapshot().html, "text/html")
        .querySelector<HTMLElement>('[data-office-kind="document"]')!;
      apply([
        {
          op: "setWordOptions",
          nodeId: root.dataset.nodeId!,
          adjustLineHeightInTable: tableGrid.checked,
        },
      ]);
    });
    const columnCount = field("栏数", selected().target.columns?.count ?? 1),
      columnGap = field("栏间距（毫米）", mm(selected().target.columns?.gap ?? 36)),
      columnSeparator = checkbox("栏间分隔线", selected().target.columns?.separator ?? false);
    columnCount.step = "1";
    columnCount.min = "1";
    const equalColumns = checkbox("等宽分栏", !selected().target.columns?.widths),
      columnWidths = field(
        "各栏宽度（磅，以逗号分隔）",
        selected().target.columns?.widths?.join(",") ?? "",
      ),
      columnSpaces = field(
        "各栏后间距（磅，以逗号分隔）",
        selected().target.columns?.spaces?.join(",") ?? "",
      );
    columnWidths.type = "text";
    columnSpaces.type = "text";
    columnWidths.style.width = columnSpaces.style.width = "180px";
    button("应用分栏", () => {
      const latest = selected(),
        targets = scope.value === "all" ? latest.metadata.sections : [latest.target],
        widths = columnWidths.value.split(",").map((value) => Number(value.trim())),
        spaces = columnSpaces.value.trim()
          ? columnSpaces.value.split(",").map((value) => Number(value.trim()))
          : widths.map(() => pt(columnGap.value));
      apply(
        targets.map((target) => ({
          op: "setWordSection",
          nodeId: target.rootNodeId,
          changes: {
            columns: {
              count: equalColumns.checked ? Number(columnCount.value) : widths.length,
              gap: pt(columnGap.value),
              separator: columnSeparator.checked,
              ...(equalColumns.checked ? {} : { widths, spaces }),
            },
          },
        })),
      );
    });
    const balanceColumns = checkbox("节末分栏自动平衡", !metadata.word?.noColumnBalance);
    button("应用分栏平衡", () => {
      const root = new DOMParser()
        .parseFromString(session.snapshot().html, "text/html")
        .querySelector<HTMLElement>('[data-office-kind="document"]')!;
      apply([
        {
          op: "setWordOptions",
          nodeId: root.dataset.nodeId!,
          noColumnBalance: !balanceColumns.checked,
        },
      ]);
    });

    const variant = choose(
      "页眉页脚类型",
      [
        ["default", "默认"],
        ["first", "首页"],
        ["even", "偶数页"],
      ],
      "default",
    );
    const effective = (kind: "header" | "footer"): string | undefined => {
      const latest = selected(),
        key = kind === "header" ? "headers" : "footers",
        index = latest.metadata.sections.indexOf(latest.target);
      for (let n = index; n >= 0; n--) {
        const id =
          latest.metadata.sections[n]![key]?.[variant.value as "default" | "first" | "even"];
        if (id) return id;
      }
      return;
    };
    const auxiliary = (kind: "header" | "footer"): void => {
      const latest = selected(),
        key = kind === "header" ? "headers" : "footers";
      let storyId = effective(kind);
      if (!storyId) {
        const index = latest.metadata.sections.indexOf(latest.target);
        apply([
          {
            op: "insertWordStory",
            nodeId: latest.target.rootNodeId,
            kind,
            variant: variant.value as "default" | "first" | "even",
            html: "<p></p>",
          },
        ]);
        const next = session.snapshot().metadata;
        if (!next.presentation) {
          section.value = next.sections[index]!.rootNodeId;
          storyId = next.sections[index]![key]?.[variant.value as "default" | "first" | "even"];
        }
      }
      if (storyId) editStory(storyId);
    };
    button("编辑页眉", () => auxiliary("header"));
    button("编辑页脚", () => auxiliary("footer"));
    const linkKind = choose(
      "链接对象",
      [
        ["header", "页眉"],
        ["footer", "页脚"],
      ],
      "header",
    );
    button("链接到前一节", () => {
      const latest = selected(),
        key = linkKind.value === "header" ? "headers" : "footers";
      if (latest.metadata.sections.indexOf(latest.target) === 0)
        throw new DocumentError("InvalidOperation", "第一节没有前一节。");
      const refs: {
        default?: string | undefined;
        first?: string | undefined;
        even?: string | undefined;
      } = { ...latest.target[key] };
      delete refs[variant.value as "default" | "first" | "even"];
      apply([{ op: "setWordSection", nodeId: latest.target.rootNodeId, changes: { [key]: refs } }]);
    });
    button("取消链接并复制", () => {
      const latest = selected(),
        kind = linkKind.value as "header" | "footer",
        storyId = effective(kind);
      apply([
        storyId
          ? {
              op: "cloneWordStory",
              nodeId: latest.target.rootNodeId,
              storyId,
              variant: variant.value as "default" | "first" | "even",
            }
          : {
              op: "insertWordStory",
              nodeId: latest.target.rootNodeId,
              kind,
              variant: variant.value as "default" | "first" | "even",
              html: "<p></p>",
            },
      ]);
    });
    const firstPage = checkbox("首页不同", selected().target.differentFirstPage ?? false),
      evenPage = checkbox("奇偶页不同", metadata.word?.evenAndOddHeaders ?? false);
    button("应用页眉页脚类型", () => {
      const root = new DOMParser()
        .parseFromString(session.snapshot().html, "text/html")
        .querySelector<HTMLElement>('[data-office-kind="document"]')!;
      apply([
        {
          op: "setWordSection",
          nodeId: selected().target.rootNodeId,
          changes: { differentFirstPage: firstPage.checked },
        },
        { op: "setWordOptions", nodeId: root.dataset.nodeId!, evenAndOddHeaders: evenPage.checked },
      ]);
    });
    const headerDistance = field("页眉距离（毫米）", mm(selected().target.headerDistance ?? 36)),
      footerDistance = field("页脚距离（毫米）", mm(selected().target.footerDistance ?? 36));
    button("应用页眉页脚距离", () =>
      apply([
        {
          op: "setWordSection",
          nodeId: selected().target.rootNodeId,
          changes: {
            headerDistance: pt(headerDistance.value),
            footerDistance: pt(footerDistance.value),
          },
        },
      ]),
    );
    const numberFormat = choose(
        "页码格式",
        [
          ["decimal", "1, 2, 3"],
          ["upperRoman", "I, II, III"],
          ["lowerRoman", "i, ii, iii"],
          ["upperLetter", "A, B, C"],
          ["lowerLetter", "a, b, c"],
        ],
        selected().target.pageNumberFormat ?? "decimal",
      ),
      start = field("起始页码", selected().target.pageNumberStart ?? 1);
    button("设置页码", () =>
      apply([
        {
          op: "setWordSection",
          nodeId: selected().target.rootNodeId,
          changes: {
            pageNumberStart: Number(start.value),
            pageNumberFormat: numberFormat.value as
              | "decimal"
              | "upperRoman"
              | "lowerRoman"
              | "upperLetter"
              | "lowerLetter",
          },
        },
      ]),
    );
    const load = (): void => {
      const setup = selected().target.pageSetup ?? selected().metadata.pageSetup;
      width.value = String(mm(setup.width));
      height.value = String(mm(setup.height));
      marginTop.value = String(mm(setup.margins.top));
      marginBottom.value = String(mm(setup.margins.bottom));
      marginLeft.value = String(mm(setup.margins.left));
      marginRight.value = String(mm(setup.margins.right));
      const target = selected().target;
      sectionStart.value = target.breakType ?? "nextPage";
      gutter.value = String(mm(target.gutter ?? 0));
      gutterPosition.value = target.gutterPosition ?? "left";
      gridType.value = target.grid?.type ?? "default";
      linePitch.value = String(target.grid?.linePitch ?? 24);
      characterSpacing.value = String((target.grid?.characterSpacing ?? 0) / 4096);
      columnCount.value = String(target.columns?.count ?? 1);
      columnGap.value = String(mm(target.columns?.gap ?? 36));
      columnSeparator.checked = target.columns?.separator ?? false;
      equalColumns.checked = !target.columns?.widths;
      columnWidths.value = target.columns?.widths?.join(",") ?? "";
      columnSpaces.value = target.columns?.spaces?.join(",") ?? "";
      firstPage.checked = target.differentFirstPage ?? false;
      headerDistance.value = String(mm(target.headerDistance ?? 36));
      footerDistance.value = String(mm(target.footerDistance ?? 36));
      start.value = String(target.pageNumberStart ?? 1);
      numberFormat.value = target.pageNumberFormat ?? "decimal";
    };
    section.addEventListener("change", load, { signal });
  }
  return { open, dispose: close };
}

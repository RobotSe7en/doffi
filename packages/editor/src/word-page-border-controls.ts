import type {
  DocumentEditingSession,
  DocumentOperation,
  DocumentSection,
  DeepReadonly,
  WordPageBorders,
} from "@doffi/core";
export function addWordPageBorderControls(
  document: Document,
  panel: HTMLElement,
  session: DocumentEditingSession,
  selected: () => {
    target: DeepReadonly<DocumentSection>;
    metadata: {
      sections: readonly DeepReadonly<DocumentSection>[];
    };
  },
  all: () => boolean,
  apply: (operations: DocumentOperation[]) => void,
  onError: (error: unknown) => void,
  signal: AbortSignal,
): () => void {
  const select = (
    name: string,
    options: readonly (readonly [string, string])[],
    initial: string,
  ): HTMLSelectElement => {
    const label = document.createElement("label"),
      input = document.createElement("select");
    label.textContent = name;
    input.setAttribute("aria-label", name);
    for (const [value, text] of options) {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = text;
      input.append(option);
    }
    input.value = initial;
    label.append(input);
    panel.append(label);
    return input;
  };
  const field = (name: string, type: string, value: string): HTMLInputElement => {
    const label = document.createElement("label"),
      input = document.createElement("input");
    label.textContent = name;
    input.type = type;
    input.setAttribute("aria-label", name);
    input.value = value;
    input.style.width = "90px";
    label.append(input);
    panel.append(label);
    return input;
  };
  const borders = selected().target.pageBorders,
    pattern = select(
      "页面边框线型",
      [
        ["none", "无"],
        ["single", "实线"],
        ["double", "双线"],
        ["dotted", "点线"],
        ["dashed", "虚线"],
        ["dashSmallGap", "短虚线"],
        ["thick", "粗实线"],
      ],
      borders?.top?.style ?? "none",
    ),
    color = field(
      "页面边框颜色",
      "color",
      "#" + (borders?.top?.color === "auto" ? "000000" : (borders?.top?.color ?? "000000")),
    ),
    width = field("页面边框宽度（磅）", "number", String(borders?.top?.width ?? 0.5)),
    space = field("页面边框距离（磅）", "number", String(borders?.top?.space ?? 24));
  width.min = "0";
  width.max = "12";
  width.step = ".125";
  space.min = "0";
  space.max = "31";
  space.step = "1";
  const offset = select(
      "页面边框相对于",
      [
        ["text", "正文"],
        ["page", "页面"],
      ],
      borders?.offsetFrom ?? "text",
    ),
    display = select(
      "页面边框页面",
      [
        ["allPages", "所有页面"],
        ["firstPage", "仅本节首页"],
        ["notFirstPage", "除本节首页"],
      ],
      borders?.display ?? "allPages",
    ),
    order = select(
      "页面边框层级",
      [
        ["front", "文字前方"],
        ["back", "文字后方"],
      ],
      borders?.zOrder ?? "front",
    );
  const flag = (name: string, value: boolean): HTMLInputElement => {
      const input = field(name, "checkbox", "");
      input.checked = value;
      return input;
    },
    shadow = flag("页面边框阴影", borders?.top?.shadow ?? false),
    frame = flag("页面边框立体效果", borders?.top?.frame ?? false),
    header = flag(
      "页面边框包含页眉",
      !session.snapshot().metadata.word?.bordersDoNotSurroundHeader,
    ),
    footer = flag(
      "页面边框包含页脚",
      !session.snapshot().metadata.word?.bordersDoNotSurroundFooter,
    ),
    sides = Object.fromEntries(
      (["top", "left", "bottom", "right"] as const).map((side, index) => [
        side,
        flag("页面边框" + ["上", "左", "下", "右"][index], borders ? Boolean(borders[side]) : true),
      ]),
    ) as Record<"top" | "left" | "bottom" | "right", HTMLInputElement>;
  const button = document.createElement("button");
  button.type = "button";
  button.textContent = "应用页面边框";
  button.addEventListener(
    "click",
    () => {
      try {
        const border: WordPageBorders = {
          offsetFrom: offset.value as WordPageBorders["offsetFrom"],
          display: display.value as WordPageBorders["display"],
          zOrder: order.value as WordPageBorders["zOrder"],
        };
        for (const side of ["top", "left", "bottom", "right"] as const)
          if (sides[side].checked)
            border[side] = {
              style: pattern.value as "single",
              color: color.value.slice(1),
              width: Number(width.value),
              space: Number(space.value),
              shadow: shadow.checked,
              frame: frame.checked,
            };
        const latest = selected(),
          targets = all() ? latest.metadata.sections : [latest.target],
          root = latest.metadata.sections[0]!.rootNodeId;
        let article = root;
        for (
          let node = session.getNode(article);
          node?.parentNodeId;
          node = session.getNode(article)
        )
          article = node.parentNodeId;
        apply([
          ...targets.map((target) => ({
            op: "setWordSection" as const,
            nodeId: target.rootNodeId,
            changes: { pageBorders: border },
          })),
          {
            op: "setWordOptions",
            nodeId: article,
            bordersDoNotSurroundHeader: !header.checked,
            bordersDoNotSurroundFooter: !footer.checked,
          },
        ]);
      } catch (error) {
        onError(error);
      }
    },
    { signal },
  );
  panel.append(button);
  return () => {
    const borders = selected().target.pageBorders,
      border = borders?.top ?? borders?.left ?? borders?.bottom ?? borders?.right;
    pattern.value = border?.style ?? "none";
    color.value = "#" + (border?.color === "auto" ? "000000" : (border?.color ?? "000000"));
    width.value = String(border?.width ?? 0.5);
    space.value = String(border?.space ?? 24);
    offset.value = borders?.offsetFrom ?? "text";
    display.value = borders?.display ?? "allPages";
    order.value = borders?.zOrder ?? "front";
    shadow.checked = border?.shadow ?? false;
    frame.checked = border?.frame ?? false;
    for (const side of ["top", "left", "bottom", "right"] as const)
      sides[side].checked = borders ? Boolean(borders[side]) : true;
  };
}

import { wordMathComponents } from "./word-math-components.js";
import { DocumentError, parseWordMath } from "@doffi/core";
import type {
  DocumentEditingSession,
  DocumentOperation,
  WordMathTree,
  WordMathNode,
  WordTextStyle,
} from "@doffi/core";
function groupMath(trees: readonly WordMathTree[]): WordMathTree {
  const nodes: WordMathNode[] = [],
    children: string[] = [];
  for (const [index, tree] of trees.entries()) {
    const prefix = "g" + index + "-";
    for (const node of tree.nodes)
      nodes.push({
        ...node,
        mathId: prefix + node.mathId,
        ...("children" in node ? { children: node.children.map((id) => prefix + id) } : {}),
      } as WordMathNode);
    children.push(prefix + tree.rootMathId);
  }
  nodes.push({ mathId: "group", kind: "equationGroup", children });
  return { rootMathId: "group", nodes };
}
export function createWordMathControls(
  parent: ShadowRoot,
  before: HTMLElement,
  session: DocumentEditingSession,
  selection: () =>
    | {
        paragraphId: string;
        offset: number;
      }
    | undefined,
  onError: (error: unknown) => void,
): {
  open(insert: boolean): void;
  bind(root: HTMLElement, signal: AbortSignal): void;
  dispose(): void;
} {
  const document = before.ownerDocument;
  let panel: HTMLElement | undefined,
    controller: AbortController | undefined,
    selected: string | undefined;
  const close = (): void => {
    controller?.abort();
    panel?.remove();
    panel = undefined;
  };
  function open(insert: boolean): void {
    close();
    const location = selection(),
      equations = session.snapshot().metadata.word?.equations ?? [],
      equation = equations.find((equation) => equation.nodeId === selected) ?? equations[0];
    if (insert && !location) throw new DocumentError("InvalidOperation", "请先在段落中放置光标。");
    if (!insert && !equation)
      throw new DocumentError("InvalidOperation", "文档中还没有可编辑公式。");
    if (!insert) selected = equation!.nodeId;
    controller = new AbortController();
    const signal = controller.signal;
    panel = document.createElement("section");
    panel.setAttribute("role", "region");
    panel.setAttribute("aria-label", "公式编辑");
    panel.style.cssText =
      "padding:12px;background:white;border:1px solid #dbe3ef;max-height:330px;overflow:auto;display:flex;flex-wrap:wrap;align-items:center;gap:8px";
    parent.insertBefore(panel, before);
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
    const field = (name: string, value: string): HTMLInputElement => {
      const label = document.createElement("label"),
        input = document.createElement("input");
      label.textContent = name;
      input.setAttribute("aria-label", name);
      input.value = value;
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
        const option = document.createElement("option");
        option.value = key;
        option.textContent = text;
        input.append(option);
      }
      input.value = value;
      label.append(input);
      panel!.append(label);
      return input;
    };
    button("关闭公式编辑", close);
    if (!insert) {
      const target = choose(
        "当前公式",
        equations.map((equation, index) => [equation.nodeId, "公式 " + (index + 1)]),
        equation!.nodeId,
      );
      target.addEventListener(
        "change",
        () => {
          selected = target.value;
          open(false);
        },
        { signal },
      );
    }
    const display = choose(
        "公式显示",
        [
          ["inline", "行内"],
          ["block", "独立行"],
        ],
        equation?.display ?? "inline",
      ),
      alignment = choose(
        "公式对齐",
        [
          ["left", "左对齐"],
          ["centerGroup", "居中"],
          ["right", "右对齐"],
        ],
        equation?.alignment ?? "centerGroup",
      ),
      font = field("公式字体", equation?.fontFamily ?? "Cambria Math"),
      inputLabel = document.createElement("label"),
      input = document.createElement("textarea");
    inputLabel.textContent = "线性公式";
    input.setAttribute("aria-label", "线性公式");
    input.value = insert ? "\\frac{x^2+1}{2}" : "";
    input.placeholder = "输入公式以替换当前结构";
    input.rows = 2;
    input.cols = 45;
    inputLabel.append(input);
    panel.append(inputLabel);
    const grouped = document.createElement("input"),
      groupLabel = document.createElement("label");
    grouped.type = "checkbox";
    grouped.setAttribute("aria-label", "多行公式组");
    grouped.checked =
      equation?.nodes.find((node) => node.mathId === equation.rootMathId)?.kind === "equationGroup";
    groupLabel.textContent = "多行公式组（每行一组）";
    groupLabel.append(grouped);
    panel.append(groupLabel);
    for (const [name, value] of [
      ["分式模板", "\\frac{a}{b}"],
      ["根式模板", "\\sqrt[3]{x}"],
      ["上下标模板", "x_1^2"],
      ["重音模板", "\\vec{x}+\\hat{y}"],
      ["矩阵模板", "\\matrix{a&b\\\\c&d}"],
      ["求和模板", "\\sum_{i=1}^{n}{i}"],
    ])
      button(name!, () => {
        input.value = value!;
        input.focus();
      });
    const applyMath = (math: WordMathTree): void => {
      if (!insert) {
        const previous = equation!.nodes.find(
            (node) => node.mathId === equation!.rootMathId,
          )?.style,
          root = math.nodes.find((node) => node.mathId === math.rootMathId)!;
        if (previous && !root.style)
          root.style = JSON.parse(JSON.stringify(previous)) as WordTextStyle;
      }
      const settings = {
        math,
        display: display.value as "inline" | "block",
        alignment: alignment.value as "left" | "right" | "centerGroup",
        fontFamily: font.value,
      };
      if (insert) {
        const previous = new Set(equations.map((equation) => equation.nodeId));
        apply([
          {
            op: "insertWordEquation",
            nodeId: location!.paragraphId,
            offset: location!.offset,
            ...settings,
          },
        ]);
        selected = session
          .snapshot()
          .metadata.word!.equations!.find((equation) => !previous.has(equation.nodeId))!.nodeId;
        open(false);
      } else {
        apply([{ op: "setWordEquation", nodeId: equation!.nodeId, ...settings }]);
        open(false);
      }
    };
    button(insert ? "插入公式" : "按线性输入重建公式", () => {
      if (!input.value.trim()) throw new DocumentError("InvalidOperation", "请输入公式。");
      if (grouped.checked) {
        display.value = "block";
        applyMath(
          groupMath(
            input.value
              .split(/\r?\n/)
              .filter((line) => line.trim())
              .map((line) => parseWordMath(line)),
          ),
        );
      } else applyMath(parseWordMath(input.value));
    });
    if (!insert) {
      button("应用公式显示设置", () =>
        applyMath({
          rootMathId: equation!.rootMathId,
          nodes: JSON.parse(JSON.stringify(equation!.nodes)) as WordMathNode[],
        }),
      );
      button("添加公式组", () => {
        if (!input.value.trim())
          throw new DocumentError("InvalidOperation", "请输入要添加的公式。");
        const old = JSON.parse(JSON.stringify(equation)) as WordMathTree,
          root = old.nodes.find((node) => node.mathId === old.rootMathId)!,
          roots = root.kind === "equationGroup" ? root.children : [old.rootMathId],
          subtrees = roots.map((id) => {
            const included = new Set<string>(),
              visit = (id: string): void => {
                included.add(id);
                const node = old.nodes.find((node) => node.mathId === id)!;
                if ("children" in node) node.children.forEach(visit);
              };
            visit(id);
            return { rootMathId: id, nodes: old.nodes.filter((node) => included.has(node.mathId)) };
          });
        display.value = "block";
        applyMath(groupMath([...subtrees, parseWordMath(input.value)]));
      });
      const formatTarget = choose(
          "公式字符格式目标",
          [
            ["whole", "整个公式"],
            ...equation!.nodes
              .filter((node) => node.mathId !== equation!.rootMathId && node.kind !== "matrixRow")
              .map(
                (node) =>
                  [node.mathId, node.kind === "text" ? node.text || "空文字" : node.kind] as const,
              ),
          ],
          "whole",
        ),
        size = field("公式字号", ""),
        color = field("公式颜色", ""),
        language = field("公式文字语言", ""),
        bold = choose(
          "公式加粗",
          [
            ["", "继承"],
            ["1", "加粗"],
            ["0", "常规"],
          ],
          "",
        ),
        italic = choose(
          "公式斜体",
          [
            ["", "继承"],
            ["1", "斜体"],
            ["0", "直立"],
          ],
          "",
        );
      const fonts = Object.fromEntries(
        (["ascii", "hAnsi", "eastAsia", "cs"] as const).map((role) => [
          role,
          field("公式字体 " + role, ""),
        ]),
      ) as Record<"ascii" | "hAnsi" | "eastAsia" | "cs", HTMLInputElement>;
      const loadStyle = (): void => {
        const node = equation!.nodes.find(
            (node) =>
              node.mathId ===
              (formatTarget.value === "whole" ? equation!.rootMathId : formatTarget.value),
          )!,
          style = node.style;
        size.value = style?.fontSize === undefined ? "" : String(style.fontSize);
        color.value = style?.color ?? "";
        language.value = style?.language ?? "";
        bold.value = style?.bold === undefined ? "" : style.bold ? "1" : "0";
        italic.value = style?.italic === undefined ? "" : style.italic ? "1" : "0";
        for (const role of ["ascii", "hAnsi", "eastAsia", "cs"] as const)
          fonts[role].value = style?.fonts?.[role] ?? "";
      };
      formatTarget.addEventListener("change", loadStyle, { signal });
      loadStyle();
      button("应用公式字符格式", () => {
        const nodes = JSON.parse(JSON.stringify(equation!.nodes)) as WordMathNode[],
          node = nodes.find(
            (node) =>
              node.mathId ===
              (formatTarget.value === "whole" ? equation!.rootMathId : formatTarget.value),
          )!,
          fontValues = Object.fromEntries(
            Object.entries(fonts)
              .filter(([, input]) => input.value.trim())
              .map(([role, input]) => [role, input.value.trim()]),
          ),
          style: WordTextStyle = {
            ...(Object.keys(fontValues).length ? { fonts: fontValues } : {}),
            ...(size.value ? { fontSize: Number(size.value) } : {}),
            ...(color.value ? { color: color.value as "#000000" } : {}),
            ...(language.value ? { language: language.value } : {}),
            ...(bold.value ? { bold: bold.value === "1" } : {}),
            ...(italic.value ? { italic: italic.value === "1" } : {}),
          };
        if (Object.keys(style).length) node.style = style;
        else delete node.style;
        if (formatTarget.value === "whole")
          for (const child of nodes) {
            if (child === node || !child.style) continue;
            for (const key of ["fontSize", "color", "bold", "italic", "language"] as const)
              if (style[key] !== undefined) delete child.style[key];
            for (const role of ["ascii", "hAnsi", "eastAsia", "cs"] as const)
              if (style.fonts?.[role] && child.style.fonts) delete child.style.fonts[role];
          }
        applyMath({ rootMathId: equation!.rootMathId, nodes });
      });
      wordMathComponents(panel, equation!, applyMath, onError);
      const textNodes = equation!.nodes.filter((node) => node.kind === "text");
      if (textNodes.length) {
        const target = choose(
            "公式文字组件",
            textNodes.map((node) => [
              node.mathId,
              node.kind === "text" ? node.text || "空白文字" : node.kind,
            ]),
            textNodes[0]!.mathId,
          ),
          text = field("公式组件文字", textNodes[0]!.kind === "text" ? textNodes[0]!.text : "");
        target.addEventListener(
          "change",
          () => {
            const node = textNodes.find((node) => node.mathId === target.value)!;
            text.value = node.kind === "text" ? node.text : "";
          },
          { signal },
        );
        const variant = choose(
          "数学字母样式",
          [
            "italic",
            "normal",
            "bold",
            "boldItalic",
            "doubleStruck",
            "script",
            "fraktur",
            "sansSerif",
            "monospace",
          ].map((value) => [value, value] as const),
          textNodes[0]!.variant ?? "italic",
        );
        target.addEventListener(
          "change",
          () => {
            variant.value =
              textNodes.find((node) => node.mathId === target.value)!.variant ?? "italic";
          },
          { signal },
        );
        button("更新公式文字组件", () => {
          const nodes = JSON.parse(JSON.stringify(equation!.nodes)) as WordMathNode[],
            node = nodes.find((node) => node.mathId === target.value)!;
          if (node.kind === "text") {
            node.text = text.value;
            node.variant = variant.value as typeof node.variant;
          }
          applyMath({ rootMathId: equation!.rootMathId, nodes });
        });
      }
      button("删除公式", () => {
        apply([{ op: "removeWordEquation", nodeId: equation!.nodeId }]);
        selected = undefined;
        close();
      });
    }
  }
  return {
    open,
    bind(root, signal) {
      root.addEventListener(
        "dblclick",
        (event) => {
          const equation = (event.target as Element).closest<HTMLElement>("[data-word-equation]");
          if (!equation) return;
          event.preventDefault();
          event.stopImmediatePropagation();
          selected = equation.dataset.nodeId;
          try {
            open(false);
          } catch (error) {
            onError(error);
          }
        },
        { signal, capture: true },
      );
    },
    dispose: close,
  };
}

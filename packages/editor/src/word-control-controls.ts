import { DocumentError } from "@doffi/core";
import type {
  DocumentEditingSession,
  DocumentOperation,
  TransactionPrepareOptions,
  WordControl,
} from "@doffi/core";
import { wordImageResource } from "./word-image-resource.js";
export function createWordControlControls(
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
  const apply = (operations: DocumentOperation[], options?: TransactionPrepareOptions): void => {
    const snapshot = session.snapshot();
    session.commit(
      session.prepare(
        {
          protocolVersion: 1,
          documentId: snapshot.manifest.documentId,
          baseRevision: snapshot.manifest.revision,
          operations,
        },
        options,
      ),
    );
  };
  const close = (): void => {
    controller?.abort();
    panel?.remove();
    panel = undefined;
  };
  const attempt = (run: () => void): void => {
    try {
      run();
    } catch (error) {
      onError(error);
    }
  };
  function open(insert: boolean): void {
    close();
    const location = selection(),
      controls = session.snapshot().metadata.word?.controls ?? [],
      control = controls.find((control) => control.nodeId === selected) ?? controls[0];
    if (insert && !location) throw new DocumentError("InvalidOperation", "请先在段落中放置光标。");
    if (!insert && !control) throw new DocumentError("InvalidOperation", "文档中还没有内容控件。");
    if (!insert) selected = control!.nodeId;
    controller = new AbortController();
    const signal = controller.signal;
    panel = document.createElement("section");
    panel.setAttribute("role", "region");
    panel.setAttribute("aria-label", "内容控件设置");
    panel.style.cssText =
      "padding:12px;background:white;border:1px solid #dbe3ef;max-height:330px;overflow:auto;display:flex;flex-wrap:wrap;align-items:center;gap:8px";
    parent.insertBefore(panel, before);
    const field = (name: string, value: string, type = "text"): HTMLInputElement => {
      const label = document.createElement("label"),
        input = document.createElement("input");
      label.textContent = name;
      input.setAttribute("aria-label", name);
      input.type = type;
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
    const button = (label: string, run: () => void): void => {
      const node = document.createElement("button");
      node.type = "button";
      node.textContent = label;
      node.addEventListener("click", () => attempt(run), { signal });
      panel!.append(node);
    };
    button("关闭内容控件设置", close);
    if (!insert) {
      const target = choose(
        "当前内容控件",
        controls.map((item, index) => [
          item.nodeId,
          item.alias ?? item.tag ?? "控件 " + (index + 1),
        ]),
        control!.nodeId,
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
    const kinds = [
      ["richText", "富文本"],
      ["text", "纯文本"],
      ["checkBox", "复选框"],
      ["dropDown", "下拉列表"],
      ["comboBox", "组合框"],
      ["date", "日期"],
      ["picture", "图片"],
      ["equation", "公式内容"],
      ["group", "分组"],
      ["citation", "引文"],
      ["bibliography", "书目"],
    ] as const;
    const kind = choose("控件类型", kinds, control?.kind ?? "richText"),
      level = choose(
        "控件位置",
        [
          ["inline", "行内"],
          ["block", "块级"],
        ],
        control?.level ?? "inline",
      );
    kind.disabled = !insert;
    level.disabled = !insert;
    const alias = field("控件标题", control?.alias ?? ""),
      tag = field("控件标记", control?.tag ?? ""),
      lock = choose(
        "控件锁定",
        [
          ["unlocked", "允许编辑和删除"],
          ["sdtLocked", "禁止删除控件"],
          ["contentLocked", "锁定内容"],
          ["sdtContentLocked", "锁定内容和控件"],
        ],
        control?.lock ?? "unlocked",
      );
    const multiline = field("允许多行", "", "checkbox");
    multiline.checked = control?.multiline ?? true;
    const itemsLabel = document.createElement("label"),
      items = document.createElement("textarea");
    itemsLabel.textContent = "选项（每行显示文字=值）";
    items.setAttribute("aria-label", "控件选项");
    items.value = (control?.items ?? [])
      .map((item) => item.displayText + "=" + item.value)
      .join("\n");
    itemsLabel.append(items);
    panel.append(itemsLabel);
    const checkedFont = field("选中符号字体", control?.checkedFont ?? ""),
      uncheckedFont = field("未选中符号字体", control?.uncheckedFont ?? ""),
      checkedSymbol = field("选中符号", String.fromCharCode(control?.checkedSymbol ?? 0x2612)),
      uncheckedSymbol = field(
        "未选中符号",
        String.fromCharCode(control?.uncheckedSymbol ?? 0x2610),
      );
    const dateFormat = field("日期格式", control?.dateFormat ?? "yyyy-MM-dd"),
      locale = field("日期语言", control?.locale ?? "zh-CN");
    const visibility = (): void => {
      itemsLabel.hidden = !["dropDown", "comboBox"].includes(kind.value);
      for (const field of [checkedFont, uncheckedFont, checkedSymbol, uncheckedSymbol])
        field.parentElement!.hidden = kind.value !== "checkBox";
      for (const field of [dateFormat, locale]) field.parentElement!.hidden = kind.value !== "date";
      multiline.parentElement!.hidden = kind.value !== "text";
    };
    visibility();
    kind.addEventListener("change", visibility, { signal });
    button(insert ? "插入内容控件" : "应用控件属性", () => {
      if (checkedSymbol.value.length !== 1 || uncheckedSymbol.value.length !== 1)
        throw new DocumentError("InvalidOperation", "勾选符号需要一个 BMP 字符。");
      const properties = {
        checkedFont: checkedFont.value || undefined,
        uncheckedFont: uncheckedFont.value || undefined,
        checkedSymbol: checkedSymbol.value.charCodeAt(0),
        uncheckedSymbol: uncheckedSymbol.value.charCodeAt(0),
        kind: kind.value as WordControl["kind"],
        level: level.value as WordControl["level"],
        alias: alias.value,
        tag: tag.value,
        lock: lock.value as NonNullable<WordControl["lock"]>,
        multiline: multiline.checked,
        dateFormat: dateFormat.value,
        locale: locale.value,
        items: items.value
          .split("\n")
          .filter(Boolean)
          .map((line) => {
            const index = line.indexOf("=");
            return index < 0
              ? { displayText: line, value: line }
              : { displayText: line.slice(0, index), value: line.slice(index + 1) };
          }),
      };
      if (insert) {
        const ids = new Set(controls.map((control) => control.nodeId));
        apply([
          {
            op: "insertWordControl",
            nodeId: location!.paragraphId,
            offset: location!.offset,
            control: properties,
            text: properties.kind === "picture" ? "选择图片" : "控件内容",
          },
        ]);
        selected = session
          .snapshot()
          .metadata.word!.controls!.find((control) => !ids.has(control.nodeId))!.nodeId;
      } else {
        const { level: _level, ...changes } = properties;
        apply([{ op: "setWordControl", nodeId: control!.nodeId, changes }]);
      }
      open(false);
    });
    if (!insert) {
      if (control!.kind === "checkBox") {
        const checked = field("控件勾选", "", "checkbox");
        checked.checked = control!.checked ?? false;
        button("应用勾选状态", () => {
          apply([{ op: "setWordControlValue", nodeId: control!.nodeId, value: checked.checked }]);
          open(false);
        });
      } else if (control!.kind === "picture") {
        const input = field("控件图片", "", "file");
        input.accept = "image/png,image/jpeg,image/webp";
        input.addEventListener(
          "change",
          () => {
            void (async () => {
              const file = input.files?.[0];
              if (!file) return;
              const revision = session.snapshot().manifest.revision,
                resource = await wordImageResource(document, file);
              if (session.snapshot().manifest.revision !== revision)
                throw new DocumentError("RevisionConflict", "读取图片期间文档已更改，请重试。");
              apply(
                [
                  {
                    op: "setWordControlPicture",
                    nodeId: control!.nodeId,
                    assetPath: resource.path,
                    alt: file.name,
                    geometry: {
                      x: 0,
                      y: 0,
                      width: resource.width,
                      height: resource.height,
                      rotation: 0,
                      flipH: false,
                      flipV: false,
                    },
                  },
                ],
                resource.options,
              );
              open(false);
            })().catch(onError);
          },
          { signal },
        );
      } else {
        const value =
          control!.kind === "dropDown"
            ? choose(
                "控件值",
                (control!.items ?? []).map((item) => [item.value, item.displayText]),
                control!.value ?? control!.items?.[0]?.value ?? "",
              )
            : field(
                "控件值",
                control!.kind === "date"
                  ? (control!.fullDate?.slice(0, 10) ?? "")
                  : (control!.value ?? session.getNode(control!.nodeId)!.text),
                control!.kind === "date" ? "date" : "text",
              );
        button("应用控件值", () => {
          apply([{ op: "setWordControlValue", nodeId: control!.nodeId, value: value.value }]);
          open(false);
        });
      }
      button("移除控件并保留内容", () => {
        apply([{ op: "removeWordControl", nodeId: control!.nodeId, preserveContent: true }]);
        close();
      });
      button("删除控件及内容", () => {
        apply([{ op: "removeWordControl", nodeId: control!.nodeId, preserveContent: false }]);
        close();
      });
    }
  }
  return {
    open,
    dispose: close,
    bind(root, signal) {
      root.addEventListener(
        "click",
        (event) => {
          const node = (event.target as Element).closest<HTMLElement>(
            "[data-word-content-control]",
          );
          if (!node) return;
          selected = node.dataset.nodeId;
          const control = session
            .snapshot()
            .metadata.word?.controls?.find((control) => control.nodeId === selected);
          if (
            control?.kind === "checkBox" &&
            !["contentLocked", "sdtContentLocked"].includes(control.lock ?? "")
          ) {
            event.preventDefault();
            attempt(() =>
              apply([
                { op: "setWordControlValue", nodeId: control.nodeId, value: !control.checked },
              ]),
            );
          }
        },
        { signal },
      );
      root.addEventListener(
        "keydown",
        (event) => {
          const node = (event.target as Element).closest<HTMLElement>(
            "[data-word-content-control=checkBox]",
          );
          if (!node || ![" ", "Enter"].includes(event.key)) return;
          event.preventDefault();
          event.stopPropagation();
          const control = session
            .snapshot()
            .metadata.word?.controls?.find((control) => control.nodeId === node.dataset.nodeId);
          if (control && !["contentLocked", "sdtContentLocked"].includes(control.lock ?? ""))
            attempt(() =>
              apply([
                { op: "setWordControlValue", nodeId: control.nodeId, value: !control.checked },
              ]),
            );
        },
        { signal, capture: true },
      );
      root.addEventListener(
        "dblclick",
        (event) => {
          const node = (event.target as Element).closest<HTMLElement>(
            "[data-word-content-control]",
          );
          if (!node || node.dataset.wordContentControl === "toc") return;
          event.preventDefault();
          selected = node.dataset.nodeId;
          attempt(() => open(false));
        },
        { signal },
      );
    },
  };
}

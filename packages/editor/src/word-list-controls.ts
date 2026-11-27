import { createAssetDescriptor, createOpaqueId, DocumentError } from "@doffi/core";
import type {
  DocumentEditingSession,
  DocumentOperation,
  TransactionPrepareOptions,
  WordNumberLevel,
} from "@doffi/core";
import { wordImageResource } from "./word-image-resource.js";
import { createWordDockedPane, replaceSameSideDock } from "./word-pane-chrome.js";

const BUILTIN_PIC_BULLETS: readonly (readonly [string, string, (ctx: CanvasRenderingContext2D) => void])[] = [
  [
    "blue-diamond",
    "蓝色菱形",
    (ctx) => {
      ctx.fillStyle = "#185abd";
      ctx.beginPath();
      ctx.moveTo(16, 2);
      ctx.lineTo(30, 16);
      ctx.lineTo(16, 30);
      ctx.lineTo(2, 16);
      ctx.closePath();
      ctx.fill();
    },
  ],
  [
    "green-check",
    "绿色对勾",
    (ctx) => {
      ctx.strokeStyle = "#107c41";
      ctx.lineWidth = 4.5;
      ctx.lineCap = "round";
      ctx.lineJoin = "round";
      ctx.beginPath();
      ctx.moveTo(5, 17);
      ctx.lineTo(13, 25);
      ctx.lineTo(27, 7);
      ctx.stroke();
    },
  ],
  [
    "gold-star",
    "金色星标",
    (ctx) => {
      ctx.fillStyle = "#e6a800";
      ctx.beginPath();
      for (let i = 0; i < 10; i++) {
        const r = i % 2 === 0 ? 14 : 6;
        const a = (Math.PI * i) / 5 - Math.PI / 2;
        const x = 16 + Math.cos(a) * r;
        const y = 16 + Math.sin(a) * r;
        if (i === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      }
      ctx.closePath();
      ctx.fill();
    },
  ],
  [
    "red-arrow",
    "橙红箭头",
    (ctx) => {
      ctx.fillStyle = "#c55a11";
      ctx.beginPath();
      ctx.moveTo(4, 10);
      ctx.lineTo(17, 10);
      ctx.lineTo(17, 4);
      ctx.lineTo(29, 16);
      ctx.lineTo(17, 28);
      ctx.lineTo(17, 22);
      ctx.lineTo(4, 22);
      ctx.closePath();
      ctx.fill();
    },
  ],
];

async function createBuiltinBulletPng(
  document: Document,
  id: string,
): Promise<{ path: string; options: TransactionPrepareOptions }> {
  const preset = BUILTIN_PIC_BULLETS.find((p) => p[0] === id) ?? BUILTIN_PIC_BULLETS[0]!;
  const canvas = document.createElement("canvas");
  canvas.width = 32;
  canvas.height = 32;
  const ctx = canvas.getContext("2d")!;
  preset[2](ctx);
  const blob = await new Promise<Blob>((resolve, reject) =>
    canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("生成图片项目符号失败。"))), "image/png"),
  );
  const bytes = new Uint8Array(await blob.arrayBuffer());
  const path = `assets/pic-bullet-${id}-${createOpaqueId().slice(0, 6)}.png`;
  return {
    path,
    options: { assets: [{ descriptor: createAssetDescriptor(path, "image/png", bytes), bytes }] },
  };
}
/** List instances share native definitions; restarting allocates an independent
 * instance and applies its continuation in the same transaction/history step. */
export function createWordListControls(
  parent: ShadowRoot,
  before: HTMLElement,
  session: DocumentEditingSession,
  paragraphs: () => readonly string[],
  onError: (error: unknown) => void,
): { open(): void; dispose(): void } {
  const document = before.ownerDocument;
  let shell: HTMLElement | undefined, panel: HTMLElement | undefined, controller: AbortController | undefined;
  const close = (): void => {
    controller?.abort();
    shell?.remove();
    shell = undefined;
    panel = undefined;
    controller = undefined;
  };
  function open(): void {
    const ids = paragraphs();
    if (!ids.length) throw new DocumentError("InvalidOperation", "请先选择段落。");
    close();
    replaceSameSideDock(parent, "right");
    controller = new AbortController();
    const signal = controller.signal;
    const chrome = createWordDockedPane(document, {
      ariaLabel: "列表设置",
      title: "项目符号和编号",
      dock: "right",
      closeLabel: "关闭列表设置",
      onClose: close,
      signal,
      width: "332px",
    });
    shell = chrome.panel;
    panel = chrome.body;
    panel.style.display = "grid";
    panel.style.gridTemplateColumns = "repeat(2,minmax(0,1fr))";
    panel.style.alignContent = "start";
    panel.style.gap = "8px 10px";
    parent.insertBefore(shell, before);
    const button = (name: string, run: () => void): void => {
      const button = document.createElement("button");
      button.type = "button";
      button.textContent = name;
      button.addEventListener(
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
      panel!.append(button);
    };
    const field = (name: string, value: string | number, type = "number"): HTMLInputElement => {
      const label = document.createElement("label"),
        input = document.createElement("input");
      label.textContent = name;
      input.setAttribute("aria-label", name);
      input.type = type;
      input.value = String(value);
      input.style.width = type === "number" ? "70px" : "120px";
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
      for (const [value, text] of options) {
        const option = document.createElement("option");
        option.value = value;
        option.textContent = text;
        input.append(option);
      }
      input.value = value;
      label.append(input);
      panel!.append(label);
      return input;
    };
    const metadata = session.snapshot().metadata.word,
      first = session.getNode(ids[0]!)!,
      original = first.attributes["data-word-numbering"],
      definition = metadata?.numbering.find((item) => item.numberingId === original),
      initialLevel = Number(first.attributes["data-word-level"] ?? 0),
      initial = definition?.levels.find((level) => level.level === initialLevel);
    const existing = choose(
        "已有列表",
        (metadata?.numbering ?? []).map((item) => [item.numberingId, item.numberingId] as const),
        original ?? "",
      ),
      level = choose(
        "列表级别",
        Array.from({ length: 9 }, (_, n) => [String(n), String(n + 1)] as const),
        String(initialLevel),
      );
    const format = choose(
        "编号格式",
        [
          ["decimal", "1, 2, 3"],
          ["lowerLetter", "a, b, c"],
          ["upperLetter", "A, B, C"],
          ["lowerRoman", "i, ii, iii"],
          ["upperRoman", "I, II, III"],
          ["bullet", "项目符号"],
          ["none", "无标记"],
        ],
        initial?.format ?? "decimal",
      ),
      template = field("编号文本", initial?.text ?? "%1.", "text"),
      start = field("起始编号", initial?.start ?? 1),
      suffix = choose(
        "编号之后",
        [
          ["tab", "制表符"],
          ["space", "空格"],
          ["nothing", "无"],
        ],
        initial?.suffix ?? "tab",
      );
    start.min = "0";
    start.step = "1";
    const indent = field("列表左缩进（磅）", (initialLevel + 1) * 24),
      hanging = field("悬挂距离（磅）", 18),
      font = field("标记字体", "", "text");
    indent.min = hanging.min = "0";
    indent.step = hanging.step = ".5";
    const linkedStyle = choose(
      "将级别链接到样式",
      [
        ["", "（无样式）"],
        ...(metadata?.styles ?? [])
          .filter((s) => s.kind === "paragraph")
          .map((s) => [s.styleId, s.name ?? s.styleId] as const),
      ],
      initial?.pStyle ?? "",
    );
    const legalNumbering = field("正规形式编号 (isLgl)", "", "checkbox");
    legalNumbering.checked = Boolean(initial?.isLgl);
    const picPreset = choose(
      "图片项目符号",
      [
        ["", "不使用图片"],
        ...BUILTIN_PIC_BULLETS.map(([id, label]) => [id, label] as const),
        ...(initial?.pictureBulletAssetPath ? [["__current__", "保留当前图片"] as const] : []),
      ],
      initial?.pictureBulletAssetPath ? "__current__" : "",
    );
    let uploadedPic: { path: string; options: TransactionPrepareOptions } | undefined;
    const picUpload = document.createElement("input");
    picUpload.type = "file";
    picUpload.accept = "image/png,image/jpeg,image/webp";
    picUpload.setAttribute("aria-label", "上传自定义图片项目符号");
    picUpload.style.width = "160px";
    picUpload.addEventListener(
      "change",
      () => {
        const file = picUpload.files?.[0];
        if (!file) return;
        void wordImageResource(document, file)
          .then((res) => {
            uploadedPic = { path: res.path, options: res.options };
            format.value = "bullet";
            template.value = "•";
            picPreset.value = "";
          })
          .catch(onError);
      },
      { signal },
    );
    panel.append(picUpload);
    const following = field("包含后续同组列表", "", "checkbox");
    following.checked = true;
    format.addEventListener(
      "change",
      () => {
        template.value =
          format.value === "bullet"
            ? "•"
            : format.value === "none"
              ? ""
              : "%" + (Number(level.value) + 1) + ".";
      },
      { signal },
    );
    const targets = (): string[] => {
      const result = [...ids];
      if (!following.checked || !original) return result;
      const last = session.getNode(ids.at(-1)!)!,
        siblings = last.parentNodeId ? session.getNode(last.parentNodeId)!.childNodeIds : [];
      for (const id of siblings.slice(siblings.indexOf(last.nodeId) + 1)) {
        const node = session.getNode(id)!;
        if (node.attributes["data-word-numbering"] !== original) break;
        result.push(id);
      }
      return result;
    };
    const apply = (
      numberingId: string,
      operations: DocumentOperation[] = [],
      prepareOptions?: TransactionPrepareOptions,
    ): void => {
      const snapshot = session.snapshot();
      for (const id of targets()) {
        const node = session.getNode(id)!;
        operations.push({
          op: "setWordList",
          nodeId: id,
          numberingId,
          level: ids.includes(id)
            ? Number(level.value)
            : Number(node.attributes["data-word-level"] ?? 0),
        });
        if (ids.includes(id))
          operations.push({
            op: "setWordParagraph",
            nodeId: id,
            properties: {
              "margin-left": Number(indent.value) + "pt",
              "text-indent": -Number(hanging.value) + "pt",
            },
          });
      }
      session.commit(
        session.prepare(
          {
            protocolVersion: 1,
            documentId: snapshot.manifest.documentId,
            baseRevision: snapshot.manifest.revision,
            operations,
          },
          prepareOptions,
        ),
      );
    };
    button("应用已有列表", () => {
      if (!existing.value) throw new DocumentError("InvalidOperation", "没有可应用的已有列表。");
      apply(existing.value);
    });
    button("继续前一列表", () => {
      const source = new DOMParser().parseFromString(session.snapshot().html, "text/html"),
        selected = source.querySelector('[data-node-id="' + CSS.escape(ids[0]!) + '"]')!,
        story = selected.closest("[data-office-story]") ?? selected.closest("article")!;
      const all = [...story.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6")],
        at = all.indexOf(selected as HTMLElement);
      const previous = all
        .slice(0, at)
        .reverse()
        .find((node) => node.dataset.wordNumbering);
      if (!previous) throw new DocumentError("InvalidOperation", "前面没有可继续的列表。");
      apply(previous.dataset.wordNumbering!);
    });
    button("应用并重新编号", () => {
      void (async () => {
        try {
          let picAssetPath: string | undefined;
          let prepareOptions: TransactionPrepareOptions | undefined;
          if (uploadedPic) {
            picAssetPath = uploadedPic.path;
            prepareOptions = uploadedPic.options;
          } else if (picPreset.value === "__current__" && initial?.pictureBulletAssetPath) {
            picAssetPath = initial.pictureBulletAssetPath;
          } else if (picPreset.value) {
            const generated = await createBuiltinBulletPng(document, picPreset.value);
            picAssetPath = generated.path;
            prepareOptions = generated.options;
          }
          const numberingId = "list-" + createOpaqueId(),
            selectedLevel = Number(level.value),
            levels: WordNumberLevel[] = Array.from({ length: 9 }, (_, n) => {
              const old = definition?.levels.find((level) => level.level === n);
              return old
                ? { ...old, cssClass: numberingId + "-" + n }
                : {
                    level: n,
                    start: 1,
                    format: "decimal",
                    text:
                      Array.from({ length: n + 1 }, (_, index) => "%" + (index + 1)).join(".") +
                      ".",
                    suffix: "tab",
                    cssClass: numberingId + "-" + n,
                  };
            });
          const nextLevel: WordNumberLevel = {
            ...levels[selectedLevel]!,
            start: Number(start.value),
            format: (picAssetPath ? "bullet" : format.value) as WordNumberLevel["format"],
            text: picAssetPath ? "•" : template.value,
            suffix: suffix.value as WordNumberLevel["suffix"],
          };
          if (picAssetPath) {
            nextLevel.pictureBulletAssetPath = picAssetPath;
            nextLevel.pictureBulletWidthPt = 10;
            nextLevel.pictureBulletHeightPt = 10;
          } else {
            delete nextLevel.pictureBulletAssetPath;
            delete nextLevel.pictureBulletWidthPt;
            delete nextLevel.pictureBulletHeightPt;
          }
          if (linkedStyle.value) nextLevel.pStyle = linkedStyle.value;
          else delete nextLevel.pStyle;
          if (legalNumbering.checked) nextLevel.isLgl = true;
          else delete nextLevel.isLgl;
          levels[selectedLevel] = nextLevel;
          const root = new DOMParser()
            .parseFromString(session.snapshot().html, "text/html")
            .querySelector<HTMLElement>('[data-office-kind="document"]')!;
          apply(
            numberingId,
            [
              {
                op: "setWordNumbering",
                nodeId: root.dataset.nodeId!,
                numberingId,
                ...(original ? { copyFromNumberingId: original } : {}),
                levels,
                levelProperties: levels.map((item) => ({
                  level: item.level,
                  properties: {
                    ...(font.value ? { "font-family": JSON.stringify(font.value) } : {}),
                    "margin-left":
                      (item.level === selectedLevel ? Number(indent.value) : (item.level + 1) * 24) +
                      "pt",
                    "text-indent": -Number(hanging.value) + "pt",
                  },
                })),
              },
            ],
            prepareOptions,
          );
        } catch (error) {
          onError(error);
        }
      })();
    });
  }
  return { open, dispose: close };
}

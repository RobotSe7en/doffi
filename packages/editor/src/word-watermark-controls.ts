import { wordImageResource } from "./word-image-resource.js";
import { DocumentError } from "@doffi/core";
import type {
  DocumentEditingSession,
  DocumentOperation,
  TransactionPrepareOptions,
} from "@doffi/core";
export function createWordWatermarkControls(
  parent: ShadowRoot,
  before: HTMLElement,
  session: DocumentEditingSession,
  sectionRoot: () => string,
  onError: (error: unknown) => void,
): {
  open(): void;
  dispose(): void;
} {
  const document = before.ownerDocument;
  let panel: HTMLElement | undefined, controller: AbortController | undefined;
  const close = (): void => {
    controller?.abort();
    panel?.remove();
    panel = undefined;
  };
  function open(): void {
    close();
    controller = new AbortController();
    const signal = controller.signal;
    panel = document.createElement("section");
    panel.setAttribute("role", "region");
    panel.setAttribute("aria-label", "水印设置");
    panel.style.cssText =
      "padding:12px;background:white;border:1px solid #dbe3ef;display:flex;flex-wrap:wrap;gap:8px";
    parent.insertBefore(panel, before);
    const field = (name: string, type: string, value: string): HTMLInputElement => {
      const label = document.createElement("label"),
        input = document.createElement("input");
      label.textContent = name;
      input.type = type;
      input.setAttribute("aria-label", name);
      input.value = value;
      label.append(input);
      panel!.append(label);
      return input;
    };
    const button = (name: string, run: () => void | Promise<void>): void => {
      const button = document.createElement("button");
      button.type = "button";
      button.textContent = name;
      button.addEventListener(
        "click",
        () => {
          try {
            const result = run();
            if (result) void result.catch(onError);
          } catch (error) {
            onError(error);
          }
        },
        { signal },
      );
      panel!.append(button);
    };
    button("关闭水印设置", close);
    const text = field("水印文字", "text", "机密"),
      font = field("水印字体", "text", "Noto Sans CJK SC"),
      size = field("水印字号（磅）", "number", "48"),
      color = field("水印颜色", "color", "#cccccc"),
      rotation = field("水印旋转（度）", "number", "-45");
    size.min = "1";
    rotation.min = "-360";
    rotation.max = "360";
    const label = document.createElement("label"),
      scope = document.createElement("select");
    label.textContent = "水印应用范围";
    scope.setAttribute("aria-label", "水印应用范围");
    for (const [value, text] of [
      ["all", "全文"],
      ["section", "本节"],
    ]) {
      const option = document.createElement("option");
      option.value = value!;
      option.textContent = text!;
      scope.append(option);
    }
    label.append(scope);
    panel.append(label);
    const apply = (
      watermark: Extract<
        DocumentOperation,
        {
          op: "setWordWatermark";
        }
      >["watermark"],
      resources?: TransactionPrepareOptions,
    ): void => {
      const snapshot = session.snapshot(),
        metadata = snapshot.metadata;
      if (metadata.presentation)
        throw new DocumentError("InvalidOperation", "水印需要 Word 文档。");
      const targets =
        scope.value === "all"
          ? metadata.sections
          : metadata.sections.filter((section) => section.rootNodeId === sectionRoot());
      if (!targets.length) throw new DocumentError("InvalidOperation", "水印目标节不存在。");
      session.commit(
        session.prepare(
          {
            protocolVersion: 1,
            documentId: snapshot.manifest.documentId,
            baseRevision: snapshot.manifest.revision,
            operations: targets.map((section) => ({
              op: "setWordWatermark",
              nodeId: section.rootNodeId,
              watermark,
            })),
          },
          resources,
        ),
      );
    };
    button("应用文字水印", () =>
      apply({
        kind: "text",
        text: text.value,
        fontFamily: font.value,
        fontSize: Number(size.value),
        color: color.value,
        rotation: Number(rotation.value),
      }),
    );
    const image = field("水印图片", "file", "");
    image.accept = "image/png,image/jpeg,image/webp";
    const scale = field("水印图片缩放（%）", "number", "100");
    scale.min = "1";
    scale.max = "1000";
    button("应用图片水印", async () => {
      const file = image.files?.[0];
      if (!file) throw new DocumentError("InvalidOperation", "请选择水印图片。");
      const revision = session.snapshot().manifest.revision,
        resource = await wordImageResource(document, file);
      if (session.snapshot().manifest.revision !== revision)
        throw new DocumentError("RevisionConflict", "读取水印图片期间文档已变化，请重试。");
      apply(
        {
          kind: "image",
          assetPath: resource.path,
          alt: file.name,
          width: (resource.width * Number(scale.value)) / 100,
          height: (resource.height * Number(scale.value)) / 100,
          rotation: Number(rotation.value),
        },
        resource.options,
      );
    });
    button("移除水印", () => apply(null));
  }
  return { open, dispose: close };
}

import { DocumentError } from "@doffi/core";
import type {
  Annotation,
  AnnotationReply,
  DeepReadonly,
  DocumentEditingSession,
  DocumentOperation,
} from "@doffi/core";
const escape = (value: string): string =>
  value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
/** Drafts retain their opening revision while the rest of the workbench updates. */
export function createWordCommentControls(
  parent: ShadowRoot,
  session: DocumentEditingSession,
  onError: (error: unknown) => void,
  onEditBody?: (annotationId: string, replyId?: string) => void,
): {
  append(card: HTMLElement, annotation: DeepReadonly<Annotation>, signal: AbortSignal): void;
  dispose(): void;
} {
  const document = parent.ownerDocument;
  let panel: HTMLElement | undefined, controller: AbortController | undefined;
  const close = (): void => {
    controller?.abort();
    panel?.remove();
    panel = undefined;
  };
  const target = (
    annotation: DeepReadonly<Annotation>,
    reply?: DeepReadonly<AnnotationReply>,
  ): string => {
    const message = reply ?? annotation;
    if (message.bodyNodeId) return message.bodyNodeId;
    const anchor = annotation.anchor;
    if (anchor.kind === "text" || anchor.kind === "node") return anchor.nodeId;
    if (anchor.kind === "text-set") return anchor.ranges[0]!.nodeId;
    if (anchor.kind === "region") return anchor.surfaceId;
    if (anchor.kind === "node-set" && anchor.nodeIds.length) return anchor.nodeIds[0]!;
    throw new DocumentError("InvalidOperation", "批注没有可用的正文或锚点。");
  };
  const open = (
    annotation: DeepReadonly<Annotation>,
    reply?: DeepReadonly<AnnotationReply>,
    adding = false,
  ): void => {
    close();
    controller = new AbortController();
    const signal = controller.signal,
      snapshot = session.snapshot();
    panel = document.createElement("section");
    panel.setAttribute("role", "region");
    panel.setAttribute(
      "aria-label",
      adding ? "添加批注回复" : reply ? "修改批注回复" : "修改批注正文",
    );
    panel.style.cssText = "padding:12px;border:1px solid #d5dbe2;background:white";
    const input = document.createElement("textarea"),
      author = document.createElement("input");
    input.setAttribute("aria-label", adding ? "回复内容" : "批注消息内容");
    input.value = adding
      ? ""
      : ((reply ?? annotation).instruction ??
        session.getNode((reply ?? annotation).bodyNodeId!)?.text ??
        "");
    author.setAttribute("aria-label", "回复作者");
    author.value = adding ? "" : ((reply ?? annotation).author?.name ?? "");
    panel.append(input);
    if (adding) panel.append(author);
    const button = (label: string, run: () => void): void => {
      const node = document.createElement("button");
      node.type = "button";
      node.textContent = label;
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
    button(adding ? "发送回复" : "保存消息", () => {
      if (!input.value.trim()) throw new DocumentError("InvalidOperation", "请填写批注内容。");
      const operation: DocumentOperation = adding
        ? {
            op: "addWordCommentReply",
            nodeId: target(annotation),
            annotationId: annotation.annotationId,
            html: "<p>" + escape(input.value) + "</p>",
            ...(author.value ? { author: { name: author.value } } : {}),
            date: new Date().toISOString(),
          }
        : {
            op: "setWordCommentBody",
            nodeId: target(annotation, reply),
            annotationId: annotation.annotationId,
            ...(reply ? { replyId: reply.replyId } : {}),
            html: "<p>" + escape(input.value) + "</p>",
          };
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: snapshot.manifest.documentId,
          baseRevision: snapshot.manifest.revision,
          operations: [operation],
        }),
      );
      close();
    });
    button("取消消息编辑", close);
    parent.append(panel);
    input.focus();
  };
  return {
    append(card, annotation, signal) {
      const button = (label: string, run: () => void, parent: HTMLElement = card): void => {
        const node = document.createElement("button");
        node.type = "button";
        node.textContent = label;
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
        parent.append(node);
      };
      const identity = document.createElement("p");
      identity.textContent = [annotation.author?.name, annotation.date].filter(Boolean).join(" · ");
      if (identity.textContent) card.prepend(identity);
      if (annotation.bodyNodeId) {
        button("修改批注正文", () => open(annotation));
        if (onEditBody) button("编辑批注富文本", () => onEditBody(annotation.annotationId));
      }
      button("回复批注", () => open(annotation, undefined, true));
      for (const reply of annotation.replies ?? []) {
        const item = document.createElement("div");
        item.dataset.replyId = reply.replyId;
        item.style.cssText = "border-left:2px solid #d5dbe2;padding-left:10px;margin-top:10px";
        const author = document.createElement("p"),
          body = document.createElement("p");
        author.textContent = [reply.author?.name ?? "回复", reply.date].filter(Boolean).join(" · ");
        body.textContent = reply.instruction ?? session.getNode(reply.bodyNodeId!)?.text ?? "";
        item.append(author, body);
        button("修改回复", () => open(annotation, reply), item);
        if (reply.bodyNodeId && onEditBody)
          button("编辑回复富文本", () => onEditBody(annotation.annotationId, reply.replyId), item);
        button(
          "删除回复",
          () => {
            const snapshot = session.snapshot();
            session.commit(
              session.prepare({
                protocolVersion: 1,
                documentId: snapshot.manifest.documentId,
                baseRevision: snapshot.manifest.revision,
                operations: [
                  {
                    op: "removeWordCommentReply",
                    nodeId: target(annotation, reply),
                    annotationId: annotation.annotationId,
                    replyId: reply.replyId,
                  },
                ],
              }),
            );
          },
          item,
        );
        card.append(item);
      }
    },
    dispose: close,
  };
}

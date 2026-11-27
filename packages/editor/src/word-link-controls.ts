import { createOpaqueId, DocumentError, getWordOutline } from "@doffi/core";
import type { DocumentEditingSession, DocumentOperation } from "@doffi/core";
import { createOfficeDialog } from "./office-dialog.js";
export interface WordReferenceRange {
  paragraphId: string;
  endNodeId?: string;
  from: number;
  to: number;
  text: string;
}
export function createWordLinkControls(
  parent: ShadowRoot,
  before: HTMLElement,
  session: DocumentEditingSession,
  selection: () => WordReferenceRange | undefined,
  focus: (nodeId: string) => void,
  onError: (error: unknown) => void,
): { open(kind: "link" | "bookmark" | "reference" | "caption"): void; dispose(): void } {
  const document = before.ownerDocument;
  let dialogUI: ReturnType<typeof createOfficeDialog> | undefined,
    panel: HTMLElement | undefined,
    controller: AbortController | undefined;
  const close = (): void => {
    controller?.abort();
    dialogUI?.close();
    dialogUI?.dialog.remove();
    dialogUI = undefined;
    panel?.remove();
    panel = undefined;
  };
  function open(kind: "link" | "bookmark" | "reference" | "caption"): void {
    close();
    const range = selection();
    if (!range) throw new DocumentError("InvalidOperation", "请先在段落中选择文字或放置光标。");
    controller = new AbortController();
    const signal = controller.signal;
    const title =
      kind === "bookmark"
        ? "书签"
        : kind === "reference"
          ? "交叉引用"
          : kind === "caption"
            ? "题注"
            : "插入超链接";
    dialogUI = createOfficeDialog(document, title);
    dialogUI.dialog.style.width = "480px";
    dialogUI.content.style.minHeight = "0";
    dialogUI.content.style.marginTop = "12px";
    panel = document.createElement("section");
    panel.setAttribute("role", "region");
    panel.setAttribute("aria-label", "链接和引用");
    panel.style.cssText = "display:flex;flex-direction:column;gap:10px";
    dialogUI.content.append(panel);
    parent.insertBefore(dialogUI.dialog, before);
    dialogUI.dialog.addEventListener("cancel", (event) => { event.preventDefault(); close(); }, { signal });
    dialogUI.show();
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
      if (name === "关闭链接和引用") dialogUI!.footer.append(node);
      else panel!.append(node);
    };
    const field = (name: string, value = ""): HTMLInputElement => {
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
      value = "",
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
    button("关闭链接和引用", close);
    const word = session.snapshot().metadata.word!,
      bookmarks = word.bookmarks;
    if (kind === "bookmark") {
      const name = field("书签名称");
      button("添加书签", () => {
        apply([
          {
            op: "insertWordBookmark",
            nodeId: range.paragraphId,
            from: range.from,
            to: range.to,
            expectedText: range.text,
            name: name.value,
            ...(range.endNodeId ? { endNodeId: range.endNodeId } : {}),
          },
        ]);
        close();
      });
      for (const bookmark of bookmarks) {
        button("重定义书签 " + bookmark.name, () => {
          apply([
            {
              op: "setWordBookmark",
              bookmarkId: bookmark.bookmarkId,
              nodeId: range.paragraphId,
              endNodeId: range.endNodeId,
              from: range.from,
              to: range.to,
              expectedText: range.text,
              name: name.value || bookmark.name,
            },
          ]);
          close();
        });
        button("定位书签 " + bookmark.name, () => focus(bookmark.startNodeId));
        button("删除书签 " + bookmark.name, () => {
          const root = new DOMParser()
            .parseFromString(session.snapshot().html, "text/html")
            .querySelector<HTMLElement>('[data-office-kind="document"]')!.dataset.nodeId!;
          apply([{ op: "removeWordBookmark", nodeId: root, bookmarkId: bookmark.bookmarkId }]);
          open("bookmark");
        });
      }
      return;
    }
    if (kind === "reference") {
      const headings = getWordOutline(session);
      const captions = (word.fields ?? [])
        .filter((f) => /^\s*SEQ\b/i.test(f.instruction))
        .map((f) => {
          let pId: string | null = f.nodeId;
          while (pId && !/^p$|^h[1-6]$/.test(session.getNode(pId)?.tagName ?? "")) {
            pId = session.getNode(pId)?.parentNodeId ?? null;
          }
          const node = pId ? session.getNode(pId) : undefined;
          return pId && node ? { nodeId: pId, text: node.text.trim() || "题注" } : undefined;
        })
        .filter((item): item is { nodeId: string; text: string } => Boolean(item));
      const notes = (word.noteReferences ?? []).map((ref, idx) => {
        const note = word.notes?.find((n) => n.noteId === ref.noteId);
        const story = note ? word.stories.find((s) => s.storyId === note.storyId) : undefined;
        const text = story ? session.getNode(story.rootNodeId)?.text.replace(/^\ufffc\s*/, "").trim().slice(0, 40) : "";
        return {
          nodeId: ref.nodeId,
          label: `${note?.kind === "endnote" ? "尾注" : "脚注"} ${idx + 1}${text ? " · " + text : ""}`,
        };
      });
      const sourceTypes: [string, string][] = [];
      if (headings.length) sourceTypes.push(["heading", "标题"]);
      if (bookmarks.length) sourceTypes.push(["bookmark", "书签"]);
      if (captions.length) sourceTypes.push(["caption", "题注"]);
      if (notes.length) sourceTypes.push(["note", "脚注/尾注"]);
      if (!sourceTypes.length)
        throw new DocumentError("InvalidOperation", "文档中暂无可引用的标题、书签、题注或注释。");

      const refKind = choose("引用类型", sourceTypes, sourceTypes[0]![0]);
      const targetSelect = choose("引用哪一个目标", [], "");
      const contentSelect = choose(
        "引用内容",
        [
          ["REF", "目标文字"],
          ["PAGEREF", "页码"],
          ["NOTEREF", "注释编号"],
        ],
        "REF",
      );
      const populateTargets = () => {
        targetSelect.replaceChildren();
        const items: [string, string][] =
          refKind.value === "bookmark"
            ? bookmarks.map((b) => [b.name, b.name])
            : refKind.value === "heading"
              ? headings.map((h) => [h.nodeId, "  ".repeat(Math.max(0, h.level - 1)) + (h.text || "空标题")])
              : refKind.value === "caption"
                ? captions.map((c) => [c.nodeId, c.text.slice(0, 50)])
                : notes.map((n) => [n.nodeId, n.label]);
        for (const [val, label] of items) {
          const opt = document.createElement("option");
          opt.value = val;
          opt.textContent = label;
          targetSelect.append(opt);
        }
        if (items.length) targetSelect.value = items[0]![0];
        if (refKind.value === "note") contentSelect.value = "NOTEREF";
        else if (contentSelect.value === "NOTEREF") contentSelect.value = "REF";
      };
      populateTargets();
      refKind.addEventListener("change", populateTargets, { signal });

      button("插入交叉引用", () => {
        const selectedVal = targetSelect.value;
        if (!selectedVal) throw new DocumentError("InvalidOperation", "请选择交叉引用目标。");
        const ops: DocumentOperation[] = [];
        let bookmarkName = selectedVal;
        let cachedText = "1";
        if (refKind.value === "bookmark") {
          const bm = bookmarks.find((b) => b.name === selectedVal);
          const startNode = bm ? session.getNode(bm.startNodeId) : undefined;
          cachedText = contentSelect.value === "PAGEREF" ? "1" : (startNode?.text.trim() || selectedVal);
        } else {
          const targetNode = session.getNode(selectedVal);
          if (!targetNode) throw new DocumentError("InvalidOperation", "目标节点已不存在。");
          const existing = bookmarks.find((b) => b.startNodeId === selectedVal);
          if (existing) {
            bookmarkName = existing.name;
          } else {
            bookmarkName = "_Ref" + createOpaqueId().replace(/[^a-zA-Z0-9]/g, "").slice(0, 9);
            const targetText = targetNode.text;
            ops.push({
              op: "insertWordBookmark",
              nodeId: selectedVal,
              from: 0,
              to: targetText.length,
              expectedText: targetText,
              name: bookmarkName,
            });
          }
          cachedText = contentSelect.value === "PAGEREF" ? "1" : (targetNode.text.trim() || "1");
        }
        ops.push({
          op: "insertWordField",
          nodeId: range.paragraphId,
          offset: range.from,
          instruction: ` ${contentSelect.value} ${bookmarkName} \\h `,
          cachedText,
        });
        apply(ops);
        close();
      });
      return;
    }
    if (kind === "caption") {
      const name = field("题注标签", "图");
      button("插入题注编号", () => {
        apply([
          {
            op: "insertWordField",
            nodeId: range.paragraphId,
            offset: range.from,
            instruction: 'SEQ "' + name.value.replaceAll('"', "") + '" \\* Arabic',
            cachedText: "1",
          },
        ]);
        close();
      });
      return;
    }
    const target = field("链接地址", "https://"),
      bookmark = choose("链接到书签", [
        ["", "外部链接"],
        ...bookmarks.map((bookmark) => [bookmark.name, bookmark.name] as const),
      ]);
    button("插入链接", () => {
      if (range.endNodeId && range.endNodeId !== range.paragraphId)
        throw new DocumentError("InvalidOperation", "链接文字请限定在同一段落。");
      if (range.from === range.to)
        throw new DocumentError("InvalidOperation", "请先选择要添加链接的文字。");
      apply([
        {
          op: "insertWordLink",
          nodeId: range.paragraphId,
          from: range.from,
          to: range.to,
          expectedText: range.text,
          ...(bookmark.value ? { bookmark: bookmark.value } : { target: target.value }),
        },
      ]);
      close();
    });
    for (const link of word.hyperlinks) {
      button("定位链接 " + (link.target ?? link.bookmark), () => focus(link.nodeId));
      button("删除链接 " + (link.target ?? link.bookmark), () => {
        apply([{ op: "removeWordLink", nodeId: link.nodeId }]);
        close();
      });
    }
  }
  return { open, dispose: close };
}

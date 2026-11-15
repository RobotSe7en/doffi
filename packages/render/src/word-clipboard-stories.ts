import { DocumentError } from "@doffi/core";
import type { ContentDocumentModel } from "@doffi/core";
export function clipboardNodeText(node: Node): string {
  if (node.nodeType === 3) return node.textContent ?? "";
  if (node.nodeType === 1) {
    const element = node as Element;
    if (element.localName === "img") return "\ufffc";
    if (element.localName === "br") return "\n";
  }
  return [...node.childNodes].map(clipboardNodeText).join("");
}
/** Traverse ownership from selected objects only. Sibling stories and source
 * body content never enter the transport graph. */
export function copyWordClipboardStories(
  holder: HTMLElement,
  model: ContentDocumentModel,
  completeNodes: ReadonlySet<string> | undefined,
): {
  roots: HTMLElement[];
  stories: {
    storyId: string;
    rootNodeId: string;
    kind: "textBox" | "footnote" | "endnote";
    holder: HTMLElement;
  }[];
  completeNodes: ReadonlySet<string> | undefined;
} {
  const word = model.snapshot().metadata.word!,
    complete = new Set(completeNodes),
    roots = [holder],
    stories: {
      storyId: string;
      rootNodeId: string;
      kind: "textBox" | "footnote" | "endnote";
      holder: HTMLElement;
    }[] = [],
    seen = new Set<string>();
  for (let index = 0; index < roots.length; index++) {
    const nodes = new Map(
      [...roots[index]!.querySelectorAll<HTMLElement>("[data-node-id]")].map((n) => [
        n.dataset.nodeId!,
        n,
      ]),
    );
    const selected = (id: string): boolean => {
      const node = nodes.get(id);
      return (
        !!node &&
        (completeNodes || index > 0
          ? complete.has(id)
          : clipboardNodeText(node) === model.getNode(id)!.text)
      );
    };
    const ownedStories: { storyId: string; kind: "textBox" | "footnote" | "endnote" }[] = [];
    for (const drawing of word.drawings ?? [])
      if ("textStoryId" in drawing && drawing.textStoryId && selected(drawing.nodeId))
        ownedStories.push({ storyId: drawing.textStoryId, kind: "textBox" });
    for (const reference of word.noteReferences ?? [])
      if (selected(reference.nodeId)) {
        const note = word.notes?.find((note) => note.noteId === reference.noteId);
        if (!note || (note.type ?? "normal") !== "normal")
          throw new DocumentError("InvalidDocument", "复制注释缺少其正常正文故事。");
        ownedStories.push({ storyId: note.storyId, kind: note.kind });
      }
    for (const owner of ownedStories) {
      if (seen.has(owner.storyId)) continue;
      const story = word.stories.find((s) => s.storyId === owner.storyId);
      if (!story || story.kind !== owner.kind)
        throw new DocumentError("InvalidDocument", "复制内容缺少其附属故事。");
      seen.add(story.storyId);
      const owned = holder.ownerDocument.createElement("div");
      owned.innerHTML = model.getNode(story.rootNodeId)!.html;
      for (const child of owned.querySelectorAll<HTMLElement>("[data-node-id]"))
        complete.add(child.dataset.nodeId!);
      roots.push(owned);
      stories.push({
        storyId: story.storyId,
        rootNodeId: story.rootNodeId,
        kind: owner.kind,
        holder: owned,
      });
    }
  }
  return { roots, stories, completeNodes: completeNodes ? complete : undefined };
}

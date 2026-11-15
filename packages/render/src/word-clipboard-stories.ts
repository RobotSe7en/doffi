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
  stories: { storyId: string; rootNodeId: string; kind: "textBox"; holder: HTMLElement }[];
  completeNodes: ReadonlySet<string> | undefined;
} {
  const word = model.snapshot().metadata.word!,
    complete = new Set(completeNodes),
    roots = [holder],
    stories: { storyId: string; rootNodeId: string; kind: "textBox"; holder: HTMLElement }[] = [],
    seen = new Set<string>();
  for (let index = 0; index < roots.length; index++) {
    const nodes = new Map(
      [...roots[index]!.querySelectorAll<HTMLElement>("[data-node-id]")].map((n) => [
        n.dataset.nodeId!,
        n,
      ]),
    );
    for (const drawing of word.drawings ?? []) {
      const node = nodes.get(drawing.nodeId);
      if (
        !node ||
        !("textStoryId" in drawing) ||
        !drawing.textStoryId ||
        seen.has(drawing.textStoryId) ||
        !(completeNodes || index > 0
          ? complete.has(drawing.nodeId)
          : clipboardNodeText(node) === model.getNode(drawing.nodeId)!.text)
      )
        continue;
      const story = word.stories.find((s) => s.storyId === drawing.textStoryId);
      if (!story || story.kind !== "textBox")
        throw new DocumentError("InvalidDocument", "复制绘图缺少其文字故事。");
      seen.add(story.storyId);
      const owned = holder.ownerDocument.createElement("div");
      owned.innerHTML = model.getNode(story.rootNodeId)!.html;
      for (const child of owned.querySelectorAll<HTMLElement>("[data-node-id]"))
        complete.add(child.dataset.nodeId!);
      roots.push(owned);
      stories.push({
        storyId: story.storyId,
        rootNodeId: story.rootNodeId,
        kind: "textBox",
        holder: owned,
      });
    }
  }
  return { roots, stories, completeNodes: completeNodes ? complete : undefined };
}

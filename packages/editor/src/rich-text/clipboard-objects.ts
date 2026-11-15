import type { ContentDocumentModel } from "@doffi/core";

/** Build a selected object forest from authored source. Coordinate ancestors
 * are empty carriers; unselected siblings and their stories are excluded. */
export function copyClipboardObjects(
  model: ContentDocumentModel,
  selectedIds: readonly string[],
  document: Document,
): { holder: HTMLElement; complete: ReadonlySet<string>; contextId: string } | undefined {
  if (!selectedIds.length) return;
  const source = new DOMParser().parseFromString(model.snapshot().html, "text/html"),
    byId = new Map(
      [...source.querySelectorAll<HTMLElement>("[data-node-id]")].map((n) => [
        n.dataset.nodeId!,
        n,
      ]),
    ),
    selected = new Set(selectedIds),
    complete = new Set<string>(),
    keep = new Set<string>(),
    roots = new Set<HTMLElement>(),
    holder = document.createElement("div");
  for (const id of selected) {
    const node = byId.get(id);
    if (!node) continue;
    for (const child of [node, ...node.querySelectorAll<HTMLElement>("[data-node-id]")]) {
      complete.add(child.dataset.nodeId!);
      keep.add(child.dataset.nodeId!);
    }
    let root = node;
    for (let parent = node.parentElement; parent; parent = parent.parentElement) {
      if (parent.hasAttribute("data-word-drawing")) root = parent;
      if (parent.matches("p,h1,h2,h3,h4,h5,h6")) break;
    }
    for (let parent: HTMLElement | null = node; parent; parent = parent.parentElement) {
      if (parent.dataset.nodeId) keep.add(parent.dataset.nodeId);
      if (parent === root) break;
    }
    roots.add(root);
  }
  const ordered = [...roots].sort((a, b) =>
    a === b ? 0 : a.compareDocumentPosition(b) & 4 ? -1 : 1,
  );
  for (const original of ordered) {
    const clone = original.cloneNode(true) as HTMLElement;
    for (const node of [...clone.querySelectorAll<HTMLElement>("[data-node-id]")].reverse())
      if (!keep.has(node.dataset.nodeId!)) node.remove();
    holder.append(clone);
  }
  const first = byId.get(selectedIds[0]!);
  return first
    ? {
        holder,
        complete,
        contextId:
          first.closest<HTMLElement>("p,h1,h2,h3,h4,h5,h6")?.dataset.nodeId ?? selectedIds[0]!,
      }
    : undefined;
}

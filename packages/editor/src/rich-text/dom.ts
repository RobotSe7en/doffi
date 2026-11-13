import type { ChangeMapping,ReconciledTextEdit } from "@doffi/core";

export type IgnoreNode = (node: Node) => boolean;
export const blockSelector = "p,h1,h2,h3,h4,h5,h6,li,td,th";
export function sourceText(node: Node, ignored: IgnoreNode): string {
  if (ignored(node)) return "";
  if (node.nodeType === 3) return node.textContent ?? "";
  if (node.nodeType === 1) {
    const element = node as Element;
    if (element.hasAttribute("data-word-tab")) return "\t";
    if (element.localName === "br") return "\n";
    if (element.localName === "img") return "\ufffc";
  }
  return [...node.childNodes].map((child) => sourceText(child, ignored)).join("");
}

/** Canonical UTF-16 position, excluding presentation children and tab leaders. */
export function offsetAt(root: Node, target: Node, offset: number, ignored: IgnoreNode): number {
  let total = 0, found = false;
  const visit = (node: Node): void => {
    if (found || ignored(node)) return;
    if (node === target) {
      total += node.nodeType === 3 ? Math.min(offset, node.textContent!.length)
        : [...node.childNodes].slice(0, offset).reduce((sum, child) => sum + sourceText(child, ignored).length, 0);
      found = true; return;
    }
    if (node.nodeType === 3 || node.nodeType === 1 && (node as Element).matches("br,img,[data-word-tab]")) {
      total += sourceText(node, ignored).length; return;
    }
    for (const child of node.childNodes) visit(child);
  };
  visit(root); return total;
}

export interface InputStep {
  readonly from: number;
  readonly to: number;
  readonly nodes: ReadonlyMap<string, { readonly text: string; readonly start: number }>;
}
function sourceIndex(root: HTMLElement, ignored: IgnoreNode): Map<string, { text: string; start: number }> {
  const result = new Map<string, { text: string; start: number }>();
  let position = 0;
  const visit = (node: Node): string => {
    if (ignored(node)) return "";
    const id = node.nodeType === 1 ? (node as HTMLElement).dataset.nodeId : undefined;
    const entry = { text: "", start: position };
    if (id && !result.has(id)) result.set(id, entry);
    if (node.nodeType === 3 || node.nodeType === 1 && (node as Element).matches("br,img,[data-word-tab]")) {
      entry.text = sourceText(node, ignored); position += entry.text.length;
    } else entry.text = [...node.childNodes].map(visit).join("");
    return entry.text;
  };
  visit(root); return result;
}
export function inputStep(root: HTMLElement, range: Range, ignored: IgnoreNode): InputStep {
  return { nodes: sourceIndex(root, ignored), from: offsetAt(root, range.startContainer, range.startOffset, ignored), to: offsetAt(root, range.endContainer, range.endOffset, ignored) };
}

/** Evidence uses the actual target range. Ambiguous structural moves receive no guessed proof. */
export function inputProofs(root: HTMLElement, step: InputStep, ignored: IgnoreNode): ReconciledTextEdit[] {
  const result: ReconciledTextEdit[] = [];
  for (const [nodeId, next] of sourceIndex(root, ignored)) {
    const previous = step.nodes.get(nodeId); if (!previous) continue;
    const value = next.text, old = previous.text;
    if (old === value) continue;
    const from = Math.max(0, Math.min(old.length, step.from - previous.start));
    const to = Math.max(from, Math.min(old.length, step.to - previous.start));
    const length = value.length - old.length + to - from;
    if (length < 0 || value.slice(0, from) !== old.slice(0, from) || value.slice(from + length) !== old.slice(to)) continue;
    result.push({ nodeId, from, to, expectedText: old.slice(from, to), text: value.slice(from, from + length) });
  }
  return result;
}

export interface CaretBookmark { readonly nodeId?: string; readonly path: readonly number[]; readonly offset: number; readonly affinity?: "forward" }
export interface SelectionBookmark { readonly anchor: CaretBookmark; readonly focus: CaretBookmark }
export function mapSelectionBookmark(value:SelectionBookmark|undefined,mapping:ChangeMapping):SelectionBookmark|undefined {
  if(!value)return;
  const point=(entry:CaretBookmark):CaretBookmark=>{
    const {nodeId,...rest}=entry;let offset=entry.offset;
    for(const edit of mapping.textEdits)if(edit.nodeId===nodeId){if(offset>=edit.to)offset+=edit.insertedLength-edit.to+edit.from;else if(offset>=edit.from)offset=edit.from+edit.insertedLength;}
    const target=nodeId&&mapping.nodes.has(nodeId)?mapping.nodes.get(nodeId):nodeId;
    return{...rest,offset,...(target?{nodeId:target}:{})};
  };
  return{anchor:point(value.anchor),focus:point(value.focus)};
}
export function bookmark(root: HTMLElement, ignored: IgnoreNode): SelectionBookmark | undefined {
  const selected = root.ownerDocument.getSelection();
  if (!selected?.anchorNode || !selected.focusNode || !root.contains(selected.anchorNode) || !root.contains(selected.focusNode)) return;
  const point = (node: Node, offset: number): CaretBookmark => {
    const parent = node.nodeType === 1 ? node as Element : node.parentElement;
    const block = parent?.closest<HTMLElement>(blockSelector), target = block && root.contains(block) ? block : root;
    const path: number[] = [];
    for (let child = target; child !== root && child.parentElement; child = child.parentElement)
      path.unshift([...child.parentElement.children].filter((item) => !ignored(item)).indexOf(child));
    return { ...(target.dataset.nodeId ? { nodeId: target.dataset.nodeId } : {}), path, offset: offsetAt(target, node, offset, ignored), ...(node.nodeType === 3 && offset === 0 ? { affinity: "forward" as const } : {}) };
  };
  return { anchor: point(selected.anchorNode, selected.anchorOffset), focus: point(selected.focusNode, selected.focusOffset) };
}
export function caretAt(root: Node, requested: number, ignored: IgnoreNode, affinity?: "forward"): [Node, number] {
  let remaining = requested, fallback: [Node, number] | undefined, empty: [Node, number] | undefined;
  const visit = (node: Node): [Node, number] | undefined => {
    if (ignored(node)) return;
    if (node.nodeType === 3) {
      if (remaining < node.textContent!.length || remaining === node.textContent!.length && affinity !== "forward") return [node, remaining];
      if (remaining === node.textContent!.length) fallback = [node, remaining];
      remaining -= node.textContent!.length; return;
    }
    if (node.nodeType === 1 && (node as Element).matches("br,img,[data-word-tab]")) {
      if (remaining === 0) return [node.parentNode!, [...node.parentNode!.childNodes].indexOf(node as ChildNode)];
      remaining--; return;
    }
    for (const child of node.childNodes) { const found = visit(child); if (found) return found; }
    // Keep an empty authored run's character formatting. Derived line fillers
    // must not place the restored caret after its paragraph's trailing break.
    if (remaining === 0 && node.nodeType === 1 && (node as Element).matches("span,strong,em,b,i,u,s,sub,sup,a,ins,del") && ![...node.childNodes].some(child => !ignored(child))) empty = [node, 0];
    return;
  };
  const point = visit(root); if (point || fallback || empty) return point ?? fallback ?? empty!;
  let end = root.childNodes.length; while (end && ignored(root.childNodes[end - 1]!)) end--;
  return [root, end];
}
export function restoreBookmark(root: HTMLElement, value: SelectionBookmark | undefined, ignored: IgnoreNode): void {
  if (!value) return;
  const point = (entry: CaretBookmark): [Node, number] => {
    let target = entry.nodeId === root.dataset.nodeId ? root
      : entry.nodeId ? root.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(entry.nodeId) + '"]') : null;
    if (!target) {
      target = root;
      for (const ordinal of entry.path) target = [...target.children].filter((item) => !ignored(item))[ordinal] as HTMLElement ?? target;
    }
    const text=sourceText(target,ignored);let offset=Math.max(0,Math.min(entry.offset,text.length));
    while(offset>0){const before=text.charCodeAt(offset-1),after=text.charCodeAt(offset),next=text.codePointAt(offset);if(!(before>=0xd800&&before<=0xdbff&&after>=0xdc00&&after<=0xdfff)&&!(next!==undefined&&(next>=0xfe00&&next<=0xfe0f||next>=0xe0100&&next<=0xe01ef)))break;offset--;}
    return caretAt(target, offset, ignored, entry.affinity);
  };
  const selected = root.ownerDocument.getSelection()!;
  selected.setBaseAndExtent(...point(value.anchor), ...point(value.focus));
}

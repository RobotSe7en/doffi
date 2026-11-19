import { DocumentError } from "@doffi/core";
import {
  viewNodeRole,
  markViewNode,
  isMeasuredViewContent,
  measureViewContent,
  hiddenFromViewSelection,
  hideFromViewSelection,
} from "./view-nodes.js";
/** DOM cloning does not copy weak ownership. Repeated stories need the same
 * source/derived distinction for geometry and canonical text offsets. */
export function cloneViewTree<T extends Node>(source: T, maximum: number): T {
  const clone = source.cloneNode(true) as T,
    queue: [Node, Node][] = [[source, clone]];
  let count = 0;
  for (let index = 0; index < queue.length; index++) {
    if (++count > maximum)
      throw new DocumentError(
        "ResourceLimitExceeded",
        "Repeated view content exceeds the rendering node budget.",
      );
    const [from, to] = queue[index]!,
      role = viewNodeRole(from);
    if (role) markViewNode(to, role);
    if (isMeasuredViewContent(from)) measureViewContent(to);
    if (hiddenFromViewSelection(from)) hideFromViewSelection(to);
    for (let child = 0; child < from.childNodes.length; child++)
      queue.push([from.childNodes[child]!, to.childNodes[child]!]);
  }
  return clone;
}

import { markViewNode } from "./view-nodes.js";

/** Each Office role is a separate inheritance layer. Point and label overrides
 * therefore keep precedence even when a series is edited later in source CSS. */
export function chartStyleProbe(scope: HTMLElement, classes: string | readonly string[] | undefined, account: (amount: number) => void, probes: HTMLElement[]): CSSStyleDeclaration {
  let parent = scope;
  const layers = typeof classes === "string" ? [classes] : classes?.length ? classes : [""];
  for (const name of layers) {
    account(1); const probe = markViewNode(scope.ownerDocument.createElement("span")); probe.className = name;
    probe.style.cssText = "position:absolute;visibility:hidden"; parent.append(probe); probes.push(probe); parent = probe;
  }
  return scope.ownerDocument.defaultView!.getComputedStyle(parent);
}

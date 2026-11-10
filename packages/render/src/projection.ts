import { hiddenFromViewSelection, ownedViewAttribute, ownedViewStyle, viewNodeRole, viewStyleSource } from "./view-nodes.js";
/** Reversible projection: layout, font fallback, assets and labels never become source. */
export class SourceProjection {
  private readonly originals = new WeakMap<Element, Element>();
  private readonly rendered = new WeakMap<Element, Element>();
  private readonly derived = new WeakSet<Node>();
  private readonly wrappers = new WeakSet<Node>();
  private readonly identities = new Map<string, { original: Element; rendered: Element }>();
  private readonly fields = new WeakMap<Element, { original: Element; display: Element }>();
  private readonly editableRoots=new WeakSet<Node>();
  readonly ignored = (node: Node): boolean => {
    for (let item: Node | null = node; item; item = item.parentNode) {
      if(this.editableRoots.has(item))return false;
      if (this.derived.has(item) || viewNodeRole(item) === "remove") return true;
    }
    return false;
  };

  /** An explicit editing command may replace a projected property (notably
   * a substituted font). Keep that intent until the same DOM batch serializes. */
  readonly setStyles = (element: HTMLElement, properties: ReadonlyMap<string, string>): boolean => {
    if (this.ignored(element) || this.wrappers.has(element) || viewNodeRole(element) === "unwrap") return false;
    const id = element.dataset.nodeId, identity = id ? this.identities.get(id) : undefined;
    const original = this.originals.get(element) ?? identity?.original, authored = original?.cloneNode(true) as HTMLElement | undefined;
    for (const [property, value] of properties) { element.style.setProperty(property, value); authored?.style.setProperty(property, value); }
    if (authored) {
      this.originals.set(element, authored);
      if (id && identity) this.identities.set(id, { original: authored, rendered: identity.rendered });
    }
    return true;
  };

  markDerived(node: Node): void { this.derived.add(node); }
  /** An editor may bind one live notes/comment instance to its exact source root.
   * Ownership is local to this projection; the global read-only view remains derived. */
  captureInstance(root:HTMLElement,source:HTMLElement):void {
    if(!root.dataset.nodeId||root.dataset.nodeId!==source.dataset.nodeId)throw new Error("An editable instance must match its source identity.");
    this.editableRoots.add(root);this.capture(root,source);
  }
  capture(root: HTMLElement, source: HTMLElement): void {
    const originals = new Map([source, ...source.querySelectorAll<HTMLElement>("[data-node-id]")].map((node) => [node.dataset.nodeId, node]));
    for (const node of [root, ...root.querySelectorAll<HTMLElement>("*")]) {
      if (this.ignored(node)) { this.derived.add(node); continue; }
      const original = originals.get(node.dataset.nodeId);
      if (!original) {
        // HTML parsers may introduce an identity-free tbody. It is source structure,
        // whereas numbering labels and measured grid columns are renderer children.
        if (viewNodeRole(node) === "remove") this.derived.add(node);
        if (viewNodeRole(node) === "unwrap") this.wrappers.add(node);
        continue;
      }
      const display = [...node.children].find((child) => viewNodeRole(child) === "remove" && child.classList.contains("sdk-field-display"));
      const authored = original.cloneNode(original.hasAttribute("data-word-tab") || Boolean(display)) as Element, view = node.cloneNode(false) as Element;
      if (display) this.fields.set(node, { original: authored, display });
      this.originals.set(node, authored); this.rendered.set(node, view);
      this.identities.set(node.dataset.nodeId!, { original: authored, rendered: view });
    }
  }
  serialize(root: HTMLElement): HTMLElement {
    return this.project(root) as HTMLElement;
  }
  shell(element: Element): Element | undefined {
    if (this.ignored(element) || this.wrappers.has(element)) return;
    return this.project(element, undefined, true) as Element;
  }
  /** Clone from live provenance, rather than guessing the origin of cloneContents nodes. */
  serializeRange(range: Range): DocumentFragment {
    let node = range.commonAncestorContainer;
    for (let parent = node.parentElement; parent; parent = parent.parentElement) if (this.fields.has(parent)) { node = parent; break; }
    const result = this.project(node, range), fragment = node.ownerDocument!.createDocumentFragment();
    if (result?.nodeType === 1 || result?.nodeType === 11) fragment.append(...result.childNodes);
    else if (result) fragment.append(result);
    return fragment;
  }
  private project(root: Node, range?: Range, shell = false): Node | undefined {
    const seen = new Set<string>();
    const clone = (node: Node): Node | undefined => {
      if (this.ignored(node)) return;
      if (range && hiddenFromViewSelection(node)) return;
      if (range && node !== root && !range.intersectsNode(node)) return;
      if (this.wrappers.has(node)) {
        const fragment = node.ownerDocument!.createDocumentFragment();
        for (const child of node.childNodes) { const value = clone(child); if (value) fragment.append(value); }
        return fragment;
      }
      if (node.nodeType === 3) {
        const from = range?.startContainer === node ? range.startOffset : 0, to = range?.endContainer === node ? range.endOffset : node.textContent!.length;
        return node.ownerDocument!.createTextNode(node.textContent!.slice(from, to));
      }
      if (node.nodeType !== 1) return node.cloneNode(false);
      const element = node as Element, result = element.cloneNode(false) as HTMLElement;
      // Range.extractContents clones partially selected ancestors; their attributes
      // still come from the same source identity and need the same reversal.
      const identity = this.identities.get(element.getAttribute("data-node-id") ?? "");
      const original = this.originals.get(element) ?? identity?.original, rendered = this.rendered.get(element) ?? identity?.rendered;
      if (original && rendered) {
        for (const name of new Set([...element.attributes, ...rendered.attributes, ...original.attributes].map((item) => item.name))) {
          if (name === "style") continue;
          if (!ownedViewAttribute(element, name) && element.getAttribute(name) !== rendered.getAttribute(name)) continue;
          const value = original.getAttribute(name);
          if (value === null) result.removeAttribute(name); else result.setAttribute(name, value);
        }
        const before = (original as HTMLElement).style, view = (rendered as HTMLElement).style;
        // CSSOM expands shorthands and exposes empty longhands for var(...).
        // Restoring an unchanged style as text preserves those source declarations.
        if (element.getAttribute("style") === rendered.getAttribute("style")) {
          const value = original.getAttribute("style");
          if (value === null) result.removeAttribute("style"); else result.setAttribute("style", value);
        } else {
          const current = (element as HTMLElement).style;
          const value = original.getAttribute("style");
          if (value === null) result.removeAttribute("style"); else result.setAttribute("style", value);
          for (const property of new Set([...current, ...view])) {
            if (ownedViewStyle(element, property) || before.getPropertyValue(property) !== view.getPropertyValue(property)
              || before.getPropertyPriority(property) !== view.getPropertyPriority(property)) continue;
            if (current.getPropertyValue(property) === view.getPropertyValue(property) && current.getPropertyPriority(property) === view.getPropertyPriority(property)) continue;
            const authored = current.getPropertyValue(property);
            if (authored) result.style.setProperty(property, authored, current.getPropertyPriority(property)); else result.style.removeProperty(property);
          }
          if (!result.style.length) {
            if (original.hasAttribute("style")) result.setAttribute("style", ""); else result.removeAttribute("style");
          }
        }
      }
      if (!original) {
        const source = viewStyleSource(element) as HTMLElement | undefined;
        if (source) {
          for (const property of [...result.style]) if (ownedViewStyle(element, property)) {
            const value = source.style.getPropertyValue(property);
            if (value) result.style.setProperty(property, value, source.style.getPropertyPriority(property)); else result.style.removeProperty(property);
          }
          if (!result.style.length && !source.hasAttribute("style")) result.removeAttribute("style");
        }
      }
      const id = result.getAttribute("data-node-id");
      if (id) { if (seen.has(id)) { result.removeAttribute("data-node-id"); result.removeAttribute("id"); } else seen.add(id); }
      if (shell && node === root) return result;
      const field = this.fields.get(element);
      if (field && range?.intersectsNode(field.display)) {
        const selected = node.ownerDocument!.createRange(); selected.selectNodeContents(field.display);
        if (range.compareBoundaryPoints(0, selected) > 0) selected.setStart(range.startContainer, range.startOffset);
        if (range.compareBoundaryPoints(2, selected) < 0) selected.setEnd(range.endContainer, range.endOffset);
        result.append(result.ownerDocument.createTextNode(selected.toString())); return result;
      }
      if (original && element.hasAttribute("data-word-tab")) {
        result.append(...[...original.childNodes].map((child) => child.cloneNode(true)));
      } else for (const child of node.childNodes) { const value = clone(child); if (value) result.append(value); }
      return result;
    };
    return clone(root);
  }
}

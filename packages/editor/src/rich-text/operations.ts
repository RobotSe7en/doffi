import { blockSelector, caretAt, inputStep, offsetAt, sourceText } from "./dom.js";
import type { IgnoreNode } from "./dom.js";
/** Shared native DOM operations; the source transaction remains the commit boundary. */
export function deleteNativeRange(range: Range): void {
    const blockAt = (node: Node): Element | null => (node.nodeType === 1 ? node as Element : node.parentElement)?.closest(blockSelector) ?? null;
    const first = blockAt(range.startContainer), last = blockAt(range.endContainer), join = first && last && first !== last && first.parentElement === last.parentElement && !first.matches("td,th,li") && !last.matches("td,th,li");
    range.deleteContents();
    if (join) {
        const offset = first.childNodes.length;
        first.append(...last.childNodes);
        last.remove();
        range.setStart(first, offset);
        range.collapse(true);
    }
}
export function nativeDeletionRange(root: HTMLElement, range: Range, kind: string, ignored: IgnoreNode): Range {
    const frame = root.ownerDocument, step = inputStep(root, range, ignored);
    if (step.from !== step.to)
        return range;
    const blockAt = (node: Node): HTMLElement | null => (node.nodeType === 1 ? node as Element : node.parentElement)?.closest<HTMLElement>(blockSelector) ?? null;
    const block = blockAt(range.startContainer);
    if (!block || block !== blockAt(range.endContainer))
        return range;
    const text = sourceText(block, ignored), offset = offsetAt(block, range.startContainer, range.startOffset, ignored), backward = kind.endsWith("Backward");
    if (backward && offset === 0 || !backward && offset === text.length) {
        let sibling = backward ? block.previousElementSibling : block.nextElementSibling;
        while (sibling && ignored(sibling))
            sibling = backward ? sibling.previousElementSibling : sibling.nextElementSibling;
        if (!block.matches("td,th,li") && sibling?.matches("p,h1,h2,h3,h4,h5,h6")) {
            const joined = frame.createRange(), first = backward ? sibling : block, last = backward ? block : sibling;
            joined.setStart(...caretAt(first, sourceText(first, ignored).length, ignored));
            joined.setEnd(...caretAt(last, 0, ignored, "forward"));
            return joined;
        }
        return range;
    }
    let from = offset, to = offset;
    if (kind.includes("Word")) {
        if (backward)
            from -= /((?:\S+\s*)|\s+)$/.exec(text.slice(0, offset))?.[0].length ?? 0;
        else
            to += /^(?:\s*\S+|\s+)/.exec(text.slice(offset))?.[0].length ?? 0;
    }
    else {
        const segmenter = typeof Intl.Segmenter === "function" ? new Intl.Segmenter(undefined, { granularity: "grapheme" }) : undefined;
        if (segmenter) {
            const part = segmenter.segment(text).containing(backward ? Math.max(0, offset - 1) : offset);
            if (part) {
                if (backward)
                    from = part.index;
                else
                    to = part.index + part.segment.length;
            }
        }
        else {
            const boundaries = [0];
            let position = 0, previous = "", regional = 0;
            for (const character of text) {
                const code = character.codePointAt(0)!, isRegional = code >= 0x1f1e6 && code <= 0x1f1ff;
                if (position && !/[\p{Mark}\uFE00-\uFE0F\u{E0100}-\u{E01EF}\u200D]/u.test(character) && previous !== "\u200d" && !(code >= 0x1f3fb && code <= 0x1f3ff) && !(isRegional && regional % 2 === 1))
                    boundaries.push(position);
                regional = isRegional ? regional + 1 : 0;
                previous = character;
                position += character.length;
            }
            boundaries.push(text.length);
            if (backward)
                from = boundaries.filter(value => value < offset).at(-1) ?? 0;
            else
                to = boundaries.find(value => value > offset) ?? text.length;
        }
    }
    const expanded = frame.createRange();
    expanded.setStart(...caretAt(block, from, ignored));
    expanded.setEnd(...caretAt(block, to, ignored, "forward"));
    return expanded;
}
/** CSS underline propagates through descendants: an inner `none` cannot cancel
 * the ancestor's line. Keep the ancestor identity and restore only the text
 * outside the chosen range, rather than changing unrelated characters. */
function clearSelectedUnderline(root: HTMLElement, range: Range, ignored: IgnoreNode): void {
    const document = root.ownerDocument, view = document.defaultView!, ancestors = new Set<HTMLElement>(), visited = new Set<HTMLElement>();
    const walker = document.createTreeWalker(root, 4);
    let node: Node | null;
    while ((node = walker.nextNode())) {
        if (ignored(node) || !range.intersectsNode(node)) continue;
        for (let element = node.parentElement; element && root.contains(element); element = element.parentElement)
            if (!visited.has(element)) { visited.add(element); if (view.getComputedStyle(element).textDecorationLine.includes("underline")) ancestors.add(element); }
    }
    const restore: { range: Range; color: string; style: string }[] = [];
    for (const ancestor of ancestors) {
        const computed = view.getComputedStyle(ancestor), color = computed.textDecorationColor, style = computed.textDecorationStyle;
        const texts = document.createTreeWalker(ancestor, 4);
        while ((node = texts.nextNode())) {
            if (ignored(node) || !node.textContent) continue;
            const length = node.textContent.length, outside = range.comparePoint(node, 0) > 0 || range.comparePoint(node, length) < 0;
            const from = outside ? length : range.startContainer === node ? range.startOffset : 0;
            const to = outside ? length : range.endContainer === node ? range.endOffset : length;
            if (from > 0) { const part = document.createRange(); part.setStart(node, 0); part.setEnd(node, from); restore.push({ range: part, color, style }); }
            if (!outside && to < length) { const part = document.createRange(); part.setStart(node, to); part.setEnd(node, length); restore.push({ range: part, color, style }); }
        }
        const remaining = computed.textDecorationLine.split(/\s+/).filter(value => value !== "underline" && value !== "none").join(" ");
        ancestor.style.setProperty("text-decoration-line", remaining || "none");
    }
    for (const saved of restore.reverse()) {
        if (saved.range.collapsed) continue;
        const span = document.createElement("span"); span.style.textDecorationLine = "underline"; span.style.textDecorationStyle = saved.style; span.style.textDecorationColor = saved.color;
        saved.range.surroundContents(span);
    }
}
export function styleNativeRange(root: HTMLElement, range: Range, properties: ReadonlyMap<string, string>, ignored: IgnoreNode, setStyles?: (node: HTMLElement, properties: ReadonlyMap<string, string>) => boolean): void {
    const frame = root.ownerDocument;
    if (properties.get("text-decoration") === "none" || properties.get("text-decoration-line") === "none") {
        const from = offsetAt(root, range.startContainer, range.startOffset, ignored), to = offsetAt(root, range.endContainer, range.endOffset, ignored);
        clearSelectedUnderline(root, range, ignored);
        range.setStart(...caretAt(root, from, ignored, "forward")); range.setEnd(...caretAt(root, to, ignored));
    }
    const walker = frame.createTreeWalker(root, 4), ranges: Range[] = [];
    let node: Node | null;
    while ((node = walker.nextNode())) {
        if (ignored(node.parentNode!) || !range.intersectsNode(node))
            continue;
        const part = frame.createRange();
        part.selectNodeContents(node);
        if (range.compareBoundaryPoints(0, part) > 0)
            part.setStart(range.startContainer, range.startOffset);
        if (range.compareBoundaryPoints(2, part) < 0)
            part.setEnd(range.endContainer, range.endOffset);
        if (!part.collapsed)
            ranges.push(part);
    }
    for (const part of ranges.reverse()) {
        const text = part.startContainer, parent = text.parentElement;
        if (text === part.endContainer && text.nodeType === 3 && part.startOffset === 0 && part.endOffset === text.textContent!.length && parent?.localName === "span" && parent.childNodes.length === 1) {
            if (setStyles?.(parent, properties)) continue;
            if (!setStyles && [...properties].every(([property]) => ["font-weight", "font-style", "text-decoration", "text-decoration-line"].includes(property) && !!parent.style.getPropertyValue(property))) {
                for (const [property, value] of properties) parent.style.setProperty(property, value); continue;
            }
        }
        const span = frame.createElement("span");
        for (const [property, value] of properties)
            span.style.setProperty(property, value);
        part.surroundContents(span);
    }
}

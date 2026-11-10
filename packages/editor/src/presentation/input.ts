import { executeClipboard } from "../clipboard-command.js";
import { readTextFormat } from "../text-format-state.js";
import type { TextFormatState } from "../text-format-state.js";
import { DocumentError } from "@doffi/core";
import type { DocumentEditingSession, DocumentOperation, ReconciledTextEdit } from "@doffi/core";
import { SourceProjection, sourceSelectionText, stripClipboardIdentities } from "@doffi/render";
import { blockSelector, bookmark, caretAt, inputProofs, inputStep, offsetAt, restoreBookmark, sourceText } from "../rich-text/dom.js";
import type { InputStep, SelectionBookmark } from "../rich-text/dom.js";
import { deleteNativeRange, nativeDeletionRange, styleNativeRange } from "../rich-text/operations.js";
export interface PresentationInputHost {
    readonly session: DocumentEditingSession;
    readonly root: HTMLElement;
    readonly editable: HTMLElement;
    readonly source: HTMLElement;
    readonly instance?: boolean;
    readonly newNotesSlideId?: string;
    readonly label: string;
    readonly typingStyle?: Map<string, string>;
    history(direction: "undo" | "redo"): void;
    save(selection: SelectionBookmark | undefined): void;
    refresh(selection: SelectionBookmark | undefined): void;
    compositionEnded(): void;
    error(reason: unknown, recovery?: string): void;
}
export interface PresentationInput {
    focus(selection?: SelectionBookmark): void;
    selection(): SelectionBookmark | undefined;
    composing(): boolean;
    committing(): boolean;
    format(property: string, value: string): void;
    formatState(): TextFormatState;
    clipboard(action: "copy" | "cut" | "paste"): Promise<void>;
    toggle(property: "bold" | "italic" | "underline"): void;
    paragraph(property: string, value: string): void;
    list(kind: "none" | "bullet" | "number"): void;
    finish(): void;
    dispose(): void;
}
function characterStyle(property: string, value: string): ReadonlyMap<string, string> {
    const result = new Map([[property, value]]);
    if (property === "font-family") for (const role of ["ascii", "h-ansi", "east-asia", "complex-script"]) result.set("--office-font-" + role, value);
    if (property === "--office-font-ascii")
        result.set("--office-font-h-ansi", value);
    if (property === "color")
        for (const [name, part] of Object.entries({ color: value, source: "direct", alpha: "1", "lum-mod": "1", "lum-off": "0", tint: "0", shade: "1" }))
            result.set("--office-text-" + name, part);
    return result;
}
function matrixText(text: string): string[][] {
    const values: string[][] = [[]];
    let value = "", quoted = false;
    for (let i = 0; i < text.length; i++) {
        const c = text[i]!;
        if (c === '"' && (quoted || !value)) {
            if (quoted && text[i + 1] === '"') {
                value += '"';
                i++;
            }
            else
                quoted = !quoted;
        }
        else if (!quoted && (c === "\t" || c === "\n" || c === "\r")) {
            values.at(-1)!.push(value);
            value = "";
            if (c !== "\t") {
                if (c === "\r" && text[i + 1] === "\n")
                    i++;
                values.push([]);
            }
        }
        else
            value += c;
    }
    if (quoted)
        throw new DocumentError("InvalidOperation", "表格粘贴包含未闭合的引号。");
    if (value || values.at(-1)!.length)
        values.at(-1)!.push(value);
    else if (values.length > 1)
        values.pop();
    return values;
}
function boundary(text: string, offset: number): boolean {
    const before = text.charCodeAt(offset - 1), after = text.charCodeAt(offset), next = text.codePointAt(offset);
    return !(before >= 0xd800 && before <= 0xdbff && after >= 0xdc00 && after <= 0xdfff) && !(offset > 0 && next !== undefined && (next >= 0xfe00 && next <= 0xfe0f || next >= 0xe0100 && next <= 0xe01ef));
}
/** One input scope over live first-party projection; it owns no document history. */
export function bindPresentationInput(host: PresentationInputHost): PresentationInput {
    const { session, root, editable, source } = host, frame = root.ownerDocument, controls = new AbortController(), projection = new SourceProjection(), revision = session.snapshot().manifest.revision;
    const attributes = new Map<HTMLElement, Map<string, string | null>>(), placeholders: Node[] = [];
    let disposed = false, composing = false, committing = false, dirty = false, compositionRevision: number | undefined, pending: {
        step: InputStep;
        length: number;
    } | undefined, saved: SelectionBookmark | undefined, proofs: ReconciledTextEdit[] = [], compositionInsertion: {
        from: number;
        to: number;
    } | undefined, plainPaste = false;
    const typing = host.typingStyle ?? new Map<string, string>();
    const paragraphStyles = new Map<string, Map<string, string>>();
    const nativeShells = new Map<string, HTMLElement>();
    const paragraphStyle = (node: HTMLElement, property: string, value: string): void => { node.style.setProperty(property, value); if (node.dataset.nodeId) {
        const properties = paragraphStyles.get(node.dataset.nodeId) ?? new Map<string, string>();
        properties.set(property, value);
        paragraphStyles.set(node.dataset.nodeId, properties);
    } };
    const attribute = (node: HTMLElement, name: string, value: string): void => { const old = attributes.get(node) ?? new Map<string, string | null>(); if (!old.has(name))
        old.set(name, node.getAttribute(name)); attributes.set(node, old); node.setAttribute(name, value); };
    attribute(editable, "contenteditable", "true");
    attribute(editable, "role", "textbox");
    attribute(editable, "aria-label", host.label);
    attribute(editable, "aria-multiline", "true");
    attribute(editable, "spellcheck", "false");
    for (const field of root.querySelectorAll<HTMLElement>("[data-office-field],.sdk-ppt-marker"))
        attribute(field, "contenteditable", "false");
    if (host.instance)
        projection.captureInstance(root, source);
    else
        projection.capture(root, source);
    if (!editable.querySelector(blockSelector) && !editable.matches(blockSelector)) {
        const p = frame.createElement("p");
        p.append(...editable.childNodes);
        editable.append(p);
        placeholders.push(p);
        if (!p.childNodes.length) {
            const filler = frame.createElement("br");
            projection.markDerived(filler);
            p.append(filler);
        }
    }
    const range = (): Range | undefined => { const selection = frame.getSelection(); if (!selection?.rangeCount)
        return; const value = selection.getRangeAt(0); return editable.contains(value.startContainer) && editable.contains(value.endContainer) ? value.cloneRange() : undefined; };
    const select = (node: Node, offset: number): void => { const value = frame.createRange(); value.setStart(node, offset); value.collapse(true); const selection = frame.getSelection()!; selection.removeAllRanges(); selection.addRange(value); };
    const remember = (): SelectionBookmark | undefined => { saved = bookmark(root, projection.ignored) ?? saved; host.save(saved); return saved; };
    const restore = (value: SelectionBookmark | undefined): void => {
        editable.focus();
        restoreBookmark(root, value, projection.ignored);
        const current = range();
        if (!current) {
            const p = editable.querySelector("p") ?? editable;
            select(p, 0);
        }
        else if (current.collapsed) {
            const parent = current.startContainer.nodeType === 1 ? current.startContainer as Element : current.startContainer.parentElement, field = parent?.closest<HTMLElement>("[data-office-field]");
            if (field && root.contains(field)) {
                const offset = offsetAt(field, current.startContainer, current.startOffset, projection.ignored);
                select(field.parentNode!, [...field.parentNode!.childNodes].indexOf(field) + (offset ? 1 : 0));
            }
        }
        remember();
    };
    const normalized = (): HTMLElement => {
        const result = projection.serialize(root);
        for (const [id, properties] of paragraphStyles) {
            const node = result.dataset.nodeId === id ? result : result.querySelector<HTMLElement>('[data-node-id="' + CSS.escape(id) + '"]');
            if (node)
                for (const [property, value] of properties)
                    node.style.setProperty(property, value);
        }
        const containers = result.querySelector("table") ? [...result.querySelectorAll<HTMLElement>("td,th")] : [result];
        for (const container of containers)
            if ([...container.childNodes].some(node => node.nodeType === 3 && node.textContent!.trim()) && !container.querySelector(blockSelector)) {
                const shell = nativeShells.get(container.dataset.nodeId ?? ""), p = shell ? shell.cloneNode(false) as HTMLElement : frame.createElement("p");
                p.append(...container.childNodes);
                container.append(p);
            }
        return result;
    };
    const commit = (operations: DocumentOperation[], baseRevision = revision): void => {
        if (baseRevision !== session.snapshot().manifest.revision)
            throw new DocumentError("RevisionConflict", "输入期间文档已变化，请在当前修订重新输入。");
        const candidate = session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision, operations });
        committing = true;
        try {
            session.commit(candidate);
        }
        catch (error) {
            session.discard(candidate);
            throw error;
        }
        finally {
            committing = false;
        }
    };
    const flush = (): void => {
        if (disposed || composing || !dirty)
            return;
        const selection = remember(), text = sourceText(root, projection.ignored), recovery = compositionInsertion ? text.slice(compositionInsertion.from, compositionInsertion.to) : undefined;
        try {
            const value = normalized(), operations: DocumentOperation[] = host.newNotesSlideId ? [{ op: "insertPresentationNotes", slideId: host.newNotesSlideId, html: value.innerHTML }] : [{ op: "reconcileFragment", nodeId: root.dataset.nodeId!, html: value.innerHTML, textEdits: proofs }];
            dirty = false;
            commit(operations, compositionRevision ?? revision);
            if (!disposed)
                host.refresh(selection);
        }
        catch (error) {
            dirty = false;
            host.refresh(selection);
            host.error(error, recovery);
        }
        compositionRevision = undefined;
        proofs = [];
        pending = undefined;
    };
    const deleteRange = (value: Range): void => {
        // Firefox selects the editable wrapper for Ctrl+A. Normalize its edges to
        // paragraph contents so replacing all text retains the first paragraph ID.
        const edge = (node: Node, offset: number, start: boolean): readonly [
            Node,
            number
        ] => {
            if (node.nodeType !== 1)
                return [node, offset];
            const element = node as HTMLElement;
            if (element.matches("p,h1,h2,h3,h4,h5,h6,li") || offset !== (start ? 0 : node.childNodes.length))
                return [node, offset];
            const blocks = [...element.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6,li")].filter(p => !projection.ignored(p)), target = start ? blocks[0] : blocks.at(-1);
            return target ? [target, start ? 0 : target.childNodes.length] : [node, offset];
        };
        const collapsed = value.collapsed, start = edge(value.startContainer, value.startOffset, collapsed ? value.startOffset === 0 : true), end = edge(value.endContainer, value.endOffset, false);
        value.setStart(...start);
        if (collapsed)
            value.collapse(true);
        else
            value.setEnd(...end);
        const cell = (node: Node) => (node.nodeType === 1 ? node as Element : node.parentElement)?.closest("td,th"), first = cell(value.startContainer), last = cell(value.endContainer);
        if (first && last && first !== last && first.closest("table") === last.closest("table")) {
            const cells = [...root.querySelectorAll("td,th")], a = cells.indexOf(first), b = cells.indexOf(last), caret = value.cloneRange();
            caret.collapse(true);
            for (const target of cells.slice(a, b + 1)) {
                const part = frame.createRange();
                part.selectNodeContents(target);
                if (target.contains(value.startContainer))
                    part.setStart(value.startContainer, value.startOffset);
                if (target.contains(value.endContainer))
                    part.setEnd(value.endContainer, value.endOffset);
                part.deleteContents();
            }
            value.setStart(caret.startContainer, caret.startOffset);
            value.collapse(true);
        }
        else
            deleteNativeRange(value);
    };
    const mutate = (value: Range, action: () => void, split = false): void => {
        if (disposed || composing)
            return;
        const step = inputStep(root, value, projection.ignored), text = sourceText(root, projection.ignored);
        if (!boundary(text, step.from) || !boundary(text, step.to)) {
            host.error(new DocumentError("InvalidOperation", "选择不能拆开 UTF-16 字符或变体序列。"));
            return;
        }
        const parent = value.startContainer.nodeType === 1 ? value.startContainer as Element : value.startContainer.parentElement, block = parent?.closest(blockSelector), moved = split && block && !block.matches("td,th,li") ? { ...step, to: Math.max(step.to, offsetAt(root, block, block.childNodes.length, projection.ignored)) } : undefined;
        try {
            action();
            const actual = inputProofs(root, step, projection.ignored), known = new Set(actual.map(edit => edit.nodeId));
            proofs.push(...actual, ...(moved ? inputProofs(root, moved, projection.ignored).filter(edit => !known.has(edit.nodeId)) : []));
            dirty = true;
            flush();
        }
        catch (error) {
            host.refresh(saved);
            host.error(error);
        }
    };
    const insert = (value: Range, node: Node): void => { deleteRange(value); value.insertNode(node); select(node.parentNode!, [...node.parentNode!.childNodes].indexOf(node as ChildNode) + 1); };
    const insertText = (value: Range, text: string): void => { let node: Node = frame.createTextNode(text); if (typing.size) {
        const span = frame.createElement("span");
        for (const [property, value] of typing)
            span.style.setProperty(property, value);
        span.append(node);
        node = span;
    } insert(value, node); };
    const paragraphBreak = (value: Range): void => {
        deleteRange(value);
        const parent = value.startContainer.nodeType === 1 ? value.startContainer as Element : value.startContainer.parentElement, block = parent?.closest<HTMLElement>(blockSelector);
        if (!block || block === root || block.matches("td,th,li")) {
            insert(value, frame.createElement("br"));
            return;
        }
        const tail = frame.createRange();
        tail.setStart(value.startContainer, value.startOffset);
        tail.setEnd(block, block.childNodes.length);
        const authored = projection.serialize(block), next = frame.createElement(block.localName);
        for (const attr of authored.attributes)
            if (!["id", "data-node-id"].includes(attr.name))
                next.setAttribute(attr.name, attr.value);
        next.append(tail.extractContents());
        block.after(next);
        select(next, 0);
    };
    let typingAnchor: string | undefined;
    const format = (property: string, value: string): void => {
        if (disposed || composing)
            return;
        restore(remember());
        const selected = range();
        if (!selected)
            return;
        const properties = characterStyle(property, value);
        if (selected.collapsed) {
            for (const [name, value] of properties)
                typing.set(name, value);
            typingAnchor = JSON.stringify(bookmark(root, projection.ignored));
            return;
        }
        const selection = remember();
        mutate(selected, () => { styleNativeRange(root, selected, properties, projection.ignored, projection.setStyles); restoreBookmark(root, selection, projection.ignored); });
    };
    const formatState = (): TextFormatState => disposed || composing ? { enabled: false } : { ...readTextFormat(root, range(), typing, projection.ignored), editingText: true };
    const toggle = (property: "bold" | "italic" | "underline"): void => {
        const active = formatState()[property] === true;
        format(property === "bold" ? "font-weight" : property === "italic" ? "font-style" : "text-decoration", property === "bold" ? active ? "400" : "700" : property === "italic" ? active ? "normal" : "italic" : active ? "none" : "underline");
    };
    const paragraphs = (): HTMLElement[] => { const selected = range(); if (!selected)
        return []; const current = selected.startContainer.nodeType === 1 ? selected.startContainer as Element : selected.startContainer.parentElement; return [...root.querySelectorAll<HTMLElement>("p,h1,h2,h3,h4,h5,h6,li")].filter(node => !projection.ignored(node) && (selected.collapsed ? node === current?.closest("p,h1,h2,h3,h4,h5,h6,li") : selected.intersectsNode(node))); };
    const paragraph = (property: string, value: string): void => { if (disposed || composing)
        return; restore(remember()); const selected = range(); if (selected)
        mutate(selected, () => { for (const p of paragraphs())
            paragraphStyle(p, property, value); }); };
    const list = (kind: "none" | "bullet" | "number"): void => { if (disposed || composing)
        return; restore(remember()); const selected = range(); if (!selected)
        return; mutate(selected, () => { for (const p of paragraphs()) {
        p.removeAttribute("data-office-bullet");
        p.removeAttribute("data-office-numbering");
        p.removeAttribute("data-office-list-none");
        if (kind === "none")
            p.dataset.officeListNone = "true";
        else if (kind === "bullet")
            p.dataset.officeBullet = "•";
        else
            p.dataset.officeNumbering = "arabicPeriod";
        paragraphStyle(p, "margin-left", kind === "none" ? "0pt" : "18pt");
        paragraphStyle(p, "text-indent", kind === "none" ? "0pt" : "-12pt");
    } }); };
    const beforeinput = (event: InputEvent): void => {
        let selected = range();
        const target = event.getTargetRanges?.()[0];
        if (target) {
            const value = frame.createRange();
            value.setStart(target.startContainer, target.startOffset);
            value.setEnd(target.endContainer, target.endOffset);
            if (editable.contains(value.startContainer) && editable.contains(value.endContainer))
                selected = value;
        }
        if (!selected) {
            event.preventDefault();
            return;
        }
        if (composing || event.isComposing || event.inputType === "insertCompositionText") {
            // Noncancelable native composition may replace a whole wrapper with a
            // text node (Firefox). Preserve the proven first paragraph shell only in
            // the serialized candidate; never reparent the live IME target.
            const parent = selected.startContainer.nodeType === 1 ? selected.startContainer as Element : selected.startContainer.parentElement, container = parent?.closest<HTMLElement>("td,th") ?? editable, first = container.querySelector<HTMLElement>("p,h1,h2,h3,h4,h5,h6,li"), id = container.dataset.nodeId ?? (container === editable ? root.dataset.nodeId : undefined);
            if (id && first?.dataset.nodeId && offsetAt(container, selected.startContainer, selected.startOffset, projection.ignored) === 0 && offsetAt(container, selected.endContainer, selected.endOffset, projection.ignored) === sourceText(container, projection.ignored).length && !nativeShells.has(id))
                nativeShells.set(id, projection.serialize(first));
            pending = { step: inputStep(root, selected, projection.ignored), length: sourceText(root, projection.ignored).length };
            return;
        }
        event.preventDefault();
        if (event.inputType === "historyUndo" || event.inputType === "historyRedo") {
            try {
                host.history(event.inputType === "historyUndo" ? "undo" : "redo");
            }
            catch (error) {
                host.error(error);
            }
            return;
        }
        if (["insertText", "insertReplacementText"].includes(event.inputType))
            mutate(selected, () => insertText(selected!, event.data ?? ""));
        else if (event.inputType.startsWith("delete")) {
            selected = nativeDeletionRange(root, selected, event.inputType, projection.ignored);
            mutate(selected, () => { deleteRange(selected!); select(selected!.startContainer, selected!.startOffset); });
        }
        else if (["insertParagraph", "insertLineBreak"].includes(event.inputType))
            mutate(selected, () => event.inputType === "insertParagraph" ? paragraphBreak(selected!) : insert(selected!, frame.createElement("br")), event.inputType === "insertParagraph");
        else
            host.error(new DocumentError("UnsupportedFeature", "当前不支持此输入操作：" + event.inputType));
    };
    const paste = (event: ClipboardEvent): void => {
        event.preventDefault();
        event.stopPropagation();
        if (!event.clipboardData || composing)
            return;
        try {
            const html = plainPaste ? "" : event.clipboardData.getData("text/html"), text = event.clipboardData.getData("text/plain");
            plainPaste = false;
            if (html.length > 1024 * 1024 || text.length > 1024 * 1024)
                throw new DocumentError("ResourceLimitExceeded", "粘贴内容超过大小限制。");
            const selected = range();
            if (!selected)
                return;
            const parent = selected.startContainer.nodeType === 1 ? selected.startContainer as Element : selected.startContainer.parentElement, cell = parent?.closest<HTMLTableCellElement>("td,th"), table = cell?.closest("table");
            if (table && root.contains(table) && /\t|\r|\n/.test(text)) {
                const occupied: HTMLTableCellElement[][] = Array.from(table.rows, () => []);
                let row = -1, column = -1;
                for (const [r, tr] of [...table.rows].entries())
                    for (const td of tr.cells) {
                        let c = 0;
                        while (occupied[r]![c])
                            c++;
                        for (let y = r; y < r + (td.rowSpan || table.rows.length - r); y++)
                            for (let x = c; x < c + td.colSpan; x++) {
                                if (!occupied[y])
                                    throw new DocumentError("InvalidOperation", "表格网格已失效。");
                                occupied[y]![x] = td;
                            }
                        if (td === cell) {
                            row = r;
                            column = c;
                        }
                    }
                const values = matrixText(text);
                if (!values.length)
                    return;
                remember();
                commit([{ op: "editPresentationTable", nodeId: root.dataset.nodeId!, change: { kind: "setCells", row, column, values } }]);
                if (!disposed)
                    host.refresh(saved);
                return;
            }
            if (!html && !text)
                return;
            const template = frame.createElement("template");
            if (html)
                template.innerHTML = html;
            else
                text.replace(/\r\n?/g, "\n").split("\n").forEach((line, index) => { if (index)
                    template.content.append(frame.createElement("br")); template.content.append(frame.createTextNode(line)); });
            if (template.content.querySelector("img,table,article,section,[data-office-kind],[data-word-field],[data-word-link]"))
                throw new DocumentError("UnsupportedFeature", "此内容不能作为幻灯片富文本粘贴，请使用纯文本或复制对象。");
            for (const node of template.content.querySelectorAll("[data-node-id],[id],[data-office-field-id]")) {
                node.removeAttribute("data-node-id");
                node.removeAttribute("id");
                node.removeAttribute("data-office-field-id");
            }
            const block = [...template.content.children].some(node => /^(p|div|h[1-6]|ul|ol|li)$/.test(node.localName));
            const operations: DocumentOperation[] = host.newNotesSlideId ? [{ op: "insertPresentationNotes", slideId: host.newNotesSlideId, html: block ? template.innerHTML : "<p>" + template.innerHTML + "</p>" }] : [{ op: "insertFragment", nodeId: cell?.dataset.nodeId ?? root.dataset.nodeId!, position: "end", html: block ? template.innerHTML : "<p>" + template.innerHTML + "</p>" }];
            const validation = session.prepare({ protocolVersion: 1, documentId: session.snapshot().manifest.documentId, baseRevision: revision, operations });
            session.discard(validation);
            mutate(selected, () => { if (block) {
                paragraphBreak(selected);
                const current = range()!, parent = current.startContainer.nodeType === 1 ? current.startContainer as Element : current.startContainer.parentElement, p = parent?.closest(blockSelector), last = template.content.lastChild;
                if (!p || !last)
                    throw new DocumentError("UnsupportedFeature", "当前位置不支持区块粘贴。");
                p.before(template.content);
                select(last, last.childNodes.length);
            }
            else {
                const last = template.content.lastChild;
                deleteRange(selected);
                selected.insertNode(template.content);
                if (last)
                    select(last.parentNode!, [...last.parentNode!.childNodes].indexOf(last) + 1);
            } });
        }
        catch (error) {
            host.error(error);
        }
    };
    const copy = (event: ClipboardEvent, cut: boolean): void => {
        if (!event.clipboardData || composing)
            return;
        const selected = range();
        if (!selected || selected.collapsed)
            return;
        event.preventDefault();
        event.stopPropagation();
        const holder = frame.createElement("div");
        holder.append(projection.serializeRange(selected));
        const p = session.snapshot().metadata.presentation!;
        for (const link of p.hyperlinks) {
            const span = holder.querySelector('[data-node-id="' + CSS.escape(link.nodeId) + '"]');
            if (span) {
                const a = frame.createElement("a");
                a.setAttribute("href", link.target ?? "#" + link.slideId);
                a.append(...span.childNodes);
                span.replaceWith(a);
            }
        }
        stripClipboardIdentities(holder);
        for (const field of holder.querySelectorAll("[data-office-field-id]"))
            field.removeAttribute("data-office-field-id");
        event.clipboardData.setData("text/html", holder.innerHTML);
        event.clipboardData.setData("text/plain", sourceSelectionText(holder));
        if (cut)
            mutate(selected, () => { deleteRange(selected); select(selected.startContainer, selected.startOffset); });
    };
    const keydown = (event: KeyboardEvent): void => {
        if (composing)
            return;
        const key = event.key.toLowerCase(), command = event.ctrlKey || event.metaKey;
        if (command && key === "v" && event.shiftKey) {
            plainPaste = true;
            return;
        }
        if (command && ["z", "y", "b", "i", "u"].includes(key)) {
            event.preventDefault();
            event.stopPropagation();
            try {
                if (key === "z" || key === "y") {
                    remember();
                    host.history(key === "y" || event.shiftKey ? "redo" : "undo");
                }
                else
                    toggle(key === "b" ? "bold" : key === "i" ? "italic" : "underline");
            }
            catch (error) {
                host.error(error);
            }
            return;
        }
        const selected = range();
        if (!selected)
            return;
        if (command && key === "a") {
            const parent = selected.startContainer.nodeType === 1 ? selected.startContainer as Element : selected.startContainer.parentElement, cell = parent?.closest("td,th");
            if (cell) {
                event.preventDefault();
                const value = frame.createRange();
                value.selectNodeContents(cell);
                const selection = frame.getSelection()!;
                selection.removeAllRanges();
                selection.addRange(value);
            }
            return;
        }
        if (event.key === "Enter") {
            event.preventDefault();
            mutate(selected, () => event.shiftKey ? insert(selected, frame.createElement("br")) : paragraphBreak(selected), !event.shiftKey);
        }
        if (event.key === "Tab") {
            event.preventDefault();
            const parent = selected.startContainer.nodeType === 1 ? selected.startContainer as Element : selected.startContainer.parentElement, cell = parent?.closest("td,th"), cells = [...root.querySelectorAll("td,th")];
            if (cell) {
                const next = cells[cells.indexOf(cell) + (event.shiftKey ? -1 : 1)];
                if (next)
                    select(...caretAt(next, 0, projection.ignored));
            }
            else
                mutate(selected, () => insertText(selected, "\t"));
        }
    };
    editable.addEventListener("beforeinput", beforeinput as EventListener, { signal: controls.signal });
    editable.addEventListener("input", () => {
        if (disposed)
            return;
        if (pending) {
            const edits = inputProofs(root, pending.step, projection.ignored);
            proofs.push(...edits);
            const edit = edits.find(edit => edit.nodeId === root.dataset.nodeId), length = sourceText(root, projection.ignored).length - pending.length + pending.step.to - pending.step.from;
            compositionInsertion = edit ? { from: edit.from, to: edit.from + edit.text.length } : { from: pending.step.from, to: pending.step.from + Math.max(0, length) };
        }
        pending = undefined;
        dirty = true;
        if (!composing)
            flush();
    }, { signal: controls.signal });
    editable.addEventListener("keydown", keydown, { signal: controls.signal });
    editable.addEventListener("paste", paste as EventListener, { signal: controls.signal });
    editable.addEventListener("copy", ((event: ClipboardEvent) => copy(event, false)) as EventListener, { signal: controls.signal });
    editable.addEventListener("cut", ((event: ClipboardEvent) => copy(event, true)) as EventListener, { signal: controls.signal });
    editable.addEventListener("drop", event => { event.preventDefault(); host.error(new DocumentError("UnsupportedFeature", "请通过粘贴或对象工具插入内容。")); }, { signal: controls.signal });
    editable.addEventListener("compositionstart", () => { composing = true; compositionRevision = revision; host.save(saved); }, { signal: controls.signal });
    editable.addEventListener("compositionend", () => queueMicrotask(() => { if (disposed)
        return; composing = false; if (compositionInsertion && typing.size && revision === session.snapshot().manifest.revision) {
        const selection = remember(), value = frame.createRange();
        value.setStart(...caretAt(root, compositionInsertion.from, projection.ignored));
        value.setEnd(...caretAt(root, compositionInsertion.to, projection.ignored));
        styleNativeRange(root, value, typing, projection.ignored, projection.setStyles);
        restoreBookmark(root, selection, projection.ignored);
    } flush(); host.compositionEnded(); }), { signal: controls.signal });
    frame.addEventListener("selectionchange", () => { if (range()) { const next = bookmark(root, projection.ignored); if (!composing && typingAnchor && next && JSON.stringify(next) !== typingAnchor) { typing.clear(); typingAnchor = undefined; } remember(); } }, { signal: controls.signal });
    return { focus: restore, selection: remember, composing: () => composing, committing: () => committing, format, formatState, toggle,
        clipboard(action) { const expected = revision, selected = JSON.stringify(remember()); return executeClipboard(root, () => restore(saved), action, () => !disposed && !composing && revision === expected && JSON.stringify(remember()) === selected); }, paragraph, list,
        finish() { if (disposed)
            return; composing = false; flush(); },
        dispose() { if (disposed)
            return; disposed = true; controls.abort(); for (const [node, values] of attributes)
            for (const [name, value] of values) {
                if (value === null)
                    node.removeAttribute(name);
                else
                    node.setAttribute(name, value);
            } if (!dirty)
            for (const node of placeholders) {
                for (const child of [...node.childNodes])
                    if (!projection.ignored(child))
                        node.parentNode?.insertBefore(child, node);
                node.parentNode?.removeChild(node);
            } proofs = []; pending = undefined; saved = undefined; if (!host.typingStyle)
            typing.clear(); paragraphStyles.clear(); nativeShells.clear(); } };
}

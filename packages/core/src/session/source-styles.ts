import { clone, generate, ident, parse, walk } from "css-tree";
import type { CssNode, Declaration, ListItem, Rule } from "css-tree";
import type { DocumentReadLimits } from "../runtime.js";
import { normalizeCss, setInlineStyle } from "../document/css.js";
import { encode, readText } from "../document/package.js";
import { invalid } from "../document/validation.js";
/** Candidate-local CSS writes use the package's existing, bounded stylesheets. */
export class SourceStyleEdit {
    constructor(private readonly files: Map<string, Uint8Array>, private readonly paths: readonly string[], private readonly limits: DocumentReadLimits) { }
    validateProperty(name: string, value: string): void { setInlineStyle("", name, value, this.limits); }
    snapshot(): {
        path: string;
        css: string;
    }[] { return this.paths.map(path => ({ path, css: normalizeCss(readText(this.files, path), "stylesheet", this.limits, { nodes: 0 }) })); }
    private append(css: string): void {
        const path = this.paths.at(-1);
        if (!path)
            invalid("Source style changes require a declared stylesheet.");
        this.files.set(path, encode(normalizeCss(readText(this.files, path) + css, "stylesheet", this.limits, { nodes: 0 })));
    }
    addClass(name: string, properties: Readonly<Record<string, string>>): void {
        this.append("." + ident.encode(name) + "{}");
        this.setClass(name, properties);
    }
    /** An alias stays in each original rule, retaining its cascade order and
     * specificity. It does not flatten inheritance or change other users. */
    aliasClass(source: string, target: string): void {
        for (const sheet of this.snapshot()) {
            const ast = parse(sheet.css, { context: "stylesheet" }); let dirty = false;
            if (ast.type !== "StyleSheet") invalid("Invalid source stylesheet.");
            for (const rule of ast.children) if (rule.type === "Rule" && rule.prelude.type === "SelectorList") {
                const aliases: CssNode[] = [];
                for (const selector of rule.prelude.children) {
                    if (selector.type !== "Selector") invalid("Invalid source class selector.");
                    let compound = 0; const matches = new Set<number>();
                    for (const node of selector.children) {
                        if (node.type === "Combinator") compound++;
                        else if (node.type === "ClassSelector" && ident.decode(node.name) === source) matches.add(compound);
                    }
                    // Only the formatted cell changes its selected class.
                    // The same token on ancestors/descendants remains source;
                    // repeated occurrences in one compound name one element.
                    for (const selected of matches) {
                        const alias = clone(selector); let compound = 0;
                        if (alias.type !== "Selector") invalid("Invalid cloned source class selector.");
                        for (const node of alias.children) {
                            if (node.type === "Combinator") compound++;
                            else if (compound === selected && node.type === "ClassSelector" && ident.decode(node.name) === source) node.name = ident.encode(target);
                        }
                        aliases.push(alias);
                    }
                }
                for (const alias of aliases) { rule.prelude.children.appendData(alias); dirty = true; }
            }
            if (dirty) this.files.set(sheet.path, encode(normalizeCss(generate(ast), "stylesheet", this.limits, { nodes: 0 })));
        }
    }
    /** A local cell edit must win over authored class/data selectors, including
     * important declarations. Generated selectors use one type component, so
     * repeated edits can retain the same specificity instead of growing it. */
    overrideLocalClass(name: string, properties: Readonly<Record<string, string>>, fontProperties: Readonly<Record<string, string>>, tag: "td" | "span"): void {
        let maximum = 0, tagsAtMaximum = 0;
        for (const sheet of this.snapshot()) walk(parse(sheet.css, { context: "stylesheet" }), node => {
            if (node.type !== "Selector") return;
            let classes = 0, tags = 0;
            for (const item of node.children) {
                if (["ClassSelector", "AttributeSelector", "PseudoClassSelector"].includes(item.type)) classes++;
                if (item.type === "TypeSelector") tags++;
            }
            if (classes > maximum) { maximum = classes; tagsAtMaximum = tags; }
            else if (classes === maximum) tagsAtMaximum = Math.max(tagsAtMaximum, tags);
        });
        const selector = ("." + ident.encode(name)).repeat(Math.max(1, maximum + Number(tagsAtMaximum > 1)));
        const declarations = (input: Readonly<Record<string, string>>): string => {
            let css = ""; for (const [property, value] of Object.entries(input)) css = setInlineStyle(css, property, value, this.limits);
            const ast = parse(css, { context: "declarationList" });
            if (ast.type !== "DeclarationList") invalid("Invalid cell appearance declarations.");
            for (const node of ast.children) if (node.type === "Declaration") node.important = true;
            return generate(ast);
        };
        this.append(tag + selector + "{" + declarations(properties) + "}" + (Object.keys(fontProperties).length ? selector + " span{" + declarations(fontProperties) + "}" : ""));
    }
    setClass(name: string, properties: Readonly<Record<string, string>>): void {
        const sheets = this.snapshot().map(({ path, css }) => { const ast = parse(css, { context: "stylesheet", parseCustomProperty: true }); if (ast.type !== "StyleSheet")
            invalid("Invalid source stylesheet."); return { path, ast, dirty: false }; });
        type Entry = {
            sheet: typeof sheets[number];
            rule: Rule;
            item: ListItem<CssNode>;
        };
        const matches = (node: CssNode): boolean => node.type === "Selector" && node.children.size === 1 && node.children.first?.type === "ClassSelector" && ident.decode(node.children.first.name) === name;
        const detach = (entry: Entry): Rule => {
            const rule = entry.rule;
            if (rule.prelude.type !== "SelectorList")
                invalid("Invalid source class rule.");
            if (rule.prelude.children.toArray().every(matches))
                return rule;
            const target = clone(rule);
            if (target.type !== "Rule" || target.prelude.type !== "SelectorList")
                invalid("Invalid source class rule.");
            rule.prelude.children.forEach((selector, item, list) => { if (matches(selector))
                list.remove(item); });
            target.prelude.children.forEach((selector, item, list) => { if (!matches(selector))
                list.remove(item); });
            if (entry.item.next)
                entry.sheet.ast.children.insertData(target, entry.item.next);
            else
                entry.sheet.ast.children.appendData(target);
            return target;
        };
        for (const [property, value] of Object.entries(properties)) {
            const parsed = parse(setInlineStyle("", property, value, this.limits), { context: "declarationList", parseCustomProperty: true });
            if (parsed.type !== "DeclarationList" || parsed.children.first?.type !== "Declaration")
                invalid("Invalid class property.");
            const incoming = parsed.children.first;
            let destination: Entry | undefined, winner: (Entry & {
                declaration: Declaration;
                index: number;
            }) | undefined;
            for (const sheet of sheets)
                sheet.ast.children.forEach((rule, item) => { if (rule.type !== "Rule" || rule.prelude.type !== "SelectorList" || !rule.prelude.children.toArray().some(matches))
                    return; const entry = { sheet, rule, item }; destination = entry; rule.block.children.toArray().forEach((node, index) => { if (node.type !== "Declaration" || ident.decode(node.property) !== property)
                    return; if (!winner || node.important || !winner.declaration.important)
                    winner = { ...entry, declaration: node, index }; }); });
            if (winner && generate(winner.declaration.value) === generate(incoming.value) && (!incoming.important || winner.declaration.important))
                continue;
            let target: Rule;
            if (winner || destination) {
                const entry = winner ?? destination!;
                target = detach(entry);
                entry.sheet.dirty = true;
            }
            else {
                const sheet = sheets.at(-1);
                if (!sheet)
                    invalid("Source style changes require a declared stylesheet.");
                const ast = parse("." + ident.encode(name) + "{}", { context: "stylesheet" });
                if (ast.type !== "StyleSheet" || ast.children.first?.type !== "Rule")
                    invalid("Invalid source class.");
                target = ast.children.first;
                sheet.ast.children.appendData(target);
                sheet.dirty = true;
            }
            if (winner) {
                const declaration = target.block.children.toArray()[winner.index];
                if (declaration?.type !== "Declaration")
                    invalid("Invalid class property mapping.");
                const value = clone(incoming.value);
                if (value.type !== "Raw" && value.type !== "Value")
                    invalid("Invalid class value mapping.");
                declaration.value = value;
                declaration.important = declaration.important || incoming.important;
            }
            else
                target.block.children.appendData(clone(incoming));
        }
        for (const sheet of sheets)
            if (sheet.dirty)
                this.files.set(sheet.path, encode(normalizeCss(generate(sheet.ast), "stylesheet", this.limits, { nodes: 0 })));
    }
    cloneClass(source: string, target: string): void {
        let css = "";
        for (const sheet of this.snapshot()) {
            const ast = parse(sheet.css, { context: "stylesheet" });
            if (ast.type !== "StyleSheet")
                invalid("Invalid source stylesheet.");
            ast.children.forEach(rule => {
                if (rule.type !== "Rule")
                    return;
                let found = false;
                walk(rule.prelude, node => {
                    if (node.type === "ClassSelector" && ident.decode(node.name) === source)
                        found = true;
                });
                if (!found)
                    return;
                const next = clone(rule);
                if (next.type !== "Rule" || next.prelude.type !== "SelectorList")
                    return;
                next.prelude.children.forEach((selector, item, list) => {
                    let contains = false, ambiguous = false;
                    walk(selector, node => {
                        if (node.type === "ClassSelector" && ident.decode(node.name) === source)
                            contains = true;
                        if (node.type === "PseudoClassSelector")
                            ambiguous = true;
                    });
                    if (!contains)
                        list.remove(item);
                    else if (ambiguous)
                        invalid("Cloning a style class in pseudo-class selectors requires an explicit source mapping.");
                });
                walk(next.prelude, node => {
                    if (node.type === "ClassSelector" && ident.decode(node.name) === source)
                        node.name = ident.encode(target);
                });
                css += generate(next);
            });
        }
        this.append(css || "." + ident.encode(target) + "{}");
    }
}

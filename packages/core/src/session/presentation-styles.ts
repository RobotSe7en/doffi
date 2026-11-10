import { ident, parse, walk } from "css-tree";
import type { PresentationEditContext } from "./presentation-edit.js";
export { SourceStyleEdit as PresentationStyleEdit } from "./source-styles.js";
export function allocatePresentationClass(ctx: PresentationEditContext, prefix: string): string {
    const classes = new Set<string>(), html = ctx.parse();
    for (const id of html.nodeIds())
        for (const name of (html.peekNode(id)!.attributes.class ?? "").split(/\s+/))
            classes.add(name);
    for (const sheet of ctx.styles.snapshot())
        walk(parse(sheet.css, { context: "stylesheet" }), node => {
            if (node.type === "ClassSelector")
                classes.add(ident.decode(node.name));
        });
    let value: string;
    do {
        value = prefix + ctx.allocate();
    } while (classes.has(value));
    return value;
}

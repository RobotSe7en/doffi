import { fontName } from "../export/values.js";
import { children } from "../export/source.js";
import { attributes, declaration, empty } from "../export/xml.js";
import { colorXml, fillXml } from "./appearance.js";
import type { NativeSurface, PresentationExportContext } from "./context.js";
import { paragraphProperties } from "./text.js";
import { rgba, unsupported } from "./values.js";
const palette = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
const mapping = ["bg1", "tx1", "bg2", "tx2", ...palette.slice(4)];
export function colorMap(ctx: PresentationExportContext, surface: NativeSurface, master = false): string {
    const raw = ctx.styles.raw(surface.node, surface), defaults: Record<string, string> = { bg1: "lt1", tx1: "dk1", bg2: "lt2", tx2: "dk2" }, attrs: Record<string, string> = {};
    let authored = false;
    for (const role of mapping) {
        const value = raw["--office-theme-" + role], match = value && /^var\(\s*--office-palette-([a-zA-Z\d]+)\s*\)$/.exec(value);
        if (match) {
            attrs[role] = match[1]!;
            authored = true;
        }
        else
            attrs[role] = defaults[role] ?? role;
    }
    return master ? empty("p:clrMap", attrs) : authored ? "<p:clrMapOvr>" + empty("a:overrideClrMapping", attrs) + "</p:clrMapOvr>" : "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>";
}
export function background(ctx: PresentationExportContext, surface: NativeSurface): string {
    const css = ctx.styles.get(surface.node, surface), own = ctx.styles.own(surface.node, surface), raw = ctx.styles.raw(surface.node, surface);
    const picture = children(surface.node).find(node => node.attrs["data-office-background-image"] === "true");
    if (picture)
        return "<p:bg><p:bgPr><a:blipFill>" + empty("a:blip", { "r:embed": ctx.image(picture, surface.part) }) +
            "<a:stretch><a:fillRect/></a:stretch></a:blipFill></p:bgPr></p:bg>";
    if (!Object.keys(own).some(key => key.startsWith("--office-fill-") || key === "background-color"))
        return "";
    const reference = css["--office-fill-reference-prefix"] ?? Array.from({ length: Math.min(1024, Number(css["--office-fill-stop-count"] ?? 0)) }, (_, index) => css["--office-fill-stop-" + index + "-reference-prefix"]).find(Boolean);
    if (reference) {
        const index = ctx.format(surface.themeId, "background", fillXml(css, "--office-fill", raw, true));
        return "<p:bg><p:bgRef" + attributes({ idx: index }) + ">" + colorXml(css, reference, raw) + "</p:bgRef></p:bg>";
    }
    const paint = css["--office-fill-kind"] ? css : css["background-color"] !== "transparent" ? { ...css, "--office-fill-kind": "solid", "--office-fill-color": css["background-color"]! } : css;
    return "<p:bg><p:bgPr>" + fillXml(paint, "--office-fill", raw) + "<a:effectLst/></p:bgPr></p:bg>";
}
export function textStyle(ctx: PresentationExportContext, surface: NativeSurface, className: string | undefined, tag: string): string {
    let result = "<" + tag + ">";
    for (let level = 0; level < 9; level++) {
        const { css, own } = ctx.styles.textDefaults(className, level, surface);
        result += "<a:lvl" + (level + 1) + "pPr" + paragraphProperties("pptx", css, own, own) + "</a:lvl" + (level + 1) + "pPr>";
    }
    return result + "</" + tag + ">";
}
export function writeThemes(ctx: PresentationExportContext): void {
    for (const theme of ctx.p.themes) {
        const css = ctx.styles.theme(theme.themeId), part = ctx.opc.part(ctx.themePart(theme.themeId), "application/vnd.openxmlformats-officedocument.theme+xml");
        part.add(declaration + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements><a:clrScheme name="Office">');
        for (const role of palette) {
            const value = css["--office-palette-" + role] ?? css["--office-theme-" + role];
            if (!value)
                unsupported("An exported theme palette is incomplete.");
            const rgb = rgba(value);
            if (rgb.alpha !== 1)
                unsupported("Native theme base colors must be opaque RGB values.");
            part.add("<a:" + role + ">" + empty("a:srgbClr", { val: rgb.rgb }) + "</a:" + role + ">");
        }
        part.add('</a:clrScheme><a:fontScheme name="Office">');
        for (const kind of ["major", "minor"]) {
            part.add("<a:" + kind + "Font>");
            for (const [tag, role] of [["latin", "ascii"], ["ea", "east-asia"], ["cs", "complex-script"]])
                part.add(empty("a:" + tag, { typeface: fontName(css["--office-" + kind + "-" + role] ?? css["--office-" + kind + "-ascii"] ?? '"sans-serif"') }));
            for (const [key, value] of Object.entries(css)) {
                const script = new RegExp("^--office-" + kind + "-script-([A-Za-z]{4})$").exec(key)?.[1];
                if (script)
                    part.add(empty("a:font", { script, typeface: fontName(value) }));
            }
            part.add("</a:" + kind + "Font>");
        }
        part.add('</a:fontScheme><a:fmtScheme name="Office">');
        for (const [kind, tag, fallback] of [["fill", "fillStyleLst", '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'], ["line", "lnStyleLst", '<a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>'], ["effect", "effectStyleLst", "<a:effectStyle><a:effectLst/></a:effectStyle>"], ["background", "bgFillStyleLst", '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>']] as const) {
            part.add("<a:" + tag + ">");
            const values = [...ctx.formatValues(theme.themeId, kind)];
            while (values.length < 3)
                values.push(fallback);
            values.forEach(value => part.add(value));
            part.add("</a:" + tag + ">");
        }
        part.add('</a:fmtScheme></a:themeElements></a:theme>');
        part.finish();
    }
}

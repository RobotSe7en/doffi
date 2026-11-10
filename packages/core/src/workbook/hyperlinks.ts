import type { WorkbookHyperlink } from "../contracts/workbook.js";
import { invalid } from "../document/validation.js";
import { parseFormula } from "./formula.js";

export function validateWorkbookHyperlink(link: Pick<WorkbookHyperlink, "url" | "location">): void {
  if (!link.url && !link.location) invalid("A workbook hyperlink requires a target.");
  if (link.url !== undefined && (!/^(https?:\/\/|mailto:)/i.test(link.url) || /[\u0000-\u0020\u007f]/.test(link.url))) invalid("A workbook hyperlink has an unsupported target.");
  if (link.location !== undefined) {
    const target = parseFormula("=" + link.location, { row: 0, column: 0 });
    if (!["name", "cell", "range"].includes(target.kind) && !(target.kind === "error" && target.value === "#REF!")) invalid("A workbook internal hyperlink must reference a name or cell range, or retain its deleted target.");
  }
}

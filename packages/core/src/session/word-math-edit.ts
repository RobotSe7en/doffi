import type { WordMetadata } from "../contracts/word.js";
import type { WordMathOperation } from "../contracts/word-math.js";
import type { DocumentEdit } from "../document/edit.js";
import type { ParsedHtml } from "../document/html.js";
import { DocumentError } from "../errors.js";
import { validateWordMath } from "../word-math.js";
export function isWordMathOperation(operation: {
  readonly op: string;
}): operation is WordMathOperation {
  return ["insertWordEquation", "setWordEquation", "removeWordEquation"].includes(operation.op);
}
export function applyWordMath(
  word: WordMetadata,
  edit: DocumentEdit,
  parse: () => ParsedHtml,
  operation: WordMathOperation,
): void {
  if (operation.op !== "removeWordEquation") validateWordMath(operation.math);
  if (
    operation.op !== "removeWordEquation" &&
    operation.math.nodes.find((node) => node.mathId === operation.math.rootMathId)?.kind ===
      "equationGroup" &&
    (operation.display ??
      (operation.op === "setWordEquation"
        ? word.equations?.find((item) => item.nodeId === operation.nodeId)?.display
        : "inline")) !== "block"
  )
    throw new DocumentError(
      "InvalidOperation",
      "A group of equations requires a display math paragraph.",
    );
  const node = parse().peekNode(operation.nodeId);
  if (!node) throw new DocumentError("InvalidOperation", "The equation source is missing.");
  const equations = (word.equations ??= []);
  if (operation.op === "insertWordEquation") {
    if (!/^p$|^h[1-6]$/.test(node.tagName))
      throw new DocumentError("InvalidOperation", "Insert an equation in a source paragraph.");
    const id = edit.insertSourceAt(
      operation.nodeId,
      operation.offset,
      '<span data-word-equation="true">\ufffc</span>',
    ).rootNodeIds[0]!;
    equations.push({
      nodeId: id,
      display: operation.display ?? "inline",
      ...operation.math,
      ...(operation.alignment ? { alignment: operation.alignment } : {}),
      ...(operation.fontFamily ? { fontFamily: operation.fontFamily } : {}),
    });
    return;
  }
  const equation = equations.find((item) => item.nodeId === operation.nodeId);
  if (!equation) throw new DocumentError("InvalidOperation", "The equation definition is missing.");
  if (operation.op === "removeWordEquation") {
    edit.removeSource(operation.nodeId);
    word.equations = equations.filter((item) => item !== equation);
    return;
  }
  equation.nodes = operation.math.nodes;
  equation.rootMathId = operation.math.rootMathId;
  if (operation.display !== undefined) equation.display = operation.display;
  if (operation.alignment !== undefined) equation.alignment = operation.alignment;
  if (operation.fontFamily !== undefined) equation.fontFamily = operation.fontFamily;
}

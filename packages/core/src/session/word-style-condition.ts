import type { WordMetadata } from "../contracts/word.js";
import type { WordOperation } from "../contracts/word-operations.js";
import type { SourceStyleEdit } from "./source-styles.js";
import { DocumentError } from "../errors.js";
type Operation = Extract<WordOperation, { op: "setWordStyle" }>;
type Style = WordMetadata["styles"][number];
/** Component CSS belongs to a definition; table contexts remain disposable. */
export function editWordTableDefinition(
  word: WordMetadata,
  definition: Style,
  previous: Style | undefined,
  styles: SourceStyleEdit,
  operation: Operation,
): void {
  const scopes = [
    ["table", "tableProperties", "removeTableProperties"],
    ["cell", "cellProperties", "removeCellProperties"],
    ["row", "rowProperties", "removeRowProperties"],
  ] as const;
  if (definition.kind !== "table") {
    if (
      operation.tableCondition ||
      scopes.some(([, patch, remove]) => operation[patch] || operation[remove])
    )
      throw new DocumentError("InvalidOperation", "Table style properties require a table style.");
    return;
  }
  const resolver = styles.resolver(),
    entries: { name: string; properties: Record<string, string> }[] = [];
  const values = (
    name: string | undefined,
    patch: Readonly<Record<string, string>> | undefined,
    remove: readonly string[] | undefined,
  ) => {
    const result = { ...(name ? resolver.specifiedClass(name) : {}) };
    for (const property of remove ?? [])
      for (const key of Object.keys(result))
        if (key === property || key.startsWith(property + "-")) delete result[key];
    Object.assign(result, patch);
    return result;
  };
  definition.components ??= {};
  for (const [scope, patch, remove] of scopes) {
    const name = definition.components[scope] ?? definition.cssClass + "-" + scope;
    entries.push({
      name,
      properties: values(previous?.components?.[scope], operation[patch], operation[remove]),
    });
    definition.components[scope] = name;
  }
  const patch = operation.tableCondition;
  if (patch) {
    if (patch.clear) {
      const legacy = word.nativeContent?.fragments.some(
        (fragment) =>
          fragment.styleId === definition.styleId &&
          !fragment.styleRegion &&
          (fragment.propertyPath.some((path) => path.endsWith(" tblStylePr")) ||
            fragment.propertyName?.endsWith(" tblStylePr")),
      );
      if (legacy)
        throw new DocumentError(
          "UnsupportedFeature",
          "Legacy native conditional properties lack a region binding and cannot be safely cleared.",
        );
      definition.conditions = definition.conditions?.filter(
        (condition) => condition.region !== patch.region,
      );
      if (word.nativeContent)
        word.nativeContent.fragments = word.nativeContent.fragments.filter(
          (fragment) =>
            fragment.styleId !== definition.styleId || fragment.styleRegion !== patch.region,
        );
    } else {
      definition.conditions ??= [];
      let condition = definition.conditions.find((condition) => condition.region === patch.region);
      if (!condition) {
        condition = { region: patch.region, cssClass: definition.cssClass + "-" + patch.region };
        definition.conditions.push(condition);
      }
      const original = previous?.conditions?.find((condition) => condition.region === patch.region);
      condition.components ??= {};
      for (const [scope, property, remove] of [
        ["p", "paragraphProperties", "removeParagraphProperties"],
        ["r", "runProperties", "removeRunProperties"],
        ...scopes,
      ] as const) {
        const name = condition.components[scope] ?? condition.cssClass + "-" + scope,
          base =
            original?.components?.[scope] ??
            (original && !original.components ? original.cssClass : undefined);
        entries.push({ name, properties: values(base, patch[property], patch[remove]) });
        condition.components[scope] = name;
      }
    }
  }
  if (entries.length) styles.replaceClasses(entries);
}

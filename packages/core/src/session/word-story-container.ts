import type { WordDocumentMetadata } from "../contracts/document.js";
import type { DocumentEdit } from "../document/edit.js";
import type { ParsedHtml } from "../document/html.js";

/** Old packages may use the article itself as their main story. Auxiliary
 * stories must be siblings; moving existing children preserves their IDs. */
export function ensureWordMainStoryContainer(
  metadata: WordDocumentMetadata,
  edit: DocumentEdit,
  parse: () => ParsedHtml,
): string {
  const source = parse(),
    article = source.nodeIds().find((id) => source.peekNode(id)!.parentNodeId === null)!,
    main = metadata.word!.stories.find((story) => story.kind === "main")!;
  if (main.rootNodeId === article) {
    const children = [...source.peekNode(article)!.childNodeIds],
      wrapper = edit.insertSource(article, '<div data-office-story="main"></div>').rootNodeIds[0]!;
    edit.moveSource(children, wrapper);
    edit.attributes(article, { "data-office-story": undefined });
    main.rootNodeId = wrapper;
    relocateWordSectionRoot(metadata, article, wrapper);
  }
  return article;
}

/** Section history and residual properties follow the semantic section when
 * a legacy shared story/section container is split into separate containers. */
export function relocateWordSectionRoot(
  metadata: WordDocumentMetadata,
  before: string,
  after: string,
): void {
  if (!metadata.sections.some((section) => section.rootNodeId === before)) return;
  for (const section of metadata.sections)
    if (section.rootNodeId === before) section.rootNodeId = after;
  for (const revision of metadata.word!.revisions)
    if (revision.nodeId === before && revision.property === "section") revision.nodeId = after;
  const name = "http://schemas.openxmlformats.org/wordprocessingml/2006/main sectPr";
  for (const fragment of metadata.word!.nativeContent?.fragments ?? [])
    if (
      fragment.nodeId === before &&
      fragment.kind === "property" &&
      (fragment.ownerName === "sectPr" ||
        fragment.propertyName === name ||
        fragment.propertyPath.includes(name))
    )
      fragment.nodeId = after;
}

import type { DocumentInteractionRequest, EditProposal } from "../contracts/interaction.js";
import { ContentOperationNameSchema } from "../contracts/operations.js";
import type { DocumentState } from "../document/state.js";
import { encode } from "../document/package.js";
import { checkLimit, freeze } from "../document/validation.js";
import { DocumentError } from "../errors.js";
import { createOpaqueId } from "../identifiers.js";
import { parseContract } from "../schema-registry.js";
import type { InteractionRequestInput } from "../runtime.js";
import { annotationNodes } from "./anchors.js";
import { expandWordContextRoots, wordInteractionContext } from "./word-interaction.js";

export const DEFAULT_INTERACTION_LIMITS = Object.freeze({
  maxBytes: 4 * 1024 * 1024,
  maxNodes: 10_000,
  maxAssets: 256,
});

function invalid(message: string): never {
  throw new DocumentError("InvalidProposal", message);
}
function unique(values: readonly string[]): boolean {
  return new Set(values).size === values.length;
}
function equal(left: unknown, right: unknown): boolean {
  return JSON.stringify(left) === JSON.stringify(right);
}

export function buildInteractionRequest(
  state: DocumentState,
  input: InteractionRequestInput,
): DocumentInteractionRequest {
  if (
    !input ||
    Object.keys(input).some(
      (key) => !["annotationIds", "supportedOperations", "limits"].includes(key),
    ) ||
    !Array.isArray(input.annotationIds) ||
    !input.annotationIds.length ||
    !unique(input.annotationIds)
  ) {
    throw new DocumentError(
      "InvalidOperation",
      "Select a nonempty, unique list of open annotations.",
    );
  }
  const limits = { ...DEFAULT_INTERACTION_LIMITS, ...input.limits };
  for (const [key, value] of Object.entries(limits)) {
    if (
      !Object.hasOwn(DEFAULT_INTERACTION_LIMITS, key) ||
      !Number.isSafeInteger(value) ||
      value <= 0
    )
      throw new DocumentError(
        "InvalidOperation",
        "Context limits must be positive safe integers with supported names.",
      );
  }
  const currentAnnotations = new Map(
    state.snapshot.annotations.map((item) => [item.annotationId, item]),
  );
  const annotations = input.annotationIds.map((id) => {
    const annotation = currentAnnotations.get(id);
    if (
      !annotation ||
      annotation.status !== "open" ||
      annotation.baseRevision !== state.snapshot.manifest.revision
    )
      throw new DocumentError(
        "AnchorNotFound",
        "Only open annotations at the current revision can be submitted.",
      );
    return annotation;
  });
  const roots = new Set<string>();
  const inline = new Set(["span", "strong", "em", "u", "s", "sub", "sup", "br", "ins", "del"]);
  for (const annotation of annotations) {
    const anchor = annotation.anchor;
    const ids =
      anchor.kind === "region" && anchor.nodeIds.length
        ? anchor.nodeIds
        : annotationNodes(annotation);
    for (const id of [
      ...ids,
      ...[annotation, ...(annotation.replies ?? [])].flatMap((message) =>
        message.bodyNodeId ? [message.bodyNodeId] : [],
      ),
    ]) {
      let node = state.html.peekNode(id);
      if (!node)
        throw new DocumentError("AnchorNotFound", "An annotation context node is missing.");
      while (inline.has(node.tagName) && node.parentNodeId)
        node = state.html.peekNode(node.parentNodeId)!;
      roots.add(node.nodeId);
    }
  }
  const contextNodes = expandWordContextRoots(state, roots, limits.maxNodes);
  const ids = state.html.nodeIds();
  const selected = ids.filter((id) => {
    if (!roots.has(id)) return false;
    for (
      let parent = state.html.peekNode(id)!.parentNodeId;
      parent;
      parent = state.html.peekNode(parent)!.parentNodeId
    )
      if (roots.has(parent)) return false;
    return true;
  });
  let nodes = 0;
  for (const id of selected) {
    const stack = [id];
    while (stack.length) {
      checkLimit(++nodes, limits.maxNodes);
      stack.push(...state.html.peekNode(stack.pop()!)!.childNodeIds);
    }
  }
  const projected = wordInteractionContext(state, contextNodes, annotations);
  checkLimit(projected.assets.length, limits.maxAssets);
  const word = projected.metadata.word;
  let semanticNodes = contextNodes.size;
  for (const equation of word?.equations ?? []) semanticNodes += equation.nodes.length;
  for (const drawing of word?.drawings ?? [])
    if (drawing.kind === "chart")
      semanticNodes +=
        drawing.chart.categories.length +
        drawing.chart.series.reduce(
          (count, series) => count + 1 + series.values.length + (series.xValues?.length ?? 0),
          0,
        );
  checkLimit(semanticNodes, limits.maxNodes);
  const supportedOperations = input.supportedOperations ?? ContentOperationNameSchema.options;
  if (
    !Array.isArray(supportedOperations) ||
    !unique(supportedOperations) ||
    supportedOperations.some((name) => !ContentOperationNameSchema.safeParse(name).success)
  )
    throw new DocumentError(
      "InvalidOperation",
      "Supported operations must be a unique list of implemented content operations.",
    );
  const request = parseContract("interaction-request", {
    protocolVersion: 1,
    requestId: createOpaqueId(),
    documentId: state.snapshot.manifest.documentId,
    baseRevision: state.snapshot.manifest.revision,
    annotations,
    context: {
      formatVersion: 1,
      kind: "document",
      blocks: selected.map((nodeId) => ({ nodeId, html: state.html.getNode(nodeId)!.html })),
      styles: state.snapshot.styles,
      ...projected,
    },
    supportedOperations,
  });
  checkLimit(encode(JSON.stringify(request)).byteLength, limits.maxBytes);
  return freeze(request);
}

/** The host supplies request bytes; IDs themselves never confer scope or authority. */
export function validateProposal(
  state: DocumentState,
  request: DocumentInteractionRequest,
  proposal: EditProposal,
): void {
  const { documentId, revision } = state.snapshot.manifest;
  if (
    request.documentId !== documentId ||
    proposal.documentId !== documentId ||
    proposal.requestId !== request.requestId
  )
    invalid("The proposal does not match this document and request.");
  if (request.baseRevision !== revision || proposal.baseRevision !== revision)
    throw new DocumentError(
      "RevisionConflict",
      "The request or proposal is based on a different document revision.",
    );
  if (
    !request.annotations.length ||
    !unique(request.annotations.map((item) => item.annotationId)) ||
    !unique(proposal.annotationIds) ||
    !unique(request.supportedOperations)
  )
    invalid("Request and proposal lists must have unique identities.");
  const annotationIds = new Set(request.annotations.map((item) => item.annotationId));
  const currentAnnotations = new Map(
    state.snapshot.annotations.map((item) => [item.annotationId, item]),
  );
  for (const annotation of request.annotations) {
    if (!equal(annotation, currentAnnotations.get(annotation.annotationId)))
      invalid("The request contains an annotation that differs from the current snapshot.");
  }
  if (proposal.annotationIds.some((id) => !annotationIds.has(id)))
    invalid("The proposal claims annotations outside its request.");
  const scope = new Set<string>();
  const blockIds = request.context.blocks.map((block) => block.nodeId);
  if (!unique(blockIds)) invalid("Request context contains duplicate blocks.");
  const roots = new Set(blockIds);
  for (const block of request.context.blocks) {
    const node = state.html.getNode(block.nodeId);
    if (!node || node.html !== block.html)
      invalid("A request block is not a complete subtree of the current source.");
    for (let parent = node.parentNodeId; parent; parent = state.html.peekNode(parent)!.parentNodeId)
      if (roots.has(parent)) invalid("Request context blocks overlap.");
    const stack = [block.nodeId];
    while (stack.length) {
      const id = stack.pop()!;
      scope.add(id);
      stack.push(...state.html.peekNode(id)!.childNodeIds);
    }
  }
  const closure = expandWordContextRoots(state, new Set(blockIds), state.html.nodeIds().length);
  if ([...closure].some((id) => !scope.has(id)))
    invalid("Request context omits an owned Word story or semantic dependency.");
  const projected = wordInteractionContext(state, scope, request.annotations);
  if (
    !equal(request.context.metadata, projected.metadata) ||
    !equal(request.context.styles, state.snapshot.styles) ||
    !equal(request.context.assets, projected.assets)
  )
    invalid("Request context differs from the current document source.");
  for (const annotation of request.annotations) {
    const anchor = annotation.anchor;
    const targets =
      anchor.kind === "region" && anchor.nodeIds.length
        ? anchor.nodeIds
        : annotationNodes(annotation);
    if (
      [
        ...targets,
        ...[annotation, ...(annotation.replies ?? [])].flatMap((message) =>
          message.bodyNodeId ? [message.bodyNodeId] : [],
        ),
      ].some((id) => !scope.has(id))
    )
      invalid("Request context does not cover every selected annotation.");
  }
  for (const operation of proposal.operations) {
    if (!request.supportedOperations.includes(operation.op))
      invalid("The proposal uses an operation not declared by the request.");
    if (!scope.has(operation.nodeId))
      invalid("A proposal operation targets a node outside the request context.");
    if ("annotationId" in operation && !annotationIds.has(operation.annotationId))
      invalid("A Word proposal cannot edit a comment outside its selected threads.");
    if (
      (operation.op === "reconcileFragment" || operation.op === "reconcileWordTracked") &&
      operation.textEdits.some((edit) => !scope.has(edit.nodeId))
    )
      invalid("Reconciled text edits must belong to the request context.");
    if (
      operation.op === "setWordSection" &&
      [
        ...Object.values(operation.changes.headers ?? {}),
        ...Object.values(operation.changes.footers ?? {}),
      ].some((id) => {
        const story = state.snapshot.metadata.word?.stories.find((story) => story.storyId === id);
        return !story || !scope.has(story.rootNodeId);
      })
    )
      invalid("Linked Word stories must belong to the request context.");
    if (operation.op === "splitWordSection") {
      const parent = state.html.peekNode(operation.nodeId)!.parentNodeId;
      if (!parent || !scope.has(parent))
        invalid("A section split requires its complete section in the request context.");
    }
    if (operation.op === "cloneWordStory") {
      const story = state.snapshot.metadata.word?.stories.find(
        (story) => story.storyId === operation.storyId,
      );
      if (!story || !scope.has(story.rootNodeId))
        invalid("Copying a Word story requires its source in the request context.");
    }
    if (operation.op === "groupWordDrawings" && operation.nodeIds.some((id) => !scope.has(id)))
      invalid("Grouping Word drawings requires all selected objects in the request context.");
    if (
      operation.op === "setWordDrawingConnector" ||
      (operation.op === "insertWordDrawing" && operation.object.kind === "connector")
    ) {
      const connector =
        operation.op === "setWordDrawingConnector"
          ? operation.connector
          : operation.object.kind === "connector"
            ? operation.object.connector
            : undefined;
      if (
        connector &&
        [connector.startNodeId, connector.endNodeId].some((id) => id && !scope.has(id))
      )
        invalid("Word connector targets must belong to the request context.");
    }
    if (
      operation.op === "replaceAsset" &&
      !request.context.assets.some((asset) => asset.path === operation.assetPath)
    )
      invalid("The proposal references a resource outside the request context.");
  }
}

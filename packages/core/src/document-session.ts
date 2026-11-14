import {refreshWordTableStyles,wordTableStyleTargets} from './session/word-table-styles.js';
import {trackWordText}from'./session/word-revisions.js';
import{pruneWordNativeResources,pruneRemovedWordImages}from'./session/word-native-resources.js';
import type { Anchor,Annotation } from "./contracts/annotations.js";
import { isPresentationCommentPoint } from "./document/presentation.js";
import type { ContentDocumentMetadata } from "./contracts/document.js";
import type { AnnotationOperation, DocumentTransaction } from "./contracts/operations.js";
import type { CandidateChange, DocumentSession, ContentDocumentModel, DocumentModel, DocumentSessionOptions, WorkbookModel, WorkbookSession } from "./runtime.js";
import { DocumentError } from "./errors.js";
import { parseContract } from "./schema-registry.js";
import { DocumentEdit } from "./document/edit.js";
import { parseHtml } from "./document/html.js";
import type { ParsedHtml } from "./document/html.js";
import { copyPackage, encode } from "./document/package.js";
import { buildDocumentState } from "./document/state.js";
import type { DocumentState } from "./document/state.js";
import { checkLimit, readLimits } from "./document/validation.js";
import { compareStates, retainedBytes } from "./session/changes.js";
import { checkTransactionSize, sessionLimits } from "./session/limits.js";
import { annotationNodes, mapAnnotations } from "./session/anchors.js";
import { buildInteractionRequest, validateProposal } from "./session/interaction.js";
import type { DocumentInteractionRequest, EditProposal } from "./contracts/interaction.js";
import type { InteractionRequestInput,PresentationInteractionRequestInput } from "./runtime.js";
import type { PresentationInteractionRequest,PresentationEditProposal } from "./contracts/presentation-interaction.js";
import { buildPresentationRequest,validatePresentationProposal,validatePresentationCandidate } from "./session/presentation-interaction.js";
import type { TransactionPrepareOptions } from "./runtime.js";
import { checkResourceUse, stageResources } from "./session/resources.js";
import { affectedPresentationSlides, guardPresentationContent, removeDeletedCommentBodies, updatePresentationMetadata, updatePresentationRecovery } from "./session/presentation.js";
import { updatePresentationConnectors } from "./session/presentation-geometry.js";
import { applyPresentationOperation, isPresentationOperation, validatePresentationTargets } from "./session/presentation-edit.js";
import { applyWordOperation, isWordOperation } from "./session/word-edit.js";
import { pruneWordComments } from './session/word-comments.js';
import { validateWordCandidate } from './session/word-interaction.js';
import { updateWordDrawingConnectors } from './word-drawing.js';
import { PresentationStyleEdit } from "./session/presentation-styles.js";
export { DEFAULT_INTERACTION_LIMITS } from "./session/interaction.js";

export { DEFAULT_DOCUMENT_SESSION_LIMITS } from "./session/limits.js";

import { createSessionLifecycle } from "./session/lifecycle.js";
import { createWorkbookSession, isWorkbookOperation } from "./session/workbook.js";
function invalid(message: string): never { throw new DocumentError("InvalidOperation", message); }

/** Remove definitions only when their complete source object was deleted. Partial damage fails validation. */
function updateWordMetadata(metadata: ContentDocumentMetadata, before: DocumentState, after: ParsedHtml): void {
  const word = metadata.word;
  if (!word) return;
  if (word.nativeContent){word.nativeContent.fragments = word.nativeContent.fragments.filter(item => !item.nodeId || after.hasNode(item.nodeId));if(word.nativeContent.identities)word.nativeContent.identities=word.nativeContent.identities.filter(item=>after.hasNode(item.nodeId));}
  word.hyperlinks = word.hyperlinks.filter((item) => after.hasNode(item.nodeId));
  word.revisions = word.revisions.filter((item) => after.hasNode(item.nodeId));
  if(word.moves)word.moves=word.moves.filter(move=>after.hasNode(move.startNodeId)||after.hasNode(move.endNodeId));
  word.tables = word.tables.filter((item) => after.hasNode(item.nodeId));
  if(word.conditionalFormats)word.conditionalFormats=word.conditionalFormats.filter(item=>after.hasNode(item.nodeId));
  if(word.drawings)word.drawings=word.drawings.filter(item=>after.hasNode(item.nodeId));
  if(word.equations)word.equations=word.equations.filter(item=>after.hasNode(item.nodeId));
  if(word.controls)word.controls=word.controls.filter(item=>after.hasNode(item.nodeId));
  word.tabs = word.tabs.filter((item) => after.hasNode(item.paragraphNodeId));
  word.bookmarks = word.bookmarks.filter((item) => after.hasNode(item.startNodeId) || after.hasNode(item.endNodeId));
  word.fields = word.fields.filter((item) => [item.nodeId, ...(item.endNodeId ? [item.endNodeId] : []), ...(item.resultNodeIds ?? [])]
    .some((id) => after.hasNode(id)));
  for (const field of word.fields) {
    const ids = [field.nodeId, ...(field.resultNodeIds ?? [])];
    if (ids.some((id) => before.html.getNode(id)?.text !== after.getNode(id)?.text)) field.dirty = true;
  }
}

function rebase(state: DocumentState, revision: number, limits: ReturnType<typeof readLimits>): DocumentState {
  const files = new Map(state.files);
  const manifest = { ...state.snapshot.manifest, revision };
  const annotations = state.snapshot.annotations.map((annotation) => annotation.status === "needs_reanchor"
    ? annotation : { ...annotation, baseRevision: revision });
  files.set("manifest.json", encode(JSON.stringify(manifest)));
  files.set(manifest.annotations, encode(JSON.stringify(annotations)));
  return buildDocumentState(files, limits);
}

/** Create an isolated session using the shared transaction/history boundary. */
export function createDocumentSession(document: WorkbookModel, options?: DocumentSessionOptions): WorkbookSession;
export function createDocumentSession(document: ContentDocumentModel, options?: DocumentSessionOptions): DocumentSession;
export function createDocumentSession(document: DocumentModel, options?: DocumentSessionOptions): DocumentSession | WorkbookSession;
export function createDocumentSession(document: DocumentModel, options: DocumentSessionOptions = {}): DocumentSession | WorkbookSession {
  if (document.kind === "workbook") return createWorkbookSession(document, options);
  const limits = readLimits(options.limits), budgets = sessionLimits(options.sessionLimits);
  const initial = buildDocumentState(copyPackage(document.serialize(), limits), limits);
  const model = (source: DocumentState): ContentDocumentModel => Object.freeze<ContentDocumentModel>({
    kind: source.snapshot.manifest.kind, snapshot: () => source.snapshot, getNode: id => source.html.getNode(id),
    serialize: () => new Map([...source.files].map(([path, bytes]) => [path, new Uint8Array(bytes)])),
  });
  const boundary = createSessionLifecycle(initial, {
    model, prepare: prepareContent,
    rebase: (source, revision) => rebase(source, revision, limits),
    compare: (before, after, mapping) => compareStates(before, after, mapping?.textEdits),
    same: (before, after) => JSON.stringify(before.snapshot) === JSON.stringify(after.snapshot),
    retained: retainedBytes, identities: source => source.html.nodeIds(),
  }, options);
  const { state, execute, prepareTransaction } = boundary;
  function applyAnnotation(operation: AnnotationOperation, annotations: Annotation[], explicit: Set<string>): void {
    const base = state().snapshot.manifest;
    if (operation.op === "addAnnotation") {
      const annotation = operation.annotation;
      if (annotations.some((item) => item.annotationId === annotation.annotationId)
        || annotation.documentId !== base.documentId || annotation.baseRevision !== base.revision) invalid("The new annotation identity or revision is invalid.");
      if (annotation.status !== "needs_reanchor" && annotationNodes(annotation).some((id) => !state().html.hasNode(id))) invalid("Annotation targets must belong to the base snapshot.");
      annotations.push(annotation);
      explicit.add(annotation.annotationId);
    } else {
      const index = annotations.findIndex((item) => item.annotationId === operation.annotationId);
      if (index < 0) invalid("The annotation operation targets a missing annotation.");
      if (operation.op === "removeAnnotation") annotations.splice(index, 1);
      else {
        const annotation = parseContract("annotation", { ...annotations[index], ...operation.changes });
        if (operation.changes.anchor) {
          if (annotation.status !== "needs_reanchor" && annotationNodes(annotation).some((id) => !state().html.hasNode(id))) invalid("Annotation targets must belong to the base snapshot.");
          explicit.add(operation.annotationId);
        }
        annotations[index] = annotation;
      }
    }
  }

  function prepareContent(base: DocumentState, transaction: DocumentTransaction, allocate: () => string,
    resolvedIds?: readonly string[], options?: TransactionPrepareOptions) {
    // Validate every target against the base before applying any operation.
    for (const operation of transaction.operations) {
      if (isWorkbookOperation(operation)) invalid("Workbook operations require a workbook session.");
      if ("nodeId" in operation && !base.html.hasNode(operation.nodeId)) invalid("An operation targets a node outside the base snapshot.");
      if (isPresentationOperation(operation)) validatePresentationTargets(base, operation);
    }
    const resources = stageResources(base, options, limits);
    const edit = new DocumentEdit(resources.state, limits, allocate);
    let annotations = parseContract("annotations", base.snapshot.annotations);
    const parsedMetadata = parseContract("document-metadata", base.snapshot.metadata);
    if ("workbook" in parsedMetadata) invalid("Content operations require Word or PPT metadata.");
    const metadata: ContentDocumentMetadata = parsedMetadata;
    const files = new Map(resources.state.files);
    const styles=new PresentationStyleEdit(files,base.snapshot.manifest.styles,limits);
    const parseEdited = (): ParsedHtml => parseHtml(edit.html(), resources.state.snapshot.manifest, files, limits, { nodes: 0 });
    const explicit = new Set<string>();
    const unresolved = new Set(base.snapshot.annotations.filter((item) => item.status === "needs_reanchor").map((item) => item.annotationId));
    for (const operation of transaction.operations) {
      if (isWorkbookOperation(operation)) invalid("Workbook operations require a workbook session.");
      else if (isPresentationOperation(operation)) {
        if (!metadata.presentation) invalid("Presentation operations cannot target a Word document.");
        applyPresentationOperation({ base, p: metadata.presentation, edit, parse: parseEdited, allocate, limits,styles,annotations,explicit }, operation);
      }
      else if (isWordOperation(operation)) {
        if (metadata.presentation) invalid("Word operations require a Word document.");
        applyWordOperation(metadata, edit, parseEdited, styles, operation, limits,annotations);
      }
      else if ("nodeId" in operation) { guardPresentationContent(base, operation);if(operation.op==="replaceText"&&metadata.word?.trackChanges)trackWordText(metadata.word,edit,parseEdited,operation);else edit.apply(operation); }
      else applyAnnotation(operation, annotations, explicit);
    }
    if(metadata.word)refreshWordTableStyles(metadata.word,edit,parseEdited,styles,wordTableStyleTargets(metadata.word,parseEdited(),base.html,transaction.operations));
    for (const annotation of annotations) {
      if (annotation.status !== "needs_reanchor" && unresolved.has(annotation.annotationId) && !explicit.has(annotation.annotationId)) {
        invalid("Reactivating an unresolved annotation requires an explicit anchor.");
      }
      if (explicit.has(annotation.annotationId) && annotation.status !== "needs_reanchor"
        && annotationNodes(annotation).some((id) => !base.html.hasNode(id))) invalid("Annotation targets must belong to the base snapshot.");
    }
    removeDeletedCommentBodies(base, annotations, edit,metadata);
    if(metadata.word)pruneWordComments(metadata.word,annotations,edit,parseEdited);
    let parsed = parseEdited();
    for(const control of base.snapshot.metadata.word?.controls??[]){
      const current=metadata.word?.controls?.find(item=>item.nodeId===control.nodeId),unlocked=transaction.operations.some(operation=>operation.op==='setWordControl'&&operation.nodeId===control.nodeId&&operation.changes.lock==='unlocked');
      const before=base.html.getNode(control.nodeId),after=parsed.getNode(control.nodeId),changed=before?.html!==after?.html;
      if(!unlocked){
        if((control.lock==='sdtLocked'||control.lock==='sdtContentLocked')&&!after)invalid('The content control container is locked.');
        if((control.lock==='contentLocked'||control.lock==='sdtContentLocked')&&(changed||current&&['checked','value','fullDate'].some(key=>Reflect.get(control,key)!==Reflect.get(current,key))))invalid('The content control contents are locked.');
      }
      if(after&&current&&changed){
        const typed=transaction.operations.some(operation=>'nodeId'in operation&&operation.nodeId===control.nodeId&&['setWordControl','setWordControlValue','setWordControlPicture','removeWordControl'].includes(operation.op));
        if(!typed&&['checkBox','dropDown','date'].includes(control.kind)&&before?.text!==after.text)invalid('Use a typed content control value operation.');
        if(!typed&&control.kind==='comboBox'&&before?.text!==after.text){current.value=after.text;current.showingPlaceholder=false;}
      }
    }
    for(const control of base.snapshot.metadata.word?.controls??[])if(control.temporary&&parsed.hasNode(control.nodeId)&&base.html.getNode(control.nodeId)?.text!==parsed.getNode(control.nodeId)?.text){edit.unwrapSource(control.nodeId);if(metadata.word?.controls)metadata.word.controls=metadata.word.controls.filter(item=>item.nodeId!==control.nodeId);}
    parsed=parseEdited();
    if(metadata.word){
      if(metadata.word.noteReferences)metadata.word.noteReferences=metadata.word.noteReferences.filter(reference=>parsed.hasNode(reference.nodeId));
      const referencedNotes=new Set((metadata.word.noteReferences??[]).map(reference=>reference.noteId));for(const note of metadata.word.notes??[])if((note.type??'normal')==='normal'&&!referencedNotes.has(note.noteId)&&base.snapshot.metadata.word?.noteReferences?.some(reference=>reference.noteId===note.noteId)){const story=metadata.word.stories.find(story=>story.storyId===note.storyId);if(story&&parsed.hasNode(story.rootNodeId))edit.removeSource(story.rootNodeId);}parsed=parseEdited();
      if(metadata.word.notes)metadata.word.notes=metadata.word.notes.filter(note=>{const story=metadata.word!.stories.find(story=>story.storyId===note.storyId);return!!story&&parsed.hasNode(story.rootNodeId);});
      // Text stories belong to their objects. Remove the complete ownership
      // closure even when a generic paragraph deletion removed the object.
      let changed=true;while(changed){changed=false;const owners=new Set((metadata.word.drawings??[]).filter(item=>parsed.hasNode(item.nodeId)).flatMap(item=>'textStoryId'in item&&item.textStoryId?[item.textStoryId]:[]));for(const story of metadata.word.stories)if(story.kind==='textBox'&&!owners.has(story.storyId)&&parsed.hasNode(story.rootNodeId)){edit.removeSource(story.rootNodeId);changed=true;}if(changed)parsed=parseEdited();}
      metadata.word.stories=metadata.word.stories.filter(story=>parsed.hasNode(story.rootNodeId));
      if(metadata.word.drawings)metadata.word.drawings=metadata.word.drawings.filter(item=>parsed.hasNode(item.nodeId));
      updateWordDrawingConnectors(metadata.word,{getNode:id=>parsed.getNode(id)});
    }
    updatePresentationMetadata(metadata, base, parsed, annotations);
    if (metadata.presentation) { updatePresentationConnectors(metadata.presentation, base.snapshot.metadata.presentation!, base.html, parsed, edit); parsed = parseEdited(); }
    files.set(base.snapshot.manifest.entrypoint, encode(parsed.html));
    checkResourceUse(resources.added, parsed, base.snapshot.manifest.entrypoint);
    const impacted = metadata.presentation ? new Set(affectedPresentationSlides(base, { html: parsed, snapshot: { ...resources.state.snapshot, metadata,styles:styles.snapshot() } },true)) : undefined;
    const affectedSurfaces = impacted ? new Set([...(base.snapshot.metadata.presentation?.slides ?? []), ...(metadata.presentation?.slides ?? [])].filter(s => impacted.has(s.slideId)).map(s => s.rootNodeId)) : undefined;
    const slideSurfaces = new Set(metadata.presentation?.slides.map(s => s.rootNodeId));
    const regionChanged = affectedSurfaces ? (surfaceId: string,anchor:Extract<Anchor,{kind:"region"}>): boolean => {
      if(isPresentationCommentPoint(anchor))return false;
      for (let id: string | null = surfaceId; id !== null; id = parsed.peekNode(id)?.parentNodeId ?? null) if (slideSurfaces.has(id)) return affectedSurfaces.has(id);
      // Old generic regions outside slide surfaces retain the conservative layout rule.
      return base.snapshot.html !== parsed.html || impacted!.size > 0;
    } : ()=>base.snapshot.html!==parsed.html||JSON.stringify(base.snapshot.metadata)!==JSON.stringify(metadata)||JSON.stringify(base.snapshot.styles)!==JSON.stringify(styles.snapshot());
    const workflows=new Map(annotations.map(annotation=>[annotation.annotationId,annotation.status]));
    mapAnnotations(annotations, base.html, parsed, edit, explicit, regionChanged);
    updatePresentationRecovery(metadata, base, annotations,workflows);
    updateWordMetadata(metadata, base, parsed);
    if(metadata.word)pruneWordNativeResources(metadata.word,base.snapshot.metadata.word?.nativeContent,files,parsed,base.snapshot.manifest.entrypoint,limits);
    if(metadata.word)pruneRemovedWordImages(metadata.word,files,base.html,parsed,base.snapshot.manifest.entrypoint);
    if (resolvedIds) {
      const resolved = new Set(resolvedIds);
      const contentChanged = parsed.html !== base.snapshot.html || JSON.stringify(metadata) !== JSON.stringify(base.snapshot.metadata)||JSON.stringify(annotations)!==JSON.stringify(base.snapshot.annotations);
      if (!contentChanged) annotations = parseContract("annotations", base.snapshot.annotations);
      else for (const annotation of annotations) {
        if (resolved.has(annotation.annotationId) && annotation.status === "open") annotation.status = "resolved";
      }
    }
    files.set(base.snapshot.manifest.metadata, encode(JSON.stringify(metadata)));
    files.set(base.snapshot.manifest.annotations, encode(JSON.stringify(annotations)));
    const next = buildDocumentState(files, limits);
    return { state: next, mapping: compareStates(base, next, edit.textEdits).mapping };
  }

  function buildRequest(input:PresentationInteractionRequestInput):PresentationInteractionRequest;
  function buildRequest(input:InteractionRequestInput):DocumentInteractionRequest;
  function buildRequest(input:InteractionRequestInput|PresentationInteractionRequestInput):DocumentInteractionRequest|PresentationInteractionRequest {
    return execute(()=>{
      checkTransactionSize(input,budgets.maxTransactionBytes,limits.maxDepth);
      if(state().snapshot.metadata.presentation){
        if(!input||!("kind"in input)||input.kind!=="presentation")throw new DocumentError("UnsupportedFeature","PPT requests require explicit kind: presentation.");
        const request=buildPresentationRequest(state(),input);checkLimit(encode(JSON.stringify(request)).byteLength,limits.maxPackageBytes);return request;
      }
      const request=buildInteractionRequest(state(),input as InteractionRequestInput);checkLimit(encode(JSON.stringify(request)).byteLength,limits.maxPackageBytes);return request;
    });
  }
  function prepareProposalRequest(request:PresentationInteractionRequest,proposal:PresentationEditProposal,options?:TransactionPrepareOptions):CandidateChange;
  function prepareProposalRequest(request:DocumentInteractionRequest,proposal:EditProposal):CandidateChange;
  function prepareProposalRequest(requestInput:DocumentInteractionRequest|PresentationInteractionRequest,proposalInput:EditProposal|PresentationEditProposal,options?:TransactionPrepareOptions):CandidateChange {
    return execute(()=>{
      checkTransactionSize(requestInput,limits.maxPackageBytes*8,limits.maxDepth+16);checkLimit(encode(JSON.stringify(requestInput)).byteLength,limits.maxPackageBytes);checkTransactionSize(proposalInput,budgets.maxTransactionBytes,limits.maxDepth);
      if(state().snapshot.metadata.presentation){
        if(requestInput?.context?.kind!=="presentation")throw new DocumentError("UnsupportedFeature","PPT proposals require their explicit presentation request contract.");
        const request=parseContract("presentation-interaction-request",requestInput),proposal=parseContract("presentation-edit-proposal",proposalInput);
        validatePresentationProposal(state(),request,proposal,options);
        const candidate=prepareTransaction({protocolVersion:1,documentId:proposal.documentId,baseRevision:proposal.baseRevision,operations:proposal.operations},proposal.annotationIds,options);
        try{validatePresentationCandidate(state(),request,candidate);}catch(error){boundary.lifecycle.discard(candidate);throw error;}return candidate;
      }
      if(options!==undefined)invalid("PPT proposal resource options require a presentation session.");
      const request=parseContract("interaction-request",requestInput),proposal=parseContract("edit-proposal",proposalInput);validateProposal(state(),request,proposal);
      const candidate=prepareTransaction({protocolVersion:1,documentId:proposal.documentId,baseRevision:proposal.baseRevision,operations:proposal.operations},proposal.annotationIds);
      try{validateWordCandidate(state(),request,candidate.model);}catch(error){boundary.lifecycle.discard(candidate);throw error;}return candidate;
    });
  }

  return Object.freeze({
    kind: initial.snapshot.manifest.kind,
    snapshot: () => state().snapshot,
    getNode: (id: string) => state().html.getNode(id),
    serialize: () => model(state()).serialize(),
    ...boundary.lifecycle,
    buildInteractionRequest: buildRequest,
    prepareProposal: prepareProposalRequest,
  });
}

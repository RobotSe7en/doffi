import { AssetDescriptorSchema } from "../contracts/document.js";
import { PresentationOperationNameSchema } from "../contracts/presentation-interaction.js";
import type { PresentationInteractionRequest, PresentationEditProposal } from "../contracts/presentation-interaction.js";
import type { DocumentState } from "../document/state.js";
import type { PresentationInteractionRequestInput, TransactionPrepareOptions, CandidateChange } from "../runtime.js";
import { PresentationDocumentMetadataSchema } from "../contracts/presentation.js";
import { DocumentError } from "../errors.js";
import { createOpaqueId } from "../identifiers.js";
import { parseContract } from "../schema-registry.js";
import { checkLimit, freeze } from "../document/validation.js";
import { encode, resolveReference } from "../document/package.js";
import { DEFAULT_INTERACTION_LIMITS } from "./interaction.js";
import { anchorNodes } from "./anchors.js";
const equal = (a: unknown, b: unknown): boolean => JSON.stringify(a) === JSON.stringify(b);
const unique = (ids: readonly string[]): boolean => new Set(ids).size === ids.length;
const globalNames = new Set(["setPresentationTheme", "insertPresentationDefinition", "updatePresentationDefinition", "removePresentationDefinition", "setPresentationSize"]);
function invalid(message: string): never { throw new DocumentError("InvalidProposal", message); }
export function buildPresentationRequest(state: DocumentState, input: PresentationInteractionRequestInput, requestId = createOpaqueId()): PresentationInteractionRequest {
    const p = state.snapshot.metadata.presentation;
    if (!p)
        throw new DocumentError("UnsupportedFormat", "PPT requests require a presentation session.");
    if (!input || Object.keys(input).some(key => !["kind", "annotationIds", "scope", "supportedOperations", "assets", "limits"].includes(key)) || input.kind !== "presentation" || !Array.isArray(input.annotationIds) || !input.annotationIds.length || !unique(input.annotationIds) || input.scope !== undefined && input.scope !== "slides" && input.scope !== "global")
        throw new DocumentError("InvalidOperation", "Select a unique, nonempty list of current PPT annotations and an explicit supported scope.");
    const limits = { ...DEFAULT_INTERACTION_LIMITS, ...input.limits };
    for (const [key, value] of Object.entries(limits))
        if (!Object.hasOwn(DEFAULT_INTERACTION_LIMITS, key) || !Number.isSafeInteger(value) || value <= 0)
            throw new DocumentError("InvalidOperation", "PPT context limits must be positive safe integers.");
    const annotations = input.annotationIds.map(id => {
        const a = state.snapshot.annotations.find(a => a.annotationId === id);
        if (!a || a.status !== "open" || a.baseRevision !== state.snapshot.manifest.revision)
            throw new DocumentError("AnchorNotFound", "Only open annotations at the current revision can be submitted.");
        return a;
    });
    const global = input.scope === "global", selected = new Set<string>(), slideByRoot = new Map(p.slides.map(s => [s.rootNodeId, s.slideId]));
    const owner = (nodeId: string): string | undefined => {
        for (let id: string | null = nodeId; id !== null; id = state.html.peekNode(id)?.parentNodeId ?? null)
            if (slideByRoot.has(id))
                return slideByRoot.get(id);
        return undefined;
    };
    if (global)
        for (const slide of p.slides)
            selected.add(slide.slideId);
    else
        for (const annotation of annotations)
            for (const id of anchorNodes(annotation.anchor)) {
                const slide = owner(id);
                if (!slide)
                    throw new DocumentError("AnchorNotFound", "Ordinary PPT requests require slide anchors; definitions require global scope.");
                selected.add(slide);
            }
    const metadata = PresentationDocumentMetadataSchema.parse(state.snapshot.metadata), source = metadata.presentation;
    source.slides = source.slides.filter(s => selected.has(s.slideId));
    const layoutIds = new Set(source.slides.map(s => s.layoutId));
    source.layouts = source.layouts.filter(l => global || layoutIds.has(l.layoutId));
    const masterIds = new Set(source.layouts.map(l => l.masterId));
    source.masters = source.masters.filter(m => global || masterIds.has(m.masterId));
    const bodyIds = new Set(annotations.flatMap(a => [a, ...a.replies ?? []]).flatMap(a => a.bodyNodeId ? [a.bodyNodeId] : []));
    source.stories = source.stories.filter(s => s.kind === "notes" ? selected.has(s.slideId) : bodyIds.has(s.rootNodeId));
    const notesIds = new Set(source.stories.flatMap(s => s.kind === "notes" && s.notesMasterId ? [s.notesMasterId] : []));
    source.notesMasters = source.notesMasters.filter(n => global || notesIds.has(n.notesMasterId));
    const themes = new Set([...source.masters, ...source.notesMasters].map(d => d.themeId));
    source.themes = source.themes.filter(t => global || themes.has(t.themeId));
    const selectedAnnotations = new Set(annotations.map(a => a.annotationId));
    source.comments = source.comments.filter(c => selectedAnnotations.has(c.annotationId));
    if (source.commentAuthors) {
        const authors = new Set(source.comments.flatMap(c => [...(c.authorId ? [c.authorId] : []), ...c.replyAuthors?.map(r => r.authorId) ?? []]));
        source.commentAuthors = source.commentAuthors.filter(a => authors.has(a.authorId));
    }
    const dependencies: PresentationInteractionRequest["context"]["dependencies"] = source.themes.map(t => ({ kind: "theme", definitionId: t.themeId, access: global ? "editable" : "read-only" }));
    for (const master of source.masters)
        dependencies.push({ kind: "master", definitionId: master.masterId, nodeId: master.rootNodeId, html: state.html.getNode(master.rootNodeId)!.html, access: global ? "editable" : "read-only" });
    for (const layout of source.layouts)
        dependencies.push({ kind: "layout", definitionId: layout.layoutId, nodeId: layout.rootNodeId, html: state.html.getNode(layout.rootNodeId)!.html, access: global ? "editable" : "read-only" });
    for (const master of source.notesMasters)
        dependencies.push({ kind: "notes-master", definitionId: master.notesMasterId, nodeId: master.rootNodeId, html: state.html.getNode(master.rootNodeId)!.html, access: global ? "editable" : "read-only" });
    for (const story of source.stories)
        dependencies.push({ kind: story.kind, definitionId: story.storyId, nodeId: story.rootNodeId, html: state.html.getNode(story.rootNodeId)!.html, access: "editable" });
    const selectedRoots = new Set(source.slides.map(s => s.rootNodeId)), blocks = state.html.nodeIds().flatMap(id => selectedRoots.has(id) ? [{ nodeId: id, slideId: slideByRoot.get(id)!, html: state.html.getNode(id)!.html }] : []), nodes = new Set<string>();
    for (const root of [...blocks, ...dependencies.filter((d): d is Exclude<typeof d, {
            kind: "theme";
        }> => d.kind !== "theme")]) {
        const stack = [root.nodeId];
        while (stack.length) {
            const id = stack.pop()!;
            nodes.add(id);
            checkLimit(nodes.size, limits.maxNodes);
            stack.push(...state.html.peekNode(id)!.childNodeIds);
        }
    }
    source.objects = source.objects.filter(o => nodes.has(o.nodeId));
    delete source.nativePreview;
    const chartIds = new Set(source.objects.flatMap(o => o.kind === "chart" ? [o.chartId] : []));
    source.charts = source.charts.filter(c => chartIds.has(c.chartId));
    const dataIds = new Set(source.charts.map(c => c.dataSetId));
    source.dataSets = source.dataSets.filter(d => dataIds.has(d.dataSetId));
    source.hyperlinks = source.hyperlinks.filter(h => nodes.has(h.nodeId));
    let count = nodes.size;
    for (const data of source.dataSets) {
        count += data.categories.length + data.series.reduce((n, s) => n + s.values.length + 1, 0);
        checkLimit(count, limits.maxNodes);
    }
    const used = new Set<string>();
    for (const id of nodes) {
        const n = state.html.peekNode(id)!;
        if (n.tagName === "img")
            used.add(resolveReference(n.attributes.src!, state.snapshot.manifest.entrypoint));
    }
    const providedAssets = (input.assets ?? []).map(a => AssetDescriptorSchema.parse(a));
    if (!unique(providedAssets.map(a => a.path)))
        throw new DocumentError("InvalidOperation", "PPT context resources must have unique paths.");
    const assets = new Map(state.snapshot.manifest.assets.filter(a => used.has(a.path)).map(a => [a.path, a]));
    for (const asset of providedAssets) {
        const previous = state.snapshot.manifest.assets.find(a => a.path === asset.path);
        if (previous && !equal(previous, asset))
            throw new DocumentError("InvalidOperation", "Authorized resources cannot replace immutable package paths.");
        assets.set(asset.path, asset);
    }
    checkLimit(assets.size, limits.maxAssets);
    const supportedOperations = input.supportedOperations ?? PresentationOperationNameSchema.options.filter(name => global || !globalNames.has(name));
    if (!Array.isArray(supportedOperations) || !unique(supportedOperations) || supportedOperations.some(name => !PresentationOperationNameSchema.safeParse(name).success || !global && globalNames.has(name)))
        throw new DocumentError("InvalidOperation", "PPT operation declarations must match their explicit scope.");
    const request = parseContract("presentation-interaction-request", { protocolVersion: 1, requestId, documentId: state.snapshot.manifest.documentId, baseRevision: state.snapshot.manifest.revision, annotations, context: { formatVersion: 1, kind: "presentation", scope: { kind: global ? "global" : "slides", slideIds: blocks.map(b => b.slideId) }, blocks, dependencies, styles: state.snapshot.styles, metadata, assets: [...assets.values()], providedAssets }, supportedOperations });
    checkLimit(encode(JSON.stringify(request)).byteLength, limits.maxBytes);
    return freeze(request);
}
export function validatePresentationProposal(state: DocumentState, request: PresentationInteractionRequest, proposal: PresentationEditProposal, options?: TransactionPrepareOptions): void {
    if (request.documentId !== state.snapshot.manifest.documentId || proposal.documentId !== request.documentId || proposal.requestId !== request.requestId)
        invalid("The PPT proposal does not match this request and document.");
    if (request.baseRevision !== state.snapshot.manifest.revision || proposal.baseRevision !== request.baseRevision)
        throw new DocumentError("RevisionConflict", "The PPT request or proposal is based on a different revision.");
    const canonical = buildPresentationRequest(state, { kind: "presentation", annotationIds: request.annotations.map(a => a.annotationId), scope: request.context.scope.kind, supportedOperations: request.supportedOperations, assets: request.context.providedAssets, limits: { maxBytes: Number.MAX_SAFE_INTEGER, maxNodes: Number.MAX_SAFE_INTEGER, maxAssets: Number.MAX_SAFE_INTEGER } }, request.requestId);
    if (!equal(request, canonical))
        invalid("The PPT request differs from the complete frozen source context.");
    if (!unique(proposal.annotationIds) || proposal.annotationIds.some(id => !request.annotations.some(a => a.annotationId === id)))
        invalid("The PPT proposal claims annotations outside this request.");
    const writable = new Set<string>(), add = (nodeId: string): void => {
        const stack = [nodeId];
        while (stack.length) {
            const id = stack.pop()!;
            writable.add(id);
            stack.push(...state.html.peekNode(id)!.childNodeIds);
        }
    };
    for (const b of request.context.blocks)
        add(b.nodeId);
    for (const d of request.context.dependencies)
        if (d.kind !== "theme" && d.access === "editable")
            add(d.nodeId);
    const p = request.context.metadata.presentation, slides = new Set(request.context.scope.slideIds), assets = new Set(request.context.assets.map(a => a.path)), global = request.context.scope.kind === "global";
    for (const resource of options?.assets ?? [])
        if (!resource?.descriptor || !request.context.providedAssets.some(a => equal(a, resource.descriptor)))
            invalid("Proposal resource bytes lack a matching host-authorized request descriptor.");
    for (const op of proposal.operations) {
        if (!request.supportedOperations.includes(op.op) || !global && globalNames.has(op.op))
            invalid("The PPT proposal uses an operation outside its declared scope.");
        const targets = [...("nodeId" in op ? [op.nodeId] : []), ...("parentNodeId" in op ? [op.parentNodeId] : []), ...("beforeNodeId" in op && op.beforeNodeId ? [op.beforeNodeId] : []), ...("nodeIds" in op ? op.nodeIds : [])];
        if (targets.some(id => !writable.has(id)))
            invalid("A proposal cannot write read-only dependencies or another slide.");
        for (const key of ["slideId", "afterSlideId", "beforeSlideId"] as const)
            if (key in op) {
                const id = (op as unknown as Record<string, unknown>)[key];
                if (typeof id === "string" && !slides.has(id))
                    invalid("A PPT proposal references a slide outside its frozen scope.");
            }
        if ("layoutId" in op && !p.layouts.some(l => l.layoutId === op.layoutId))
            invalid("A proposal references a layout outside its dependencies.");
        if ("themeId" in op && !p.themes.some(t => t.themeId === op.themeId))
            invalid("A proposal references a theme outside its dependencies.");
        if ("notesMasterId" in op && op.notesMasterId && !p.notesMasters.some(m => m.notesMasterId === op.notesMasterId))
            invalid("A proposal references a notes master outside its dependencies.");
        if ("annotationId" in op && !request.annotations.some(a => a.annotationId === op.annotationId))
            invalid("A proposal cannot edit a thread outside its request.");
        if (op.op === "replaceAsset" && !assets.has(op.assetPath) || op.op === "insertPresentationObject" && op.object.kind === "image" && !assets.has(op.object.assetPath))
            invalid("A PPT proposal references an unauthorized resource.");
        if (op.op === "insertPresentationObject" && op.object.kind === "connector" && [op.object.startNodeId, op.object.endNodeId].some(id => id && !writable.has(id)) || op.op === "setPresentationConnector" && [op.changes.start, op.changes.end].some(end => end && !writable.has(end.nodeId)))
            invalid("A connector cannot bind outside the proposal scope.");
        const anchor = op.op === "addPresentationComment" ? op.anchor : op.op === "updatePresentationComment" ? op.changes.anchor : undefined;
        if (anchor && anchorNodes(anchor).some(id => !writable.has(id)))
            invalid("A proposal comment anchor is outside its writable content.");
    }
}
export function validatePresentationCandidate(state: DocumentState, request: PresentationInteractionRequest, candidate: CandidateChange): void {
    const assets = new Set(request.context.assets.map(a => a.path));
    for (const change of candidate.changes) {
        const node = candidate.model.getNode(change.nodeId);
        if (node?.tagName === "img" && !assets.has(resolveReference(node.attributes.src!, candidate.preview.manifest.entrypoint)))
            invalid("The PPT candidate introduces an image outside its frozen resources.");
    }
    if (request.context.scope.kind === "global")
        return;
    const existing = new Set(state.snapshot.metadata.presentation!.slides.map(s => s.slideId)), allowed = new Set(request.context.scope.slideIds);
    if (candidate.affectedSlideIds?.some(id => existing.has(id) && !allowed.has(id)))
        invalid("The PPT candidate indirectly changes a slide outside the frozen request scope.");
}

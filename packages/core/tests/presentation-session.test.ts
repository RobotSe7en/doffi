import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, test } from "vitest";
import { Ajv2020 } from "ajv/dist/2020.js";
import { createAssetDescriptor, createDocumentSession, createPresentation, getContractJsonSchema, parseContract, readPresentationGeometry } from "@doffi/core";
import type { Annotation, DocumentOperation, DocumentSession, DocumentSessionOptions, PresentationDocumentMetadata, PresentationGeometry, TransactionPrepareOptions } from "@doffi/core";
import { repository } from "./fixtures.js";

const encode = (value: string) => new TextEncoder().encode(value), decode = (value: Uint8Array | undefined) => new TextDecoder().decode(value);
function example(): Map<string, Uint8Array> {
  return new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json", "assets/sample.png"].map(path => [path, new Uint8Array(readFileSync(join(repository, "examples/presentation", path)))]));
}
const json = (files: Map<string, Uint8Array>, path: string, value: unknown) => files.set(path, encode(JSON.stringify(value)));
const metadata = (files: Map<string, Uint8Array>) => JSON.parse(decode(files.get("metadata.json"))) as PresentationDocumentMetadata;
function twoSlides(): Map<string, Uint8Array> {
  const files = example(), data = metadata(files);
  data.presentation.slides.push({ slideId: "slide-2", rootNodeId: "slide-root-2", layoutId: "layout-blank" });
  data.presentation.objects.push({ nodeId: "second-text", kind: "text" });
  json(files, "metadata.json", data);
  files.set("content.html", encode(decode(files.get("content.html")).replace("</article>", '<section data-node-id="slide-root-2" data-office-kind="slide" data-office-story="slide"><div data-node-id="second-text" data-office-kind="text" data-office-x="20" data-office-y="20" data-office-width="200" data-office-height="50"><p data-node-id="second-p">另一页</p></div></section></article>')));
  return files;
}
function session(files = example(), options?: DocumentSessionOptions): DocumentSession { return createDocumentSession(openDocument(files), options); }
const geometry = (s: DocumentSession, id: string): PresentationGeometry => readPresentationGeometry(s.getNode(id)!.attributes)!;
function prepare(s: DocumentSession, operations: DocumentOperation[], options?: TransactionPrepareOptions) {
  return s.prepare({ protocolVersion: 1, documentId: s.snapshot().manifest.documentId, baseRevision: s.snapshot().manifest.revision, operations }, options);
}
function region(id: string, surfaceId: string): Annotation {
  return { annotationId: id, documentId: "presentation-example", baseRevision: 0, status: "open", instruction: "检查位置", anchor: { kind: "region", surfaceId, nodeIds: [], rect: { x: 10, y: 10, width: 30, height: 20 }, layoutKey: "layout-0" } };
}
const move = (s: DocumentSession, id: string, dx: number): DocumentOperation => ({ op: "setPresentationGeometry", nodeId: id, geometry: { ...geometry(s, id), x: geometry(s, id).x + dx } });
const fail = (action: () => unknown, code = "InvalidOperation") => expect(action).toThrowError(expect.objectContaining({ code }));

describe("PPT transactions share atomic candidates, revisions and history", () => {
  test("keeps candidates isolated, maps UTF-16, rejects stale/foreign/consumed candidates and reopens source", () => {
    const s = session(), other = session(), initial = s.snapshot();
    const first = prepare(s, [{ op: "replaceText", nodeId: "milestone-p", from: 0, to: 0, expectedText: "", text: "中文𠀀😀葛󠄀" }, move(s, "milestone", 36)]);
    expect(s.snapshot()).toEqual(initial);
    const stale = prepare(s, [move(s, "heading", 4)]);
    fail(() => other.commit(first));
    expect(s.commit(first)).toMatchObject({ revision: 1, changed: true, affectedSlideIds: ["slide-1"] });
    expect(s.getAnnotations()[0]!.anchor).toMatchObject({ kind: "text", from: 9, to: 14, quote: "本季度上线" });
    fail(() => s.commit(stale), "RevisionConflict"); fail(() => s.commit(first));
    expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
    expect(s.undo().revision).toBe(2); expect(s.snapshot().html).toBe(initial.html);
    expect(s.redo().revision).toBe(3); expect(s.snapshot().html).toBe(first.preview.html);
    expect(other.snapshot()).toEqual(initial);
  });
  test("later invalid geometry/text/required references do not change source, resources or redo", () => {
    const s = session(); s.commit(prepare(s, [move(s, "heading", 10)])); s.undo();
    const snapshot = s.snapshot(), history = s.history();
    for (const operation of [
      { ...move(s, "milestone", 10), geometry: { ...geometry(s, "milestone"), width: 0 } },
      { op: "replaceText", nodeId: "heading-p", from: 0, to: 1, expectedText: "不匹配", text: "x" },
      { op: "removeNode", nodeId: "layout-root-blank" },
      { op: "removeNode", nodeId: "slide-root-1" },
    ] as DocumentOperation[]) fail(() => prepare(s, [move(s, "heading", 22), operation]));
    expect(s.snapshot()).toEqual(snapshot); expect(s.history()).toEqual(history);
    expect(s.redo().changed).toBe(true);
  });
  test("numeric geometry no-ops, empty batches and round trips preserve omission, revision and redo", () => {
    const s = session(), original = geometry(s, "heading");
    s.commit(prepare(s, [move(s, "heading", 10)])); s.undo();
    for (const operations of [[], [move(s, "heading", 0)], [move(s, "heading", 12), { op: "setPresentationGeometry", nodeId: "heading", geometry: original }] as DocumentOperation[]]) {
      expect(s.commit(prepare(s, operations))).toMatchObject({ changed: false, revision: 2, affectedSlideIds: [] });
      expect(s.history()).toMatchObject({ redoSteps: 1 });
    }
    expect(s.getNode("heading")!.attributes["data-office-rotation"]).toBeUndefined();
  });
  test("rejects format mismatch, CSS dimensions, role/geometry reconciliation and unimplemented protocol adapters", () => {
    const s = session(), word = createDocumentSession(openDocument(new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map(path => [path, new Uint8Array(readFileSync(join(repository, "examples/report", path)))]))));
    for (const operation of [move(s, "heading", 1), { op: "setPresentationSize", size: { unit: "pt", width: 720, height: 540 } }] as DocumentOperation[]) fail(() => prepare(word, operation.op === "setPresentationGeometry" ? [{ ...operation, nodeId: "n-title" }] : [operation]));
    fail(() => prepare(s, [{ op: "setStyle", nodeId: "heading", property: "width", value: "100pt" }]));
    fail(() => prepare(s, [{ op: "reconcileFragment", nodeId: "delivery-group", html: s.getNode("group-label")!.html.replace('data-office-x="120"', 'data-office-x="12"'), textEdits: [] }]));
    fail(() => s.buildInteractionRequest({ annotationIds: ["review-1"] }), "UnsupportedFeature");
    expect(s.snapshot().manifest.revision).toBe(0);
  });
  test("strict schemas include format operations without exposing them in Word proposals", () => {
    const operations: DocumentOperation[] = [{ op: "setPresentationGeometry", nodeId: "a", geometry: { x: -5, y: 0, width: 10, height: 10, rotation: 12, flipH: true, flipV: false } }, { op: "setPresentationSize", size: { unit: "pt", width: 720, height: 540 } }];
    const validate = new Ajv2020({ strict: false }).compile(getContractJsonSchema("document-operation"));
    for (const op of operations) { expect(validate(op)).toBe(true); expect(parseContract("document-operation", op)).toEqual(op); expect(validate({ ...op, extra: true })).toBe(false); }
    fail(() => parseContract("content-operation", operations[0]));
    expect(JSON.stringify(getContractJsonSchema("edit-proposal"))).not.toContain("setPresentationGeometry");
  });
  test("metadata-only size changes report pages and events, and undo restores exact dimensions", () => {
    const s = session(twoSlides()), events: unknown[] = []; s.subscribe(e => events.push(parseContract("document-event", e)));
    const before = s.snapshot().html, candidate = prepare(s, [{ op: "setPresentationSize", size: { unit: "pt", width: 720, height: 540 } }]);
    expect(candidate.preview.html).toBe(before); expect(candidate.affectedSlideIds).toEqual(["slide-1", "slide-2"]);
    expect(candidate.affectedNodeIds).toEqual(expect.arrayContaining(["slide-root-1", "slide-root-2"]));
    s.commit(candidate); expect(events[0]).toMatchObject({ type: "documentChanged", revision: 1, affectedSlideIds: ["slide-1", "slide-2"] });
    s.undo(); expect(s.snapshot().metadata.presentation!.size.width).toBe(960);
  });
  test("retains stable regions on unrelated slides and notes, invalidates referencing slides for global style", () => {
    const files = twoSlides(); json(files, "annotations.json", [region("r1", "slide-root-1"), region("r2", "slide-root-2")]);
    const data = metadata(files); data.presentation.comments = []; json(files, "metadata.json", data);
    const s = session(files);
    let c = prepare(s, [{ op: "setStyle", nodeId: "speaker-notes-p", property: "font-size", value: "20pt" }]);
    expect(c.affectedSlideIds).toEqual([]); s.commit(c);
    expect(s.getAnnotations().map(a => a.status)).toEqual(["open", "open"]);
    c = prepare(s, [move(s, "heading", 10)]); s.commit(c);
    expect(c.affectedSlideIds).toEqual(["slide-1"]); expect(s.getAnnotations().map(a => a.status)).toEqual(["needs_reanchor", "open"]);
    expect(s.getAnnotations()[0]).toMatchObject({ baseRevision: 1, anchor: region("r1", "slide-root-1").anchor });
    s.undo(); s.commit(prepare(s, [{ op: "setStyle", nodeId: "presentation-root", property: "font-size", value: "22pt" }]));
    expect(s.getAnnotations().map(a => a.status)).toEqual(["needs_reanchor", "needs_reanchor"]);
  });
  test("inherited placeholder geometry materializes only an explicit local override", () => {
    const files = example(), data = metadata(files); data.presentation.slides[0]!.layoutId = "layout-body";
    data.presentation.objects.find(o => o.nodeId === "heading")!.placeholder = { type: "title", index: 0 }; json(files, "metadata.json", data);
    files.set("content.html", encode(decode(files.get("content.html")).replace(/ data-office-x="48" data-office-y="24" data-office-width="864" data-office-height="64"/, "")));
    const s = session(files); expect(readPresentationGeometry(s.getNode("heading")!.attributes)).toBeUndefined();
    const g = geometry(s, "layout-title"), c = prepare(s, [{ op: "setPresentationGeometry", nodeId: "heading", geometry: g }]);
    expect(c.affectedSlideIds).toEqual(["slide-1"]); s.commit(c); expect(geometry(s, "heading")).toEqual(g);
    s.undo(); expect(readPresentationGeometry(s.getNode("heading")!.attributes)).toBeUndefined();
  });
  test("a definition change affects its referencing slides without invalidating another layout", () => {
    const files = twoSlides(), data = metadata(files); data.presentation.slides[0]!.layoutId = "layout-body";
    data.presentation.comments = []; json(files, "metadata.json", data); json(files, "annotations.json", [region("r1", "slide-root-1"), region("r2", "slide-root-2")]);
    const s = session(files), c = prepare(s, [move(s, "layout-title", 12)]);
    expect(c.affectedSlideIds).toEqual(["slide-1"]); s.commit(c); expect(s.getAnnotations().map(a => a.status)).toEqual(["needs_reanchor", "open"]);
    s.undo(); expect(s.getAnnotations().map(a => a.status)).toEqual(["open", "open"]);
  });
  test("bound connector caches update atomically when target or parent group moves", () => {
    const files = example(), data = metadata(files), link = data.presentation.objects.find(o => o.nodeId === "milestone-link")!;
    if (link.kind !== "connector") throw Error(); link.endNodeId = "group-diamond"; link.startSite = 1; link.endSite = 0; json(files, "metadata.json", data);
    const s = session(files), original = geometry(s, "milestone-link");
    const c = prepare(s, [move(s, "milestone", 100), { op: "setPresentationGeometry", nodeId: "delivery-group", geometry: { ...geometry(s, "delivery-group"), rotation: 90, flipH: true }, childCoordinates: { x: 0, y: 0, width: 800, height: 280 } }]);
    expect(c.affectedNodeIds).toContain("milestone-link"); s.commit(c);
    const g = geometry(s, "milestone-link"); expect(g).not.toEqual(original); expect(g.rotation).toBe(0);
    // RoundRect left site (148,170); group-local diamond top (50,0), then scale .5,
    // horizontal flip and 90-degree rotation around the group center => (318,505).
    expect(g.x).toBeCloseTo(148); expect(g.y).toBeCloseTo(170); expect(g.width).toBeCloseTo(170); expect(g.height).toBeCloseTo(335);
    s.undo(); expect(geometry(s, "milestone-link")).toEqual(original);
  });
  test("invalid connected sites and singular group extents fail the complete batch", () => {
    const files = example(), data = metadata(files), link = data.presentation.objects.find(o => o.nodeId === "milestone-link")!;
    if (link.kind !== "connector") throw Error(); link.startSite = 100; json(files, "metadata.json", data);
    const s = session(files), before = s.snapshot(); fail(() => prepare(s, [move(s, "milestone", 1)])); expect(s.snapshot()).toEqual(before);
    fail(() => prepare(s, [{ op: "setPresentationGeometry", nodeId: "delivery-group", geometry: geometry(s, "delivery-group"), childCoordinates: { x: 0, y: 0, width: 0, height: 10 } }]), "InvalidOperation");
  });
  test("complete object removal cleans chart/data/link references and detaches bound endpoints", () => {
    const s = session(); const c = prepare(s, [{ op: "removeNode", nodeId: "progress-chart" }, { op: "removeNode", nodeId: "milestone" }]); s.commit(c);
    const p = s.snapshot().metadata.presentation!; expect(p.charts).toEqual([]); expect(p.dataSets).toEqual([]);
    expect(p.objects.find(o => o.nodeId === "milestone-link")).not.toHaveProperty("startNodeId");
    expect(s.getAnnotations()[0]).toMatchObject({ status: "needs_reanchor", anchor: { nodeId: "milestone-p", quote: "本季度上线" } });
    s.undo(); expect(s.snapshot().metadata.presentation!.charts).toHaveLength(1); expect(s.getAnnotations()[0]!.status).toBe("open");
  });
  test("removing a thread removes its owned rich stories and undo restores them with native format", () => {
    const s = session(), before = s.snapshot(); s.commit(prepare(s, [{ op: "removeAnnotation", annotationId: "review-1" }]));
    expect(s.getAnnotations()).toEqual([]); expect(s.getNode("comment-body")).toBeUndefined(); expect(s.getNode("reply-body")).toBeUndefined();
    expect(s.snapshot().metadata.presentation!.stories.map(story => story.kind)).toEqual(["notes"]);
    expect(s.snapshot().metadata.presentation!.comments).toEqual([]); s.undo(); expect(s.snapshot().html).toBe(before.html);
  });
  test("preserves closed native workflow state through recovery and explicit reanchor", () => {
    const files = example(), annotation = JSON.parse(decode(files.get("annotations.json"))) as Annotation[]; annotation[0]!.status = "closed"; json(files, "annotations.json", annotation);
    const s = session(files); s.commit(prepare(s, [{ op: "removeNode", nodeId: "milestone" }]));
    expect(s.snapshot().metadata.presentation!.comments[0]).toMatchObject({ format: "modern", recoveryStatus: "closed" });
    fail(() => prepare(s, [{ op: "updateAnnotation", annotationId: "review-1", changes: { status: "closed" } }]));
    s.commit(prepare(s, [{ op: "updateAnnotation", annotationId: "review-1", changes: { status: "closed", anchor: { kind: "node", nodeId: "heading" } } }]));
    expect(s.snapshot().metadata.presentation!.comments[0]).not.toHaveProperty("recoveryStatus");
  });
  test("records modern format for new annotations and retains imported classic metadata", () => {
    const files = example(), data = metadata(files); data.presentation.comments[0]!.format = "classic"; data.presentation.commentFormat = "classic"; json(files, "metadata.json", data);
    const s = session(files); s.commit(prepare(s, [{ op: "addAnnotation", annotation: { ...region("new", "slide-root-1"), anchor: { kind: "node", nodeId: "heading" } } }]));
    expect(s.snapshot().metadata.presentation!.comments.map(c => c.format)).toEqual(["classic", "modern"]);
  });
  test("PPT foreground changes replace imported native color modifiers together", () => {
    const s = session(); s.commit(prepare(s, [{ op: "setStyle", nodeId: "heading-p", property: "color", value: "#008800" }]));
    expect(s.getNode("heading-p")!.attributes.style).toContain("--office-text-color:#008800");
    expect(s.getNode("heading-p")!.attributes.style).toContain("--office-text-source:direct");
  });
  test("review: an empty batch preserves valid legacy annotations without native registry entries", () => {
    const files = example(), data = metadata(files); data.presentation.comments = []; json(files, "metadata.json", data);
    const s = session(files), before = s.snapshot(); expect(s.commit(prepare(s, [])).changed).toBe(false); expect(s.snapshot()).toEqual(before);
  });
  test("review: a region on an object container cannot retain a stale rectangle after its slide changes", () => {
    const files = example(); json(files, "annotations.json", [region("review-1", "milestone")]);
    const s = session(files); s.commit(prepare(s, [move(s, "milestone", 10)])); expect(s.getAnnotations()[0]!.status).toBe("needs_reanchor");
  });
  test("review: a CSS margin cannot provide an alternative object coordinate", () => {
    const s = session(); fail(() => prepare(s, [{ op: "setStyle", nodeId: "heading", property: "margin-left", value: "24pt" }]));
    expect(s.snapshot().manifest.revision).toBe(0);
  });
  test("review: text replacement cannot silently cross a native metadata-only hyperlink boundary", () => {
    const files = example(), data = metadata(files); data.presentation.hyperlinks.push({ nodeId: "link-run", target: "https://example.com" }); json(files, "metadata.json", data);
    files.set("content.html", encode(decode(files.get("content.html")).replace('data-node-id="heading-p" style="font-size:30pt">项目进展', 'data-node-id="heading-p" style="font-size:30pt">项目<span data-node-id="link-run">进展</span>')));
    const s = session(files); fail(() => prepare(s, [{ op: "replaceText", nodeId: "heading-p", from: 0, to: 4, expectedText: "项目进展", text: "全部更新" }]));
    expect(s.snapshot().metadata.presentation!.hyperlinks).toHaveLength(1);
  });
});

describe("authorized resources share the PPT transaction boundary", () => {
  function resource() {
    const bytes = new Uint8Array(readFileSync(join(repository, "examples/presentation/assets/sample.png")));
    return { bytes, descriptor: createAssetDescriptor("assets/new.png", "image/png", bytes) };
  }
  test("copies bytes, previews source+manifest together, and undoes/redoes one image/resource revision", () => {
    const s = session(), input = resource(), original = s.snapshot();
    const c = prepare(s, [{ op: "replaceAsset", nodeId: "project-image-source", assetPath: input.descriptor.path }], { assets: [input] });
    input.bytes.fill(0); expect(s.snapshot()).toEqual(original); expect(c.preview.manifest.assets).toHaveLength(2);
    s.commit(c); expect(s.getNode("project-image-source")!.attributes.src).toBe("assets/new.png"); expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
    s.undo(); expect(s.serialize().has("assets/new.png")).toBe(false); expect(s.snapshot().manifest.assets).toHaveLength(1);
    s.redo(); expect(s.serialize().has("assets/new.png")).toBe(true); expect(s.history().undoSteps).toBe(1);
  });
  test("bad digest, unused bytes, duplicate paths, path collisions, missing resources and a later failure leave no artifacts", () => {
    const s = session(), before = s.snapshot(), input = resource();
    const op: DocumentOperation = { op: "replaceAsset", nodeId: "project-image-source", assetPath: input.descriptor.path };
    fail(() => prepare(s, [op] ), "AssetUnavailable");
    fail(() => prepare(s, [op], { assets: [{ ...input, descriptor: { ...input.descriptor, sha256: "0".repeat(64) } }] }));
    fail(() => prepare(s, [], { assets: [input] })); fail(() => prepare(s, [op], { assets: [input, input] }));
    for (const path of ["content.html", "content.html/child", "assets"]) fail(() => prepare(s, [], { assets: [{ bytes: input.bytes, descriptor: createAssetDescriptor(path, "image/png", input.bytes) }] }));
    fail(() => prepare(s, [op, { op: "removeNode", nodeId: "layout-root-blank" }], { assets: [input] }));
    expect(s.snapshot()).toEqual(before); expect(s.serialize().has(input.descriptor.path)).toBe(false); expect(s.history().undoSteps).toBe(0);
  });
  test("immutable existing descriptors deduplicate and unavailable caller bytes never overwrite source", () => {
    const s = session(), bytes = new Uint8Array(s.serialize().get("assets/sample.png")!), descriptor = createAssetDescriptor("assets/sample.png", "image/png", bytes);
    expect(s.commit(prepare(s, [], { assets: [{ descriptor, bytes }] })).changed).toBe(false);
    const updated = new Uint8Array(bytes); updated[20] = updated[20]! + 1;
    const changed = createAssetDescriptor("assets/sample.png", "image/png", updated);
    fail(() => prepare(s, [], { assets: [{ descriptor: changed, bytes: updated }] }));
  });
  test("resource and candidate byte/history quotas cannot partly commit", () => {
    const input = resource(), s = session(example(), { sessionLimits: { maxCandidateBytes: 10 } });
    fail(() => prepare(s, [{ op: "replaceAsset", nodeId: "project-image-source", assetPath: input.descriptor.path }], { assets: [input] }), "ResourceLimitExceeded");
    expect(s.snapshot().manifest.revision).toBe(0); expect(s.serialize().has(input.descriptor.path)).toBe(false);
    const limited = session(example(), { sessionLimits: { maxHistoryBytes: 10 } });
    const c = prepare(limited, [move(limited, "heading", 10)]); fail(() => limited.commit(c), "ResourceLimitExceeded"); expect(limited.snapshot().manifest.revision).toBe(0);
  });
  test("new blank presentations support independent sessions without introducing another authority", () => {
    const source = openDocument(createPresentation({ documentId: "new" })), a = createDocumentSession(source), b = createDocumentSession(source);
    a.commit(prepare(a, [{ op: "setPresentationSize", size: { unit: "pt", width: 720, height: 540 } }]));
    expect(a.snapshot().metadata.presentation!.size.width).toBe(720); expect(b.snapshot().metadata.presentation!.size.width).toBe(960); expect(source.snapshot().manifest.revision).toBe(0);
  });
});

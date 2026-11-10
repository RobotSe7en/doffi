import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { expect, test } from "vitest";
import { createDocumentSession, createPresentation, PresentationDocumentMetadataSchema, readPresentationGeometry } from "@doffi/core";
import type { DocumentOperation, DocumentSession, PresentationGeometry } from "@doffi/core";
import { repository } from "./fixtures.js";

const geometry: PresentationGeometry = { x: 40, y: 50, width: 240, height: 100, rotation: 0, flipH: false, flipV: false };
function blank() { return createDocumentSession(openDocument(createPresentation({ documentId: "edit-ppt" }))); }
function example() { return createDocumentSession(openDocument(new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json", "assets/sample.png"].map(path => [path, new Uint8Array(readFileSync(join(repository, "examples/presentation", path)))])))); }
function prepare(s: DocumentSession, operations: DocumentOperation[]) { const { documentId, revision } = s.snapshot().manifest; return s.prepare({ protocolVersion: 1, documentId, baseRevision: revision, operations }); }
const commit = (s: DocumentSession, operations: DocumentOperation[]) => s.commit(prepare(s, operations));
const p = (s: DocumentSession) => s.snapshot().metadata.presentation!;
const fail = (action: () => unknown, code = "InvalidOperation") => expect(action).toThrowError(expect.objectContaining({ code }));

test("table merged-cell edits preserve surviving identities and move spanning origins across removed rows",()=>{
  const s=example(),before=s.snapshot();
  commit(s,[{op:"editPresentationTable",nodeId:"progress-table",change:{kind:"mergeCells",row:0,column:0,rows:2,columns:1}}]);
  expect(s.getNode("table-h1")!.attributes.rowspan).toBe("2");expect(s.getNode("table-h1")!.text).toBe("阶段研发");expect(s.getNode("table-c1")).toBeUndefined();expect(s.getNode("table-c1-p")).toBeDefined();
  commit(s,[{op:"editPresentationTable",nodeId:"progress-table",change:{kind:"insertRows",at:1,count:1}}]);expect(s.getNode("table-h1")!.attributes.rowspan).toBe("3");
  commit(s,[{op:"editPresentationTable",nodeId:"progress-table",change:{kind:"deleteRows",at:0,count:1}}]);expect(s.getNode("table-h1")!.attributes.rowspan).toBe("2");expect(s.getNode("table-h1")!.parentNodeId).not.toBe("table-r1");
  commit(s,[{op:"editPresentationTable",nodeId:"progress-table",change:{kind:"splitCell",row:0,column:0}}]);expect(s.getNode("table-h1")!.attributes.rowspan).toBeUndefined();expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
  s.undo();s.undo();s.undo();s.undo();expect(s.snapshot().html).toBe(before.html);
});

test("column, TSV matrix and size edits are atomic and honor merged-cell bounds",()=>{
  const s=example();
  commit(s,[{op:"editPresentationTable",nodeId:"progress-table",change:{kind:"insertColumns",at:1,count:2}},{op:"editPresentationTable",nodeId:"progress-table",change:{kind:"setCells",row:0,column:1,values:[["甲","乙"],["中文𠮷󠄀","4"]]}}]);
  const table=p(s).objects.find(o=>o.nodeId==="progress-table")!;if(table.kind!=="table")throw Error();expect(table.grid).toHaveLength(4);expect(s.getNode("table")!.text).toContain("中文𠮷󠄀4");
  commit(s,[{op:"editPresentationTable",nodeId:table.nodeId,change:{kind:"deleteColumns",at:1,count:2}}]);expect(s.getNode("table")!.text).toBe("阶段进度研发完成");
  const before=s.snapshot();fail(()=>commit(s,[{op:"editPresentationTable",nodeId:table.nodeId,change:{kind:"deleteRows",at:0,count:2}}]));expect(s.snapshot()).toEqual(before);
  commit(s,[{op:"editPresentationTable",nodeId:table.nodeId,change:{kind:"resize",grid:[120,240],rowHeights:[40,60]}}]);expect(readPresentationGeometry(s.getNode(table.nodeId)!.attributes)).toMatchObject({width:360,height:100});expect(s.getNode("table-r1")!.attributes.style).toContain("height:40pt");
  const resized=s.snapshot();fail(()=>commit(s,[{op:"editPresentationTable",nodeId:table.nodeId,change:{kind:"mergeCells",row:0,column:0,rows:2,columns:1}},{op:"editPresentationTable",nodeId:table.nodeId,change:{kind:"mergeCells",row:1,column:0,rows:1,columns:2}}]));expect(s.snapshot()).toEqual(resized);
});

test("notes creation and text fitting use allocated source identities and shared history",()=>{
  const s=blank();commit(s,[{op:"insertPresentationNotes",slideId:"slide-1",html:"<p>中文备注</p>"}]);const note=p(s).stories.find(story=>story.kind==="notes")!;expect(s.getNode(note.rootNodeId)!.text).toBe("中文备注");fail(()=>commit(s,[{op:"insertPresentationNotes",slideId:"slide-1",html:"<p>重复</p>"}]));s.undo();expect(p(s).stories).toEqual([]);
  commit(s,[{op:"insertPresentationObject",parentNodeId:"slide-root-1",geometry,object:{kind:"text",text:"正文"}}]);const id=s.getNode("slide-root-1")!.childNodeIds[0]!;
  commit(s,[{op:"setPresentationTextProperties",nodeId:id,changes:{autoFit:"normal",fontScale:.8,lineSpacingReduction:.1}}]);commit(s,[{op:"setPresentationTextProperties",nodeId:id,changes:{autoFit:"shape"}}]);expect(p(s).objects.find(o=>o.nodeId===id)).toMatchObject({text:{autoFit:"shape",fontScale:1,lineSpacingReduction:0}});
  commit(s,[{op:"setPresentationTextProperties",nodeId:id,changes:{autoFit:null,fontScale:null,lineSpacingReduction:null}}]);expect(p(s).objects.find(o=>o.nodeId===id)).not.toHaveProperty("text");
});

test("creates source-ordered slides and inherited placeholders, reorders without reanchoring stable regions", () => {
  const s = blank(), original = s.snapshot();
  commit(s, [{ op: "insertSlide", layoutId: "layout-body", afterSlideId: "slide-1" }]);
  const created = p(s).slides.find(slide => slide.slideId !== "slide-1")!;
  expect(s.getNode(created.rootNodeId)!.childNodeIds).toHaveLength(2);
  for (const id of s.getNode(created.rootNodeId)!.childNodeIds) expect(p(s).objects.find(o => o.nodeId === id)!.placeholder).toBeDefined();
  commit(s, [{ op: "addAnnotation", annotation: { annotationId: "region", documentId: "edit-ppt", baseRevision: 1, status: "open", instruction: "位置", anchor: { kind: "region", surfaceId: created.rootNodeId, nodeIds: [], rect: { x: 20, y: 20, width: 20, height: 20 }, layoutKey: "measured" } } }]);
  const c = prepare(s, [{ op: "moveSlide", slideId: created.slideId, beforeSlideId: "slide-1" }]);
  expect(c.affectedSlideIds).toEqual([]); s.commit(c); expect(s.getAnnotations()[0]!.status).toBe("open");
  const order = s.getNode("presentation-root")!.childNodeIds.filter(id => p(s).slides.some(slide => slide.rootNodeId === id));
  expect(order).toEqual([created.rootNodeId, "slide-root-1"]);
  expect(p(s).slides[0]!.slideId).toBe("slide-1");
  s.undo(); s.undo(); s.undo(); expect(s.snapshot().html).toBe(original.html);
});

test("slide property changes are atomic and the last slide cannot be removed", () => {
  const s = blank(), original = s.snapshot();
  fail(() => commit(s, [{ op: "setSlideProperties", slideId: "slide-1", changes: { name: "更名" } }, { op: "removeSlide", slideId: "slide-1" }]));
  expect(s.snapshot()).toEqual(original);
  commit(s, [{ op: "setSlideProperties", slideId: "slide-1", changes: { name: "封面", hidden: true } }]);
  expect(p(s).slides[0]).toMatchObject({ name: "封面", hidden: true }); s.undo(); expect(p(s).slides[0]).not.toHaveProperty("hidden");
});

test("duplicates slide objects, notes and independent chart data without cloning review discussions", () => {
  const s = example(), before = s.snapshot(), original = p(s).slides[0]!;
  commit(s, [{ op: "duplicateSlide", slideId: original.slideId }]);
  const copied = p(s).slides.find(slide => slide.slideId !== original.slideId)!;
  expect(s.getNode(copied.rootNodeId)!.text).toBe(s.getNode(original.rootNodeId)!.text);
  expect(p(s).stories.filter(story => story.kind === "notes")).toHaveLength(2);
  expect(s.getAnnotations()).toHaveLength(1); expect(p(s).comments).toHaveLength(1);
  expect(p(s).charts).toHaveLength(2); expect(p(s).dataSets).toHaveLength(2);
  expect(p(s).charts[1]!.dataSetId).not.toBe(p(s).charts[0]!.dataSetId);
  const newIds = new Set(s.getNode(copied.rootNodeId)!.childNodeIds);
  for (const object of p(s).objects.filter(o => newIds.has(o.nodeId))) expect(object).not.toHaveProperty("nativeId");
  const connector = p(s).objects.find(o => o.kind === "connector" && newIds.has(o.nodeId))!;
  expect(connector).toMatchObject({ kind: "connector" });
  if (connector.kind !== "connector") throw Error(); expect(newIds.has(connector.startNodeId!)).toBe(true); expect(newIds.has(connector.endNodeId!)).toBe(true);
  expect(s.serialize().size).toBe(6); expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
  s.undo(); expect(s.snapshot().html).toBe(before.html); s.redo(); expect(p(s).slides).toHaveLength(2);
});

test("slide removal deletes owned notes and objects while retaining recoverable review bodies", () => {
  const s = example(); commit(s, [{ op: "insertSlide", layoutId: "layout-blank" }]);
  commit(s, [{ op: "removeSlide", slideId: "slide-1" }]);
  expect(p(s).slides).toHaveLength(1); expect(s.getNode("speaker-notes")).toBeUndefined(); expect(s.getNode("heading")).toBeUndefined();
  expect(s.getNode("comment-body")).toBeDefined(); expect(s.getAnnotations()[0]!.status).toBe("needs_reanchor");
  expect(p(s).stories.filter(story => story.kind === "comment").every(story => story.slideId === undefined)).toBe(true);
  s.undo(); expect(s.getNode("speaker-notes")).toBeDefined(); expect(s.getAnnotations()[0]!.status).toBe("open");
});

test("creates text, native geometry and merged-table-ready objects with core identities", () => {
  const s = blank();
  commit(s, [
    { op: "insertPresentationObject", parentNodeId: "slide-root-1", geometry, object: { kind: "text", text: "中文𠀀😀葛󠄀\n第二行", style: "font-size:24pt" } },
    { op: "insertPresentationObject", parentNodeId: "slide-root-1", geometry: { ...geometry, y: 200 }, object: { kind: "geometry", preset: "roundRect", text: "里程碑" } },
    { op: "insertPresentationObject", parentNodeId: "slide-root-1", geometry: { ...geometry, x: 400 }, object: { kind: "table", rows: 2, columns: 3 } },
  ]);
  const nodes = s.getNode("slide-root-1")!.childNodeIds; expect(nodes).toHaveLength(3);
  expect(new Set(nodes).size).toBe(3); expect(nodes.every(id => id.startsWith("edit-"))).toBe(true);
  expect(s.getNode(nodes[0]!)!.text).toBe("中文𠀀😀葛󠄀\n第二行");
  expect(p(s).objects.find(o => o.nodeId === nodes[2])).toMatchObject({ kind: "table", grid: [80,80,80] });
  expect(readPresentationGeometry(s.getNode(nodes[0]!)!.attributes)).toEqual(geometry);
  s.undo(); expect(s.getNode("slide-root-1")!.childNodeIds).toEqual([]); s.redo(); expect(s.getNode("slide-root-1")!.childNodeIds).toEqual(nodes);
});

test("object creation errors do not consume source, resources or history and reject guessed same-batch identities", () => {
  const s = blank(), before = s.snapshot();
  fail(() => prepare(s, [{ op: "insertPresentationObject", parentNodeId: "slide-root-1", geometry, object: { kind: "text", text: "临时" } }, { op: "insertPresentationObject", parentNodeId: "edit-1", geometry, object: { kind: "text", text: "猜测" } }]));
  fail(() => prepare(s, [{ op: "insertPresentationObject", parentNodeId: "slide-root-1", geometry, object: { kind: "text", text: "x", style: "position:fixed" } }]), "UnsupportedFeature");
  expect(s.snapshot()).toEqual(before); expect(s.history().undoSteps).toBe(0);
});

test("image crop and shape preset changes preserve identity and share one undo record", () => {
  const s = example(), before = s.snapshot();
  commit(s, [{ op: "setPresentationCrop", nodeId: "project-image", crop: { left: .1, right: .2, top: .1, bottom: 0 } }, { op: "setPresentationPreset", nodeId: "milestone", preset: "diamond" }]);
  expect(s.getNode("project-image")!.attributes["data-office-crop-left"]).toBe("0.1"); expect(p(s).objects.find(o => o.nodeId === "milestone")).toMatchObject({ preset: "diamond" });
  fail(() => prepare(s, [{ op: "setPresentationCrop", nodeId: "project-image", crop: { left: .8, right: .5, top: 0, bottom: 0 } }]));
  s.undo(); expect(s.snapshot().html).toBe(before.html); expect(p(s).objects.find(o => o.nodeId === "milestone")).toMatchObject({ preset: "roundRect" });
});

test("default properties, crop and last-slide moves remain no-ops and shape round trips retain authored adjustments",()=>{
  const s=example(),files=new Map(s.serialize()),source=new TextDecoder().decode(files.get("content.html"));
  files.set("content.html",new TextEncoder().encode(source.replace('data-node-id="milestone"','data-node-id="milestone" data-office-adjust-adj="25000"')));
  const edited=createDocumentSession(openDocument(files)),before=edited.snapshot();
  const receipt=commit(edited,[{op:"setSlideProperties",slideId:"slide-1",changes:{hidden:false}},{op:"moveSlide",slideId:"slide-1"},{op:"setPresentationCrop",nodeId:"project-image",crop:{left:0,top:0,right:0,bottom:0}},{op:"setPresentationPreset",nodeId:"milestone",preset:"roundRect"},{op:"setPresentationPreset",nodeId:"milestone",preset:"diamond"},{op:"setPresentationPreset",nodeId:"milestone",preset:"roundRect"}]);
  expect(receipt.changed).toBe(false);expect(edited.snapshot()).toEqual(before);expect(edited.history().undoSteps).toBe(0);
});

test("multiselect alignment, equal-gap distribution and layering preserve identities and no-op history", () => {
  const s = blank(); commit(s, [10,100,300].map((x,index) => ({ op:"insertPresentationObject",parentNodeId:"slide-root-1",geometry:{...geometry,x,width:[40,80,60][index]!},object:{kind:"text",text:String(index)} })));
  const ids = s.getNode("slide-root-1")!.childNodeIds;
  commit(s,[{op:"distributePresentationObjects",nodeIds:[...ids].reverse(),axis:"horizontal"}]);
  expect(readPresentationGeometry(s.getNode(ids[1]!)!.attributes)!.x).toBeCloseTo(135); s.undo();
  commit(s,[{op:"alignPresentationObjects",nodeIds:[ids[1]!,ids[2]!],alignment:"left",relativeTo:"slide"}]);
  expect(readPresentationGeometry(s.getNode(ids[1]!)!.attributes)!.x).toBe(0); expect(readPresentationGeometry(s.getNode(ids[2]!)!.attributes)!.x).toBe(0); s.undo();
  commit(s,[{op:"reorderPresentationObjects",nodeIds:[ids[2]!],position:"back"}]); expect(s.getNode("slide-root-1")!.childNodeIds).toEqual([ids[2],ids[0],ids[1]]);
  expect(commit(s,[{op:"reorderPresentationObjects",nodeIds:[ids[2]!],position:"back"}]).changed).toBe(false);
  s.undo(); expect(s.getNode("slide-root-1")!.childNodeIds).toEqual(ids);
});

test("object copies receive independent identities and datasets, while existing images share immutable bytes", () => {
  const s = example(), old = new Set(p(s).objects.map(o=>o.nodeId));
  commit(s,[{op:"duplicatePresentationObjects",nodeIds:["project-image","progress-chart"],offset:{x:20,y:10}}]);
  const added = p(s).objects.filter(o=>!old.has(o.nodeId)); expect(added).toHaveLength(2); expect(p(s).dataSets).toHaveLength(2);
  const image = added.find(o=>o.kind==="image")!; expect(readPresentationGeometry(s.getNode(image.nodeId)!.attributes)!.x).toBe(68);
  expect(s.getNode(s.getNode(image.nodeId)!.childNodeIds[0]!)!.attributes.src).toBe("assets/sample.png"); expect(s.snapshot().manifest.assets).toHaveLength(1);
  expect(s.getAnnotations()).toHaveLength(1); s.undo(); expect(p(s).dataSets).toHaveLength(1);
});

test("grouping preserves object identities and ungrouping a unit transform restores all local geometry", () => {
  const s = example(), ids=["group-diamond","group-label"], original=ids.map(id=>readPresentationGeometry(s.getNode(id)!.attributes)!);
  commit(s,[{op:"groupPresentationObjects",nodeIds:ids}]);
  const group = p(s).objects.find(o=>o.kind==="group" && o.nodeId!=="delivery-group")!;
  expect(s.getNode(group.nodeId)!.childNodeIds).toEqual(ids); expect(s.getNode("delivery-group")!.childNodeIds).toEqual([group.nodeId]);
  commit(s,[{op:"ungroupPresentationObject",nodeId:group.nodeId}]);
  expect(s.getNode("delivery-group")!.childNodeIds).toEqual(ids);
  for(const [i,id] of ids.entries()) expect(readPresentationGeometry(s.getNode(id)!.attributes)).toEqual(original[i]);
  s.undo(); expect(s.getNode(group.nodeId)!.childNodeIds).toEqual(ids); s.undo(); expect(s.getNode(group.nodeId)).toBeUndefined();
});

test("ungrouping rotated and flipped groups keeps glyph orientation and refuses actual shear atomically", () => {
  const s=example(), g=readPresentationGeometry(s.getNode("delivery-group")!.attributes)!;
  commit(s,[{op:"setPresentationGeometry",nodeId:"delivery-group",geometry:{...g,rotation:90,flipH:true}}]);
  commit(s,[{op:"ungroupPresentationObject",nodeId:"delivery-group"}]);
  const child=readPresentationGeometry(s.getNode("group-diamond")!.attributes)!;
  expect(child.x).toBeCloseTo(218); expect(child.y).toBeCloseTo(430); expect(child.width).toBeCloseTo(100); expect(child.height).toBeCloseTo(100); expect(child.rotation).toBeCloseTo(90); expect(child.flipH).toBe(true); expect(child.flipV).toBe(false);
  s.undo(); s.undo();
  commit(s,[{op:"setPresentationGeometry",nodeId:"delivery-group",geometry:{...g,width:g.width*2}},{op:"setPresentationGeometry",nodeId:"group-label",geometry:{...readPresentationGeometry(s.getNode("group-label")!.attributes)!,rotation:30}}]);
  const before=s.snapshot(); fail(()=>prepare(s,[{op:"ungroupPresentationObject",nodeId:"delivery-group"}]),"UnsupportedFeature"); expect(s.snapshot()).toEqual(before);
});

test("connector endpoint and routing edits update source caches atomically, detach without a jump and reject invalid sites",()=>{
  const s=example();commit(s,[{op:"setPresentationConnector",nodeId:"milestone-link",changes:{routing:"elbow",end:{nodeId:"group-label",site:1}}}]);
  const connector=p(s).objects.find(o=>o.nodeId==="milestone-link")!;expect(connector).toMatchObject({routing:"elbow",endNodeId:"group-label",endSite:1});
  expect(readPresentationGeometry(s.getNode("milestone-link")!.attributes)).toMatchObject({x:148,y:120,width:20,height:190});
  const geometry=readPresentationGeometry(s.getNode("milestone-link")!.attributes)!;
  commit(s,[{op:"setPresentationConnector",nodeId:"milestone-link",changes:{start:null}}]);expect(readPresentationGeometry(s.getNode("milestone-link")!.attributes)).toEqual(geometry);
  const before=s.snapshot();fail(()=>prepare(s,[{op:"setPresentationConnector",nodeId:"milestone-link",changes:{end:{nodeId:"group-label",site:999}}}]));expect(s.snapshot()).toEqual(before);
  s.undo();s.undo();expect(p(s).objects.find(o=>o.nodeId==="milestone-link")).toMatchObject({routing:"straight",endNodeId:"delivery-group",startNodeId:"milestone"});
});


test("allocated PPT identities do not collide with imported metadata or review identities",()=>{
  const files=new Map(createPresentation({documentId:"seeded-identities"})),decode=new TextDecoder(),encode=new TextEncoder();
  const metadata=PresentationDocumentMetadataSchema.parse(JSON.parse(decode.decode(files.get("metadata.json"))));metadata.presentation.themes[0]!.themeId="edit-1";metadata.presentation.masters[0]!.themeId="edit-1";files.set("metadata.json",encode.encode(JSON.stringify(metadata)));
  files.set("annotations.json",encode.encode(JSON.stringify([{annotationId:"edit-2",documentId:"seeded-identities",baseRevision:0,status:"open",instruction:"保留",anchor:{kind:"node",nodeId:"slide-root-1"}}])));
  const s=createDocumentSession(openDocument(files));commit(s,[{op:"insertPresentationObject",parentNodeId:"slide-root-1",geometry,object:{kind:"text",text:"新对象"}},{op:"addPresentationComment",anchor:{kind:"node",nodeId:"slide-root-1"},html:"<p>新批注</p>"}]);
  const identities=[...p(s).objects.map(o=>o.nodeId),...s.getAnnotations().map(a=>a.annotationId),...p(s).themes.map(t=>t.themeId)];expect(new Set(identities).size).toBe(identities.length);expect(s.getAnnotations().find(a=>a.annotationId==="edit-2")!.instruction).toBe("保留");expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
});

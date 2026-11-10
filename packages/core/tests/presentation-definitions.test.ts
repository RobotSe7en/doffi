import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import { expect,test } from "vitest";
import { createDocumentSession, createPresentation, resolvePresentationGeometry } from "@doffi/core";
import type { DocumentOperation,DocumentSession } from "@doffi/core";
const blank=()=>createDocumentSession(openDocument(createPresentation({documentId:"definitions"})));
const p=(s:DocumentSession)=>s.snapshot().metadata.presentation!;
function prepare(s:DocumentSession,operations:DocumentOperation[]){const{documentId,revision}=s.snapshot().manifest;return s.prepare({protocolVersion:1,documentId,baseRevision:revision,operations});}
const commit=(s:DocumentSession,operations:DocumentOperation[])=>s.commit(prepare(s,operations));

test("new slide placeholders inherit live layout geometry and retain explicit local overrides until reset",()=>{
  const s=blank();commit(s,[{op:"insertSlide",layoutId:"layout-body"}]);const slide=p(s).slides[1]!,id=s.getNode(slide.rootNodeId)!.childNodeIds[0]!,base=resolvePresentationGeometry(s,"layout-title");
  expect(s.getNode(id)!.attributes["data-office-x"]).toBeUndefined();
  const candidate=prepare(s,[{op:"setPresentationGeometry",nodeId:"layout-title",geometry:{...base,x:90}}]);expect(candidate.affectedSlideIds).toEqual([slide.slideId]);s.commit(candidate);expect(resolvePresentationGeometry(s,id).x).toBe(90);
  commit(s,[{op:"setPresentationGeometry",nodeId:id,geometry:{...resolvePresentationGeometry(s,id),x:50}}]);commit(s,[{op:"setPresentationGeometry",nodeId:"layout-title",geometry:{...base,x:120}}]);expect(resolvePresentationGeometry(s,id).x).toBe(50);
  commit(s,[{op:"resetPresentationOverrides",nodeId:id,geometry:true}]);expect(resolvePresentationGeometry(s,id).x).toBe(120);s.undo();expect(resolvePresentationGeometry(s,id).x).toBe(50);
});

test("switching between centered and ordinary title layouts keeps the title identity and live inheritance",()=>{
  const s=blank();commit(s,[{op:"insertSlide",layoutId:"layout-title"}]);const slide=p(s).slides[1]!,id=s.getNode(slide.rootNodeId)!.childNodeIds[0]!;commit(s,[{op:"insertFragment",nodeId:id,position:"end",html:"<p>保留标题</p>"}]);
  commit(s,[{op:"setPresentationSlideLayout",slideId:slide.slideId,layoutId:"layout-body"}]);expect(s.getNode(id)!.text).toBe("保留标题");expect(resolvePresentationGeometry(s,id)).toEqual(resolvePresentationGeometry(s,"layout-title"));expect(p(s).objects.filter(o=>s.getNode(o.nodeId)!.parentNodeId===slide.rootNodeId&&o.placeholder?.index===0)).toHaveLength(1);expect(s.getNode(id)!.attributes["data-office-x"]).toBeUndefined();expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
});

test("theme cloning is isolated, multi-page changes report their actual dependencies, and no-op edits stay no-op",()=>{
  const s=blank(),before=s.snapshot();let candidate=prepare(s,[{op:"insertPresentationDefinition",kind:"theme",copyFromId:"theme-default"}]);expect(candidate.affectedSlideIds).toEqual([]);s.commit(candidate);const theme=p(s).themes[1]!;
  commit(s,[{op:"insertPresentationDefinition",kind:"master",themeId:theme.themeId}]);const master=p(s).masters[1]!;commit(s,[{op:"insertPresentationDefinition",kind:"layout",masterId:master.masterId,name:"第二主题"}]);const layout=p(s).layouts.at(-1)!;
  commit(s,[{op:"insertSlide",layoutId:layout.layoutId},{op:"insertSlide",layoutId:layout.layoutId}]);const ids=p(s).slides.slice(1).map(s=>s.slideId);
  candidate=prepare(s,[{op:"setPresentationTheme",themeId:theme.themeId,properties:{"--office-theme-accent1":"#123456"}}]);expect(candidate.affectedSlideIds).toEqual(ids);s.commit(candidate);
  expect(commit(s,[{op:"setPresentationTheme",themeId:theme.themeId,properties:{"--office-theme-accent1":"#123456"}}]).changed).toBe(false);
  expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());for(let i=0;i<5;i++)s.undo();expect(s.snapshot().html).toBe(before.html);expect(s.snapshot().styles).toEqual(before.styles);
});

test("theme color edits update palette slots without bypassing master color maps, and Latin aliases stay paired",()=>{
  const s=blank();commit(s,[{op:"setStyle",nodeId:"master-root-default",property:"--office-theme-accent1",value:"var(--office-palette-accent2)"},{op:"setPresentationTheme",themeId:"theme-default",properties:{"--office-theme-accent2":"#123456","--office-major-ascii":'"Liberation Serif"'}}]);
  const css=s.snapshot().styles.map(s=>s.css).join("");expect(css).toContain("--office-palette-accent2:#123456");expect(css).toContain("--office-theme-accent2:var(--office-palette-accent2)");expect(s.getNode("master-root-default")!.attributes.style).toContain("--office-theme-accent1:var(--office-palette-accent2)");expect(css).toContain('--office-major-h-ansi:"Liberation Serif"');
  const before=s.snapshot();expect(()=>commit(s,[{op:"setPresentationTheme",themeId:"theme-default",properties:{"--office-theme-accent2":"not-a-color"}}])).toThrow();expect(s.snapshot()).toEqual(before);
  commit(s,[{op:"insertPresentationDefinition",kind:"notes-master"},{op:"insertPresentationNotes",slideId:"slide-1",html:"<p>备注正文</p>"}]);const id=p(s).notesMasters[0]!.notesMasterId;commit(s,[{op:"setPresentationNotesMaster",slideId:"slide-1",notesMasterId:null}]);expect(p(s).stories[0]).not.toHaveProperty("notesMasterId");commit(s,[{op:"setPresentationNotesMaster",slideId:"slide-1",notesMasterId:id}]);expect(p(s).stories[0]).toMatchObject({notesMasterId:id});expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
});

test("layout replacement preserves unmatched text, copied definitions get new native identities, and required references cannot dangle",()=>{
  const s=blank();commit(s,[{op:"insertSlide",layoutId:"layout-body"}]);const slide=p(s).slides[1]!,id=s.getNode(slide.rootNodeId)!.childNodeIds[0]!,g=resolvePresentationGeometry(s,id);
  commit(s,[{op:"insertFragment",nodeId:id,position:"end",html:"<p>用户正文</p>"}]);const before=s.snapshot();expect(()=>commit(s,[{op:"removePresentationDefinition",kind:"layout",definitionId:"layout-body"}])).toThrow();expect(s.snapshot()).toEqual(before);
  commit(s,[{op:"removePresentationDefinition",kind:"layout",definitionId:"layout-body",replacementId:"layout-blank"}]);expect(s.getNode(id)!.text).toBe("用户正文");expect(resolvePresentationGeometry(s,id)).toEqual(g);s.undo();expect(s.snapshot().html).toBe(before.html);
  commit(s,[{op:"insertPresentationDefinition",kind:"layout",copyFromId:"layout-body",name:"副本"}]);const copy=p(s).layouts.at(-1)!;expect(s.getNode(copy.rootNodeId)!.childNodeIds).toHaveLength(2);expect(s.getNode(copy.rootNodeId)!.childNodeIds.some(id=>["layout-title","layout-body"].includes(id))).toBe(false);
  const checkpoint=s.snapshot();expect(()=>commit(s,[{op:"setPresentationSlideLayout",slideId:slide.slideId,layoutId:"guessed"}])).toThrow();expect(()=>commit(s,[{op:"setPresentationTheme",themeId:"theme-default",properties:{"--office-theme-accent1":"red;color:blue"}}])).toThrow();expect(s.snapshot()).toEqual(checkpoint);
});


test("review: repeated theme changes replace declarations instead of accumulating stylesheet rules",()=>{
  const s=blank();commit(s,[{op:"setPresentationTheme",themeId:"theme-default",properties:{"--office-theme-accent1":"#123456"}}]);const length=s.snapshot().styles.map(s=>s.css).join("").length;
  for(let i=0;i<20;i++)commit(s,[{op:"setPresentationTheme",themeId:"theme-default",properties:{"--office-theme-accent1":i%2?"#123456":"#abcdef"}}]);expect(s.snapshot().styles.map(s=>s.css).join("").length).toBe(length);expect(commit(s,[{op:"setPresentationTheme",themeId:"theme-default",properties:{"--office-theme-accent1":"#123456"}}]).changed).toBe(false);expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
});

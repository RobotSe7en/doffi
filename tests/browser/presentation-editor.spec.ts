import { openSdkHost } from './sdk-host.js';
import { readFileSync,readdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { build } from "vite";
import { expect,test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { createDocumentSession,createPresentation,openDocument } from "@doffi/core";
import type { DocumentPackage } from "@doffi/core";
import { renderDocumentPdf } from "@doffi/node";
import { importPptx,readDocumentZip } from "@doffi/office";
import {writeOpcZip} from "../../packages/office/src/document-zip.js";
import type { PresentationEditorHost } from "./fixtures/presentation-editor-host.js";
declare global{interface Window{presentationEditorHost:PresentationEditorHost}}
let code:string;
const selectAll="ControlOrMeta+a",jumpToEnd=process.platform==="darwin"?"Meta+ArrowDown":"Control+End";
test.beforeEach(async({page})=>{await page.setViewportSize({width:1600,height:1000});});
test.beforeAll(async()=>{
  const result=await build({configFile:false,logLevel:"silent",build:{write:false,target:"es2022",lib:{entry:fileURLToPath(new URL("./fixtures/presentation-editor-host.ts",import.meta.url)),formats:["es"]}}});
  const output=Array.isArray(result)?result[0]!.output:"output" in result?result.output:[];code=output.flatMap(item=>item.type==="chunk"?[item.code]:[]).join("\n");
});
function example():DocumentPackage{return new Map([...readdirSync("examples/presentation").filter(name=>/\.(json|html|css)$/.test(name)),"assets/sample.png"].map(path=>[path,new Uint8Array(readFileSync("examples/presentation/"+path))]));}
async function mount(page:Page,files:DocumentPackage=createPresentation({documentId:"editor"})):Promise<void>{await openSdkHost(page);await page.addScriptTag({type:"module",content:code});await page.waitForFunction(()=>!!window.presentationEditorHost);await page.evaluate(files=>window.presentationEditorHost.mount(files),[...files].map(([path,bytes])=>[path,[...bytes]] as [string,number[]]));}
async function settled(page:Page):Promise<void>{await page.evaluate(()=>window.presentationEditorHost.view!.ready);}
const frame=(page:Page)=>page.frameLocator("iframe[data-sdk-presentation-canvas]");

test("a large deck composes a distant slide on demand and edits its shape without composing the whole deck", async ({ page }) => {
  const session = createDocumentSession(openDocument(createPresentation({ documentId: "windowed-slides" }), { kind: "presentation" }));
  const apply = (operations: Parameters<typeof session.prepare>[0]["operations"]) => session.commit(session.prepare({
    protocolVersion: 1, documentId: "windowed-slides", baseRevision: session.snapshot().manifest.revision, operations,
  }));
  apply(Array.from({ length: 27 }, () => ({ op: "insertSlide" as const, layoutId: "layout-blank" })));
  const last = session.snapshot().metadata.presentation!.slides.at(-1)!;
  const geometry = { x: 25, y: 30, width: 180, height: 70, rotation: 0, flipH: false, flipV: false };
  apply([{ op: "insertPresentationObject", parentNodeId: last.rootNodeId, geometry, object: { kind: "geometry", preset: "rect", text: "" } }]);
  const shape = session.snapshot().metadata.presentation!.objects.find(item => item.kind === "geometry")!;
  await mount(page, session.serialize());
  session.dispose();
  const canvas = frame(page), slide = canvas.locator('[data-node-id="' + last.rootNodeId + '"]');
  await expect(slide.locator("svg")).toHaveCount(0);
  await page.getByRole("navigation", { name: "幻灯片与对象导航" }).locator(".slide-list button").last().click();
  await expect(slide).toBeVisible();
  await expect(slide.locator("svg").first()).toBeVisible();
  await page.evaluate(({ nodeId, geometry }) => {
    window.presentationEditorHost.view!.perform([{ op: "setPresentationGeometry", nodeId, geometry }]);
  }, { nodeId: shape.nodeId, geometry: { ...geometry, x: 60 } });
  await settled(page);
  await expect(slide.locator('[data-node-id="' + shape.nodeId + '"]')).toHaveCSS("left", "80px");
  expect(await page.evaluate(() => window.presentationEditorHost.view!.layout().pages.length)).toBe(28);
  expect(await page.evaluate(() => window.presentationEditorHost.errors)).toEqual([]);
});

test("presentation object geometry changes keep unchanged image resources available", async ({ page }) => {
  const files = importPptx(readFileSync("examples/pptx/07-images.pptx"), { documentId: "presentation-image-reuse" });
  await mount(page, files);
  const image = frame(page).locator('[data-office-kind="image"] img').filter({ visible: true }).first();
  await expect(image).toBeVisible();
  const { id, url } = await image.evaluate(node => ({
    id: node.closest<HTMLElement>('[data-office-kind="image"]')!.dataset.nodeId!,
    url: (node as HTMLImageElement).src,
  }));
  await page.evaluate(id => {
    const { view, session } = window.presentationEditorHost, attrs = session!.getNode(id)!.attributes;
    view!.perform([{ op: "setPresentationGeometry", nodeId: id, geometry: {
      x: Number(attrs["data-office-x"]) + 10, y: Number(attrs["data-office-y"]),
      width: Number(attrs["data-office-width"]), height: Number(attrs["data-office-height"]),
      rotation: Number(attrs["data-office-rotation"] ?? 0), flipH: attrs["data-office-flip-h"] === "true",
      flipV: attrs["data-office-flip-v"] === "true",
    } }]);
  }, id);
  await settled(page);
  expect(await image.getAttribute("src")).toBe(url);
  expect(await image.evaluate(async node => {
    await (node as HTMLImageElement).decode();
    return (node as HTMLImageElement).naturalWidth;
  })).toBeGreaterThan(0);
});

test("normal view switches one canvas with live thumbnails and keyboard navigation without editing the document", async ({ page, browserName }) => {
  const files = importPptx(new Uint8Array(readFileSync("examples/pptx/13-multi-slide.pptx")), { documentId: "normal-view" });
  await mount(page, files);
  const nav = page.getByRole("navigation", { name: "幻灯片与对象导航" });
  const cards = nav.locator(".slide-list button");
  const slides = frame(page).locator('article > [data-office-kind="slide"]');
  const slideCount = await cards.count();
  expect(slideCount).toBeGreaterThan(2);
  await expect(slides.filter({ visible: true })).toHaveCount(1);
  await expect(cards.first()).toHaveAttribute("aria-pressed", "true");
  await expect(cards.first().frameLocator("iframe").locator('[data-office-kind="slide"]')).toBeVisible();
  const canvasBox = await page.locator("iframe[data-sdk-presentation-canvas]").boundingBox();
  const firstSlideBox = await slides.first().boundingBox();
  expect(canvasBox!.height).toBeLessThan(firstSlideBox!.height * 2);
  const before = await page.evaluate(() => window.presentationEditorHost.session!.snapshot());
  await cards.nth(1).click();
  await expect(slides.nth(1)).toBeVisible();
  await expect(slides.first()).toBeHidden();
  await cards.nth(1).press("End");
  await expect(cards.last()).toBeFocused();
  await expect(slides.last()).toBeVisible();
  await cards.last().press("PageUp");
  await expect(slides.nth(slideCount - 2)).toBeVisible();
  await cards.nth(slideCount - 2).press("Home");
  await expect(cards.first()).toBeFocused();
  await expect(slides.first()).toBeVisible();
  await slides.first().click({ position: { x: 8, y: 8 } });
  await page.keyboard.press("ArrowRight");
  await expect(slides.nth(1)).toBeVisible();
  await page.keyboard.press("PageDown");
  await expect(slides.nth(2)).toBeVisible();
  await page.keyboard.press("ArrowLeft");
  await expect(slides.nth(1)).toBeVisible();
  await expect(slides.filter({ visible: true })).toHaveCount(1);
  const report = await page.evaluate(() => window.presentationEditorHost.view!.layout());
  expect(report.pages).toHaveLength(slideCount);
  expect(await page.evaluate(() => window.presentationEditorHost.session!.snapshot())).toEqual(before);
  if (browserName === "chromium") {
    const serialized = await page.evaluate(() => [...window.presentationEditorHost.session!.serialize()].map(([path, bytes]) => [path, [...bytes]] as [string, number[]]));
    const pdf = await renderDocumentPdf(openDocument(new Map(serialized.map(([path, bytes]) => [path, new Uint8Array(bytes)])), { kind: "presentation" }), {
      ...(process.env.DOFFI_BROWSER_EXECUTABLE ? { browserExecutable: process.env.DOFFI_BROWSER_EXECUTABLE } : {}),
      ...(process.env.DOFFI_BROWSER_CHANNEL ? { browserChannel: process.env.DOFFI_BROWSER_CHANNEL } : {}),
    });
    expect(pdf.report.pages).toHaveLength(slideCount);
    expect(Buffer.from(pdf.bytes).toString("latin1").match(/\/Type\s*\/Page\b/g)).toHaveLength(slideCount);
  }
});

test("native PPT text typing, paragraphs, formatting and empty shapes keep source identities and history",async({page})=>{
  await mount(page);await page.getByRole("button",{name:"文本框",exact:true}).click();await settled(page);
  await page.getByRole("button",{name:"编辑文字",exact:true}).click();
  const input=frame(page).getByRole("textbox",{name:"幻灯片文字输入"});await input.press(selectAll);await input.pressSequentially("中文ABC");await input.press("Enter");await input.pressSequentially("第二段");await settled(page);
  expect(await page.evaluate(()=>{const h=window.presentationEditorHost,id=h.session!.snapshot().metadata.presentation!.objects.find(o=>o.kind==="text"&&h.session!.getNode(o.nodeId)!.parentNodeId==="slide-root-1")!.nodeId;return h.session!.getNode(id)!.text;})).toBe("中文ABC第二段");
  await input.press(selectAll);await page.getByRole("button",{name:"加粗",exact:true}).click();await settled(page);
  expect(await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().html)).toContain("font-weight:700");
  await page.getByRole("button",{name:"完成输入",exact:true}).click();await page.getByRole("button",{name:"形状",exact:true}).click();await settled(page);await page.getByRole("button",{name:"编辑文字",exact:true}).click();await frame(page).getByRole("textbox").pressSequentially("空形状");await settled(page);
  expect(await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().html)).toContain("空形状");expect(await page.evaluate(()=>window.presentationEditorHost.errors)).toEqual([]);
  expect(await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().html)).not.toMatch(/contenteditable|sdk-ppt|data-sdk-/);
});

test("editing a Chinese presentation reuses glyph probes and keeps current font diagnostics", async ({ page }) => {
  await mount(page, importPptx(readFileSync("examples/pptx/13-multi-slide.pptx"), { documentId: "ppt-font-probe-reuse" }));
  const result = await page.evaluate(async () => {
    const { view, session, errors } = window.presentationEditorHost;
    const iframe = view!.element.shadowRoot!.querySelector<HTMLIFrameElement>("iframe[data-sdk-presentation-canvas]")!;
    const win = iframe.contentWindow as Window & typeof globalThis;
    const context = win.CanvasRenderingContext2D.prototype, read = context.getImageData;
    const before = view!.layout().fonts;
    let readbacks = 0;
    context.getImageData = function (...args: Parameters<typeof read>) { readbacks++; return read.apply(this, args); };
    try {
      view!.perform([{ op: "setSlideProperties", slideId: view!.activeSlide(), changes: { name: "编辑后的页名" } }]);
      await view!.ready;
      return { readbacks, before, after: view!.layout().fonts, revision: session!.snapshot().manifest.revision, errors };
    } finally { context.getImageData = read; }
  });
  expect(result.before.length).toBeGreaterThan(0);
  expect(result).toMatchObject({ readbacks: 0, after: result.before, revision: 1, errors: [] });
});

test("new and imported notes edit the registered story through the same undo boundary",async({page})=>{
  await mount(page);await page.getByRole("button",{name:"编辑备注",exact:true}).click();expect(await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().manifest.revision)).toBe(0);
  await frame(page).getByRole("textbox",{name:"演讲者备注输入"}).pressSequentially("讲者中文");await settled(page);
  expect(await page.evaluate(()=>{const s=window.presentationEditorHost.session!,note=s.snapshot().metadata.presentation!.stories.find(s=>s.kind==="notes")!;return s.getNode(note.rootNodeId)!.text;})).toBe("讲者中文");
  await page.getByRole("button",{name:"完成输入",exact:true}).click();await page.getByRole("button",{name:"撤销",exact:true}).click();await settled(page);
  expect(await page.evaluate(()=>{const s=window.presentationEditorHost.session!,note=s.snapshot().metadata.presentation!.stories.find(s=>s.kind==="notes")!;return s.getNode(note.rootNodeId)!.text;})).toBe("讲者中");
  await mount(page,example());await page.getByRole("button",{name:"编辑备注",exact:true}).click();const input=frame(page).getByRole("textbox",{name:"演讲者备注输入"});await input.press(jumpToEnd);await input.pressSequentially("新增");await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.getNode("speaker-notes")?.text)).toContain("新增");expect(await page.evaluate(()=>window.presentationEditorHost.errors)).toEqual([]);
});

test("PPT real IME input commits once and rejects concurrent source changes",async({page,browserName})=>{
  test.skip(browserName!=="chromium","CDP composition is Chromium-specific; other engines cover native input and history.");
  await mount(page);await page.getByRole("button",{name:"文本框",exact:true}).click();await settled(page);await page.getByRole("button",{name:"编辑文字",exact:true}).click();const input=frame(page).getByRole("textbox");await input.press(selectAll);
  const cdp=await page.context().newCDPSession(page);await cdp.send("Input.imeSetComposition",{text:"中",selectionStart:1,selectionEnd:1});expect(await page.evaluate(()=>window.presentationEditorHost.session!.history().undoSteps)).toBe(1);await cdp.send("Input.imeSetComposition",{text:"中文",selectionStart:2,selectionEnd:2});await cdp.send("Input.insertText",{text:"中文"});await settled(page);
  expect(await page.evaluate(()=>window.presentationEditorHost.session!.history().undoSteps)).toBe(2);expect(await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().html)).toContain("中文");
  await cdp.send("Input.imeSetComposition",{text:"迟到",selectionStart:2,selectionEnd:2});await page.evaluate(()=>window.presentationEditorHost.view!.perform([{op:"setSlideProperties",slideId:window.presentationEditorHost.view!.activeSlide(),changes:{name:"并发"}}]));await cdp.send("Input.insertText",{text:"迟到"});await settled(page);
  await expect(page.getByRole("alert")).toContainText("输入期间文档已变化");expect(await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().html)).not.toContain("迟到");
});

test("table text, TSV paste, spanning row operations and undo use the native table source",async({page})=>{
  await mount(page,example());await page.evaluate(()=>window.presentationEditorHost.view!.editText("progress-table"));
  const first=frame(page).locator('[contenteditable="true"] [data-node-id="table-h1-p"]');await first.click();await page.keyboard.press("Control+a");await page.keyboard.insertText("列名");await settled(page);
  expect(await page.evaluate(()=>window.presentationEditorHost.session!.getNode("table-h1")!.text)).toBe("列名");
  const clipboard=await frame(page).locator('[contenteditable="true"] [data-node-id="table-h1"]').evaluate(node=>{const p=node.querySelector("p")!,range=document.createRange();range.selectNodeContents(p);range.collapse(true);const selection=document.getSelection()!;selection.removeAllRanges();selection.addRange(range);const data=new DataTransfer();data.setData("text/plain","甲\t乙\n丙\t丁");const event=new ClipboardEvent("paste",{clipboardData:data,bubbles:true,cancelable:true});// Firefox sanitizes synthetic ClipboardEvent data; supply the host-authorized test payload explicitly.
Object.defineProperty(event,"clipboardData",{value:data});const result={before:data.getData("text/plain"),event:event.clipboardData?.getData("text/plain")};node.closest("[contenteditable]")!.dispatchEvent(event);return result;});await settled(page);expect(clipboard).toEqual({before:"甲\t乙\n丙\t丁",event:"甲\t乙\n丙\t丁"});
  expect(await page.evaluate(()=>window.presentationEditorHost.session!.getNode("table")!.text)).toBe("甲乙丙丁");
  await page.getByLabel("选区行数",{exact:true}).fill("2");await page.getByRole("button",{name:"合并单元格",exact:true}).click();await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.getNode("table-h1")!.attributes.rowspan)).toBe("2");
  await page.getByRole("button",{name:"拆分单元格",exact:true}).click();await settled(page);await page.getByRole("button",{name:"撤销",exact:true}).click();await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.getNode("table-h1")!.attributes.rowspan)).toBe("2");expect(await page.evaluate(()=>window.presentationEditorHost.errors)).toEqual([]);
});

test("definition canvases edit real layout sources and staged changes list affected slides before one-step apply",async({page})=>{
  await mount(page);await page.getByLabel("新页布局").selectOption("layout-body");await page.getByRole("button",{name:"新增幻灯片",exact:true}).click();await settled(page);
  await page.getByLabel("定义画布",{exact:true}).selectOption("layout-root-body");await page.getByRole("button",{name:"打开定义",exact:true}).click();await settled(page);
  await expect(frame(page).locator('[data-node-id="layout-root-body"]')).toBeVisible();await page.evaluate(()=>window.presentationEditorHost.view!.selectObjects(["layout-title"]));
  await page.getByLabel("X（pt）",{exact:true}).fill("100");await page.getByLabel("X（pt）",{exact:true}).press("Tab");await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.getNode("layout-title")!.attributes["data-office-x"])).toBe("100");
  await page.getByRole("button",{name:"编辑文字",exact:true}).click();await frame(page).getByRole("textbox").pressSequentially("布局标题");await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.getNode("layout-title")!.text)).toBe("布局标题");await page.getByRole("button",{name:"完成输入",exact:true}).click();
  await page.getByRole("button",{name:"返回幻灯片",exact:true}).click();await settled(page);const before=await page.evaluate(()=>window.presentationEditorHost.session!.history().undoSteps);
  await page.getByLabel("主题颜色值",{exact:true}).fill("#123456");await page.getByRole("button",{name:"预览主题颜色",exact:true}).click();expect(await page.evaluate(()=>window.presentationEditorHost.session!.history().undoSteps)).toBe(before);await expect(page.getByText(/待应用：影响 2 页/)).toBeVisible();
  await page.getByRole("button",{name:"应用定义修改",exact:true}).click();await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.history().undoSteps)).toBe(before+1);await page.getByRole("button",{name:"撤销",exact:true}).click();await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().styles.map(s=>s.css).join(""))).not.toContain("#123456");expect(await page.evaluate(()=>window.presentationEditorHost.errors)).toEqual([]);
});

test("chart data tools update native constants and first-party graphics through shared history",async({page})=>{
  await mount(page);await page.getByRole("button",{name:"插入图表",exact:true}).click();await settled(page);await expect(page.getByRole("table",{name:"图表数据"})).toBeVisible();
  await page.getByLabel("系列 1 分类 1",{exact:true}).fill("35");await page.getByLabel("系列 1 分类 1",{exact:true}).press("Tab");await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().metadata.presentation!.dataSets[0]!.series[0]!.values)).toEqual([35,20]);
  const before=await page.evaluate(()=>window.presentationEditorHost.session!.history().undoSteps);await page.getByRole("button",{name:"新增系列",exact:true}).click();await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.history().undoSteps)).toBe(before+1);expect(await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().metadata.presentation!.dataSets[0]!.series)).toHaveLength(2);
  await page.getByLabel("图表类型",{exact:true}).selectOption("line");await settled(page);await expect(frame(page).locator('[data-sdk-chart="line"]')).toHaveCount(1);
  await page.getByLabel("系列新颜色",{exact:true}).fill("#123456");await page.getByRole("button",{name:"设置系列颜色",exact:true}).click();await settled(page);expect(await frame(page).locator("svg[data-sdk-chart] path").evaluateAll(nodes=>nodes.map(node=>getComputedStyle(node).stroke.replace(/\s/g,"")))).toContain("rgb(18,52,86)");
  await page.getByRole("button",{name:"撤销",exact:true}).click();await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.errors)).toEqual([]);
});

for(const format of["classic","modern"])test(`${format} comment rich input, replies, status, location and deletion use the core history`,async({page})=>{
  await mount(page,example());await page.evaluate(()=>window.presentationEditorHost.view!.selectObjects(["heading"]));await page.getByLabel("新批注格式",{exact:true}).selectOption(format);await page.getByLabel("新批注正文",{exact:true}).fill("新增批注");await page.getByRole("button",{name:"添加批注",exact:true}).click();await settled(page);
  const input=frame(page).getByRole("textbox",{name:"批注正文输入"});await input.press(jumpToEnd);await input.pressSequentially("补充");await settled(page);const id=await page.evaluate(()=>window.presentationEditorHost.session!.getAnnotations().at(-1)!.annotationId);
  expect(await page.evaluate(id=>{const s=window.presentationEditorHost.session!;return s.getNode(s.getAnnotations().find(a=>a.annotationId===id)!.bodyNodeId!)!.text;},id)).toBe("新增批注补充");await page.getByRole("button",{name:"完成输入",exact:true}).click();
  const thread=page.locator("fieldset section").filter({has:page.locator("p",{hasText:"新增批注补充"})});await thread.getByLabel("回复正文",{exact:true}).fill("新增回复");await thread.getByRole("button",{name:"添加回复",exact:true}).click();await settled(page);await frame(page).getByRole("textbox",{name:"批注正文输入"}).pressSequentially("追加");await settled(page);await page.getByRole("button",{name:"完成输入",exact:true}).click();
  await thread.getByRole("button",{name:"解决批注",exact:true}).click();await settled(page);expect(await page.evaluate(id=>window.presentationEditorHost.session!.getAnnotations().find(a=>a.annotationId===id)!.status,id)).toBe("resolved");await thread.getByRole("button",{name:"删除回复",exact:true}).click();await settled(page);await page.getByRole("button",{name:"撤销",exact:true}).click();await settled(page);expect(await page.evaluate(id=>window.presentationEditorHost.session!.getAnnotations().find(a=>a.annotationId===id)!.replies,id)).toHaveLength(1);expect(await page.evaluate(()=>window.presentationEditorHost.errors)).toEqual([]);
});

test("slide regions survive page order, unrelated content and notes, then recover after related geometry undo",async({page},testInfo)=>{
  await mount(page,example());
  const measurements=async()=>frame(page).locator('[data-node-id="slide-root-1"]').evaluate(root=>{const origin=root.getBoundingClientRect();return[root,...root.querySelectorAll<HTMLElement>("[data-node-id],[data-office-script],.sdk-ppt-marker,svg,text,path")].map(node=>{const r=node.getBoundingClientRect(),s=getComputedStyle(node);return[node.getAttribute("data-node-id"),node.localName,r.left-origin.left,r.top-origin.top,r.width,r.height,s.fontFamily,s.fontSize,s.fontWeight,s.fontStyle,s.lineHeight,s.letterSpacing,node.getAttribute("d")];});});
  const before=await measurements();
  const key=await page.evaluate(()=>{const h=window.presentationEditorHost;h.view!.selectRegion(h.view!.activeSlide(),{x:30,y:25,width:250,height:50});const anchor=JSON.parse(JSON.stringify(h.view!.selection()!.anchor));h.view!.perform([{op:"addAnnotation",annotation:{annotationId:"region-test",documentId:h.session!.snapshot().manifest.documentId,baseRevision:0,status:"open",instruction:"区域",anchor}}]);return anchor.layoutKey as string;});await settled(page);
  await page.evaluate(()=>{const h=window.presentationEditorHost;h.view!.perform([{op:"insertSlide",layoutId:"layout-blank"}]);h.view!.perform([{op:"moveSlide",slideId:h.session!.snapshot().metadata.presentation!.slides[0]!.slideId}]);h.view!.perform([{op:"replaceText",nodeId:"speaker-notes-p",from:0,to:0,expectedText:"",text:"备注变化"}]);});await settled(page);
  const after=await measurements();await testInfo.attach("region-measurements",{body:JSON.stringify({before,after}),contentType:"application/json"});
  expect(await page.evaluate(()=>window.presentationEditorHost.session!.getAnnotations().find(a=>a.annotationId==="region-test")!.status)).toBe("open");await expect(page.getByText("区域布局或字体已变化，请重新选择区域定位批注。",{exact:true})).toBeHidden();
  expect(await page.evaluate(()=>{const h=window.presentationEditorHost;h.view!.selectRegion(h.session!.snapshot().metadata.presentation!.slides[0]!.slideId,{x:30,y:25,width:250,height:50});const anchor=h.view!.selection()!.anchor;return anchor.kind==="region"?anchor.layoutKey:"";})).toBe(key);
  await page.evaluate(()=>window.presentationEditorHost.view!.perform([{op:"setStyle",nodeId:"heading-p",property:"font-size",value:"40pt"}]));await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.getAnnotations().find(a=>a.annotationId==="region-test")!.status)).toBe("needs_reanchor");await page.getByRole("button",{name:"撤销",exact:true}).click();await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.getAnnotations().find(a=>a.annotationId==="region-test")!.status)).toBe("open");expect(await page.evaluate(()=>window.presentationEditorHost.errors)).toEqual([]);
});

test("region drag is transient, binds one source slide and cancels without history",async({page})=>{
  await mount(page,example());await page.getByRole("button",{name:"框选区域",exact:true}).click();const slide=frame(page).locator('[data-node-id="slide-root-1"]');await slide.scrollIntoViewIfNeeded();const box=await slide.boundingBox();
  await page.mouse.move(box!.x+50,box!.y+60);await page.mouse.down();await page.mouse.move(box!.x+150,box!.y+110,{steps:3});expect(await page.evaluate(()=>window.presentationEditorHost.session!.history().undoSteps)).toBe(0);await page.mouse.up();
  const selection=await page.evaluate(()=>window.presentationEditorHost.view!.selection());expect(selection!.anchor).toMatchObject({kind:"region",surfaceId:"slide-root-1",rect:{width:75,height:37.5}});expect(await page.evaluate(()=>window.presentationEditorHost.session!.history().undoSteps)).toBe(0);
  await page.mouse.move(box!.x+60,box!.y+60);await page.mouse.down();await page.mouse.move(box!.x+170,box!.y+130,{steps:3});await page.keyboard.press("Escape");await page.mouse.up();expect(await page.evaluate(()=>window.presentationEditorHost.session!.history().undoSteps)).toBe(0);expect(await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().html)).not.toContain("data-sdk-ppt");
});

test("PPT text keeps exact repeated-text anchors and deletes CJK variation sequences as a whole",async({page},testInfo)=>{
  await mount(page,example());await page.evaluate(()=>{const h=window.presentationEditorHost;h.view!.perform([{op:"replaceText",nodeId:"heading-p",from:0,to:4,expectedText:"项目进展",text:"甲甲甲甲甲甲甲甲"}]);h.view!.perform([{op:"addAnnotation",annotation:{annotationId:"exact-input",documentId:h.session!.snapshot().manifest.documentId,baseRevision:1,status:"open",instruction:"定位",anchor:{kind:"text",nodeId:"heading-p",from:4,to:6,quote:"甲甲"}}}]);h.view!.editText("heading");});await settled(page);
  const caret=async(offset:number):Promise<void>=>{await frame(page).locator('[contenteditable="true"] [data-node-id="heading-p"]').evaluate((root,offset)=>{const walker=document.createTreeWalker(root,NodeFilter.SHOW_TEXT);let node:Node|null,remaining=offset;while((node=walker.nextNode())){if(remaining<=node.textContent!.length){const range=document.createRange();range.setStart(node,remaining);range.collapse(true);const selection=document.getSelection()!;selection.removeAllRanges();selection.addRange(range);return;}remaining-=node.textContent!.length;}throw new Error("Caret target is missing.");},offset);};
  await caret(1);await page.keyboard.insertText("甲");await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.getAnnotations().find(a=>a.annotationId==="exact-input")!.anchor)).toMatchObject({kind:"text",from:5,to:7,quote:"甲甲"});
  const input=frame(page).getByRole("textbox",{name:"幻灯片文字输入"});await input.press(selectAll);await page.keyboard.insertText("甲𠮷󠄀乙");await settled(page);const state=await page.evaluate(()=>({node:window.presentationEditorHost.session!.getNode("heading-p"),html:window.presentationEditorHost.session!.snapshot().html,errors:window.presentationEditorHost.errors}));await testInfo.attach("replaced-text-state",{body:JSON.stringify(state),contentType:"application/json"});expect(state.node,JSON.stringify({heading:state.html.match(/<div data-node-id="heading".*?<\/div>/)?.[0],errors:state.errors})).toBeDefined();await caret(5);await input.press("Backspace");await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.getNode("heading-p")!.text)).toBe("甲乙");
  await input.press("ControlOrMeta+z");await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.getNode("heading-p")!.text)).toBe("甲𠮷󠄀乙");const revision=await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().manifest.revision);await caret(2);await input.evaluate(node=>node.dispatchEvent(new InputEvent("beforeinput",{inputType:"insertText",data:"X",bubbles:true,cancelable:true})));await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().manifest.revision)).toBe(revision);await expect(page.getByRole("alert")).toContainText("UTF-16");
});

test("slide tools create inherited pages and preserve source order, properties and one-step history",async({page})=>{
  await mount(page);await page.getByLabel("新页布局").selectOption("layout-body");await page.getByRole("button",{name:"新增幻灯片",exact:true}).click();await settled(page);
  await expect(frame(page).locator('[data-office-kind="slide"]')).toHaveCount(2);
  expect(await page.evaluate(()=>window.presentationEditorHost.session!.history().undoSteps)).toBe(1);
  await page.getByLabel("幻灯片名称",{exact:true}).fill("议程");await page.getByLabel("幻灯片名称",{exact:true}).press("Tab");await settled(page);
  await page.getByLabel("隐藏幻灯片",{exact:true}).check();await settled(page);await expect(page.getByRole("navigation").getByRole("button",{name:/议程.*隐藏/})).toBeVisible();
  await page.getByRole("button",{name:"幻灯片上移",exact:true}).click();await settled(page);
  expect(await page.evaluate(()=>{const h=window.presentationEditorHost,p=h.session!.snapshot().metadata.presentation!;return h.session!.getNode(h.session!.getNode(p.slides[0]!.rootNodeId)!.parentNodeId!)!.childNodeIds.filter(id=>p.slides.some(slide=>slide.rootNodeId===id))[0]===p.slides.find(slide=>slide.name==="议程")!.rootNodeId;})).toBe(true);
  await page.getByRole("button",{name:"撤销",exact:true}).click();await settled(page);
  expect(await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().manifest.revision)).toBe(5);
});

for(const zoom of [1,.5])test(`pointer drag at zoom ${zoom} commits one local-pt transaction and cancellation leaves source untouched`,async({page})=>{
  await mount(page,example());await page.evaluate(zoom=>window.presentationEditorHost.view!.setZoom(zoom),zoom);
  const object=frame(page).locator('[data-node-id="milestone"]'),before=await object.boundingBox();
  await page.mouse.move(before!.x+before!.width/2,before!.y+before!.height/2);await page.mouse.down();await page.mouse.move(before!.x+before!.width/2+40,before!.y+before!.height/2+20,{steps:5});
  expect(await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().manifest.revision)).toBe(0);
  await page.mouse.up();await settled(page);
  expect(await page.evaluate(()=>window.presentationEditorHost.session!.history().undoSteps)).toBe(1);
  const x=await page.evaluate(()=>Number(window.presentationEditorHost.session!.getNode("milestone")!.attributes["data-office-x"]));expect(x).toBeCloseTo(48+30/zoom,4);
  const moved=await object.boundingBox();await page.mouse.move(moved!.x+moved!.width/2,moved!.y+moved!.height/2);await page.mouse.down();await page.mouse.move(moved!.x+moved!.width/2+30,moved!.y+moved!.height/2+30,{steps:3});await page.keyboard.press("Escape");await page.mouse.up();
  expect(await page.evaluate(()=>window.presentationEditorHost.session!.history().undoSteps)).toBe(1);expect(await page.evaluate(()=>Number(window.presentationEditorHost.session!.getNode("milestone")!.attributes["data-office-x"]))).toBe(x);
  await page.getByRole("button",{name:"撤销",exact:true}).click();await settled(page);expect((await object.boundingBox())!.x).toBeCloseTo(before!.x,1);
});

test("multiselect, grouping, keyboard movement, resize and rotation share the core history",async({page})=>{
  await mount(page,example());await page.evaluate(()=>window.presentationEditorHost.view!.selectObjects(["group-diamond","group-label"]));
  await page.getByRole("button",{name:"组合",exact:true}).click();await settled(page);await page.getByRole("button",{name:"取消组合",exact:true}).click();await settled(page);
  await page.evaluate(()=>window.presentationEditorHost.view!.selectObjects(["milestone"]));
  const object=frame(page).locator('[data-node-id="milestone"]');await object.click();await page.keyboard.press("ArrowRight");await settled(page);
  expect(await page.evaluate(()=>Number(window.presentationEditorHost.session!.getNode("milestone")!.attributes["data-office-x"]))).toBeCloseTo(49,4);
  const handle=frame(page).getByRole("button",{name:"调整对象大小",exact:true}),box=await handle.boundingBox();await page.mouse.move(box!.x+box!.width/2,box!.y+box!.height/2);await page.mouse.down();await page.mouse.move(box!.x+box!.width/2+40,box!.y+box!.height/2+20,{steps:4});await page.mouse.up();await settled(page);
  expect(await page.evaluate(()=>Number(window.presentationEditorHost.session!.getNode("milestone")!.attributes["data-office-width"]))).toBeCloseTo(230,4);
  await page.getByLabel("旋转（度）",{exact:true}).fill("30");await page.getByLabel("旋转（度）",{exact:true}).press("Tab");await settled(page);
  expect(await page.evaluate(()=>Number(window.presentationEditorHost.session!.getNode("milestone")!.attributes["data-office-rotation"]))).toBe(30);
  expect(await page.evaluate(()=>window.presentationEditorHost.errors)).toEqual([]);
});

test("representable imported group flattening preserves measured shapes and glyph orientation",async({page},testInfo)=>{
  const files=importPptx(new Uint8Array(readFileSync("examples/pptx/05-groups.pptx")),{documentId:"native-groups"});await mount(page,files);
  const model=openDocument(files),p=model.snapshot().metadata.presentation!,group=p.objects.find(o=>o.kind==="group"&&model.getNode(o.nodeId)!.parentNodeId===p.slides[0]!.rootNodeId)!;
  const boxes=async()=>frame(page).locator('[data-office-kind="slide"] [data-office-kind="text"]').evaluateAll(nodes=>nodes.map(node=>{const b=node.getBoundingClientRect(),text=node.querySelector<HTMLElement>(".sdk-ppt-text")!,s=getComputedStyle(text.querySelector("[data-office-script]")!);return{x:b.x,y:b.y,width:b.width,height:b.height,fontSize:s.fontSize,text:text.textContent};}));
  const before=await boxes();await page.evaluate(id=>window.presentationEditorHost.view!.selectObjects([id]),group.nodeId);await page.getByRole("button",{name:"取消组合",exact:true}).click();await settled(page);
  const after=await boxes();expect(after).toHaveLength(before.length);
  await testInfo.attach("group-flattening-bounds",{body:JSON.stringify({before,after}),contentType:"application/json"});
  // Native geometry is checked precisely in core/Office tests. Display bounds
  // remain within half a CSS pixel after changing the transform hierarchy.
  for(const[index,box]of after.entries()){const old=before[index]!;expect(box.text).toBe(old.text);expect(box.fontSize).toBe(old.fontSize);for(const key of["x","y","width","height"]as const)expect(Math.abs(box[key]-old[key]),JSON.stringify({before,after,key})).toBeLessThan(.5);}
});

test("nested transformed group drags use the actual parent space and concurrent commits cancel pending gestures",async({page})=>{
  await mount(page,example());await page.evaluate(()=>{const h=window.presentationEditorHost,node=h.session!.getNode("delivery-group")!,a=node.attributes;h.view!.perform([{op:"setPresentationGeometry",nodeId:node.nodeId,geometry:{x:Number(a["data-office-x"]),y:Number(a["data-office-y"]),width:800,height:210,rotation:25,flipH:true,flipV:false}}]);});await settled(page);
  const object=frame(page).locator('[data-node-id="group-label"]'),box=await object.boundingBox();await page.mouse.move(box!.x+box!.width/2,box!.y+box!.height/2);await page.mouse.down();await page.mouse.move(box!.x+box!.width/2+40,box!.y+box!.height/2+20,{steps:5});await page.mouse.up();await settled(page);
  const a=await page.evaluate(()=>window.presentationEditorHost.session!.getNode("group-label")!.attributes),r=25*Math.PI/180;
  expect(Number(a["data-office-x"])).toBeCloseTo(120-(30*Math.cos(r)+15*Math.sin(r))/2,3);expect(Number(a["data-office-y"])).toBeCloseTo((-30*Math.sin(r)+15*Math.cos(r))/1.5,3);
  const milestone=frame(page).locator('[data-node-id="milestone"]'),start=await milestone.boundingBox();await page.mouse.move(start!.x+start!.width/2,start!.y+start!.height/2);await page.mouse.down();await page.mouse.move(start!.x+start!.width/2+30,start!.y+start!.height/2+20,{steps:3});
  await page.evaluate(()=>window.presentationEditorHost.view!.perform([{op:"setStyle",nodeId:"heading-p",property:"font-size",value:"31pt"}]));await settled(page);await page.mouse.up();
  expect(await page.evaluate(()=>Number(window.presentationEditorHost.session!.getNode("milestone")!.attributes["data-office-x"]))).toBe(48);expect(await page.evaluate(()=>window.presentationEditorHost.session!.history().undoSteps)).toBe(3);
});

test("image insertion, native crop and undo restore the authorized resource manifest",async({page})=>{
  await mount(page,example());await page.locator('input[type="file"]').setInputFiles("examples/presentation/assets/sample.png");
  await expect.poll(()=>page.evaluate(()=>window.presentationEditorHost.session!.snapshot().manifest.assets.length)).toBe(2);await settled(page);
  await page.getByLabel("左裁剪（%）",{exact:true}).fill("25");await page.getByLabel("左裁剪（%）",{exact:true}).press("Tab");await settled(page);
  expect(await page.evaluate(()=>{const h=window.presentationEditorHost,anchor=h.view!.selection()!.anchor;return anchor.kind==="node"?h.session!.getNode(anchor.nodeId)!.attributes["data-office-crop-left"]:undefined;})).toBe("0.25");
  await page.getByRole("button",{name:"撤销",exact:true}).click();await settled(page);await page.getByRole("button",{name:"撤销",exact:true}).click();await settled(page);
  expect(await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().manifest.assets.length)).toBe(1);
  expect(await page.evaluate(()=>window.presentationEditorHost.errors)).toEqual([]);
});

test("appearance edits preserve gradient stops and unrelated line alpha through atomic style transactions",async({page})=>{
  await mount(page,importPptx(new Uint8Array(readFileSync("examples/pptx/04-shapes.pptx")),{documentId:"appearance"}));
  const node=frame(page).locator('[data-office-kind="geometry"]:has(linearGradient)').first(),id=(await node.getAttribute("data-node-id"))!;
  await page.evaluate(id=>{const h=window.presentationEditorHost;h.view!.selectObjects([id]);h.view!.perform([
    {op:"setStyle",nodeId:id,property:"--office-fill-stop-count",value:"3"},{op:"setStyle",nodeId:id,property:"--office-fill-stop-1-position",value:"0.5"},{op:"setStyle",nodeId:id,property:"--office-fill-stop-1-color",value:"#ff0000"},
    {op:"setStyle",nodeId:id,property:"--office-fill-stop-2-position",value:"1"},{op:"setStyle",nodeId:id,property:"--office-fill-stop-2-color",value:"#ffffff"},{op:"setStyle",nodeId:id,property:"--office-line-alpha",value:"0.5"}]);},id);await settled(page);
  await page.getByLabel("不透明度（%）",{exact:true}).fill("40");await page.getByLabel("不透明度（%）",{exact:true}).press("Tab");await settled(page);
  await expect(node.locator("linearGradient stop")).toHaveCount(3);expect(await node.locator("linearGradient stop").nth(1).getAttribute("offset")).toBe("0.5");
  await page.getByLabel("渐变结束颜色",{exact:true}).evaluate(input=>{(input as HTMLInputElement).value="#00ff00";input.dispatchEvent(new Event("change",{bubbles:true}));});await settled(page);
  expect(await node.locator("linearGradient stop").last().getAttribute("stop-color")).toBe("rgba(0,255,0,0.4)");
  await page.getByLabel("轮廓宽度（pt）",{exact:true}).fill("2.5");await page.getByLabel("轮廓宽度（pt）",{exact:true}).press("Tab");await settled(page);
  expect(await node.locator(":scope>svg").locator(":scope>path").first().getAttribute("stroke-width")).toBe("2.5");
  expect(await page.evaluate(id=>window.presentationEditorHost.session!.getNode(id)!.attributes.style,id)).toContain("--office-line-alpha:0.5");
  await page.getByLabel("基础外阴影",{exact:true}).check();await settled(page);expect(await node.evaluate(node=>getComputedStyle(node).filter)).toContain("drop-shadow");
});

test("connector controls edit native bindings and routing while rejecting invalid endpoints",async({page})=>{
  await mount(page,example());await page.evaluate(()=>window.presentationEditorHost.view!.selectObjects(["milestone-link"]));
  await page.getByLabel("终点对象",{exact:true}).selectOption("group-label");await settled(page);
  await page.getByLabel("连接线类型",{exact:true}).selectOption("elbow");await settled(page);
  expect(await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().metadata.presentation!.objects.find(o=>o.nodeId==="milestone-link"))).toMatchObject({routing:"elbow",endNodeId:"group-label"});
  await page.getByLabel("起点对象",{exact:true}).selectOption("");await settled(page);
  const before=await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().manifest.revision);
  await page.getByLabel("终点连接位置",{exact:true}).fill("999");await page.getByLabel("终点连接位置",{exact:true}).press("Tab");
  expect(await page.evaluate(()=>window.presentationEditorHost.session!.snapshot().manifest.revision)).toBe(before);await expect(page.getByRole("alert")).toBeVisible();
});

test("PPT undo maps the live caret before continuing beside a CJK variation sequence",async({page})=>{
  await mount(page,example());await page.evaluate(()=>{const h=window.presentationEditorHost;h.view!.perform([{op:"replaceText",nodeId:"heading-p",from:0,to:4,expectedText:"项目进展",text:"甲𠮷󠄀乙"}]);h.view!.editText("heading");});await settled(page);
  const input=frame(page).getByRole("textbox",{name:"幻灯片文字输入"});await frame(page).locator('[data-node-id="heading-p"]').evaluate(p=>{const walker=document.createTreeWalker(p,NodeFilter.SHOW_TEXT),text=walker.nextNode()!,range=document.createRange();range.setStart(text,1);range.collapse(true);const s=document.getSelection()!;s.removeAllRanges();s.addRange(range);});
  await page.keyboard.insertText("X");await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.getNode("heading-p")!.text)).toBe("甲X𠮷󠄀乙");await input.press("Control+z");await settled(page);await page.keyboard.insertText("Y");await settled(page);expect(await page.evaluate(()=>window.presentationEditorHost.session!.getNode("heading-p")!.text)).toBe("甲Y𠮷󠄀乙");expect(await page.evaluate(()=>window.presentationEditorHost.errors)).toEqual([]);
});

test("table cell formatting applies to the selected grid range and keeps source cell identities through undo",async({page})=>{
  await mount(page,example());await page.evaluate(()=>window.presentationEditorHost.view!.selectObjects(["progress-table"]));await page.getByLabel("选区列数",{exact:true}).fill("2");await page.getByLabel("单元格填充颜色",{exact:true}).fill("#123456");await page.getByRole("button",{name:"设置单元格填充",exact:true}).click();await settled(page);
  await expect(frame(page).locator('[data-node-id="table-h1"]')).toHaveCSS("background-color","rgb(18, 52, 86)");await expect(frame(page).locator('[data-node-id="table-h2"]')).toHaveCSS("background-color","rgb(18, 52, 86)");
  await page.getByLabel("单元格垂直对齐",{exact:true}).selectOption("bottom");await page.getByRole("button",{name:"设置单元格垂直对齐",exact:true}).click();await settled(page);await expect(frame(page).locator('[data-node-id="table-h1"]')).toHaveCSS("vertical-align","bottom");
  await page.getByLabel("单元格边框颜色",{exact:true}).fill("#abcdef");await page.getByRole("button",{name:"设置单元格边框",exact:true}).click();await settled(page);await expect(frame(page).locator('[data-node-id="table-h1"]')).toHaveCSS("border-left-color","rgb(171, 205, 239)");
  await page.getByRole("button",{name:"撤销",exact:true}).click();await settled(page);await expect(frame(page).locator('[data-node-id="table-h1"]')).toHaveCSS("background-color","rgb(18, 52, 86)");expect(await page.evaluate(()=>window.presentationEditorHost.session!.getNode("table-h1")!.tagName)).toBe("th");expect(await page.evaluate(()=>window.presentationEditorHost.errors)).toEqual([]);
});

test("PPT text wrapping is a source override and can return to inherited layout behavior",async({page})=>{
  await mount(page,example());await page.evaluate(()=>window.presentationEditorHost.view!.selectObjects(["heading"]));await page.getByLabel("文本换行",{exact:true}).selectOption("false");await settled(page);expect(await page.evaluate(()=>{const object=window.presentationEditorHost.session!.snapshot().metadata.presentation!.objects.find(o=>o.nodeId==="heading")!;return object.kind==="text"?object.text?.wrap:undefined;})).toBe(false);
  await page.getByLabel("文本换行",{exact:true}).selectOption("inherit");await settled(page);expect(await page.evaluate(()=>{const object=window.presentationEditorHost.session!.snapshot().metadata.presentation!.objects.find(o=>o.nodeId==="heading")!;return object.kind==="text"?object.text?.wrap:undefined;})).toBeUndefined();expect(await page.evaluate(()=>window.presentationEditorHost.errors)).toEqual([]);
});

test("review: theme editing respects important shared rules without changing the other selector",async({page})=>{
  const files=new Map(example()),decode=new TextDecoder(),encode=new TextEncoder();files.set("styles.css",encode.encode(decode.decode(files.get("styles.css"))+'.theme-default,.shared-theme{--office-palette-accent1:#112233!important;--office-theme-accent1:var(--office-palette-accent1)!important}'));
  await mount(page,files);await page.evaluate(()=>window.presentationEditorHost.view!.perform([{op:"setStyle",nodeId:"milestone",property:"--office-fill-kind",value:"solid"},{op:"setStyle",nodeId:"milestone",property:"--office-fill-color",value:"var(--office-theme-accent1)"}]));await settled(page);await expect(frame(page).locator('[data-node-id="milestone"]>svg>path')).toHaveCSS("fill","rgb(17, 34, 51)");
  await page.getByLabel("主题颜色值",{exact:true}).fill("#abcdef");await page.getByRole("button",{name:"预览主题颜色",exact:true}).click();await page.getByRole("button",{name:"应用定义修改",exact:true}).click();await settled(page);await expect(frame(page).locator('[data-node-id="milestone"]>svg>path')).toHaveCSS("fill","rgb(171, 205, 239)");
  const other=await frame(page).locator("body").evaluate(body=>{const probe=document.createElement("span");probe.className="shared-theme";probe.style.color="var(--office-palette-accent1)";body.append(probe);const value=getComputedStyle(probe).color;probe.remove();return value;});expect(other).toBe("rgb(17, 34, 51)");expect(await page.evaluate(()=>window.presentationEditorHost.errors)).toEqual([]);
});

test("review: opening a definition canvas does not treat hidden slide regions as a changed font environment",async({page})=>{
  await mount(page,example());await page.evaluate(()=>{const h=window.presentationEditorHost;h.view!.selectRegion(h.view!.activeSlide(),{x:30,y:25,width:250,height:50});h.view!.perform([{op:"addAnnotation",annotation:{annotationId:"definition-region",documentId:h.session!.snapshot().manifest.documentId,baseRevision:0,status:"open",instruction:"区域",anchor:JSON.parse(JSON.stringify(h.view!.selection()!.anchor))}}]);});await settled(page);
  await page.evaluate(()=>{const h=window.presentationEditorHost;h.view!.setDefinitionSurface(h.session!.snapshot().metadata.presentation!.layouts[0]!.rootNodeId);});await settled(page);await expect(page.getByText("区域布局或字体已变化，请重新选择区域定位批注。",{exact:true})).toBeHidden();expect(await page.evaluate(()=>window.presentationEditorHost.session!.getAnnotations().find(a=>a.annotationId==="definition-region")!.status)).toBe("open");
  await page.evaluate(()=>window.presentationEditorHost.view!.setDefinitionSurface());await settled(page);await expect(page.getByText("区域布局或字体已变化，请重新选择区域定位批注。",{exact:true})).toBeHidden();expect(await page.evaluate(()=>window.presentationEditorHost.errors)).toEqual([]);
});


test("review: editing native theme paints keeps visible alpha and applies direct colors",async({page})=>{
  const parts=new Map(readDocumentZip(readFileSync("examples/pptx/04-shapes.pptx"))),decoder=new TextDecoder(),encoder=new TextEncoder(),reference='<p:style><a:lnRef idx="1"><a:srgbClr val="204060"><a:alpha val="35000"/></a:srgbClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr></a:fillRef><a:effectRef idx="0"><a:srgbClr val="000000"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>';
  parts.set("ppt/slides/slide1.xml",encoder.encode(decoder.decode(parts.get("ppt/slides/slide1.xml")).replace(/<p:sp>[\s\S]*?<\/p:sp>/,shape=>shape.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/,"").replace(/<a:ln\b[\s\S]*?<\/a:ln>/,"").replace("<p:txBody>",reference+"<p:txBody>"))));await mount(page,importPptx(writeOpcZip(parts),{documentId:"theme-paint-edit"}));
  const id=await page.evaluate(()=>{const h=window.presentationEditorHost,p=h.session!.snapshot().metadata.presentation!,id=p.objects.find(o=>o.kind==="geometry"&&h.session!.getNode(o.nodeId)!.parentNodeId===p.slides[0]!.rootNodeId)!.nodeId;h.view!.selectObjects([id]);return id;}),path=frame(page).locator('[data-node-id="'+id+'"]>svg>path').first();await expect(path).toHaveAttribute("fill","rgba(68,114,196,0.5)");
  await page.getByLabel("不透明度（%）",{exact:true}).fill("75");await page.getByLabel("不透明度（%）",{exact:true}).press("Tab");await settled(page);await expect(path).toHaveAttribute("fill","rgba(68,114,196,0.75)");await page.getByRole("button",{name:"撤销",exact:true}).click();await settled(page);
  await page.getByLabel("填充颜色",{exact:true}).fill("#123456");await page.getByLabel("填充颜色",{exact:true}).press("Tab");await settled(page);await expect(path).toHaveAttribute("fill","rgba(18,52,86,0.5)");
  await page.getByLabel("轮廓颜色",{exact:true}).fill("#abcdef");await page.getByLabel("轮廓颜色",{exact:true}).press("Tab");await settled(page);await expect(path).toHaveAttribute("stroke","rgba(171,205,239,0.35)");expect(await page.evaluate(()=>window.presentationEditorHost.errors)).toEqual([]);
});

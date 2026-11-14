import { openSdkHost } from './sdk-host.js';
import { fileURLToPath } from "node:url";
import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { build } from "vite";
import type { DocumentPackage } from "@doffi/core";
import type { DocumentPreviewOptions } from "@doffi/editor";
import { fieldsFiles, longTable, pagedFiles } from "./fixtures/pagination-files.js";
import type { PaginationHost } from "./fixtures/pagination-host.js";
import { createDocumentSession,createWordDocument,openDocument as openWordDocument } from '@doffi/core';
declare global { interface Window { paginationHost: PaginationHost } }
let code: string;
test.beforeAll(async () => {
  const result = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: {
    entry: fileURLToPath(new URL("./fixtures/pagination-host.ts", import.meta.url)), formats: ["es"],
  } } });
  const output = Array.isArray(result) ? result[0]!.output : "output" in result ? result.output : [];
  code = output.flatMap((item) => item.type === "chunk" ? [item.code] : []).join("\n");
});
async function mount(page: Page, files: DocumentPackage, mode: "preview" | "edit" | "interaction" = "preview", options: DocumentPreviewOptions = {}): Promise<void> {
  await openSdkHost(page); await page.addScriptTag({ type: "module", content: code }); await page.waitForFunction(() => !!window.paginationHost);
  await page.evaluate(({ entries, mode, options }) => window.paginationHost.mount(entries, mode, options), {
    entries: [...files].map(([path, bytes]) => [path, [...bytes]] as [string, number[]]), mode, options,
  });
}
function columnFiles(table=false):DocumentPackage{
 const session=createDocumentSession(openWordDocument(createWordDocument({documentId:'columns',pageSetup:{unit:'pt',width:300,height:240,margins:{top:24,right:24,bottom:24,left:24}}}),{kind:'document'}));
 const content=table?'<table style="width:100pt;border-collapse:collapse"><tr data-word-repeat-header="true"><th style="border:1pt solid black"><p>HEADER</p></th></tr>'+Array.from({length:32},(_,n)=>'<tr><td style="border:1pt solid black;padding:0"><p>ROW-'+n+'</p></td></tr>').join('')+'</table>':'<p style="line-height:16pt;margin:0">'+Array.from({length:40},(_,n)=>'LINE-'+n+' 😀').join('<br>')+'</p>';
 session.commit(session.prepare({protocolVersion:1,documentId:'columns',baseRevision:0,operations:[{op:'replaceFragment',nodeId:'section-root-1',html:content},{op:'setWordSection',nodeId:'section-root-1',changes:{columns:{count:2,gap:24,separator:true}}}]}));return session.serialize();
}
test('bookmark and hyperlink panels apply typed source changes and retain their text selection',async({page})=>{
 const session=createDocumentSession(openWordDocument(createWordDocument({documentId:'link-controls'}),{kind:'document'}));session.commit(session.prepare({protocolVersion:1,documentId:'link-controls',baseRevision:0,operations:[{op:'replaceText',nodeId:'paragraph-1',from:0,to:0,expectedText:'',text:'链接文字'}]}));await mount(page,session.serialize(),'edit');const frame=page.frameLocator('iframe'),paragraph=frame.locator('[data-node-id="paragraph-1"]');await paragraph.evaluate(node=>{(node.closest('[data-office-editor]')as HTMLElement).focus();const range=document.createRange();range.selectNodeContents(node);const selection=document.getSelection()!;selection.removeAllRanges();selection.addRange(range);});await page.evaluate(()=>(window.paginationHost.view as import('@doffi/editor').DocumentTextEditor).execute({type:'word',action:'bookmarks'}));await page.getByLabel('书签名称',{exact:true}).fill('LinkedText');await page.getByRole('button',{name:'添加书签',exact:true}).click();await page.evaluate(()=>(window.paginationHost.view as import('@doffi/editor').DocumentTextEditor).execute({type:'word',action:'insertLink'}));await page.getByLabel('链接地址',{exact:true}).fill('https://example.com');await page.getByRole('button',{name:'插入链接',exact:true}).click();const result=await page.evaluate(()=>({snapshot:window.paginationHost.session!.snapshot(),text:window.paginationHost.session!.getNode('paragraph-1')!.text,errors:window.paginationHost.errors}));expect(result.snapshot.metadata.word!.bookmarks[0]!.name).toBe('LinkedText');expect(result.snapshot.metadata.word!.hyperlinks[0]!.target).toBe('https://example.com');expect(result.text).toBe('链接文字');expect(result.snapshot.html).not.toMatch(/contenteditable|sdk-page/);expect(result.errors).toEqual([]);
});
test('bookmark references and caption fields derive current values and explicit cache updates share undo history',async({page})=>{
 const session=createDocumentSession(openWordDocument(createWordDocument({documentId:'reference-fields'}),{kind:'document'}));session.commit(session.prepare({protocolVersion:1,documentId:'reference-fields',baseRevision:0,operations:[{op:'replaceText',nodeId:'paragraph-1',from:0,to:0,expectedText:'',text:'引用文字'},{op:'insertWordBookmark',nodeId:'paragraph-1',from:0,to:4,expectedText:'引用文字',name:'Target'},{op:'insertFragment',nodeId:'paragraph-1',position:'after',html:'<p>引用：</p>'}]}));const paragraph=session.getNode('section-root-1')!.childNodeIds[1]!;session.commit(session.prepare({protocolVersion:1,documentId:'reference-fields',baseRevision:1,operations:[{op:'insertWordField',nodeId:paragraph,instruction:'REF Target \\h',cachedText:'cache'},{op:'insertWordField',nodeId:paragraph,instruction:'PAGEREF Target \\h',cachedText:'cache'},{op:'insertWordField',nodeId:paragraph,instruction:'SEQ Figure',cachedText:'cache'}]}));await mount(page,session.serialize(),'edit');const frame=page.frameLocator('iframe');await expect(frame.locator('.sdk-field-display').nth(0)).toHaveText('引用文字');await expect(frame.locator('.sdk-field-display').nth(1)).toHaveText('1');await expect(frame.locator('.sdk-field-display').nth(2)).toHaveText('1');const before=await page.evaluate(()=>window.paginationHost.session!.snapshot());expect(before.html).toContain('cache');await page.evaluate(()=>window.paginationHost.updateFields());const after=await page.evaluate(()=>window.paginationHost.session!.snapshot());expect(after.html).not.toContain('cache');expect(after.html).not.toContain('sdk-field-display');await page.evaluate(()=>window.paginationHost.session!.undo());expect((await page.evaluate(()=>window.paginationHost.session!.snapshot())).html).toBe(before.html);expect(await page.evaluate(()=>window.paginationHost.errors)).toEqual([]);
});
function notesFiles(long=false):DocumentPackage{
 const session=createDocumentSession(openWordDocument(createWordDocument({documentId:'word-notes',pageSetup:{unit:'pt',width:300,height:240,margins:{top:24,right:24,bottom:24,left:24}}}),{kind:'document'}));session.commit(session.prepare({protocolVersion:1,documentId:'word-notes',baseRevision:0,operations:[{op:'replaceText',nodeId:'paragraph-1',from:0,to:0,expectedText:'',text:'正文前后'},{op:'insertWordNote',nodeId:'paragraph-1',offset:2,kind:'footnote',text:long?Array.from({length:60},(_,index)=>'长脚注 '+index).join('\n'):'脚注文字'},{op:'insertWordNote',nodeId:'paragraph-1',offset:0,kind:'endnote',text:'尾注文字',customMark:'*'}]}));return session.serialize();
}
test('footnotes reserve page space, display editable auxiliary stories and retain native reference atoms',async({page})=>{
 await mount(page,notesFiles(),'edit');const frame=page.frameLocator('iframe');await expect(frame.locator('[data-word-note-reference]')).toHaveCount(2);await expect(frame.locator('.sdk-page-notes')).toHaveCount(1);await expect(frame.locator('.sdk-page-notes')).toContainText('脚注文字');const before=await page.evaluate(()=>window.paginationHost.session!.snapshot());await frame.locator('[data-node-id="paragraph-1"]').click();await page.evaluate(()=>(window.paginationHost.view as import('@doffi/editor').DocumentTextEditor).execute({type:'word',action:'referenceSettings'}));await page.getByRole('button',{name:'编辑 ',exact:false}).first().click();const input=frame.locator('[data-office-editor] p').first();await input.click();await page.keyboard.press('End');await page.keyboard.insertText('补充');await page.evaluate(()=>(window.paginationHost.view as import('@doffi/editor').DocumentTextEditor).execute({type:'word',action:'mainStory'}));await expect(frame.locator('.sdk-page-notes')).toContainText('补充');const result=await page.evaluate(()=>({snapshot:window.paginationHost.session!.snapshot(),text:window.paginationHost.session!.getNode('paragraph-1')!.text,errors:window.paginationHost.errors}));expect(result.snapshot.metadata.word!.noteReferences).toEqual(before.metadata.word!.noteReferences);expect(result.snapshot.html).not.toMatch(/sdk-page-notes|sdk-word-notes|contenteditable/);expect(result.errors).toEqual([]);
});
test('long footnotes continue across measured pages without truncating the saved source',async({page})=>{
 await mount(page,notesFiles(true));const frame=page.frameLocator('iframe');expect(await frame.locator('.sdk-page-notes').count()).toBeGreaterThan(2);const report=await page.evaluate(()=>window.paginationHost.layout());expect(report.pages.length).toBeGreaterThan(2);await expect(frame.locator('.sdk-word-notes')).toContainText('长脚注 59');expect(await page.evaluate(()=>window.paginationHost.errors)).toEqual([]);
});
function drawingFiles():DocumentPackage{
 const session=createDocumentSession(openWordDocument(createWordDocument({documentId:'word-drawings'}),{kind:'document'}));session.commit(session.prepare({protocolVersion:1,documentId:'word-drawings',baseRevision:0,operations:[{op:'replaceText',nodeId:'paragraph-1',from:0,to:0,expectedText:'',text:'前后正文'},{op:'insertWordDrawing',nodeId:'paragraph-1',offset:1,geometry:{x:30,y:20,width:120,height:75,rotation:0,flipH:false,flipV:false},layout:{mode:'floating',wrap:'none'},object:{kind:'shape',preset:'roundRect',text:'形状文字'}},{op:'insertWordDrawing',nodeId:'paragraph-1',offset:0,geometry:{x:230,y:20,width:120,height:75,rotation:0,flipH:false,flipV:false},layout:{mode:'floating',wrap:'none'},object:{kind:'textBox',text:'文本框文字'}}]}));return session.serialize();
}
test('Word drawing composition preserves source atoms while neighbouring native input edits',async({page})=>{
 await mount(page,drawingFiles(),'edit');const frame=page.frameLocator('iframe'),paragraph=frame.locator('[data-node-id="paragraph-1"]');await expect(frame.locator('.sdk-word-drawing-text')).toHaveCount(2);await expect(frame.locator('.sdk-word-drawing-text').first()).toContainText('文本框文字');expect(await frame.locator('[data-word-drawing="shape"]').evaluate(node=>Math.round(node.getBoundingClientRect().width))).toBe(160);
 await paragraph.evaluate(node=>{const walker=document.createTreeWalker(node,4);let text=walker.nextNode();while(text&&text.parentElement?.closest('[data-word-drawing]'))text=walker.nextNode();if(!text)throw new Error('Missing surrounding source text');const range=document.createRange();range.setStart(text,1);range.collapse(true);(node.closest('[data-office-editor]')as HTMLElement).focus();const selection=document.getSelection()!;selection.removeAllRanges();selection.addRange(range);});await page.keyboard.insertText('EDIT');
 const result=await page.evaluate(()=>({snapshot:window.paginationHost.session!.snapshot(),text:window.paginationHost.session!.getNode('paragraph-1')!.text,errors:window.paginationHost.errors}));expect(result.text).toContain('EDIT');expect(result.snapshot.metadata.word!.drawings).toHaveLength(2);expect(result.snapshot.html).not.toMatch(/sdk-word|sdk-preset|contenteditable|<svg|position:absolute/);expect(result.errors).toEqual([]);
});
test('Word objects drag and resize through one-step history and edit their text in the shared story session',async({page})=>{
 await mount(page,drawingFiles(),'edit');const frame=page.frameLocator('iframe'),shape=frame.locator('[data-word-drawing="shape"]'),before=await page.evaluate(()=>window.paginationHost.session!.snapshot()),object=before.metadata.word!.drawings!.find(item=>item.kind==='shape')!,box=await shape.boundingBox();expect(box).not.toBeNull();await page.mouse.move(box!.x+20,box!.y+20);await page.mouse.down();await page.mouse.move(box!.x+60,box!.y+40,{steps:4});await page.mouse.up();const moved=await page.evaluate(()=>window.paginationHost.session!.snapshot().metadata.word!.drawings!.find(item=>item.kind==='shape')!.geometry);expect(moved.x).toBeCloseTo(object.geometry.x+30);expect(moved.y).toBeCloseTo(object.geometry.y+15);await page.evaluate(()=>window.paginationHost.session!.undo());expect((await page.evaluate(()=>window.paginationHost.session!.snapshot())).html).toBe(before.html);
 await shape.click({position:{x:10,y:10}});await page.getByRole('button',{name:'编辑对象文字',exact:true}).click();const text=frame.locator('[data-office-editor] p').first();await text.click();await page.keyboard.press('End');await page.keyboard.insertText('编辑');await page.evaluate(()=>(window.paginationHost.view as import('@doffi/editor').DocumentTextEditor).execute({type:'word',action:'mainStory'}));await expect(shape.locator('.sdk-word-drawing-text')).toContainText('形状文字编辑');expect(await page.evaluate(()=>window.paginationHost.errors)).toEqual([]);
});
test('physical columns keep one source paragraph, precise selection and reversible rich input',async({page})=>{
 await mount(page,columnFiles(),'edit');const before=await page.evaluate(()=>window.paginationHost.session!.snapshot()),main=before.metadata.word!.stories.find(story=>story.kind==='main')!;
 const frame=page.frameLocator('iframe'),paragraph=frame.locator('[data-office-editor] p').first();const positions=await paragraph.evaluate(element=>{const base=element.closest('.sdk-page-flow')!.getBoundingClientRect(),walker=document.createTreeWalker(element,4),result:{x:number;y:number}[]=[];for(let node=walker.nextNode();node;node=walker.nextNode()){const range=document.createRange();range.selectNodeContents(node);for(const rect of range.getClientRects())if(rect.width>.1)result.push({x:rect.x-base.x,y:rect.y-base.y});}return result;});expect(positions.some(rect=>rect.x>150&&rect.y<256)).toBe(true);expect(positions.some(rect=>rect.y>344)).toBe(true);
 await paragraph.evaluate(element=>{const walker=document.createTreeWalker(element,4),base=element.closest('.sdk-page-flow')!.getBoundingClientRect();for(let node=walker.nextNode();node;node=walker.nextNode()){const range=document.createRange();range.selectNodeContents(node);if(range.getBoundingClientRect().left-base.left>150){range.setStart(node,0);range.collapse(true);const selection=document.getSelection()!;selection.removeAllRanges();selection.addRange(range);(element.closest('[data-office-editor]') as HTMLElement).focus();return;}}throw new Error('Missing second-column text.');});await page.keyboard.press('Enter');await page.keyboard.insertText('EDIT');
 const result=await page.evaluate(()=>({snapshot:window.paginationHost.session!.snapshot(),text:window.paginationHost.session!.getNode(window.paginationHost.session!.snapshot().metadata.word!.stories.find(story=>story.kind==='main')!.rootNodeId)!.text,errors:window.paginationHost.errors}));expect(result.text.replace('EDIT','')).toBe(await page.evaluate(id=>window.paginationHost.model!.getNode(id)!.text,main.rootNodeId));expect(result.snapshot.html).not.toMatch(/position:relative|sdk-page|sdk-repeat-header/);expect(result.errors).toEqual([]);
});
test('table header projections repeat at every physical column without entering source',async({page})=>{
 await mount(page,columnFiles(true));const frame=page.frameLocator('iframe'),headers=frame.locator('.sdk-repeat-header');expect(await headers.count()).toBeGreaterThan(1);const boxes=await headers.evaluateAll(nodes=>nodes.map(node=>({x:node.getBoundingClientRect().left,y:node.getBoundingClientRect().top})));expect(boxes.some((box,index)=>boxes.some((other,ordinal)=>index!==ordinal&&Math.abs(box.y-other.y)<1&&Math.abs(box.x-other.x)>100))).toBe(true);expect((await page.evaluate(()=>window.paginationHost.model!.snapshot())).html).not.toContain('sdk-repeat-header');
});
test('native list keyboard input adjusts levels, exits an empty list and shares undo history',async({page})=>{
 const session=createDocumentSession(openWordDocument(createWordDocument({documentId:'list-input'}),{kind:'document'}));session.commit(session.prepare({protocolVersion:1,documentId:'list-input',baseRevision:0,operations:[{op:'replaceText',nodeId:'paragraph-1',from:0,to:0,expectedText:'',text:'First'}]}));await mount(page,session.serialize(),'edit');const frame=page.frameLocator('iframe'),first=frame.locator('[data-node-id="paragraph-1"]');await first.click();await page.evaluate(()=>(window.paginationHost.view as import('@doffi/editor').DocumentTextEditor).execute({type:'word',action:'list',kind:'numbering'}));await page.keyboard.press('End');await page.keyboard.press('Enter');await page.keyboard.press('Tab');await expect(frame.locator('[data-office-editor] p').nth(1)).toHaveAttribute('data-word-level','1');await page.keyboard.press('Shift+Tab');await expect(frame.locator('[data-office-editor] p').nth(1)).toHaveAttribute('data-word-level','0');await page.keyboard.press('Enter');await expect(frame.locator('[data-office-editor] p').nth(1)).not.toHaveAttribute('data-word-numbering');await page.keyboard.press('Control+z');await expect(frame.locator('[data-office-editor] p').nth(1)).toHaveAttribute('data-word-numbering');expect(await page.evaluate(()=>window.paginationHost.errors)).toEqual([]);
});
test('dragging a source selection at the outer viewport edge scrolls and extends its range',async({page})=>{
 await mount(page,columnFiles(),'edit');const frame=page.frameLocator('iframe'),first=frame.locator('[data-office-editor] p').first(),box=await first.boundingBox();expect(box).not.toBeNull();const stage=page.locator('.stage'),bounds=await stage.boundingBox();expect(bounds).not.toBeNull();await page.mouse.move(box!.x+8,box!.y+12);await page.mouse.down();await page.mouse.move(box!.x+40,bounds!.y+bounds!.height-3,{steps:5});await expect.poll(()=>stage.evaluate(node=>node.scrollTop)).toBeGreaterThan(30);await page.mouse.up();expect(await first.evaluate(()=>document.getSelection()!.toString().length)).toBeGreaterThan(20);expect(await page.evaluate(()=>window.paginationHost.errors)).toEqual([]);
});
test('clear formatting uses source defaults and formatting marks survive input without entering source',async({page})=>{
 const session=createDocumentSession(openWordDocument(createWordDocument({documentId:'reset-format'}),{kind:'document'}));session.commit(session.prepare({protocolVersion:1,documentId:'reset-format',baseRevision:0,operations:[{op:'replaceText',nodeId:'paragraph-1',from:0,to:0,expectedText:'',text:'a b 😀'}]}));await mount(page,session.serialize(),'edit');const frame=page.frameLocator('iframe'),paragraph=frame.locator('[data-node-id="paragraph-1"]');await paragraph.evaluate(node=>{const range=document.createRange();range.selectNodeContents(node);const selection=document.getSelection()!;selection.removeAllRanges();selection.addRange(range);(node.closest('[data-office-editor]') as HTMLElement).focus();});
 await page.evaluate(()=>{const editor=window.paginationHost.view as import('@doffi/editor').DocumentTextEditor;editor.execute({type:'format',property:'font-size',value:'32pt'});editor.execute({type:'format',property:'font-family',value:'"Liberation Mono"'});editor.execute({type:'word',action:'clearFormat'});editor.execute({type:'word',action:'showMarks',visible:true});});
 expect(await paragraph.evaluate(node=>[...node.querySelectorAll<HTMLElement>('span')].filter(span=>span.textContent?.trim()).some(span=>Math.abs(parseFloat(getComputedStyle(span).fontSize)-16)<.05))).toBe(true);await expect(paragraph.locator('.sdk-space-mark')).toHaveCount(2);await paragraph.click();await page.keyboard.press('End');await page.keyboard.insertText(' next');await expect(paragraph.locator('.sdk-space-mark')).toHaveCount(3);const snapshot=await page.evaluate(()=>window.paginationHost.session!.snapshot());expect(snapshot.html).not.toMatch(/data-doffi-show-marks|sdk-space-mark|data-sdk-space-mark/);expect(await page.evaluate(()=>window.paginationHost.session!.getNode('paragraph-1')!.text)).toBe('a b 😀 next');expect(await page.evaluate(()=>window.paginationHost.errors)).toEqual([]);
});
test('source projection strips late owned wrappers while preserving bound roots and unowned lookalikes', async ({ page }) => {
  await mount(page, pagedFiles('<p data-node-id="text">文档</p>'));
  const result = await page.evaluate(() => window.paginationHost.lateProjection());
  expect(result.html).toBe(result.original); expect(result.instance).toBe(result.original); expect(result.shellMissing).toBe(true);
  expect(result.copy).toBe('<span data-node-id="run" style="color:#123456">原文</span>');
  expect(result.unowned).toContain('<span class="sdk-inline-flow" style="display: block;">源内容</span>');
});
test("read-only pagination requires no editing session and preserves the complete package source", async ({ page }) => {
  const files = pagedFiles(Array.from({ length: 40 }, (_, index) => '<p data-node-id="p-' + index + '">LINE-' + index + "</p>").join(""));
  await mount(page, files);
  const state = await page.evaluate(() => ({ report: window.paginationHost.layout(), session: !!window.paginationHost.session, html: window.paginationHost.model!.snapshot().html }));
  expect(state.session).toBe(false); expect(state.report.pages).toHaveLength(4);
  const frame = page.frameLocator("iframe"); await expect(frame.locator('[contenteditable="true"]')).toHaveCount(0);
  await expect(frame.locator("script,input,textarea")).toHaveCount(0);
  expect(state.html).not.toMatch(/sdk-page|column-height|contenteditable/);
  expect(state.html).toContain("LINE-39");
  const geometry = await frame.locator('[data-node-id="p-12"]').evaluate((node) => ({ y: node.getBoundingClientRect().top, text: node.textContent }));
  expect(geometry.y).toBeGreaterThan(300); expect(state.report.nodePages["p-12"]).toEqual(["page-2"]);
});
test("one paragraph spans pages and maps a cross-page selection to one canonical UTF-16 range", async ({ page }) => {
  await mount(page, pagedFiles('<p data-node-id="long">' + Array.from({ length: 40 }, (_, index) => "LINE-" + index + " 😀").join("<br>") + "</p>"), "interaction");
  const frame = page.frameLocator("iframe");
  await frame.locator('[data-node-id="long"]').evaluate((node) => {
    const range = document.createRange(); range.selectNodeContents(node); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); document.dispatchEvent(new Event("selectionchange"));
  });
  await expect.poll(() => page.evaluate(() => window.paginationHost.interaction!.selection()?.anchor)).toEqual(await page.evaluate(() => ({ kind: "text", nodeId: "long", from: 0,
    to: window.paginationHost.session!.getNode("long")!.text.length, quote: window.paginationHost.session!.getNode("long")!.text })));
  expect((await page.evaluate(() => window.paginationHost.layout().nodePages.long))!.length).toBeGreaterThan(1);
});
test("hiding a preview preserves its accepted pages and revealing it remeasures the visible canvas", async ({ page }) => {
  const errors: string[] = []; page.on('pageerror', error => errors.push(error.message));
  await mount(page, pagedFiles(Array.from({ length: 40 }, (_, index) => `<p data-node-id="p-${index}">LINE-${index}</p>`).join('')));
  const original = await page.evaluate(() => window.paginationHost.layout()); expect(original.pages).toHaveLength(4);
  const hidden = await page.evaluate(async () => { (document.body.firstElementChild as HTMLElement).hidden = true; return window.paginationHost.view!.settled(); });
  expect(hidden.pages).toEqual(original.pages); expect(hidden.nodePages).toEqual(original.nodePages);
  await page.evaluate(() => { (document.body.firstElementChild as HTMLElement).hidden = false; });
  await expect.poll(() => page.evaluate(async () => (await window.paginationHost.view!.settled()).pages)).toEqual(original.pages);
  await expect.poll(() => page.frameLocator('iframe').locator('[data-node-id="p-12"]').evaluate(node => node.getBoundingClientRect().top)).toBeGreaterThan(300);
  expect(errors).toEqual([]);
});
test("editing reflows pages immediately and undo restores layout without persisting page wrappers", async ({ page }) => {
  await mount(page, pagedFiles('<p data-node-id="a">Start</p>'), "edit");
  const before = await page.evaluate(() => window.paginationHost.layout().pages.length);
  await page.frameLocator("iframe").locator('[data-node-id="a"]').evaluate((node) => {
    const range = document.createRange(); range.selectNodeContents(node); range.collapse(false); const selection = document.getSelection()!; selection.removeAllRanges(); selection.addRange(range); (document.querySelector("[data-office-editor]") as HTMLElement).focus();
  });
  await page.keyboard.insertText(" long text".repeat(300));
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBeGreaterThan(before);
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).not.toMatch(/sdk-page|column-height|contenteditable/);
  await page.keyboard.press("Control+z");
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBe(before);
  expect(await page.evaluate(() => window.paginationHost.session!.getNode("a")!.text)).toBe("Start");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("repeated table headers remain a display projection while body cells can be edited", async ({ page }) => {
  await mount(page, longTable(), "edit");
  const pages = await page.evaluate(() => window.paginationHost.layout().pages.length);
  expect(pages).toBeGreaterThan(1);
  await expect(page.frameLocator("iframe").locator(".sdk-repeat-header")).toHaveCount(pages - 1);
  await page.frameLocator("iframe").locator('[data-node-id="cell-20"]').first().evaluate((node) => {
    const range = document.createRange(); range.selectNodeContents(node); range.collapse(false); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); (document.querySelector("[data-office-editor]") as HTMLElement).focus();
  });
  await page.keyboard.insertText(" EDITED");
  expect(await page.evaluate(() => window.paginationHost.session!.getNode("cell-20")!.text)).toBe("UNIQUE-20 EDITED");
  const html = await page.evaluate(() => window.paginationHost.session!.snapshot().html);
  expect(html.match(/REPEATED HEADER/g)).toHaveLength(1); expect(html).not.toMatch(/sdk-repeat|box-decoration-break|contenteditable/);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("page fields and TOC are derived; explicit cache update is one core undo step", async ({ page }) => {
  await mount(page, fieldsFiles(), "interaction");
  const state = await page.evaluate(() => ({ report: window.paginationHost.layout(), html: window.paginationHost.session!.snapshot().html, revision: window.paginationHost.session!.snapshot().manifest.revision }));
  expect(state.report.pages).toHaveLength(2);
  expect(state.report.fields.find((field) => field.kind === "NUMPAGES")?.value).toBe("2");
  expect(state.report.fields.find((field) => field.kind === "TOC")?.value).toBe("Alpha\t1\nBeta\t2");
  expect(state.html).not.toContain("sdk-field");
  await page.evaluate(() => window.paginationHost.updateFields());
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().manifest.revision)).toBe(state.revision + 1);
  expect(await page.evaluate(() => window.paginationHost.session!.history().undoSteps)).toBe(1);
  await page.evaluate(() => window.paginationHost.session!.undo());
  expect(await page.evaluate(() => window.paginationHost.session!.snapshot().html)).toBe(state.html);
});
test("continuous preview remains available and page budgets fail explicitly", async ({ page }) => {
  const files = pagedFiles(Array.from({ length: 40 }, (_, index) => '<p data-node-id="p-' + index + '">LINE-' + index + "</p>").join(""));
  await mount(page, files, "preview", { layout: "continuous" });
  expect(await page.evaluate(() => window.paginationHost.layout().layout)).toBe("continuous");
  await expect(page.frameLocator("iframe").locator(".sdk-paper")).toHaveCount(0);
  await expect(mount(page, files, "preview", { limits: { maxPages: 1 } })).rejects.toThrow("page budget");
});

const prefix = (count: number): string => Array.from({ length: count }, (_, index) => `<p data-node-id="prefix-${index}">LINE</p>`).join("");
async function atStart(page: Page, id: string): Promise<void> {
  await page.frameLocator("iframe").locator(`[data-node-id="${id}"]`).first().evaluate((node) => {
    const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT), text = walker.nextNode()!;
    const range = document.createRange(); range.setStart(text, 0); range.collapse(true);
    const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); (document.querySelector("[data-office-editor]") as HTMLElement).focus();
  });
}
test("headings stay with following text and short paragraphs respect orphan and widow limits", async ({ page }) => {
  await mount(page, pagedFiles(prefix(11) + '<h1 data-node-id="heading" style="font-size:12pt;line-height:16pt;break-after:avoid">HEADING</h1><p data-node-id="after">FOLLOWING</p>'));
  expect(await page.evaluate(() => window.paginationHost.layout().nodePages.heading)).toEqual(["page-2"]);
  expect(await page.evaluate(() => window.paginationHost.layout().nodePages.after)).toEqual(["page-2"]);
  await mount(page, pagedFiles(prefix(11) + '<p data-node-id="short" style="orphans:2;widows:2">FIRST<br>SECOND<br>THIRD</p>'));
  expect(await page.evaluate(() => window.paginationHost.layout().nodePages.short)).toEqual(["page-2"]);
  await mount(page, pagedFiles(prefix(9) + '<p data-node-id="four" style="orphans:2;widows:2">FIRST<br>SECOND<br>THIRD<br>FOURTH</p>'));
  const y = await page.frameLocator("iframe").locator('[data-node-id="four"]').evaluate((node) => [...node.childNodes].filter((item) => item.nodeType === 3 && item.textContent).map((item) => {
    const range = document.createRange(); range.selectNodeContents(item); return range.getBoundingClientRect().top;
  }));
  expect(y[1]! - y[0]!).toBeLessThan(30); expect(y[2]! - y[1]!).toBeGreaterThan(60); expect(y[3]! - y[2]!).toBeLessThan(30);
});
test("unidentified hard page breaks stay canonical after edit, save and reopen", async ({ page }) => {
  await mount(page, pagedFiles('<p data-node-id="a">AAA<br data-word-break="page">BBB</p>'), "edit");
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBe(2);
  await atStart(page, "a"); await page.keyboard.insertText("X");
  expect(await page.evaluate(() => window.paginationHost.session!.getNode("a")!.text)).toBe("XAAA\nBBB");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
  const entries = await page.evaluate(() => [...window.paginationHost.session!.serialize()].map(([path, bytes]) => [path, [...bytes]] as [string, number[]]));
  await page.evaluate((files) => window.paginationHost.mount(files, "preview"), entries);
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBe(2);
  expect(await page.evaluate(() => window.paginationHost.model!.snapshot().html)).not.toMatch(/display:\s*none|sdk-page-gap/);
});
test("direct container text is paginated without losing its source identity", async ({ page }) => {
  await mount(page, pagedFiles('<div data-node-id="long">' + Array.from({ length: 40 }, (_, index) => "LINE-" + index).join("<br>") + "</div>"));
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBeGreaterThan(1);
  expect(await page.evaluate(() => window.paginationHost.layout().nodePages.long!.length)).toBeGreaterThan(1);
  expect(await page.frameLocator("iframe").locator('[data-node-id="long"]').count()).toBe(1);
});
test("section metadata order never reorders source or saved edits", async ({ page }) => {
  const files = pagedFiles('<section data-node-id="first"><p data-node-id="a">A</p></section><section data-node-id="second"><p data-node-id="b">B</p></section>');
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json"))); metadata.sections = [{ sectionId: "second", rootNodeId: "second", breakType: "nextPage" }, { sectionId: "first", rootNodeId: "first" }];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  await mount(page, files, "edit"); await atStart(page, "a"); await page.keyboard.insertText("X");
  const state = await page.evaluate(() => ({ text: window.paginationHost.session!.getNode("root")!.text, html: window.paginationHost.session!.snapshot().html, errors: window.paginationHost.errors }));
  expect(state.text).toBe("XAB"); expect(state.html.indexOf('data-node-id="first"')).toBeLessThan(state.html.indexOf('data-node-id="second"')); expect(state.errors).toEqual([]);
});
test("deletion without target-range support crosses view gaps and preserves graphemes", async ({ page }) => {
  await mount(page, pagedFiles('<p data-node-id="a">AB漢󠄀😀</p><p data-node-id="b">NEXT</p>'), "edit");
  await page.frameLocator("iframe").locator('[data-node-id="a"]').evaluate((node) => {
    Object.defineProperty(InputEvent.prototype, "getTargetRanges", { configurable: true, value: undefined });
    const range = document.createRange(); range.selectNodeContents(node); range.collapse(false); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range); (document.querySelector("[data-office-editor]") as HTMLElement).focus();
  });
  await page.keyboard.press("Backspace"); await page.keyboard.press("Backspace");
  expect(await page.evaluate(() => window.paginationHost.session!.getNode("a")!.text)).toBe("AB");
  await atStart(page, "b"); await page.keyboard.press("Backspace");
  expect(await page.evaluate(() => window.paginationHost.session!.getNode("a")!.text)).toBe("ABNEXT");
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
});
test("long table rows fragment while continued pages reserve repeated-header space", async ({ page }) => {
  const lines = Array.from({ length: 30 }, (_, index) => "CELL-" + index).join("<br>");
  await mount(page, pagedFiles('<table data-node-id="table"><tr data-node-id="header" data-word-repeat-header="true"><th data-node-id="h1">HEAD</th><th data-node-id="h2">HEAD2</th></tr><tr data-node-id="long-row"><td data-node-id="left">' + lines + '</td><td data-node-id="right">' + lines + '</td></tr><tr data-node-id="tail"><td data-node-id="tail-left">END</td><td data-node-id="tail-right">END2</td></tr></table>'));
  const report = await page.evaluate(() => window.paginationHost.layout()); expect(report.pages.length).toBeGreaterThan(2);
  await expect(page.frameLocator("iframe").locator(".sdk-repeat-header")).toHaveCount(report.pages.length - 1);
  const overlaps = await page.frameLocator("iframe").locator('[data-node-id="left"]').evaluate((node) => {
    const headerRects = [...document.querySelectorAll(".sdk-repeat-header")].map((header) => header.getBoundingClientRect()), walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT); let item; const bad: string[] = [];
    while ((item = walker.nextNode())) { if (!item.textContent) continue; const range = document.createRange(); range.selectNodeContents(item); for (const rect of range.getClientRects()) if (headerRects.some((header) => rect.top < header.bottom - 0.5 && rect.bottom > header.top + 0.5)) bad.push(item.textContent); }
    return bad;
  });
  expect(overlaps).toEqual([]);
});
test("mixed container inline text remains visible and authored copy excludes page furniture", async ({ page }) => {
  await mount(page, pagedFiles('<div data-node-id="mixed">' + Array.from({ length: 30 }, (_, index) => "DIRECT-" + index).join("<br>") + '<p data-node-id="end">END</p></div>'));
  expect(await page.evaluate(() => window.paginationHost.layout().pages.length)).toBeGreaterThan(2);
  const copy = await page.frameLocator("iframe").locator('[data-node-id="mixed"]').evaluate((node) => {
    const range = document.createRange(); range.selectNodeContents(node); const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range);
    const data = new Map<string, string>(), event = new Event("copy", { bubbles: true, cancelable: true });
    Object.defineProperty(event, "clipboardData", { value: { setData: (type: string, value: string) => data.set(type, value) } }); document.dispatchEvent(event);
    return { html: data.get("text/html") ?? "", text: data.get("text/plain") ?? "", prevented: event.defaultPrevented };
  });
  expect(copy.prevented).toBe(true); expect(copy.html).not.toMatch(/sdk-page|sdk-inline-flow|contenteditable|clip-path/); expect(copy.text).toContain("DIRECT-29"); expect(copy.text).toContain("END");
});
test("host interaction events work while CSP blocks executable document content", async ({ page }) => {
  await mount(page, pagedFiles('<p data-node-id="a">SAFE</p>'));
  await page.evaluate(() => {
    (window as unknown as { sourceExecuted: number; sourceViolations: number }).sourceExecuted = 0;
    (window as unknown as { sourceViolations: number }).sourceViolations = 0;
    const frame = document.querySelector("iframe")!.contentDocument!;
    frame.addEventListener("securitypolicyviolation", (event) => { if (event.violatedDirective.startsWith("script-src")) (window as unknown as { sourceViolations: number }).sourceViolations++; });
    const script = frame.createElement("script"); script.textContent = "parent.sourceExecuted++"; frame.body.append(script);
    const image = frame.createElement("img"); image.setAttribute("onerror", "parent.sourceExecuted++"); image.src = "invalid-source:"; frame.body.append(image);
  });
  await expect.poll(() => page.evaluate(() => (window as unknown as { sourceViolations: number }).sourceViolations)).toBeGreaterThanOrEqual(2);
  expect(await page.evaluate(() => (window as unknown as { sourceExecuted: number }).sourceExecuted)).toBe(0);
  expect(await page.frameLocator("iframe").locator('meta[http-equiv="Content-Security-Policy"]').getAttribute("content")).toContain("default-src 'none'");
});
test("overheight headers stay outside body text and report their clipping", async ({ page }) => {
  const files = pagedFiles('<section data-node-id="main" data-office-story="main"><section data-node-id="sec"><p data-node-id="body-text">BODY FIRST LINE</p></section></section><section data-node-id="head" data-office-story="header"><p data-node-id="header-a">HEADER 1</p><p data-node-id="header-b">HEADER 2</p><p data-node-id="header-c">HEADER 3</p></section>');
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json"))); metadata.pageSetup = { ...metadata.pageSetup, height: 300, margins: { top: 72, bottom: 72, left: 24, right: 24 } };
  metadata.sections = [{ sectionId: "s", rootNodeId: "sec", headers: { default: "hdr" }, headerDistance: 36 }];
  metadata.word = { stories: [{ storyId: "main", kind: "main", rootNodeId: "main" }, { storyId: "hdr", kind: "header", rootNodeId: "head" }], styles: [], numbering: [], hyperlinks: [], bookmarks: [], fields: [], tabs: [], tables: [], revisions: [] };
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata))); await mount(page, files);
  expect((await page.evaluate(() => window.paginationHost.layout().diagnostics)).some((item) => item.code === "ContentOverflow" && item.nodeId === "head")).toBe(true);
  const bounds = await page.frameLocator("iframe").locator("body").evaluate(() => ({ header: document.querySelector(".sdk-paper .sdk-page-story")!.getBoundingClientRect().bottom, body: document.querySelector('[data-node-id="body-text"]')!.getBoundingClientRect().top }));
  expect(bounds.header).toBeLessThanOrEqual(bounds.body + 0.5);
});
test("copy keeps nested block boundaries and source attributes cannot impersonate the editor", async ({ page }) => {
  const files = pagedFiles('<section data-node-id="nested"><p data-node-id="first">FIRST</p><p data-node-id="second">SECOND</p></section><table data-node-id="grid"><tr data-node-id="r"><td data-node-id="c">C</td><td data-node-id="d">D</td></tr></table>');
  files.set("content.html", new TextEncoder().encode(new TextDecoder().decode(files.get("content.html")).replace('data-node-id="root"', 'data-node-id="root" data-office-editor="source-data"')));
  await mount(page, files);
  const copied = await page.frameLocator("iframe").locator('[data-node-id="root"]').evaluate((node) => {
    const range = document.createRange(); range.selectNodeContents(node); document.getSelection()!.removeAllRanges(); document.getSelection()!.addRange(range);
    const values: Record<string, string> = {}, event = new Event("copy", { bubbles: true, cancelable: true }); Object.defineProperty(event, "clipboardData", { value: { setData: (type: string, value: string) => { values[type] = value; } } }); document.dispatchEvent(event);
    return { text: values["text/plain"], html: values["text/html"], prevented: event.defaultPrevented, editable: (node as HTMLElement).isContentEditable };
  });
  expect(copied.editable).toBe(false); expect(copied.prevented).toBe(true); expect(copied.text).toBe("FIRST\nSECOND\nC\tD"); expect(copied.html).not.toContain("sdk-page");
});
for (const mode of ["preview", "edit"] as const) test("partial field display copy stays within the selected range: " + mode, async ({ page }) => {
  await mount(page, fieldsFiles(), mode);
  const copied = await page.frameLocator("iframe").locator(".sdk-field-display").filter({ hasText: "Alpha" }).first().evaluate((node) => {
    const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT), text = walker.nextNode()!;
    const range = document.createRange(); range.setStart(text, 0); range.setEnd(text, 3); document.getSelection()!.removeAllRanges(); document.getSelection()!.addRange(range);
    const values: Record<string, string> = {}, event = new Event("copy", { bubbles: true, cancelable: true }); Object.defineProperty(event, "clipboardData", { value: { setData: (type: string, value: string) => { values[type] = value; } } });
    node.dispatchEvent(event); return values;
  });
  expect(copied["text/plain"]).toBe("Alp"); expect(copied["text/html"]).not.toContain("Beta"); expect(copied["text/html"]).not.toMatch(/sdk-field|data-word-field|data-node-id/);
});
test("disposing a preview during settled invalidates its pending report", async ({ page }) => {
  await mount(page, pagedFiles('<p data-node-id="a">TEXT</p>'));
  const result = await page.evaluate(async () => {
    const view = window.paginationHost.view as import("@doffi/editor").DocumentPreview;
    const pending = view.settled(); await Promise.resolve(); view.dispose();
    return pending.then(() => "success", (error: { code: string }) => error.code);
  });
  expect(result).toBe("InvalidOperation");
});
test("odd-page section breaks create a blank sheet using the new geometry and restart printed numbering", async ({ page }) => {
  const files = pagedFiles('<section data-node-id="first"><p data-node-id="a">FIRST</p></section><section data-node-id="second"><p data-node-id="b">SECOND</p></section>');
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")));
  metadata.sections = [{ sectionId: "first", rootNodeId: "first" }, { sectionId: "second", rootNodeId: "second", breakType: "oddPage", pageNumberStart: 8, pageNumberFormat: "lowerRoman", pageSetup: { ...metadata.pageSetup, width: 240, height: 300 } }];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata))); await mount(page, files);
  const report = await page.evaluate(() => window.paginationHost.layout());
  expect(report.pages.map((item) => [item.blank, item.setup.width, item.setup.height, item.label])).toEqual([[false, 300, 240, "1"], [true, 240, 300, "2"], [false, 240, 300, "viii"]]);
  expect(report.nodePages.a).toEqual(["page-1"]); expect(report.nodePages.b).toEqual(["page-3"]);
});
test("unitless inherited Word line spacing works without Typed OM and preserves absolute overrides", async ({ page }) => {
  await page.addInitScript(() => { Object.defineProperty(HTMLElement.prototype, "computedStyleMap", { configurable: true, value: undefined }); });
  const files = pagedFiles('<p data-node-id="auto" data-word-line-rule="auto">AUTO</p><p data-node-id="absolute" data-word-line-rule="auto" style="line-height:30px">ABSOLUTE</p>');
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json")));
  metadata.word = { stories: [{ storyId: "main", kind: "main", rootNodeId: "root" }], styles: [], numbering: [], hyperlinks: [], bookmarks: [], fields: [], tabs: [], tables: [], revisions: [] };
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata)));
  files.set("content.html", new TextEncoder().encode(new TextDecoder().decode(files.get("content.html")).replace('data-office-kind="document"', 'data-office-kind="document" data-office-story="main"')));
  files.set("styles.css", new TextEncoder().encode('article{font-family:"Liberation Serif",serif;font-size:12pt}p{line-height:1.5;margin:0}'));
  await mount(page, files);
  const measured = await page.frameLocator("iframe").locator('[data-node-id="auto"]').evaluate((node) => {
    const probe = document.createElement("span"); probe.style.cssText = "position:absolute;display:block;line-height:normal;white-space:pre;font:inherit"; probe.style.lineHeight = "normal"; probe.textContent = "汉M"; node.append(probe);
    const normal = probe.getBoundingClientRect().height; probe.remove(); return { line: parseFloat(getComputedStyle(node).lineHeight), normal };
  });
  expect(measured.line).toBeCloseTo(measured.normal * 1.5, 1);
  await expect(page.frameLocator("iframe").locator('[data-node-id="absolute"]')).toHaveCSS("line-height", "30px");
});
for (const layout of ["pages", "continuous"] as const) test("editing preserves source content before, between and after sections: " + layout, async ({ page }) => {
  const files = pagedFiles('<p data-node-id="lead">L</p><section data-node-id="first"><p data-node-id="a">A</p></section><p data-node-id="middle">M</p><section data-node-id="second"><p data-node-id="b">B</p></section><p data-node-id="tail">T</p>');
  const metadata = JSON.parse(new TextDecoder().decode(files.get("metadata.json"))); metadata.sections = [{ sectionId: "first", rootNodeId: "first" }, { sectionId: "second", rootNodeId: "second", breakType: "nextPage" }];
  files.set("metadata.json", new TextEncoder().encode(JSON.stringify(metadata))); await mount(page, files, "edit", { layout });
  await atStart(page, "a"); await page.keyboard.insertText("X");
  const state = await page.evaluate(() => ({ text: window.paginationHost.session!.getNode("root")!.text, html: window.paginationHost.session!.snapshot().html, errors: window.paginationHost.errors }));
  expect(state.text).toBe("LXAMBT"); expect(state.errors).toEqual([]);
  expect(state.html).not.toMatch(/sdk-page|edit-\d|<div/);
  const structure = await page.evaluate(() => window.paginationHost.session!.getNode("root")!.childNodeIds);
  expect(structure).toEqual(["lead", "first", "middle", "second", "tail"]);
  await page.keyboard.press("Control+z"); expect(await page.evaluate(() => window.paginationHost.session!.getNode("root")!.text)).toBe("LAMBT");
});
test("formatted table whitespace does not become extra plain-text rows or cells", async ({ page }) => {
  await mount(page, pagedFiles('<table data-node-id="table">\n <tbody>\n <tr data-node-id="row">\n  <td data-node-id="c">C</td>\n  <td data-node-id="d">D</td>\n </tr>\n </tbody>\n</table>'));
  const copied = await page.frameLocator("iframe").locator('[data-node-id="table"]').evaluate((node) => {
    const range = document.createRange(); range.selectNode(node); document.getSelection()!.removeAllRanges(); document.getSelection()!.addRange(range);
    const values: Record<string, string> = {}, event = new Event("copy", { bubbles: true, cancelable: true }); Object.defineProperty(event, "clipboardData", { value: { setData: (type: string, value: string) => { values[type] = value; } } }); node.dispatchEvent(event); return values;
  });
  expect(copied["text/plain"]).toBe("C\tD");
});

import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { readFileSync, mkdirSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { build } from "vite";
import { importPptx } from "@doffi/office";

if (process.platform !== "darwin") throw new Error("Safari acceptance requires macOS and enabled safaridriver.");
const origin = "http://127.0.0.1:4173", endpoint = "http://127.0.0.1:4445";
const root = fileURLToPath(new URL("../", import.meta.url));
const server = spawn(process.execPath, [root + "node_modules/vite/bin/vite.js", "preview", "--host", "127.0.0.1", "--port", "4173", "--strictPort"], { cwd: root + "apps/playground", stdio: "inherit" });
const driver = spawn("/usr/bin/safaridriver", ["--port", "4445"], { stdio: "inherit" });
let session;
const pause = () => new Promise((resolve) => setTimeout(resolve, 100));
async function ready(url) {
  for (let pass = 0; pass < 100; pass++) { try { if ((await fetch(url, { signal: AbortSignal.timeout(2000) })).ok) return; } catch {} await pause(); }
  throw new Error("Acceptance service did not become ready: " + url);
}
async function command(method, path, data) {
  const response = await fetch(endpoint + path, { method, ...(data === undefined ? {} : { headers: { "Content-Type": "application/json" }, body: JSON.stringify(data) }), signal: AbortSignal.timeout(30_000) });
  const result = await response.json();
  if (!response.ok || result.value?.error) throw new Error(JSON.stringify(result)); return result.value;
}
const execute = (script, args = []) => command("POST", `/session/${session}/execute/sync`, { script, args });
const iframeExpression = "((view)=>view.element.localName==='iframe'?view.element:view.element.shadowRoot.querySelector('iframe'))(window.paginationHost.view)";
function files(body) {
  const entries = new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map((path) => [path, readFileSync(root + "examples/report/" + path)]));
  entries.set("content.html", Buffer.from(entries.get("content.html").toString().replace(/<article[\s\S]*<\/article>/, '<article data-node-id="root" data-office-kind="document">' + body + "</article>")));
  const metadata = JSON.parse(entries.get("metadata.json")); metadata.pageSetup = { unit: "pt", width: 300, height: 240, margins: { top: 24, bottom: 24, left: 24, right: 24 } };
  entries.set("metadata.json", Buffer.from(JSON.stringify(metadata))); entries.set("annotations.json", Buffer.from("[]"));
  entries.set("styles.css", Buffer.from('article{font-family:serif;font-size:12pt}p,td,th{line-height:16pt;margin:0;padding:0}table{border-collapse:collapse;width:240pt}tr{break-inside:avoid}'));
  return [...entries].map(([path, bytes]) => [path, [...bytes]]);
}
async function mount(body, mode) {
  const result = await command("POST", `/session/${session}/execute/async`, { script: "const done=arguments[arguments.length-1];window.paginationHost.mount(arguments[0],arguments[1]).then(()=>done({ok:true}),error=>done({error:String(error)}));", args: [files(body), mode] });
  assert.equal(result.ok, true, JSON.stringify(result));
}
try {
  await Promise.all([ready(origin), ready(endpoint + "/status")]);
  const created = await command("POST", "/session", { capabilities: { alwaysMatch: { browserName: "safari" } } }); session = created.sessionId;
  assert.ok(session); console.log("Safari capabilities", JSON.stringify(created.capabilities));
  await command("POST", `/session/${session}/url`, { url: origin });
  const result = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: { entry: root + "tests/browser/fixtures/pagination-host.ts", formats: ["iife"], name: "PaginationAcceptance" } } });
  const chunks = (Array.isArray(result) ? result[0].output : result.output).filter((item) => item.type === "chunk").map((item) => item.code).join("\n");
  await execute(chunks);
  await mount(Array.from({ length: 40 }, (_, index) => `<p data-node-id="p-${index}">LINE-${index}</p>`).join(""), "preview");
  assert.equal(await execute("return window.paginationHost.layout().pages.length"), 4);
  assert.equal(await execute("return !!window.paginationHost.session"), false);
  await mount('<p data-node-id="a">START</p>', "edit");
  const elementKey = "element-6066-11e4-a52e-4f735466cecf";
  const iframe = await execute("return " + iframeExpression);
  await command("POST", `/session/${session}/frame`, { id: iframe });
  await execute("const root=document.querySelector('[data-office-editor]'),p=document.querySelector('[data-node-id=a]');root.focus();const r=document.createRange();r.setStart(p.firstChild,0);r.collapse(true);const s=document.getSelection();s.removeAllRanges();s.addRange(r);");
  const editable = await command("POST", `/session/${session}/element`, { using: "css selector", value: "[data-office-editor]" });
  await command("POST", `/session/${session}/element/${editable[elementKey]}/value`, { text: "X", value: ["X"] });
  await command("POST", `/session/${session}/frame`, { id: null });
  assert.equal(await execute("return window.paginationHost.session.getNode('a').text"), "XSTART");
  await execute("window.paginationHost.session.undo()");
  assert.equal(await execute("return window.paginationHost.session.getNode('a').text"), "START");
  const lines = Array.from({ length: 40 }, (_, index) => `TEXT-${index}😀`);
  const longText = lines.join("\n");
  await mount('<p data-node-id="long">' + lines.join("<br>") + "</p>", "interaction");
  await execute("const d=" + iframeExpression + ".contentDocument,r=d.createRange();r.selectNodeContents(d.querySelector('[data-node-id=long]'));d.getSelection().removeAllRanges();d.getSelection().addRange(r);d.dispatchEvent(new Event('selectionchange'));");
  for (let pass = 0; pass < 50 && !await execute("return !!window.paginationHost.interaction.selection()"); pass++) await pause();
  const selected = await execute("return window.paginationHost.interaction.selection()?.anchor");
  assert.deepEqual(selected, { kind: "text", nodeId: "long", from: 0, to: longText.length, quote: longText });
  assert.ok(await execute("return window.paginationHost.layout().pages.length>1"));
  assert.deepEqual(await execute("return window.paginationHost.errors"), []);
  await execute("window.paginationHost.view.dispose();window.paginationHost.session?.dispose();");
  const pptBuild = await build({ configFile: false, logLevel: "silent", build: { write: false, minify: false, target: "es2022", lib: { entry: root + "tests/browser/fixtures/presentation-host.ts", formats: ["iife"], name: "PresentationAcceptance" } } });
  await execute((Array.isArray(pptBuild) ? pptBuild[0].output : pptBuild.output).filter(item => item.type === "chunk").map(item => item.code).join("\n"));
  const ppt = importPptx(readFileSync(root + "examples/pptx/09-charts.pptx"), { documentId: "safari-presentation" });
  const mounted = await command("POST", `/session/${session}/execute/async`, { script: "const done=arguments[arguments.length-1];window.presentationHost.mount(arguments[0]).then(()=>done({ok:true}),error=>done({mountError:{message:String(error),stack:error?.stack}}));", args: [[...ppt].map(([path, data]) => [path, [...data]])] });
  assert.equal(mounted.ok, true, JSON.stringify(mounted));
  assert.equal(await execute("return window.presentationHost.view.layout().pages.length"), 3);
  assert.equal(await execute("return window.presentationHost.view.element.contentDocument.querySelectorAll('svg[data-sdk-chart]').length"), 3);
  assert.equal(await execute("return window.presentationHost.view.element.contentDocument.querySelectorAll('[contenteditable=true]').length"), 0);
  assert.equal(await execute("return /<svg|data-sdk/.test(window.presentationHost.projected())"), false);
  await execute("window.presentationHost.view.dispose()");
  const pptFiles = [...["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json", "assets/sample.png"].map(path => [path, [...readFileSync(root + "examples/presentation/" + path)]])];
  async function installFixture(entry, react = false) {
    const fixture = await build({ configFile: false, logLevel: "silent", define: { "process.env.NODE_ENV": JSON.stringify("production") }, ...(react ? { resolve: { alias: { react: root + "apps/playground/node_modules/react", "react-dom": root + "apps/playground/node_modules/react-dom" } } } : {}), build: { write: false, target: "es2022", lib: { entry: root + "tests/browser/fixtures/" + entry, formats: ["iife"], name: "PptEditingAcceptance" } } });
    await execute((Array.isArray(fixture) ? fixture[0].output : fixture.output).filter(item => item.type === "chunk").map(item => item.code).join("\n"));
  }
  await installFixture("presentation-editor-host.ts");
  const editMount = await command("POST", `/session/${session}/execute/async`, { script: "const done=arguments[arguments.length-1];window.presentationEditorHost.mount(arguments[0]).then(()=>done({ok:true}),error=>done({error:String(error)}));", args: [pptFiles] });
  assert.equal(editMount.ok, true, JSON.stringify(editMount));
  await execute("window.presentationEditorHost.view.editText('heading')");
  const pptIframe = await execute("return window.presentationEditorHost.view.element.shadowRoot.querySelector('iframe[data-sdk-presentation-canvas]')");
  await command("POST", `/session/${session}/frame`, { id: pptIframe });
  await execute("const input=document.querySelector('[contenteditable=true]'),p=document.querySelector('[data-node-id=heading-p]');input.focus();const range=document.createRange();range.selectNodeContents(p);range.collapse(true);const selection=document.getSelection();selection.removeAllRanges();selection.addRange(range);");
  const pptInput = await command("POST", `/session/${session}/element`, { using: "css selector", value: "[contenteditable=true]" });
  await command("POST", `/session/${session}/element/${pptInput[elementKey]}/value`, { text: "S", value: ["S"] });
  await command("POST", `/session/${session}/frame`, { id: null });
  assert.equal(await execute("return window.presentationEditorHost.session.getNode('heading-p').text"), "S项目进展");
  await execute("window.presentationEditorHost.session.undo()");
  assert.equal(await execute("return window.presentationEditorHost.session.getNode('heading-p').text"), "项目进展");
  assert.deepEqual(await execute("return window.presentationEditorHost.errors"), []);
  await execute("window.presentationEditorHost.view.dispose();window.presentationEditorHost.session.dispose()");
  await installFixture("presentation-workbench-host.ts", true);
  const workbench = await command("POST", `/session/${session}/execute/async`, { script: "const done=arguments[arguments.length-1];window.presentationWorkbenchHost.mount(arguments[0],'js').then(()=>window.presentationWorkbenchHost.view.submitAnnotations(['review-1'])).then(()=>window.presentationWorkbenchHost.view.ready).then(()=>done({ok:true}),error=>done({error:String(error)}));", args: [pptFiles] });
  assert.equal(workbench.ok, true, JSON.stringify(workbench));
  assert.equal(await execute("return window.presentationWorkbenchHost.session.getNode('heading-p').text"), "项目进展");
  assert.equal(await execute("return window.presentationWorkbenchHost.view.element.shadowRoot.querySelector('iframe[title=\"PPT 候选只读预览\"]').contentDocument.querySelector('[data-node-id=heading-p]').textContent"), "提案标题");
  await execute("window.presentationWorkbenchHost.view.applyProposal()");
  assert.equal(await execute("return window.presentationWorkbenchHost.session.getNode('heading-p').text"), "提案标题");
  assert.equal(await execute("return window.presentationWorkbenchHost.session.history().undoSteps"), 1);
  await execute("window.presentationWorkbenchHost.session.undo()");
  assert.equal(await execute("return window.presentationWorkbenchHost.session.getNode('heading-p').text"), "项目进展");
  assert.deepEqual(await execute("return window.presentationWorkbenchHost.errors"), []);
  await execute("window.presentationWorkbenchHost.dispose()");
  await installFixture("interaction-host.ts");
  const workspaceFiles = ["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map(path => [path, [...readFileSync(root + "examples/report/" + path)]]);
  const workspace = await command("POST", `/session/${session}/execute/async`, { script: "const done=arguments[arguments.length-1];window.interactionHost.mount(arguments[0],false,'workspace').then(()=>done({ok:true}),error=>done({error:String(error)}));", args: [workspaceFiles] });
  assert.equal(workspace.ok, true, JSON.stringify(workspace));
  await execute("const d=window.interactionHost.view.element.shadowRoot.querySelector('iframe').contentDocument,r=d.createRange(),p=d.querySelector('[data-node-id=n-summary]');r.setStart(p.firstChild,2);r.setEnd(p.firstChild,8);d.getSelection().removeAllRanges();d.getSelection().addRange(r);d.dispatchEvent(new Event('selectionchange'));");
  for (let pass = 0; pass < 50 && !await execute("return window.interactionHost.view.element.shadowRoot.querySelector('[role=dialog]').hidden===false"); pass++) await pause();
  assert.equal(await execute("return window.interactionHost.view.element.shadowRoot.querySelector('[role=dialog]').hidden"), false);
  const composer = await execute("return window.interactionHost.view.element.shadowRoot.querySelector('[role=dialog] textarea')");
  await command("POST", `/session/${session}/element/${composer[elementKey]}/value`, { text: "Safari 批注", value: [..."Safari 批注"] });
  await execute("[...window.interactionHost.view.element.shadowRoot.querySelectorAll('[role=dialog] button')].find(node=>node.textContent==='添加批注').click()");
  assert.equal(await execute("return window.interactionHost.session.snapshot().manifest.revision"), 1);
  assert.equal(await execute("return window.interactionHost.session.getAnnotations().at(-1).instruction"), "Safari 批注");
  await execute("window.interactionHost.session.undo();window.interactionHost.view.dispose();window.interactionHost.view.dispose()");
  assert.equal(await execute("return window.interactionHost.session.snapshot().manifest.revision"), 2);
  assert.equal(await execute("return document.querySelectorAll('iframe').length"), 0);
  await execute("window.interactionHost.session.dispose()");
  mkdirSync(root + ".cache", { recursive: true }); writeFileSync(root + ".cache/safari-result.json", JSON.stringify({ capabilities: created.capabilities, checks: ["pure preview pages", "native input", "undo", "cross-page selection", "presentation fixed pages", "three native charts", "reversible presentation projection", "presentation disposal", "PPT native input/history", "PPT frozen request/candidate/apply/undo", "PPT editor/workbench disposal", "workspace floating comment/input/undo/disposal"] }, null, 2));
  console.log("Safari acceptance passed: Word input/history/selection and PPT preview/input/proposals/history/disposal.");
} catch (error) {
  if (session) { try { mkdirSync(root + ".cache", { recursive: true }); writeFileSync(root + ".cache/safari-failure.png", Buffer.from(await command("GET", `/session/${session}/screenshot`), "base64")); } catch {} }
  throw error;
} finally {
  try { if (session) await command("DELETE", `/session/${session}`); } finally { driver.kill(); server.kill(); }
}

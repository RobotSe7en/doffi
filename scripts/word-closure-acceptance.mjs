/** External, packed public-API runtime acceptance. No workspace source aliases. */
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { createServer } from "node:http";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { extname, join, resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { chromium, expect } from "@playwright/test";
import { build } from "vite";
import { packPackages } from "./pack.mjs";

const root = resolve(import.meta.dirname, ".."), cache = join(root, ".cache/word-closure-acceptance");
const temporary = mkdtempSync(join(tmpdir(), "doffi packed word closure-"));
const run = (exe, args, cwd = root, env = process.env) => execFileSync(exe, args, { cwd, env, encoding: "utf8", timeout: 60_000, maxBuffer: 8 * 1024 * 1024 }).trim();
const digest = (bytes) => createHash("sha256").update(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)).digest("hex");
const pnpm = process.env.npm_execpath;
assert.ok(pnpm, "Run acceptance through pnpm.");
rmSync(cache, { recursive: true, force: true }); mkdirSync(cache, { recursive: true });
const sources = Object.fromEntries(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map((name) => [name, readFileSync(join(root, "examples/report", name), "utf8")]));
const proposal = JSON.parse(readFileSync(join(root, "examples/report.proposal.json"), "utf8"));
const browserCache = process.env.PLAYWRIGHT_BROWSERS_PATH ?? join(root, ".playwright");
const shell = readdirSync(browserCache).find((name) => name.startsWith("chromium_headless_shell-"));
const executable = process.env.DOFFI_BROWSER_EXECUTABLE ?? (shell ? join(browserCache, shell, "chrome-headless-shell-linux64/chrome-headless-shell") : undefined);
assert.ok(executable && existsSync(executable), "Install the pinned Chromium headless shell first.");
let browser;
const servers = [], cases = [];
try {
  const archives = packPackages(join(temporary, "archives")); assert.equal(archives.length, 5);
  const dependencies = {};
  for (const module of ["core", "render", "editor", "office", "node"]) {
    const manifest = JSON.parse(readFileSync(join(root, "packages", module, "package.json"), "utf8"));
    const archive = archives.find((path) => path.endsWith(manifest.name.replace("@", "").replace("/", "-") + "-" + manifest.version + ".tgz"));
    assert.ok(archive); dependencies[manifest.name] = "file:" + archive;
  }
  const store = run(process.execPath, [pnpm, "store", "path"]);
  async function consumer(kind) {
    const directory = join(temporary, kind + " consumer"); mkdirSync(directory);
    const deps = { ...dependencies };
    if (kind === "react") Object.assign(deps, { react: "19.3.1", "react-dom": "19.3.1" });
    writeFileSync(join(directory, "package.json"), JSON.stringify({ private: true, type: "module", dependencies: deps, pnpm: { overrides: dependencies } }));
    run(process.execPath, [pnpm, "install", "--prefer-offline", "--ignore-scripts", "--ignore-workspace", "--no-frozen-lockfile", "--store-dir", store], directory);
    assert.equal(existsSync(join(directory, "node_modules/react")), kind === "react");
    writeFileSync(join(directory, "index.html"), '<meta charset="utf-8"><div id="preview"></div><div id="workbench"></div><script type="module" src="/entry.js"></script>');
    writeFileSync(join(directory, "entry.js"), `
import { openDocument, createDocumentSession } from '@doffi/core';
import { mountDocumentPreview } from '@doffi/render';
import { mountDocumentWorkbench } from '@doffi/editor';
import { exportDocx, importDocx, writeDocumentZip } from '@doffi/office';
${kind === "react" ? "import { createElement } from 'react'; import { createRoot } from 'react-dom/client'; import { DocumentWorkbench } from '@doffi/editor/react';" : ""}
const files = new Map(Object.entries(${JSON.stringify(sources)}).map(([path, text]) => [path, new TextEncoder().encode(text)]));
const original = openDocument(files);
const model = ${kind === "react" ? "openDocument(importDocx(exportDocx(original).bytes, {documentId:'packed-docx'}))" : "original"};
const preview = mountDocumentPreview(document.getElementById('preview'), model, {layout:'pages',title:'独立只读预览'});
await preview.ready;
const session = createDocumentSession(model);
const nodes = [...model.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)].map(match => model.getNode(match[1]));
const title = nodes.find(node => /^h[1-6]$|^p$/.test(node.tagName) && node.text === '项目实施报告');
const summary = nodes.find(node => node.tagName === 'p' && node.text === '项目计划于第三季度完成。');
let view;
const options = {session,layout:'pages',onSubmitAnnotations(request){window.submitted=request;},onError(error){window.hostError=error.message;}};
${kind === "react" ? "const reactRoot=createRoot(document.getElementById('workbench')); reactRoot.render(createElement(DocumentWorkbench,{...options,onViewChange(value){view=value;}}));" : "view=mountDocumentWorkbench(document.getElementById('workbench'),session,options);"}
window.host = {titleId:title.nodeId,summaryId:summary.nodeId,session,preview,
  ready:()=>Boolean(view), revision:()=>session.snapshot().manifest.revision,
  proposal(){const request=view.request()??view.buildRequest(); const fixed=${JSON.stringify(proposal)}; view.loadProposal(request,{...fixed,requestId:request.requestId,documentId:request.documentId,baseRevision:request.baseRevision,annotationIds:request.annotations.map(a=>a.annotationId),operations:fixed.operations.map(op=>({...op,nodeId:op.nodeId==='n-summary'?summary.nodeId:title.nodeId}))});},
  reanchor(){const annotations=session.getAnnotations(), text=session.getNode(summary.nodeId).text;
    const operations=annotations.map(a=>({op:'updateAnnotation',annotationId:a.annotationId,changes:a.status==='needs_reanchor'?{anchor:{kind:'text',nodeId:summary.nodeId,from:0,to:text.length,quote:text},status:'resolved'}:{status:'resolved'}}));
    session.commit(session.prepare({protocolVersion:1,documentId:session.snapshot().manifest.documentId,baseRevision:session.snapshot().manifest.revision,operations}));},
  save(){const native=exportDocx(session);return {docx:[...native.bytes],report:native.report,zip:[...writeDocumentZip(session.serialize())],snapshot:session.snapshot()};},
  dispose(){${kind === "react" ? "reactRoot.unmount();" : "view.dispose();"}preview.dispose();session.dispose();}
};
`);
    await build({ root: directory, configFile: false, logLevel: "silent", build: { target: "es2022", outDir: "dist" } });
    const serving = join(directory, "dist");
    const server = createServer((request, response) => {
      const path = resolve(serving, "." + new URL(request.url, "http://localhost").pathname);
      if (!path.startsWith(serving + "/")) { response.writeHead(404).end(); return; }
      const file = path.endsWith("/") ? join(path, "index.html") : path;
      try { response.setHeader("Content-Type", ({ ".html": "text/html", ".js": "text/javascript", ".css": "text/css" })[extname(file)] ?? "application/octet-stream"); response.end(readFileSync(file)); }
      catch { response.writeHead(404).end(); }
    });
    await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve)); servers.push(server);
    return { directory, url: "http://127.0.0.1:" + server.address().port + "/index.html" };
  }
  browser = await chromium.launch({ executablePath: executable, headless: true });
  for (const kind of ["javascript", "react"]) {
    const { directory, url } = await consumer(kind), page = await browser.newPage(), errors = [];
    page.on("pageerror", (error) => errors.push(error.message));
    await page.goto(url); await page.waitForFunction(() => window.host?.ready());
    const { titleId, summaryId } = await page.evaluate(() => ({ titleId: window.host.titleId, summaryId: window.host.summaryId }));
    const readonly = page.frameLocator('iframe[title="独立只读预览"]');
    await expect(readonly.locator('[contenteditable="true"]')).toHaveCount(0);
    await expect(readonly.locator(`[data-node-id="${titleId}"]`).first()).toHaveText("项目实施报告");
    await page.getByRole("button", { name: "编辑正文", exact: true }).click();
    const input = page.frameLocator('iframe[title="正文编辑"]');
    await input.locator(`[data-node-id="${titleId}"]`).evaluate((element) => {
      const range = document.createRange(); range.selectNodeContents(element); range.collapse(false);
      const selection = document.getSelection(); selection.removeAllRanges(); selection.addRange(range);
      document.querySelector('[data-office-editor]').focus(); document.dispatchEvent(new Event("selectionchange"));
    });
    await page.keyboard.insertText("（人工）");
    await expect.poll(() => page.evaluate(() => window.host.revision())).toBe(1);
    await expect(input.locator(`[data-node-id="${titleId}"]`)).toHaveText("项目实施报告（人工）");
    await page.getByRole("button", { name: "返回预览与批注", exact: true }).click();
    await page.getByRole("button", { name: "统一提交批注", exact: true }).click();
    await page.waitForFunction(() => window.submitted?.baseRevision === 1);
    await page.evaluate(() => window.host.proposal());
    await page.getByRole("button", { name: "整批应用提案", exact: true }).click();
    await expect.poll(() => page.evaluate(() => window.host.revision())).toBe(2);
    const incomplete = await page.evaluate(() => window.host.save());
    assert.equal(incomplete.report.unexportedAnnotations.length, 1);
    assert.equal(incomplete.report.unexportedAnnotations[0].reason, "needs_reanchor");
    await page.getByRole("button", { name: "撤销", exact: true }).click();
    await expect.poll(() => page.evaluate(() => window.host.revision())).toBe(3);
    await page.getByRole("button", { name: "重做", exact: true }).click();
    await expect.poll(() => page.evaluate(() => window.host.revision())).toBe(4);
    // Explicit host/user action, never an automatic semantic guess.
    await page.evaluate(() => window.host.reanchor());
    const saved = await page.evaluate(() => window.host.save());
    assert.equal(saved.snapshot.manifest.revision, 5); assert.deepEqual(saved.report.unexportedAnnotations, []);
    assert.equal(saved.report.exportedAnnotationIds.length, 2);
    for (const extension of ["docx", "zip"]) writeFileSync(join(cache, kind + "." + extension), new Uint8Array(saved[extension]));
    writeFileSync(join(cache, kind + ".report.json"), JSON.stringify(saved.report, null, 2));
    await page.screenshot({ path: join(cache, kind + ".png"), fullPage: true });
    await page.evaluate(() => window.host.dispose()); assert.deepEqual(errors, []); await page.close();
    // This program executes from outside the repository and resolves only tarball exports.
    writeFileSync(join(directory, "node.mjs"), `
import assert from 'node:assert/strict'; import {readFileSync,writeFileSync} from 'node:fs';
import {openDocument} from '@doffi/core'; import {readDocumentZip,importDocx} from '@doffi/office'; import {runConversionProcess} from '@doffi/node/dev'; import {renderDocumentPdf} from '@doffi/node';
const directory=${JSON.stringify(cache)}, kind=${JSON.stringify(kind)}, bytes=new Uint8Array(readFileSync(directory+'/'+kind+'.zip'));
const model=openDocument(readDocumentZip(bytes));assert.equal(model.snapshot().manifest.revision,5);
assert.equal(model.getNode(${JSON.stringify(titleId)}).text,'项目实施报告（人工）');
assert.ok(model.getNode(${JSON.stringify(summaryId)}).text.includes('第四季度'));assert.ok(model.snapshot().annotations.every(a=>a.status==='resolved'));
const docx=new Uint8Array(readFileSync(directory+'/'+kind+'.docx')), native=openDocument(importDocx(docx,{documentId:'packed-native'}));
assert.equal(native.snapshot().annotations.length,2);assert.ok(native.snapshot().annotations.every(a=>a.status==='resolved'));
const pdf=await renderDocumentPdf(model,{browserExecutable:${JSON.stringify(executable)}});writeFileSync(directory+'/'+kind+'.pdf',pdf.bytes);
assert.equal(pdf.report.revision,5);
const processResult=await runConversionProcess({protocolVersion:1,taskId:'packed-closure',format:'pdf',input:{format:'docx',documentId:'packed-cli',bytesBase64:Buffer.from(docx).toString('base64')}},{browserExecutable:${JSON.stringify(executable)}});
assert.equal(processResult.ok,true);assert.equal(processResult.output.report.documentId,'packed-cli');assert.equal(processResult.output.mimeType,'application/pdf');
writeFileSync(directory+'/'+kind+'.cli.pdf',Buffer.from(processResult.output.bytesBase64,'base64'));
writeFileSync(directory+'/'+kind+'.node.json',JSON.stringify({engine:pdf.engine,pages:pdf.report.pages.length,revision:pdf.report.revision}));
`);
    run(process.execPath, ["node.mjs"], directory);
    for (const suffix of [".pdf", ".cli.pdf"]) {
      const text = run("pdftotext", [join(cache, kind + suffix), "-"]).replace(/\s/g, "");
      assert.ok(text.includes("项目实施报告（人工）")); assert.ok(text.includes("第四季度"));
    }
    const reopened = join(cache, "reopened-" + kind); mkdirSync(reopened);
    run("libreoffice", ["-env:UserInstallation=" + pathToFileURL(join(temporary, "lo-" + kind)).href, "--headless", "--convert-to", "docx:Office Open XML Text", "--outdir", reopened, join(cache, kind + ".docx")]);
    writeFileSync(join(directory, "reopen.mjs"), `import assert from 'node:assert/strict';import {readFileSync} from 'node:fs';import {openDocument} from '@doffi/core';import {importDocx} from '@doffi/office';const model=openDocument(importDocx(new Uint8Array(readFileSync(${JSON.stringify(join(reopened, kind + ".docx"))})),{documentId:'office-reopened'}));const snapshot=model.snapshot();assert.equal(snapshot.annotations.length,2);assert.ok(snapshot.annotations.every(a=>a.status==='resolved'));const text=model.getNode(snapshot.metadata.word.stories.find(s=>s.kind==='main').rootNodeId).text;assert.ok(text.includes('项目实施报告（人工）'));assert.ok(text.includes('第四季度'));`);
    run(process.execPath, ["reopen.mjs"], directory);
    cases.push({ host: kind, input: kind === "react" ? "docx" : "html-package", revision: 5, resolvedComments: 2, docxSha256: digest(saved.docx), zipSha256: digest(saved.zip), ...JSON.parse(readFileSync(join(cache, kind + ".node.json"), "utf8")) });
    console.log("Packed " + kind + " preview, input, proposal, history, native DOCX, Node PDF/CLI and Office reopen passed.");
  }
  writeFileSync(join(cache, "result.json"), JSON.stringify({ environment: { node: process.version, chromium: browser.version(), libreOffice: run("libreoffice", ["--version"]), word: "not-verified", wps: "not-verified", fonts: ["Noto Sans CJK SC", "Liberation Serif"].map((family) => ({family,sha256:digest(readFileSync(run("fc-match", ["-f", "%{file}", family])))})) }, archives: archives.map(path=>({name:path.split('/').pop(),sha256:digest(readFileSync(path))})), cases }, null, 2) + "\n");
} finally {
  if (browser) await browser.close();
  await Promise.all(servers.map(server => new Promise(resolve => server.close(resolve))));
  rmSync(temporary, { recursive: true, force: true });
}

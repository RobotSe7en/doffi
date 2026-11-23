/** Actual native Office acceptance. Development tools only; no SDK I/O adapter. */
import assert from "node:assert/strict";
import { execFileSync, spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { join, resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { createDocumentSession, openDocument } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";
const root = resolve(import.meta.dirname, ".."), cache = join(root, ".cache/docx-export-acceptance");
const run = (exe, args) => execFileSync(exe, args, { cwd: root, encoding: "utf8", timeout: 55_000, maxBuffer: 8 * 1024 * 1024 }).trim();
const digest = (bytes) => createHash("sha256").update(bytes).digest("hex");
const mainText = (model) => model.getNode(model.snapshot().metadata.word.stories.find((story) => story.kind === "main").rootNodeId).text;
const imageAssets = snapshot => new Set(snapshot.manifest.assets.filter(asset => asset.mimeType.startsWith("image/")).map(asset => asset.sha256)).size;
for (const name of ["source", "source-reopened", "native", "reopened", "pdf"]) { rmSync(join(cache, name), { recursive: true, force: true }); mkdirSync(join(cache, name), { recursive: true }); }
const poppler = spawnSync("pdftotext", ["-v"], { encoding: "utf8", timeout: 10_000 }); assert.equal(poppler.status, 0);
const environment = { libreOffice: run("libreoffice", ["--version"]), poppler: (poppler.stdout + poppler.stderr).trim().split("\n")[0], fonts: [], word: "not-verified", wps: "not-verified" };
for (const family of ["Noto Sans CJK SC", "Noto Serif CJK SC", "Liberation Serif"]) {
  assert.ok(run("fc-match", ["-f", "%{family}", family]).includes(family), "The acceptance font must be installed: " + family);
  const file = run("fc-match", ["-f", "%{file}", family]); environment.fonts.push({ family, sha256: digest(readFileSync(file)) });
}
const cases = [];
for (const producer of ["", "producer/"]) for (const name of readdirSync(join(root, "examples/docx", producer)).filter((name) => name.endsWith(".docx"))) {
  const bytes = new Uint8Array(readFileSync(join(root, "examples/docx", producer, name))), source = openDocument(importDocx(bytes, { documentId: "acceptance" }));
  const filename = (producer ? "producer-" : "") + name, native = exportDocx(source);
  assert.deepEqual(native.report.unexportedAnnotations, []);
  assert.deepEqual(exportDocx(source).bytes, native.bytes);
  writeFileSync(join(cache, "source", filename), bytes); writeFileSync(join(cache, "native", filename), native.bytes);
  cases.push({ filename, source, inputSha256: digest(bytes), exportSha256: digest(native.bytes), edited: false });
}
const edited = createDocumentSession(cases.find((item) => item.filename === "08-comments-revisions.docx").source);
const paragraph = [...edited.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)].map((match) => edited.getNode(match[1])).find((node) => node.tagName === "p" && node.text);
const operations = [{ op: "replaceText", nodeId: paragraph.nodeId, from: 0, to: 0, expectedText: "", text: "DOCX重新编辑确认 " },
  { op: "insertFragment", nodeId: paragraph.nodeId, position: "after", html: '<div>INLINE <strong>A</strong> <em>B</em></div><table style="width:200pt;border-collapse:collapse"><tr><td style="border:.5pt solid #336699;background-color:#eef4ff;padding:6pt">CELL <strong>A</strong> B</td></tr></table>' },
  ...edited.getAnnotations().map((annotation) => ({ op: "updateAnnotation", annotationId: annotation.annotationId, changes: { status: "resolved" } }))];
edited.commit(edited.prepare({ protocolVersion: 1, documentId: "acceptance", baseRevision: 0, operations }));
const native = exportDocx(edited); writeFileSync(join(cache, "native/edited-resolved.docx"), native.bytes);
cases.push({ filename: "edited-resolved.docx", source: openDocument(edited.serialize()), inputSha256: null, exportSha256: digest(native.bytes), edited: true }); edited.dispose();
for (const [property, active, variant] of [["font-style", "italic", "Italic"], ["font-weight", "bold", "Bold"]]) {
  const encoder = new TextEncoder(), decoder = new TextDecoder();
  const files = new Map(["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map((name) => [name, new Uint8Array(readFileSync(join(root, "examples/report", name)))]));
  const body = '<article data-node-id="root" data-office-kind="document"><div data-node-id="story" data-office-story="main"><section data-node-id="section"><p data-node-id="p" class="P" data-word-style="P"><span data-node-id="s" class="C" data-word-style="C">TOGGLE_' + variant.toUpperCase() + '</span></p></section></div></article>';
  files.set("content.html", encoder.encode(decoder.decode(files.get("content.html")).replace(/<article[\s\S]*<\/article>/, body)));
  files.set("styles.css", encoder.encode('article{font-family:"Liberation Serif";font-size:12pt}.P{' + property + ':' + active + '}.C{' + property + ':' + active + '}'));
  files.set("annotations.json", encoder.encode("[]"));
  const metadata = JSON.parse(decoder.decode(files.get("metadata.json"))); metadata.sections = [{ sectionId: "section", rootNodeId: "section" }];
  metadata.word = { stories: [{ storyId: "main", kind: "main", rootNodeId: "story" }], styles: [{ styleId: "P", kind: "paragraph", cssClass: "P", paragraphProperties: [], runProperties: [property] }, { styleId: "C", kind: "character", cssClass: "C", runProperties: [property] }], numbering: [], hyperlinks: [], bookmarks: [], fields: [], tabs: [], revisions: [], tables: [] };
  files.set("metadata.json", encoder.encode(JSON.stringify(metadata)));
  const source = openDocument(files), output = exportDocx(source), filename = "toggle-" + active + ".docx";
  writeFileSync(join(cache, "native", filename), output.bytes); cases.push({ filename, source, inputSha256: null, exportSha256: digest(output.bytes), edited: false, variant });
}
const profile = join(cache, "profile-" + process.pid), lo = ["-env:UserInstallation=" + pathToFileURL(profile).href, "--headless"];
try {
  const inputs = cases.map((item) => join(cache, "native", item.filename));
  run("libreoffice", [...lo, "--convert-to", "docx", "--outdir", join(cache, "source-reopened"), ...cases.filter((item) => item.inputSha256 !== null).map((item) => join(cache, "source", item.filename))]);
  run("libreoffice", [...lo, "--convert-to", "docx", "--outdir", join(cache, "reopened"), ...inputs]);
  run("libreoffice", [...lo, "--convert-to", "pdf", "--outdir", join(cache, "pdf"), ...inputs]);
  const records = [];
  for (const item of cases) {
    const reopenedBytes = new Uint8Array(readFileSync(join(cache, "reopened", item.filename))), reopened = openDocument(importDocx(reopenedBytes, { documentId: "reopened" })), before = item.source.snapshot(), after = reopened.snapshot();
    assert.equal(after.metadata.sections.length, before.metadata.sections.length, item.filename + " sections");
    if (!before.metadata.word.fields.length) assert.equal(mainText(reopened), mainText(item.source), item.filename + " native text");
    // LibreOffice itself normalizes unsupported nested IF instructions and TOC
    // switches. Compare with an independent reopen of the original producer;
    // exact native controls/instructions are checked before Office in unit tests.
    const baseline = item.edited || !before.metadata.word.fields.length ? item.source : openDocument(importDocx(new Uint8Array(readFileSync(join(cache, "source-reopened", item.filename))), { documentId: "producer-baseline" }));
    assert.deepEqual(after.metadata.word.fields.map((field) => field.instruction.trim()), baseline.snapshot().metadata.word.fields.map((field) => field.instruction.trim()), item.filename + " producer field baseline");
    if (before.metadata.word.fields.length) assert.equal(mainText(reopened), mainText(baseline), item.filename + " producer field results");
    assert.deepEqual(after.metadata.word.revisions.map(({ kind, author, date }) => ({ kind, author, date })), before.metadata.word.revisions.map(({ kind, author, date }) => ({ kind, author, date })), item.filename + " revisions");
    assert.deepEqual(after.annotations.map(({ status, author }) => ({ status, author })), before.annotations.map(({ status, author }) => ({ status, author })), item.filename + " comments");
    assert.equal(imageAssets(after), imageAssets(before), item.filename + " image parts");
    const pdf = join(cache, "pdf", item.filename.replace(/\.docx$/, ".pdf")), info = run("pdfinfo", [pdf]), pages = Number(/Pages:\s+(\d+)/.exec(info)?.[1]);
    assert.ok(pages > 0, item.filename + " PDF pages");
    const text = run("pdftotext", [pdf, "-"]); assert.ok(text.trim(), item.filename + " searchable text");
    if (item.variant) assert.ok(run("pdffonts", [pdf]).includes("LiberationSerif-" + item.variant), item.filename + " actual Office font variant");
    if (item.edited) {
      assert.ok(text.replace(/\s/g, "").includes("DOCX重新编辑确认")); assert.ok(after.annotations.every((annotation) => annotation.status === "resolved"));
      const ids = [...after.html.matchAll(/data-node-id="([^"]+)"/g)].map((match) => reopened.getNode(match[1]));
      const cell = ids.find((node) => node.tagName === "td" && node.text === "CELL A B"); assert.ok(cell, "edited inline cell text");
      assert.equal(cell.childNodeIds.map((id) => reopened.getNode(id)).filter((node) => node.tagName === "p").length, 1, "edited inline cell remains one paragraph after Office");
      assert.ok(!cell.childNodeIds.map((id) => reopened.getNode(id)).filter((node) => node.tagName === "p").some((node) => node.attributes.style?.includes("border-")), "cell border must not become an inner paragraph border after Office");
      assert.ok(ids.some((node) => node.tagName === "p" && node.text === "INLINE A B"), "edited inline container remains one paragraph after Office");
    }
    if (!item.edited) {
      const reference = JSON.parse(readFileSync(join(root, "examples/docx/reference.json"), "utf8"));
      const entries = reference.fixtures ?? reference.records ?? reference.documents;
      if (entries) { const expected = entries.find((entry) => entry.file === item.filename.replace(/^producer-/, "")); if (expected) assert.equal(pages, expected.pages, item.filename + " page count"); }
    }
    const images = run("pdfimages", ["-list", pdf]).split("\n").filter((line) => /^\s*\d+\s+\d+\s+image\s/.test(line)).length;
    assert.ok(images >= imageAssets(before), item.filename + " rendered images");
    run("pdftoppm", ["-f", "1", "-singlefile", "-r", "96", "-png", pdf, pdf.replace(/\.pdf$/, "")]);
    records.push({ file: item.filename, inputSha256: item.inputSha256, exportSha256: item.exportSha256, reopenedSha256: digest(reopenedBytes), pages, images, comments: after.annotations.length, resolvedComments: after.annotations.filter((annotation) => annotation.status === "resolved").length });
  }
  writeFileSync(join(cache, "result.json"), JSON.stringify({ environment, cases: records }, null, 2) + "\n");
  process.stdout.write("Native DOCX reopen, editable text, fields, revisions, comments, PDF pages and images passed for " + records.length + " cases.\n");
} finally { rmSync(profile, { recursive: true, force: true }); }

/** Exercise the five packed SDKs in independent browser and Node hosts. */
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, join, resolve } from "node:path";
import { chromium, expect } from "@playwright/test";
import { build, preview } from "vite";
import { packPackages } from "./pack.mjs";

const root = resolve(import.meta.dirname, "..");
const cache = join(root, ".cache/workbook-closure-acceptance");
const fixtures = join(import.meta.dirname, "fixtures/workbook-closure");
const source = join(root, "examples/xlsx/09-classic-notes.xlsx");
const pnpm = process.env.npm_execpath;
assert(pnpm, "Run workbook closure through pnpm.");
const browserCache = process.env.PLAYWRIGHT_BROWSERS_PATH ?? join(root, ".playwright");
const shell = existsSync(browserCache) && readdirSync(browserCache).find(name => /^chromium_headless_shell-\d+$/.test(name));
const executable = process.env.DOFFI_BROWSER_EXECUTABLE ?? (shell && join(browserCache, shell, "chrome-headless-shell-linux64/chrome-headless-shell"));
assert(executable && existsSync(executable), "Install the pinned Chromium headless shell first.");

function run(command, args, cwd, options = {}) {
  return execFileSync(command, args, { cwd, encoding: "utf8", timeout: 90_000, maxBuffer: 8 * 1024 * 1024, ...options });
}
function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}
function saveJson(path, value) {
  writeFileSync(path, JSON.stringify(value, null, 2) + "\n");
}

async function editInBrowser(browser, url, output) {
  const page = await browser.newPage({ viewport: { width: 1280, height: 1000 } });
  const errors = [];
  page.on("pageerror", error => errors.push(error.message));
  try {
    // Playwright needs access to the closed SDK toolbar roots for real input.
    await page.addInitScript(() => {
      const attachShadow = Element.prototype.attachShadow;
      Element.prototype.attachShadow = function (options) {
        return attachShadow.call(this, { ...options, mode: "open" });
      };
    });
    await page.goto(url);
    await page.waitForFunction(() => Boolean(window.host));
    const layout = await page.evaluate(() => window.host.selectPreviewCell());
    assert.equal(layout.kind, "workbook");
    assert.equal(layout.revision, 0);
    assert.equal(layout.storedCellCount, 2);
    const readonly = page.getByRole("region", { name: "独立工作簿只读预览", exact: true });
    const value = readonly.getByRole("textbox", { name: "单元格值或公式", exact: true });
    await expect(value).toHaveValue("经典备注");
    await expect(value).toHaveAttribute("readonly", "");

    const inspect = () => page.evaluate(() => window.host.inspect());
    const address = page.getByRole("textbox", { name: "选择单元格或区域", exact: true });
    const formula = page.getByRole("textbox", { name: "单元格或公式输入", exact: true });
    await address.fill("A5");
    await address.press("Enter");
    await expect(address).toHaveValue("A5");
    await formula.fill("37");
    await formula.press("Control+Enter");
    await expect.poll(inspect).toMatchObject({ revision: 1, input: 37, formula: null, undoSteps: 1, errors: [] });

    await page.getByLabel("请求写入范围", { exact: true }).selectOption("global");
    await page.getByRole("button", { name: "提交工作簿批注", exact: true }).click();
    await expect(page.getByLabel("工作簿提案差异")).toBeVisible();
    assert.equal((await inspect()).formula, null);
    await page.getByRole("button", { name: "整批应用工作簿提案", exact: true }).click();
    await expect.poll(inspect).toMatchObject({ revision: 2, formula: "=A5*3", undoSteps: 2, errors: [] });
    await page.getByRole("button", { name: "撤销", exact: true }).click();
    await expect.poll(inspect).toMatchObject({ revision: 3, input: 37, formula: null, undoSteps: 1, errors: [] });
    await page.getByRole("button", { name: "重做", exact: true }).click();
    await expect.poll(inspect).toMatchObject({ revision: 4, input: 37, formula: "=A5*3", undoSteps: 2, errors: [] });

    const { xlsx, zip, report } = await page.evaluate(() => window.host.save());
    assert.equal(report.revision, 4);
    assert.deepEqual(report.formulaCaches, { written: 1, omitted: 0 });
    assert.equal(report.exportedAnnotationIds.length, 1);
    assert.deepEqual(report.unexportedAnnotations, []);
    writeFileSync(join(output, "edited.xlsx"), new Uint8Array(xlsx));
    writeFileSync(join(output, "edited.zip"), new Uint8Array(zip));
    saveJson(join(output, "xlsx-report.json"), report);
    await page.screenshot({ path: join(output, "host.png"), fullPage: true });
    await page.evaluate(() => window.host.dispose());
    await expect(page.locator("iframe")).toHaveCount(0);
    await expect(readonly).toHaveCount(0);
    assert.deepEqual(errors, []);
  } finally {
    await page.close();
  }
}

function verifyNodeAndCli(directory, output) {
  run(process.execPath, ["node.mjs", output, executable], directory);
  const report = JSON.parse(readFileSync(join(output, "print-report.json"), "utf8"));
  const workbook = { sheets: [{ sheetId: report.pages[0].sheetId, areas: [report.pages[0].range] }] };
  const requests = [
    { format: "pdf", input: { format: "xlsx", documentId: "packed-native" }, file: "edited.xlsx", revision: 0 },
    { format: "png", input: { format: "document-package-zip", documentId: "packed-workbook" }, file: "edited.zip", revision: 4 },
  ];
  const frames = requests.map(({ format, input, file }) => JSON.stringify({
    protocolVersion: 1, taskId: "packed-" + format, format,
    input: { ...input, bytesBase64: readFileSync(join(output, file)).toString("base64") }, options: { workbook },
  }));
  const responses = run(join(directory, "node_modules/.bin/doffi"), [], directory, {
    input: frames.join("\n") + "\n", env: { ...process.env, DOFFI_CHROMIUM_EXECUTABLE: executable },
  }).trim().split("\n").map(line => JSON.parse(line));
  assert.equal(responses.length, requests.length);
  for (const [index, { input, format, revision }] of requests.entries()) {
    const result = responses[index];
    assert.equal(result.ok, true, JSON.stringify(result.error));
    assert.equal(result.taskId, "packed-" + format);
    assert.equal(result.output.report.kind, "workbook");
    assert.equal(result.output.report.revision, revision);
    assert.equal(result.output.report.documentId, input.documentId);
    assert.deepEqual(result.output.report.pages, report.pages);
    assert.equal(result.output.report.formulaCells.valid, 1);
    writeFileSync(join(output, "cli." + format), Buffer.from(result.output.bytesBase64, "base64"));
  }
  for (const filename of ["edited.pdf", "cli.pdf"]) {
    const text = run("pdftotext", ["-layout", join(output, filename), "-"], directory);
    assert.match(text, /\b37\s+111\b/);
  }
  for (const filename of ["edited.png", "cli.png"]) {
    const png = readFileSync(join(output, filename));
    assert.deepEqual([...png.subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
    assert(png.readUInt32BE(16) > 0 && png.readUInt32BE(20) > 0);
  }
}

const temporary = mkdtempSync(join(tmpdir(), "doffi packed workbook-"));
let browser;
try {
  rmSync(cache, { recursive: true, force: true });
  mkdirSync(cache, { recursive: true });
  const archives = packPackages(join(temporary, "archives"));
  assert.equal(archives.length, 5);
  const dependencies = {};
  for (const name of ["core", "render", "editor", "office", "node"]) {
    const manifest = JSON.parse(readFileSync(join(root, "packages", name, "package.json"), "utf8"));
    const archive = archives.find(path => basename(path) === `doffi-${name}-${manifest.version}.tgz`);
    assert(archive, `Missing packed ${manifest.name}.`);
    dependencies[manifest.name] = "file:" + archive;
  }
  browser = await chromium.launch({ executablePath: executable, headless: true });
  for (const kind of ["javascript", "react"]) {
    const directory = join(temporary, kind), output = join(cache, kind);
    mkdirSync(join(directory, "public"), { recursive: true });
    mkdirSync(output);
    const react = kind === "react" ? { react: "19.3.1", "react-dom": "19.3.1" } : {};
    saveJson(join(directory, "package.json"), {
      private: true, type: "module", dependencies: { ...dependencies, ...react }, pnpm: { overrides: dependencies },
    });
    run(process.execPath, [pnpm, "install", "--prefer-offline", "--ignore-scripts", "--ignore-workspace", "--no-frozen-lockfile"], directory);
    const installed = readdirSync(join(directory, "node_modules/.pnpm"));
    for (const name of ["react", "react-dom"]) {
      assert.equal(installed.some(path => path.startsWith(name + "@")), kind === "react");
    }
    copyFileSync(source, join(directory, "public/source.xlsx"));
    for (const file of ["browser.mjs", "node.mjs", kind + ".mjs"]) copyFileSync(join(fixtures, file), join(directory, file));
    writeFileSync(join(directory, "index.html"), `<meta charset="utf-8"><div id="preview"></div><div id="workbench"></div><script type="module" src="./${kind}.mjs"></script>`);
    const vite = { root: directory, configFile: false, logLevel: "silent" };
    await build({ ...vite, build: { target: "es2022" } });
    const server = await preview({ ...vite, preview: { host: "127.0.0.1", port: 0, strictPort: true } });
    try {
      await editInBrowser(browser, `http://127.0.0.1:${server.httpServer.address().port}/`, output);
    } finally {
      await server.close();
    }
    verifyNodeAndCli(directory, output);
    console.log(`Packed ${kind}: browser edit, proposal, undo/redo, native and package reopen, Node PDF/PNG and CLI passed.`);
  }
  saveJson(join(cache, "acceptance.json"), {
    sourceHead: process.env.DOFFI_SOURCE_HEAD ?? run("git", ["rev-parse", "HEAD"], root).trim(),
    source: { file: basename(source), sha256: sha256(source) },
    archives: archives.map(path => ({ file: basename(path), sha256: sha256(path) })),
    chromium: browser.version(),
    consumers: ["javascript", "react"],
  });
} finally {
  try { await browser?.close(); }
  finally { rmSync(temporary, { recursive: true, force: true }); }
}

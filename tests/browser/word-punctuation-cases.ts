import { writeFileSync } from "node:fs";
import { expect, test } from "@playwright/test";
import type { Locator, Page, TestInfo } from "@playwright/test";
import { openDocument } from "@doffi/core";
import type { DocumentPackage } from "@doffi/core";
import type { DocumentTextEditor } from "@doffi/editor";
import type { DocumentPreview } from "@doffi/render";
import { exportDocx, importDocx } from "@doffi/office";
import {
  body,
  mutateDocx,
  styles,
  wordNamespace,
} from "../../packages/office/tests/regression-fixtures.js";

type Mount = (page: Page, files: DocumentPackage, mode?: "preview" | "edit") => Promise<void>;
function run(text: string, properties = ""): string {
  return (
    "<w:r>" +
    (properties ? "<w:rPr>" + properties + "</w:rPr>" : "") +
    "<w:t>" +
    text +
    "</w:t></w:r>"
  );
}
function para(width: number, content: string, properties = ""): string {
  return (
    "<w:p><w:pPr>" +
    properties +
    '<w:ind w:right="' +
    (2880 - 15 * width) +
    '"/></w:pPr>' +
    content +
    "</w:p>"
  );
}
function fixture(paragraphs: string, settings = "", grid = false): DocumentPackage {
  const bytes = mutateDocx("01-text-numbering.docx", {
    "word/document.xml": (source) =>
      body(
        source,
        paragraphs +
          '<w:sectPr><w:pgSz w:w="4320" w:h="5400"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360" w:gutter="0"/>' +
          (grid ? '<w:docGrid w:type="snapToChars" w:linePitch="360" w:charSpace="0"/>' : "") +
          "</w:sectPr>",
      ),
    "word/settings.xml": () =>
      '<w:settings xmlns:w="' + wordNamespace + '">' + settings + "</w:settings>",
    "word/styles.xml": () =>
      styles(
        '<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/><w:rPr>' +
          '<w:rFonts w:ascii="Noto Sans CJK SC" w:hAnsi="Noto Sans CJK SC" w:eastAsia="Noto Sans CJK SC"/>' +
          '<w:sz w:val="24"/><w:lang w:eastAsia="zh-CN"/></w:rPr></w:style>',
      ),
  });
  return importDocx(bytes, { documentId: "word-punctuation", unsupportedContent: "reject" });
}
const hangingOff = '<w:overflowPunct w:val="0"/>';
const paragraphs = (page: Page) =>
  page.frameLocator("iframe").locator(".sdk-page-flow p, .sdk-column-flow p");

function geometry(paragraph: Locator) {
  return paragraph.evaluate((root) => {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT),
      range = document.createRange(),
      left = root.getBoundingClientRect().left,
      width = root.getBoundingClientRect().width,
      marks: { text: string; x: number; end: number; y: number; height: number }[] = [];
    for (let text = walker.nextNode(); text; text = walker.nextNode()) {
      if (text.parentElement?.closest(".sdk-word-kinsoku-boundary,[data-sdk-word-autospace]"))
        continue;
      for (const { segment, index } of new Intl.Segmenter("und", {
        granularity: "grapheme",
      }).segment(text.textContent!)) {
        range.setStart(text, index);
        range.setEnd(text, index + segment.length);
        const rect = range.getBoundingClientRect();
        if (rect.width > 0.1 && rect.height > 0.1)
          marks.push({
            text: segment,
            x: rect.left - left,
            end: rect.right - left,
            y: rect.top,
            height: rect.height,
          });
      }
    }
    return { marks, width };
  });
}
async function capture(page: Page, info: TestInfo) {
  const { entries, fonts, html } = await page.evaluate(() => {
    const host = window.paginationHost,
      source = host.session!.snapshot().html;
    return {
      entries: [...host.session!.serialize()].map(([name, bytes]) => [name, [...bytes]] as const),
      html: source,
      fonts: host.layout().fonts,
    };
  });
  expect(html).not.toMatch(/sdk-word-punctuation|sdk-word-kinsoku-boundary|\u2060/);
  expect(await page.evaluate(() => window.paginationHost.errors)).toEqual([]);
  expect(
    fonts.some((font) => font.requestedFamily === "Noto Sans CJK SC" && !font.substituted),
  ).toBe(true);
  writeFileSync(info.outputPath("font-report.json"), JSON.stringify(fonts, null, 2));
  await page.screenshot({ path: info.outputPath("punctuation.png") });
  const model = openDocument(
      new Map(entries.map(([name, bytes]) => [name, Uint8Array.from(bytes)])),
      { kind: "document" },
    ),
    bytes = exportDocx(model).bytes;
  writeFileSync(info.outputPath("punctuation.docx"), bytes);
  for (let pass = 0, native = bytes; pass < 2; pass++) {
    const next = openDocument(
      importDocx(native, { documentId: "punctuation-reopened", unsupportedContent: "reject" }),
      { kind: "document" },
    );
    expect(next.snapshot().metadata.word!.typography).toEqual(
      model.snapshot().metadata.word!.typography,
    );
    expect(next.snapshot().metadata.word!.doNotWrapTextWithPunct).toBe(
      model.snapshot().metadata.word!.doNotWrapTextWithPunct,
    );
    native = exportDocx(next).bytes;
  }
}
export function registerWordPunctuationTests(mount: Mount): void {
  test("R19 hangs one closing mark across runs and compresses actual line starts", async ({
    page,
  }, info) => {
    await mount(
      page,
      fixture(
        para(33, run("甲乙") + run("，") + run("丙丁"), "<w:overflowPunct/>") +
          para(33, run("甲乙") + run("，") + run("丙丁"), hangingOff) +
          para(36, run("（") + run("甲乙"), hangingOff + "<w:topLinePunct/>") +
          para(36, run("（") + run("甲乙"), hangingOff + '<w:topLinePunct w:val="0"/>') +
          para(33, run("甲乙，。丙丁"), "<w:overflowPunct/>") +
          para(33, run("甲乙（丙丁"), hangingOff + "<w:topLinePunct/>"),
      ),
      "edit",
    );
    const p = paragraphs(page);
    const { marks: on, width } = await geometry(p.nth(0)),
      { marks: off } = await geometry(p.nth(1)),
      { marks: leading } = await geometry(p.nth(2)),
      { marks: plain } = await geometry(p.nth(3)),
      { marks: consecutive, width: pairWidth } = await geometry(p.nth(4));
    expect(on[2]!.y).toBeCloseTo(on[0]!.y, 1);
    expect(on[2]!.end).toBeGreaterThan(width + 8);
    expect(on[3]!.y).toBeGreaterThan(on[2]!.y + 8);
    expect(off[2]!.y).toBeGreaterThan(off[0]!.y + 8);
    expect(leading[1]!.y).toBeCloseTo(leading[0]!.y, 1);
    expect(plain[1]!.y).toBeCloseTo(plain[0]!.y, 1);
    expect(plain[1]!.x - leading[1]!.x).toBeGreaterThan(5);
    expect(consecutive.filter((mark) => mark.end > pairWidth + 1).length).toBeLessThanOrEqual(1);
    const { marks: wrapped } = await geometry(p.nth(5));
    expect(wrapped[2]!.y).toBeGreaterThan(wrapped[1]!.y + 8);
    expect(wrapped[2]!.x).toBeLessThan(-5);
    await capture(page, info);
    const entries = await page.evaluate(() =>
      [...window.paginationHost.session!.serialize()].map(
        ([path, bytes]) => [path, [...bytes]] as const,
      ),
    );
    await page.evaluate((entries) => window.paginationHost.mount(entries, "preview"), entries);
    expect(
      await page.evaluate(() => {
        const print = (window.paginationHost.view as DocumentPreview).printDocument();
        return Array.from(print.querySelectorAll<HTMLElement>(".sdk-word-punctuation")).some(
          (span) => span.textContent === "，" && span.style.inlineSize === "0px",
        );
      }),
    ).toBe(true);
    await page.evaluate(
      (entries) => window.paginationHost.mount(entries, "preview", { layout: "continuous" }),
      entries,
    );
    const continuous = await geometry(
      page.frameLocator("iframe").locator('[data-office-kind="document"] p').first(),
    );
    expect(continuous.marks[2]!.end).toBeGreaterThan(continuous.width + 8);
  });

  test("R19 replaces language-local kinsoku defaults and page controls edit the rules in one undo step", async ({
    page,
  }, info) => {
    const settings =
      '<w:strictFirstAndLastChars w:val="0"/><w:noLineBreaksAfter w:lang="zh-CN" w:val="乙"/><w:noLineBreaksBefore w:lang="zh-CN" w:val="丁"/>';
    await mount(
      page,
      fixture(
        para(33, run("甲乙") + run("丙戊"), hangingOff) +
          para(33, run("甲丙") + run("丁戊"), hangingOff) +
          para(33, run("甲丙，乙"), hangingOff) +
          para(33, run("甲乙丙丁", '<w:lang w:eastAsia="zh-TW"/>'), hangingOff) +
          para(33, run("甲乙ぁ丙", '<w:lang w:eastAsia="ja-JP"/>'), hangingOff) +
          para(96, run("甲，丙"), hangingOff),
        settings,
      ),
      "edit",
    );
    const p = paragraphs(page);
    for (const n of [0, 1]) {
      const { marks } = await geometry(p.nth(n));
      expect(marks[1]!.y).toBeCloseTo(marks[2]!.y, 1);
      expect(marks[1]!.y).toBeGreaterThan(marks[0]!.y + 8);
    }
    const { marks: replaced } = await geometry(p.nth(2)),
      { marks: other } = await geometry(p.nth(3)),
      { marks: standard } = await geometry(p.nth(4)),
      { marks: uncompressed } = await geometry(p.nth(5));
    expect(replaced[1]!.y).toBeCloseTo(replaced[0]!.y, 1);
    expect(replaced[2]!.y).toBeGreaterThan(replaced[1]!.y + 8);
    expect(other[1]!.y).toBeCloseTo(other[0]!.y, 1);
    expect(standard[2]!.y).toBeGreaterThan(standard[1]!.y + 8);
    expect(
      await p.first().evaluate((paragraph) => {
        const editor = paragraph.closest<HTMLElement>("[data-office-editor]")!;
        editor.focus();
        const selection = document.getSelection()!,
          range = document.createRange();
        range.selectNodeContents(paragraph);
        selection.removeAllRanges();
        selection.addRange(range);
        const data = new DataTransfer(),
          event = new ClipboardEvent("copy", { bubbles: true, cancelable: true });
        Object.defineProperty(event, "clipboardData", { value: data });
        editor.dispatchEvent(event);
        return data.getData("text/plain").trim();
      }),
    ).toBe("甲乙丙戊");

    await page.evaluate(() =>
      (window.paginationHost.view as DocumentTextEditor).execute({
        type: "word",
        action: "pageSettings",
      }),
    );
    await expect(page.getByLabel("行首禁则语言", { exact: true })).toHaveValue("zh-CN");
    const beforeCharacters = page.getByLabel("行首禁则字符", { exact: true });
    await beforeCharacters.press("Control+a");
    await beforeCharacters.press("X");
    await expect(beforeCharacters).toHaveValue("X");
    await page.getByLabel("自定义行尾禁则", { exact: true }).uncheck();
    await page.getByLabel("东亚字符压缩", { exact: true }).selectOption("compressPunctuation");
    await page.getByLabel("日文禁则", { exact: true }).selectOption("1");
    await page.getByRole("button", { name: "应用东亚排版设置" }).click();
    await page.evaluate(() => window.paginationHost.view!.settled());
    expect(
      await page.evaluate(
        () => window.paginationHost.session!.snapshot().metadata.word!.typography,
      ),
    ).toEqual({
      characterSpacingControl: "compressPunctuation",
      strictFirstAndLastChars: true,
      noLineBreaksBefore: { language: "zh-CN", characters: "X" },
    });
    const { marks: changed } = await geometry(p.nth(0));
    expect(changed[1]!.y).toBeCloseTo(changed[0]!.y, 1);
    const { marks: strict } = await geometry(p.nth(4)),
      { marks: compressed } = await geometry(p.nth(5));
    expect(strict[2]!.y).toBeCloseTo(strict[1]!.y, 1);
    expect(uncompressed[2]!.x - compressed[2]!.x).toBeGreaterThan(5);
    await capture(page, info);
    await page.evaluate(() => window.paginationHost.session!.undo());
    await page.evaluate(() => window.paginationHost.view!.settled());
    const { marks: reverted } = await geometry(p.nth(0));
    expect(reverted[1]!.y).toBeCloseTo(reverted[2]!.y, 1);
  });

  test("R19 grid compatibility and mixed-size punctuation follow visible review text and preserve input", async ({
    page,
  }, info) => {
    await mount(
      page,
      fixture(
        para(33, run("甲乙") + run("，", '<w:sz w:val="32"/>') + run("丙丁"), "<w:overflowPunct/>"),
        "<w:trackRevisions/><w:compat><w:doNotWrapTextWithPunct/></w:compat>",
        true,
      ),
      "edit",
    );
    const p = paragraphs(page).first(),
      id = (await p.getAttribute("data-node-id"))!;
    const before = await geometry(p);
    expect(before.marks[2]!.y).toBeGreaterThan(before.marks[0]!.y + 8);
    await page.evaluate((id) => {
      const s = window.paginationHost.session!,
        { manifest } = s.snapshot(),
        root = s.getNode(s.snapshot().metadata.word!.stories[0]!.rootNodeId)!.parentNodeId!;
      s.commit(
        s.prepare({
          protocolVersion: 1,
          documentId: manifest.documentId,
          baseRevision: manifest.revision,
          operations: [
            { op: "setWordOptions", nodeId: root, doNotWrapTextWithPunct: false },
            { op: "trackWordText", nodeId: id, from: 1, to: 2, expectedText: "乙", text: "" },
          ],
        }),
      );
      (window.paginationHost.view as DocumentTextEditor).execute({
        type: "word",
        action: "review",
      });
    }, id);
    const review = page.getByLabel("修订显示", { exact: true });
    await review.selectOption("original");
    await page.evaluate(() => window.paginationHost.view!.settled());
    const original = await geometry(p);
    expect(original.marks[2]!.end).toBeGreaterThan(original.width + 8);
    await review.selectOption("final");
    await page.evaluate(() => window.paginationHost.view!.settled());
    const final = await geometry(p);
    expect(final.marks.map((mark) => mark.text).join("")).toBe("甲，丙丁");
    expect(final.marks[1]!.end).toBeLessThanOrEqual(final.width + 4);
    await p.evaluate((paragraph) => {
      (paragraph.closest("[data-office-editor]") as HTMLElement).focus();
      const comma = [...paragraph.querySelectorAll(".sdk-word-punctuation")].find(
        (span) => span.textContent === "，",
      )!;
      const range = document.createRange();
      range.selectNodeContents(comma);
      range.collapse(false);
      getSelection()!.removeAllRanges();
      getSelection()!.addRange(range);
    });
    await page.keyboard.insertText("新");
    await expect
      .poll(() => page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, id))
      .toBe("甲乙，新丙丁");
    await capture(page, info);
    await page.evaluate(() => window.paginationHost.session!.undo());
    expect(await page.evaluate((id) => window.paginationHost.session!.getNode(id)!.text, id)).toBe(
      "甲乙，丙丁",
    );
  });
}

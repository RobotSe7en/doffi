import { findWordText, getWordOutline, getWordTextStatistics } from "@doffi/core";
import type { Anchor, ContentDocumentModel, WordTextMatch } from "@doffi/core";
import { createWordDockedPane, replaceSameSideDock } from "./word-pane-chrome.js";

/** Reading commands never prepare an editing candidate or change document history. */
export function createWordReadingControls(
  parent: ShadowRoot,
  before: HTMLElement,
  model: ContentDocumentModel,
  focus: (anchor: Anchor, revision: number) => void,
  onError: (error: unknown) => void,
): {
  open(action: "find" | "outline" | "wordCount"): void;
  refresh(): void;
  dispose(): void;
} {
  const document = before.ownerDocument;
  let panel: HTMLElement | undefined, controller: AbortController | undefined;
  let refresh: (() => void) | undefined;
  let resultControls: AbortController | undefined;
  const close = (): void => {
    resultControls?.abort();
    controller?.abort();
    panel?.remove();
    panel = undefined;
    refresh = undefined;
  };
  function open(action: "find" | "outline" | "wordCount"): void {
    close();
    const dock = action === "wordCount" ? "floating" : "left";
    replaceSameSideDock(parent, dock);
    controller = new AbortController();
    const signal = controller.signal;
    const chrome = createWordDockedPane(document, {
      ariaLabel: action === "find" ? "全文查找" : action === "outline" ? "文档导航" : "字数统计",
      title: action === "find" ? "在文档中查找" : action === "outline" ? "导航" : "字数统计",
      dock,
      closeLabel: "关闭导航",
      onClose: close,
      signal,
    });
    panel = chrome.panel;
    const body = chrome.body;
    const button = (label: string, run: () => void, target = body, listenerSignal = signal) => {
      const element = document.createElement("button");
      element.type = "button";
      element.textContent = label;
      element.addEventListener(
        "click",
        () => {
          try {
            run();
          } catch (error) {
            onError(error);
          }
        },
        { signal: listenerSignal },
      );
      target.append(element);
      return element;
    };
    const results = document.createElement("div");
    results.style.cssText = "display:flex;flex-direction:column;gap:4px";
    if (action === "find") {
      const searchCard = document.createElement("div");
      searchCard.className = "word-pane-section";
      const query = document.createElement("input"),
        summary = document.createElement("p");
      summary.className = "word-pane-muted";
      query.type = "search";
      query.placeholder = "搜索文档…";
      query.setAttribute("aria-label", "查找内容");
      const matchCase = document.createElement("input"),
        wholeWord = document.createElement("input");
      const opts = document.createElement("div");
      opts.className = "word-pane-row";
      for (const [input, text] of [
        [matchCase, "区分大小写"],
        [wholeWord, "全字匹配"],
      ] as const) {
        input.type = "checkbox";
        const label = document.createElement("label");
        label.append(input, text);
        opts.append(label);
      }
      let matches: readonly WordTextMatch[] = [],
        current = -1,
        revision = -1;
      const select = (index: number) => {
        const match = matches[index];
        if (!match) return;
        const ranges = (match.ranges ?? [match]).map((range) => ({
          nodeId: range.nodeId,
          from: range.from,
          to: range.to,
          quote: range.expectedText,
        }));
        focus(
          ranges.length === 1 ? { kind: "text", ...ranges[0]! } : { kind: "text-set", ranges },
          revision,
        );
        current = index;
        summary.textContent = `${matches.length} 处匹配 · 第 ${index + 1} 处`;
      };
      refresh = () => {
        resultControls?.abort();
        resultControls = new AbortController();
        revision = model.snapshot().manifest.revision;
        matches = findWordText(model, query.value, {
          matchCase: matchCase.checked,
          wholeWord: wholeWord.checked,
        });
        current = -1;
        results.replaceChildren();
        summary.textContent =
          matches.length + " 处匹配" + (matches.length > 200 ? "（列表显示前 200 处）" : "");
        for (const [index, match] of matches.slice(0, 200).entries()) {
          const item = button(
            index + 1 + " · " + match.expectedText,
            () => select(index),
            results,
            resultControls.signal,
          );
          item.className = "word-style-item";
        }
      };
      const next = (direction: number) => {
        if (revision !== model.snapshot().manifest.revision) refresh?.();
        if (matches.length)
          select(
            current < 0
              ? direction > 0
                ? 0
                : matches.length - 1
              : (current + direction + matches.length) % matches.length,
          );
      };
      query.addEventListener("input", () => refresh?.(), { signal });
      for (const input of [matchCase, wholeWord])
        input.addEventListener("change", () => refresh?.(), { signal });
      query.addEventListener(
        "keydown",
        (event) => {
          if (event.key !== "Enter") return;
          event.preventDefault();
          try {
            next(event.shiftKey ? -1 : 1);
          } catch (error) {
            onError(error);
          }
        },
        { signal },
      );
      const navRow = document.createElement("div");
      navRow.className = "word-pane-row";
      button("上一处", () => next(-1), navRow);
      button("下一处", () => next(1), navRow).className = "word-btn-primary";
      searchCard.append(query, opts, navRow, summary);
      body.append(searchCard);
      queueMicrotask(() => {
        if (!signal.aborted) query.focus();
      });
    } else
      refresh = () => {
        resultControls?.abort();
        resultControls = new AbortController();
        results.replaceChildren();
        if (action === "outline") {
          const revision = model.snapshot().manifest.revision;
          for (const heading of getWordOutline(model)) {
            const item = button(
              heading.text || "空标题",
              () => focus({ kind: "node", nodeId: heading.nodeId }, revision),
              results,
              resultControls.signal,
            );
            item.className = "word-style-item";
            item.style.paddingLeft = 8 + (heading.level - 1) * 14 + "px";
          }
          if (!results.childElementCount) {
            const empty = document.createElement("p");
            empty.className = "word-pane-muted";
            empty.textContent = "文档中没有标题。";
            results.append(empty);
          }
        } else {
          const stats = getWordTextStatistics(model);
          const table = document.createElement("table");
          table.className = "word-stat-table";
          for (const [k, v] of [
            ["字数", String(stats.words)],
            ["字符数（计空格）", String(stats.characters)],
            ["字符数（不计空格）", String(stats.charactersWithoutSpaces)],
            ["段落", String(stats.paragraphs)],
          ]) {
            const tr = document.createElement("tr");
            tr.innerHTML = `<td>${k}</td><td>${v}</td>`;
            table.append(tr);
          }
          const summary = document.createElement("p");
          summary.className = "word-pane-muted";
          summary.textContent = `字数 ${stats.words} · 字符数 ${stats.characters} · 不计空格 ${stats.charactersWithoutSpaces} · 段落 ${stats.paragraphs}`;
          results.append(table, summary);
        }
      };
    body.append(results);
    parent.insertBefore(panel, before);
    refresh();
  }
  return { open, refresh: (): void => refresh?.(), dispose: close };
}

import { expandWordRevisionGroups } from "@doffi/core";
import type { DocumentEditingSession, DocumentOperation } from "@doffi/core";
import { createWordDockedPane, replaceSameSideDock } from "./word-pane-chrome.js";

export type WordRevisionDisplay = "all" | "simple" | "final" | "original";

export function createWordReviewControls(
  parent: ShadowRoot,
  before: HTMLElement,
  session: DocumentEditingSession,
  onError: (error: unknown) => void,
  display: () => WordRevisionDisplay,
  setDisplay: (mode: WordRevisionDisplay) => void,
  focus: (nodeId: string) => void,
): {
  open(): void;
  dispose(): void;
} {
  const document = before.ownerDocument;
  let panel: HTMLElement | undefined, controller: AbortController | undefined;
  const close = (): void => {
    controller?.abort();
    panel?.remove();
    panel = undefined;
  };
  function open(): void {
    close();
    replaceSameSideDock(parent, "left");
    controller = new AbortController();
    const signal = controller.signal,
      snapshot = session.snapshot(),
      word = snapshot.metadata.word,
      root = new DOMParser()
        .parseFromString(snapshot.html, "text/html")
        .querySelector<HTMLElement>('[data-office-kind="document"]')!.dataset.nodeId!;
    const chrome = createWordDockedPane(document, {
      ariaLabel: "Word 修订审阅",
      title: "审阅窗格",
      dock: "left",
      closeLabel: "关闭修订审阅",
      onClose: close,
      signal,
      width: "300px",
    });
    panel = chrome.panel;
    const body = chrome.body;
    parent.insertBefore(panel, before);
    const apply = (operations: DocumentOperation[]): void => {
      const current = session.snapshot();
      session.commit(
        session.prepare({
          protocolVersion: 1,
          documentId: current.manifest.documentId,
          baseRevision: current.manifest.revision,
          operations,
        }),
      );
      open();
    };
    const button = (
      name: string,
      run: () => void,
      container: HTMLElement = body,
      primary = false,
    ): void => {
      const node = document.createElement("button");
      node.type = "button";
      node.textContent = name;
      if (primary) node.className = "word-btn-primary";
      node.addEventListener(
        "click",
        () => {
          try {
            run();
          } catch (error) {
            onError(error);
          }
        },
        { signal },
      );
      container.append(node);
    };

    const settingsCard = document.createElement("div");
    settingsCard.className = "word-pane-section";
    const tracking = document.createElement("input"),
      author = document.createElement("input"),
      label = document.createElement("label"),
      authorLabel = document.createElement("label");
    tracking.type = "checkbox";
    tracking.checked = word?.trackChanges ?? false;
    tracking.setAttribute("aria-label", "记录修订");
    label.append(tracking, "记录修订");
    authorLabel.textContent = "修订作者";
    author.setAttribute("aria-label", "修订作者");
    author.value = word?.trackingAuthor ?? "";
    authorLabel.append(author);
    const modeLabel = document.createElement("label");
    modeLabel.textContent = "修订显示";
    const mode = document.createElement("select");
    mode.setAttribute("aria-label", "修订显示");
    for (const [value, name] of [
      ["all", "显示所有标记"],
      ["simple", "显示简单标记"],
      ["final", "无标记（最终状态）"],
      ["original", "显示原始状态"],
    ] as const) {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = name;
      mode.append(option);
    }
    mode.value = display();
    mode.addEventListener("change", () => setDisplay(mode.value as WordRevisionDisplay), {
      signal,
    });
    modeLabel.append(mode);
    const settingsActions = document.createElement("div");
    settingsActions.className = "word-pane-row";
    button(
      "应用修订设置",
      () =>
        apply([
          { op: "setWordReview", nodeId: root, trackChanges: tracking.checked, author: author.value },
        ]),
      settingsActions,
      true,
    );
    button("刷新修订列表", open, settingsActions);
    settingsCard.append(label, authorLabel, modeLabel, settingsActions);
    body.append(settingsCard);

    const filterCard = document.createElement("div");
    filterCard.className = "word-pane-section";
    const authorsLabel = document.createElement("label");
    authorsLabel.textContent = "筛选审阅者";
    const authors = document.createElement("select");
    authors.setAttribute("aria-label", "筛选修订作者");
    for (const name of [
      "全部作者",
      ...new Set(word?.revisions.map((revision) => revision.author ?? "")),
    ]) {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name || "未署名";
      authors.append(option);
    }
    authorsLabel.append(authors);
    const grouping = document.createElement("p");
    grouping.className = "word-pane-muted";
    grouping.textContent = "移动的来源与目标、同组替换会一起接受或拒绝。";
    const bulkRow = document.createElement("div");
    bulkRow.className = "word-pane-row";
    const all = (decision: "accept" | "reject"): void =>
      apply([
        {
          op: "reviewWordRevisions",
          nodeId: root,
          decision,
          ...(authors.selectedIndex ? { author: authors.value } : {}),
        },
      ]);
    button("接受全部显示的修订", () => all("accept"), bulkRow, true);
    button("拒绝全部显示的修订", () => all("reject"), bulkRow);
    filterCard.append(authorsLabel, bulkRow, grouping);
    body.append(filterCard);

    const list = document.createElement("div");
    list.style.cssText = "display:flex;flex-direction:column;gap:8px";
    body.append(list);
    const labels = {
      insert: "插入",
      delete: "删除",
      moveFrom: "移动来源",
      moveTo: "移动目标",
      format: "格式更改",
      paragraphInsert: "插入段落标记",
      paragraphDelete: "删除段落标记",
      paragraphMoveFrom: "移动来源段落标记",
      paragraphMoveTo: "移动目标段落标记",
      rowInsert: "插入行",
      rowDelete: "删除行",
      cellInsert: "插入单元格",
      cellDelete: "删除单元格",
      nativeInsert: "插入原生区块",
    };
    const render = (): void => {
      list.replaceChildren();
      const visible = new Set(
        expandWordRevisionGroups(
          word?.revisions ?? [],
          (word?.revisions ?? [])
            .filter((revision) => (revision.author ?? "") === authors.value)
            .map((revision) => revision.revisionId ?? revision.nodeId),
        ),
      );
      let count = 0;
      for (const revision of word?.revisions ?? []) {
        if (authors.selectedIndex && !visible.has(revision)) continue;
        count++;
        const row = document.createElement("div");
        row.className = "word-pane-section";
        const text = document.createElement("p");
        text.style.cssText = "margin:0;font-size:12px;line-height:1.5;word-break:break-word";
        text.textContent =
          labels[revision.kind] +
          " · " +
          (revision.author || "未署名") +
          " · " +
          (session.getNode(revision.nodeId)?.text.slice(0, 80) || "标记");
        const actions = document.createElement("div");
        actions.className = "word-pane-row";
        button("定位修订", () => focus(revision.nodeId), actions);
        button(
          "接受此修订",
          () =>
            apply([
              {
                op: "reviewWordRevision",
                nodeId: revision.nodeId,
                revisionId: revision.revisionId ?? revision.nodeId,
                decision: "accept",
              },
            ]),
          actions,
        );
        button(
          "拒绝此修订",
          () =>
            apply([
              {
                op: "reviewWordRevision",
                nodeId: revision.nodeId,
                revisionId: revision.revisionId ?? revision.nodeId,
                decision: "reject",
              },
            ]),
          actions,
        );
        row.append(text, actions);
        list.append(row);
      }
      if (!count) {
        const empty = document.createElement("p");
        empty.className = "word-pane-muted";
        empty.textContent = "当前文档暂无修订记录。";
        list.append(empty);
      }
    };
    authors.addEventListener("change", render, { signal });
    render();
  }
  return { open, dispose: close };
}

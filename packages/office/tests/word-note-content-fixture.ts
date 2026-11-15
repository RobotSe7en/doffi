import { readFileSync } from "node:fs";
import {
  createWordDocument,
  createDocumentSession,
  openDocument,
  createAssetDescriptor,
  SourceStyleResolver,
  DEFAULT_DOCUMENT_READ_LIMITS,
} from "@doffi/core";
import type { DocumentOperation, DocumentEditingSession, WordContentSemantics } from "@doffi/core";
import { exportDocx, importDocx } from "@doffi/office";

export function wordNoteContentFixture(native = false): {
  session: DocumentEditingSession;
  footId: string;
  endId: string;
  paragraphId: string;
} {
  const documentId = "note-copy-source";
  let session = createDocumentSession(
    openDocument(createWordDocument({ documentId }), { kind: "document" }),
  );
  const apply = (
    operations: DocumentOperation[],
    options?: Parameters<typeof session.prepare>[1],
  ) =>
    session.commit(
      session.prepare(
        {
          protocolVersion: 1,
          documentId,
          baseRevision: session.snapshot().manifest.revision,
          operations,
        },
        options,
      ),
    );
  apply([
    {
      op: "replaceText",
      nodeId: "paragraph-1",
      from: 0,
      to: 0,
      expectedText: "",
      text: "LEFT😀MID禰\u{e0100}RIGHT",
    },
    {
      op: "insertFragment",
      nodeId: "section-root-1",
      position: "end",
      html: "<p>PRIVATE-BODY</p>",
    },
  ]);
  const privateP = session.getNode("section-root-1")!.childNodeIds.at(-1)!;
  apply([
    {
      op: "insertWordNote",
      nodeId: "paragraph-1",
      offset: 4,
      kind: "footnote",
      text: "FOOT-BODY😀",
      customMark: "*",
    },
    {
      op: "insertWordNote",
      nodeId: "paragraph-1",
      offset: 13,
      kind: "endnote",
      text: "END-BODY禰\u{e0100}",
    },
    {
      op: "insertWordNote",
      nodeId: "paragraph-1",
      offset: 0,
      kind: "footnote",
      text: "PRIVATE-SAME-PARAGRAPH-NOTE",
    },
    {
      op: "insertWordNote",
      nodeId: privateP,
      offset: 0,
      kind: "endnote",
      text: "PRIVATE-OTHER-PARAGRAPH-NOTE",
    },
    {
      op: "setWordNoteSeparator",
      nodeId: "document-root",
      kind: "footnote",
      type: "separator",
      text: "PRIVATE-SEPARATOR-NOTE",
    },
  ]);
  const find = (text: string) =>
    session
      .snapshot()
      .metadata.word!.notes!.find((note) =>
        session
          .getNode(
            session.snapshot().metadata.word!.stories.find((s) => s.storyId === note.storyId)!
              .rootNodeId,
          )!
          .text.includes(text),
      )!;
  const foot = find("FOOT-BODY"),
    end = find("END-BODY"),
    footRoot = session
      .snapshot()
      .metadata.word!.stories.find((s) => s.storyId === foot.storyId)!.rootNodeId,
    endRoot = session
      .snapshot()
      .metadata.word!.stories.find((s) => s.storyId === end.storyId)!.rootNodeId;
  const footP = session.getNode(footRoot)!.childNodeIds[0]!,
    endP = session.getNode(endRoot)!.childNodeIds[0]!;
  const picture = new Uint8Array(
      readFileSync(new URL("../../../examples/xlsx/reference/13-date-1904-1.png", import.meta.url)),
    ),
    path = "assets/note-owned.png",
    geometry = { x: 0, y: 0, width: 24, height: 16, rotation: 0, flipH: false, flipV: false };
  apply(
    [
      { op: "insertWordField", nodeId: footP, offset: 0, instruction: " PAGE ", cachedText: "7" },
      {
        op: "insertWordDrawing",
        nodeId: footP,
        offset: 0,
        geometry,
        layout: { mode: "inline" },
        object: { kind: "image", assetPath: path, alt: "Owned note image" },
      },
      {
        op: "insertWordDrawing",
        nodeId: footP,
        offset: 0,
        geometry: { ...geometry, width: 90, height: 36 },
        layout: { mode: "inline" },
        object: { kind: "textBox", text: "NOTE-OWNED-TEXTBOX" },
      },
      { op: "insertFragment", nodeId: footRoot, position: "end", html: "<p>LIST-BODY</p>" },
      {
        op: "insertWordField",
        nodeId: endP,
        offset: 0,
        instruction: " NUMPAGES ",
        cachedText: "3",
      },
      { op: "insertWordTable", nodeId: endP, position: "after", rows: 1, columns: 1 },
    ],
    { assets: [{ descriptor: createAssetDescriptor(path, "image/png", picture), bytes: picture }] },
  );
  const listP = session.getNode(footRoot)!.childNodeIds.at(-1)!,
    table = session.snapshot().metadata.word!.tables[0]!,
    tableP = [...session.snapshot().html.matchAll(/data-node-id="([^"]+)"/g)]
      .map((m) => session.getNode(m[1]!)!)
      .find(
        (n) =>
          n.tagName === "p" &&
          n.text === "" &&
          (() => {
            for (let id: string | null = n.nodeId; id; id = session.getNode(id)!.parentNodeId)
              if (id === table.nodeId) return true;
            return false;
          })(),
      )!.nodeId;
  apply([
    {
      op: "setWordNumbering",
      nodeId: "document-root",
      numberingId: "note-list",
      levels: [
        {
          level: 0,
          start: 4,
          format: "upperRoman",
          text: "%1)",
          suffix: "tab",
          cssClass: "note-numbering-marker",
        },
      ],
    },
    { op: "setWordList", nodeId: listP, numberingId: "note-list", level: 0 },
    {
      op: "setWordTabs",
      nodeId: footP,
      defaultInterval: 36,
      stops: [{ position: 72, alignment: "right", leader: "dot" }],
    },
    { op: "replaceText", nodeId: tableP, from: 0, to: 0, expectedText: "", text: "NOTE-CELL" },
  ]);
  const linkFrom = session.getNode(endP)!.text.indexOf("END-BODY");
  apply([
    {
      op: "insertWordLink",
      nodeId: endP,
      from: linkFrom,
      to: linkFrom + 8,
      expectedText: "END-BODY",
      target: "https://example.com/note",
    },
  ]);
  if (native)
    session = createDocumentSession(
      openDocument(
        importDocx(exportDocx(session).bytes, { documentId, unsupportedContent: "reject" }),
        { kind: "document" },
      ),
    );
  const refs = session.snapshot().metadata.word!.noteReferences!,
    word = session.snapshot().metadata.word!;
  const reference = (text: string) =>
    refs.find((ref) =>
      session
        .getNode(
          word.stories.find(
            (s) => s.storyId === word.notes!.find((n) => n.noteId === ref.noteId)!.storyId,
          )!.rootNodeId,
        )!
        .text.includes(text),
    )!.nodeId;
  const footId = reference("FOOT-BODY"),
    endId = reference("END-BODY");
  let paragraph = session.getNode(footId)!;
  while (paragraph.tagName !== "p") paragraph = session.getNode(paragraph.parentNodeId!)!;
  return { session, footId, endId, paragraphId: paragraph.nodeId };
}

export function wordNoteContentGraph(
  source: DocumentEditingSession,
  referenceIds: readonly string[],
): { html: string; semantics: WordContentSemantics } {
  const word = source.snapshot().metadata.word!,
    ids = new Set<string>(),
    storyIds = new Set<string>(),
    noteIds = new Set<string>(),
    pending = [...referenceIds];
  while (pending.length) {
    const id = pending.pop()!;
    if (ids.has(id)) continue;
    ids.add(id);
    pending.push(...source.getNode(id)!.childNodeIds);
    const ref = word.noteReferences?.find((r) => r.nodeId === id);
    if (ref) {
      const note = word.notes!.find((n) => n.noteId === ref.noteId)!;
      noteIds.add(note.noteId);
      if (!storyIds.has(note.storyId)) {
        storyIds.add(note.storyId);
        pending.push(word.stories.find((s) => s.storyId === note.storyId)!.rootNodeId);
      }
    }
    const drawing = word.drawings?.find((d) => d.nodeId === id);
    if (
      drawing &&
      "textStoryId" in drawing &&
      drawing.textStoryId &&
      !storyIds.has(drawing.textStoryId)
    ) {
      storyIds.add(drawing.textStoryId);
      pending.push(word.stories.find((s) => s.storyId === drawing.textStoryId)!.rootNodeId);
    }
  }
  const numbers = new Set(
      [...ids].flatMap((id) =>
        source.getNode(id)!.attributes["data-word-numbering"]
          ? [source.getNode(id)!.attributes["data-word-numbering"]!]
          : [],
      ),
    ),
    numbering = word.numbering.filter((n) => numbers.has(n.numberingId)),
    css = new SourceStyleResolver(
      source.snapshot().styles.map((s) => s.css),
      DEFAULT_DOCUMENT_READ_LIMITS,
    );
  return {
    html: referenceIds.map((id) => source.getNode(id)!.html).join(""),
    semantics: JSON.parse(
      JSON.stringify({
        notes: word.notes?.filter((n) => noteIds.has(n.noteId)),
        noteReferences: word.noteReferences?.filter((r) => ids.has(r.nodeId)),
        stories: word.stories
          .filter((s) => storyIds.has(s.storyId))
          .map((s) => ({ ...s, html: source.getNode(s.rootNodeId)!.html })),
        fields: word.fields.filter((f) => ids.has(f.nodeId)),
        hyperlinks: word.hyperlinks.filter((l) => ids.has(l.nodeId)),
        drawings: word.drawings?.filter((d) => ids.has(d.nodeId)),
        tables: word.tables.filter((t) => ids.has(t.nodeId)),
        tabs: word.tabs.filter((t) => ids.has(t.paragraphNodeId)),
        numbering,
        extraClasses: Object.fromEntries(
          numbering.flatMap((n) =>
            n.levels.map((l) => [l.cssClass, css.specifiedClass(l.cssClass)]),
          ),
        ),
      }),
    ) as WordContentSemantics,
  };
}

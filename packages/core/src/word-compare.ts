import { createWordDocument } from "./word.js";
import { openDocument } from "./document-model.js";
import { encode } from "./document/package.js";
import type { ContentDocumentModel, DocumentPackage } from "./runtime.js";
import type { WordMetadata } from "./contracts/word.js";
import { wordQueryParagraphs } from "./word-query-view.js";

export interface CompareWordDocumentsOptions {
  readonly documentId?: string;
  readonly title?: string;
  readonly author?: string;
  readonly date?: string;
}

interface DiffOp<T> {
  readonly kind: "equal" | "insert" | "delete";
  readonly value: T;
}

function lcsDiff<T>(a: readonly T[], b: readonly T[], equals: (x: T, y: T) => boolean): DiffOp<T>[] {
  const m = a.length,
    n = b.length;
  if (m === 0) return b.map((value) => ({ kind: "insert", value }));
  if (n === 0) return a.map((value) => ({ kind: "delete", value }));
  // Guard quadratic DP for very large inputs with greedy prefix/suffix trim.
  let start = 0;
  while (start < m && start < n && equals(a[start]!, b[start]!)) start++;
  let endA = m - 1,
    endB = n - 1;
  while (endA >= start && endB >= start && equals(a[endA]!, b[endB]!)) {
    endA--;
    endB--;
  }
  const midA = a.slice(start, endA + 1),
    midB = b.slice(start, endB + 1);
  const ops: DiffOp<T>[] = a.slice(0, start).map((value) => ({ kind: "equal", value }));
  if (midA.length * midB.length > 400_000) {
    for (const value of midA) ops.push({ kind: "delete", value });
    for (const value of midB) ops.push({ kind: "insert", value });
  } else {
    const rows = midA.length + 1,
      cols = midB.length + 1;
    const dp = new Int32Array(rows * cols);
    for (let i = midA.length - 1; i >= 0; i--) {
      for (let j = midB.length - 1; j >= 0; j--) {
        dp[i * cols + j] = equals(midA[i]!, midB[j]!)
          ? dp[(i + 1) * cols + (j + 1)]! + 1
          : Math.max(dp[(i + 1) * cols + j]!, dp[i * cols + (j + 1)]!);
      }
    }
    let i = 0,
      j = 0;
    while (i < midA.length && j < midB.length) {
      if (equals(midA[i]!, midB[j]!)) {
        ops.push({ kind: "equal", value: midA[i]! });
        i++;
        j++;
      } else if (dp[(i + 1) * cols + j]! >= dp[i * cols + (j + 1)]!) {
        ops.push({ kind: "delete", value: midA[i]! });
        i++;
      } else {
        ops.push({ kind: "insert", value: midB[j]! });
        j++;
      }
    }
    while (i < midA.length) ops.push({ kind: "delete", value: midA[i++]! });
    while (j < midB.length) ops.push({ kind: "insert", value: midB[j++]! });
  }
  for (let k = endA + 1; k < m; k++) ops.push({ kind: "equal", value: a[k]! });
  return ops;
}

function tokenizeText(text: string): string[] {
  return text.match(/\p{Script=Han}|[A-Za-z0-9_]+|\s+|[^\s]/gu) ?? [];
}

function escapeHtml(text: string): string {
  return text
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

/** Compare two Word documents and produce a validated package containing native
 * tracked insertion/deletion revisions (exportable to DOCX for Microsoft Word). */
export function compareWordDocuments(
  original: ContentDocumentModel,
  revised: ContentDocumentModel,
  options: CompareWordDocumentsOptions = {},
): DocumentPackage {
  const author = options.author ?? "文档比较";
  const date = (options.date ?? new Date().toISOString()).replace(/\.\d+(?=(?:Z|[+-]\d\d:\d\d)$)/, "");
  const left = wordQueryParagraphs(original, { revisionView: "final" }).map((p) => p.text);
  const right = wordQueryParagraphs(revised, { revisionView: "final" }).map((p) => p.text);
  const paraDiff = lcsDiff(left, right, (a, b) => a === b);

  let serial = 100;
  const nextId = (prefix: string) => `${prefix}${serial++}`;
  const revisions: WordMetadata["revisions"] = [];
  const paragraphsHtml: string[] = [];

  const emitChunk = (kind: "insert" | "delete", text: string): string => {
    if (!text) return "";
    const id = nextId("cmp");
    const tag = kind === "insert" ? "ins" : "del";
    revisions.push({
      nodeId: id,
      sourceId: String(serial),
      revisionId: id,
      kind,
      author,
      date,
    });
    return `<${tag} data-node-id="${id}" data-word-revision="${kind}">${escapeHtml(text)}</${tag}>`;
  };

  for (let i = 0; i < paraDiff.length; ) {
    const entry = paraDiff[i]!;
    if (entry.kind === "equal") {
      paragraphsHtml.push(`<p data-node-id="${nextId("p")}">${escapeHtml(entry.value)}</p>`);
      i++;
      continue;
    }
    // Pair contiguous deleted and inserted paragraphs for word-level inline diff.
    const dels: string[] = [];
    const ins: string[] = [];
    while (i < paraDiff.length && paraDiff[i]!.kind !== "equal") {
      if (paraDiff[i]!.kind === "delete") dels.push(paraDiff[i]!.value);
      else ins.push(paraDiff[i]!.value);
      i++;
    }
    const paired = Math.min(dels.length, ins.length);
    for (let p = 0; p < paired; p++) {
      const tokens = lcsDiff(tokenizeText(dels[p]!), tokenizeText(ins[p]!), (a, b) => a === b);
      let body = "";
      let k = 0;
      while (k < tokens.length) {
        const kind = tokens[k]!.kind;
        let combined = "";
        while (k < tokens.length && tokens[k]!.kind === kind) {
          combined += tokens[k]!.value;
          k++;
        }
        body += kind === "equal" ? escapeHtml(combined) : emitChunk(kind, combined);
      }
      paragraphsHtml.push(`<p data-node-id="${nextId("p")}">${body}</p>`);
    }
    for (let p = paired; p < dels.length; p++) {
      paragraphsHtml.push(`<p data-node-id="${nextId("p")}">${emitChunk("delete", dels[p]!)}</p>`);
    }
    for (let p = paired; p < ins.length; p++) {
      paragraphsHtml.push(`<p data-node-id="${nextId("p")}">${emitChunk("insert", ins[p]!)}</p>`);
    }
  }

  const base = createWordDocument({
    documentId: options.documentId ?? "word-compare-result",
    title: options.title ?? "文档比较结果",
  });
  const baseModel = openDocument(base, { kind: "document" });
  const snap = baseModel.snapshot();
  const sectionId = snap.metadata.sections?.[0]?.rootNodeId;
  if (!sectionId) {
    throw new Error("Unexpected missing Word section in newly created comparison package");
  }
  const html = snap.html.replace(
    new RegExp(`(<section[^>]*data-node-id="${sectionId}"[^>]*>)[\\s\\S]*?(</section>)`),
    `$1${paragraphsHtml.length ? paragraphsHtml.join("") : `<p data-node-id="${nextId("p")}"></p>`}$2`,
  );
  const metadata = JSON.parse(JSON.stringify(snap.metadata)) as typeof snap.metadata & {
    word: WordMetadata;
  };
  metadata.word.revisions = revisions;
  metadata.word.trackChanges = true;
  metadata.word.trackingAuthor = author;

  const files = new Map(base);
  files.set(snap.manifest.entrypoint, encode(html));
  files.set("metadata.json", encode(JSON.stringify(metadata, null, 2)));
  // Validate complete package through the canonical model boundary before returning.
  return openDocument(files, { kind: "document" }).serialize();
}

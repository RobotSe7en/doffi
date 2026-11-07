import { ident, parse, walk } from "css-tree";
import type { WorkbookManifest } from "../contracts/document.js";
import type { DocumentReadLimits, WorkbookReadLimits, WorkbookModel, WorkbookCell, WorkbookWindow, WorkbookSnapshot } from "../runtime.js";
import { parseContract } from "../schema-registry.js";
import { normalizeCss } from "../document/css.js";
import { assetDigest, validateAssets } from "../document/assets.js";
import { readJson, readText, encode } from "../document/package.js";
import { checkLimit, freeze, invalid } from "../document/validation.js";
import { parseHtml } from "../document/html.js";
import { validateWorkbookMetadata } from "./metadata.js";
import { validateWorkbookContent } from "./content.js";
import { createFormulaTemplatePool } from "./formula.js";
import { CHUNK_ROWS, CHUNK_COLUMNS, MAX_SHEET_COLUMNS, MAX_SHEET_ROWS, indexWorkbookChunk } from "./chunk-index.js";
import type { ChunkIndex } from "./chunk-index.js";
import { WorkbookReferences } from "./reference.js";
import { WorkbookCellReadCache } from "./cell-read-cache.js";
import { WorkbookChunkTextCache } from "./chunk-text-cache.js";
import { registerWorkbookSource, workbookSource } from "./source.js";

export const DEFAULT_WORKBOOK_READ_LIMITS: WorkbookReadLimits = Object.freeze({
  maxCells: 250_000, maxFormulas: 25_000, maxWindowCells: 5000, maxFormulaTemplates: 25_000, maxFormulaKeyCodeUnits: 4 * 1024 * 1024,
});
function readLimits(input: Partial<WorkbookReadLimits> = {}): WorkbookReadLimits {
  for (const [key, value] of Object.entries(input)) if (!Object.hasOwn(DEFAULT_WORKBOOK_READ_LIMITS, key) || !Number.isSafeInteger(value) || value <= 0) invalid("Workbook limits must be positive safe integers with supported names.");
  return Object.freeze({ ...DEFAULT_WORKBOOK_READ_LIMITS, ...input });
}

/** Receives an isolated package copy. Indexed chunks take sole ownership of their
 * stored bytes; the remaining map retains only non-chunk package files. */
export function openWorkbookModel(files: Map<string, Uint8Array>, manifest: WorkbookManifest, documentLimits: DocumentReadLimits, inputLimits?: Partial<WorkbookReadLimits>, previous?: WorkbookModel): WorkbookModel {
  const steps = openWorkbookModelSteps(files, manifest, documentLimits, inputLimits, previous);
  for (;;) { const next = steps.next(); if (next.done) return next.value ?? invalid("Workbook model loading did not complete."); }
}

/** Chunk checkpoints keep candidate validation private until the final model is
 * registered. The synchronous reader consumes these exact same steps. */
export function* openWorkbookModelSteps(files: Map<string, Uint8Array>, manifest: WorkbookManifest, documentLimits: DocumentReadLimits, inputLimits?: Partial<WorkbookReadLimits>, previous?: WorkbookModel): Generator<{ completed: number; total: number }, WorkbookModel | undefined> {
  const limits = readLimits(inputLimits), metadata = parseContract("document-metadata", readJson(files, manifest.metadata));
  if (!("workbook" in metadata)) invalid("Manifest kind and workbook metadata must match.");
  const annotations = parseContract("annotations", readJson(files, manifest.annotations));
  validateAssets(manifest.assets, files);
  const cssBudget = { nodes: 0 }, classes = new Set<string>();
  const styles = manifest.styles.map(path => ({ path, css: normalizeCss(readText(files, path), "stylesheet", documentLimits, cssBudget) }));
  for (const style of styles) walk(parse(style.css), node => { if (node.type === "ClassSelector") classes.add(ident.decode(node.name)); });
  const { formats, cellStyleClasses, items } = validateWorkbookMetadata(metadata.workbook, classes, documentLimits.maxNodes);
  const bindingKey = JSON.stringify([[...classes].sort(), [...cellStyleClasses].sort(), [...formats].sort()]);
  const prior = previous ? workbookSource(previous) : undefined;
  const priorDescriptors = new Map(previous?.snapshot().manifest.chunks.map(chunk => [chunk.path, chunk]) ?? []);
  const canReuse = prior?.bindingKey === bindingKey && prior.limits.maxNodes <= documentLimits.maxNodes
    && prior.limits.maxDepth <= documentLimits.maxDepth;
  const content = manifest.content === undefined ? undefined : parseHtml(readText(files, manifest.content), { kind: "workbook-content", entrypoint: manifest.content, styles: manifest.styles, assets: manifest.assets }, files, documentLimits, cssBudget);
  validateWorkbookContent(metadata.workbook, manifest, annotations, content, classes, documentLimits.maxNodes, items);
  const sheets = new Map(metadata.workbook.sheets.map(sheet => [sheet.sheetId, new Map<number, ChunkIndex>()]));
  const chunks: { path: string; index: ChunkIndex }[] = [], formulas = createFormulaTemplatePool(limits.maxFormulaTemplates, limits.maxFormulaKeyCodeUnits);
  const textCache = new WorkbookChunkTextCache();
  let cellCount = 0, formulaCount = 0;
  for (const chunk of manifest.chunks) {
    const sheet = sheets.get(chunk.sheetId), key = chunk.row / CHUNK_ROWS * (MAX_SHEET_COLUMNS / CHUNK_COLUMNS) + chunk.column / CHUNK_COLUMNS;
    if (!sheet || sheet.has(key)) invalid("Workbook chunks must have unique positions on existing worksheets.");
    const bytes = files.get(chunk.path);
    if (!bytes) invalid("Workbook chunk bytes do not match their digest.");
    checkLimit(bytes.byteLength, documentLimits.maxFileBytes);
    // Only private, already-owned immutable buffers carry a digest proof.
    // New bytes and changed declarations still hash before index reuse; style
    // bindings, positions and every cumulative quota remain independent.
    const unchanged = prior?.files.get(chunk.path) === bytes && priorDescriptors.get(chunk.path)?.sha256 === chunk.sha256;
    if (!unchanged && assetDigest(bytes) !== chunk.sha256) invalid("Workbook chunk bytes do not match their digest.");
    const known = canReuse && unchanged ? prior.chunks.get(chunk.path) : undefined;
    const samePosition = known && known.position.sheetId === chunk.sheetId && known.position.row === chunk.row && known.position.column === chunk.column;
    const index = samePosition ? known : indexWorkbookChunk(bytes, chunk, documentLimits.maxFileBytes, { takeOwnership: true, formulas, textCache, cssClasses: classes, cellStyleClasses, numberFormats: formats, maxNodes: documentLimits.maxNodes, maxDepth: documentLimits.maxDepth });
    if (samePosition) for (const expression of index.formulaTemplates) formulas.include(expression);
    cellCount += index.cellCount; formulaCount += index.formulaCount;
    checkLimit(cellCount, limits.maxCells); checkLimit(formulaCount, limits.maxFormulas);
    sheet.set(key, index); chunks.push({ path: chunk.path, index });
    yield { completed: chunks.length, total: manifest.chunks.length };
  }
  const decodedCells = new WorkbookCellReadCache(Math.min(5000, limits.maxWindowCells));
  for (const sheet of metadata.workbook.sheets) for (const table of sheet.tables ?? []) if (table.header) {
    for (const [index, column] of table.columns.entries()) {
      const cell = readWindow({ sheetId: sheet.sheetId, row: table.range.from.row, column: table.range.from.column + index, rows: 1, columns: 1 })[0];
      if (cell?.formula || cell?.value.type !== "text" || cell.value.value !== column.name) invalid("A table column name must match its constant text header cell.");
    }
  }
  const snapshot: WorkbookSnapshot = freeze({ manifest, metadata, annotations, styles, ...(content ? { content: { path: manifest.content!, html: content.html } } : {}) });
  const assetPaths = new Set(manifest.assets.map(asset => asset.path));
  const references = new WorkbookReferences(snapshot.metadata.workbook);
  if (content) files.set(manifest.content!, encode(content.html));
  files.set("manifest.json", encode(JSON.stringify(manifest)));
  files.set(manifest.metadata, encode(JSON.stringify(metadata))); files.set(manifest.annotations, encode(JSON.stringify(annotations)));
  for (const style of styles) files.set(style.path, encode(style.css));
  let total = 0;
  for (const bytes of files.values()) { checkLimit(bytes.byteLength, documentLimits.maxFileBytes); total += bytes.byteLength; }
  checkLimit(total, documentLimits.maxPackageBytes);
  function readWindow(window: WorkbookWindow): readonly WorkbookCell[] {
    if (Object.keys(window).some(key => !["sheetId", "row", "column", "rows", "columns"].includes(key))
      || ![window.row, window.column, window.rows, window.columns].every(Number.isSafeInteger)
      || window.row < 0 || window.column < 0 || window.rows <= 0 || window.columns <= 0
      || window.row + window.rows > MAX_SHEET_ROWS || window.column + window.columns > MAX_SHEET_COLUMNS) invalid("Invalid workbook window.");
    checkLimit(window.rows * window.columns, limits.maxWindowCells);
    const sheet = sheets.get(window.sheetId); if (!sheet) invalid("A workbook query references an unknown worksheet.");
    const cells: WorkbookCell[] = [];
    for (let row = Math.floor(window.row / CHUNK_ROWS); row <= Math.floor((window.row + window.rows - 1) / CHUNK_ROWS); row++) {
      for (let column = Math.floor(window.column / CHUNK_COLUMNS); column <= Math.floor((window.column + window.columns - 1) / CHUNK_COLUMNS); column++) {
        const chunk = sheet.get(row * (MAX_SHEET_COLUMNS / CHUNK_COLUMNS) + column);
        for (const source of chunk?.readWindow(window) ?? []) cells.push(decodedCells.read(window.sheetId, source));
      }
    }
    cells.sort((a, b) => a.row - b.row || a.column - b.column);
    return Object.freeze(cells);
  }
  const model: WorkbookModel = Object.freeze({
    kind: "workbook", snapshot: () => snapshot, getNode: (nodeId: string) => content?.getNode(nodeId),
    readWindow,
    getAsset(path: string, maxBytes = documentLimits.maxFileBytes): Uint8Array | undefined {
      if (!Number.isSafeInteger(maxBytes) || maxBytes < 1) invalid("Workbook asset byte limits must be positive safe integers.");
      if (!assetPaths.has(path)) return undefined;
      const bytes = files.get(path)!; checkLimit(bytes.byteLength, maxBytes); return new Uint8Array(bytes);
    },
    resolveReference: references.resolve.bind(references),
    getCell: (sheetId: string, row: number, column: number) => readWindow({ sheetId, row, column, rows: 1, columns: 1 })[0],
    serialize: () => new Map(Array.from(files, ([path, bytes]) => [path, new Uint8Array(bytes)] as const)),
  });
  registerWorkbookSource(model, { files, chunks: new Map(chunks.map(chunk => [chunk.path, chunk.index])), classes, bindingKey,
    limits: documentLimits, workbookLimits: limits, ...(content ? { content } : {}) });
  return model;
}

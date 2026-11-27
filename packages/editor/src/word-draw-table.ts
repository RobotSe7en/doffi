import { DocumentError, type DocumentEditingSession, type DocumentOperation } from "@doffi/core";

export type WordTableToolMode = "off" | "draw" | "erase";

function locateCellGrid(
  session: DocumentEditingSession,
  tableId: string,
  cellId: string,
): { row: number; column: number; rows: number; columns: number } | undefined {
  const table = session.getNode(tableId);
  if (!table) return undefined;
  const rowIds = table.childNodeIds.flatMap((id) => {
    const node = session.getNode(id)!;
    return node.tagName === "tr" ? [id] : node.childNodeIds;
  });
  const occupied: string[][] = Array.from({ length: rowIds.length }, () => []);
  for (let r = 0; r < rowIds.length; r++) {
    for (const id of session.getNode(rowIds[r]!)!.childNodeIds) {
      const cell = session.getNode(id)!;
      let c = 0;
      while (occupied[r]![c]) c++;
      const rows = Number(cell.attributes.rowspan ?? 1) || rowIds.length - r;
      const columns = Number(cell.attributes.colspan ?? 1);
      for (let y = r; y < r + rows; y++)
        for (let x = c; x < c + columns; x++) occupied[y]![x] = id;
      if (id === cellId) return { row: r, column: c, rows, columns };
    }
  }
  return undefined;
}

/** Interactive Word Draw Table & Border Eraser canvas modes. All geometry
 * overlays are view-only; every table edit commits through one core transaction. */
export function createWordDrawTableController(
  session: DocumentEditingSession,
  callbacks: {
    onModeChange(mode: WordTableToolMode): void;
    onError(error: unknown): void;
  },
): {
  mode(): WordTableToolMode;
  setMode(next: WordTableToolMode): void;
  bind(frame: Document, root: HTMLElement, signal: AbortSignal): void;
} {
  let activeMode: WordTableToolMode = "off";
  let currentFrame: Document | undefined;

  const syncCursor = () => {
    if (!currentFrame) return;
    if (activeMode === "off") delete currentFrame.documentElement.dataset.wordTableTool;
    else currentFrame.documentElement.dataset.wordTableTool = activeMode;
  };

  const commit = (operations: DocumentOperation[]) => {
    const snapshot = session.snapshot();
    session.commit(
      session.prepare({
        protocolVersion: 1,
        documentId: snapshot.manifest.documentId,
        baseRevision: snapshot.manifest.revision,
        operations,
      }),
    );
  };

  function setMode(next: WordTableToolMode): void {
    if (activeMode === next) return;
    activeMode = next;
    syncCursor();
    callbacks.onModeChange(activeMode);
  }

  function bind(frame: Document, root: HTMLElement, signal: AbortSignal): void {
    currentFrame = frame;
    if (!frame.head.querySelector("style[data-sdk-table-tool]")) {
      const style = frame.createElement("style");
      style.dataset.sdkTableTool = "true";
      style.textContent =
        `html[data-word-table-tool="draw"] *{cursor:crosshair!important}` +
        `html[data-word-table-tool="erase"] *{cursor:not-allowed!important}` +
        `.sdk-draw-table-box{position:fixed;pointer-events:none;z-index:9999;border:1.5px dashed #185abd;background:rgba(24,90,189,0.08);box-sizing:border-box}` +
        `.sdk-draw-table-box[data-mode="erase"]{border-color:#c00000;background:rgba(192,0,0,0.08)}`;
      frame.head.append(style);
    }
    syncCursor();

    let drag:
      | {
          startX: number;
          startY: number;
          startCell?: HTMLElement | null;
          startTable?: HTMLElement | null;
          startParagraph?: HTMLElement | null;
          box: HTMLDivElement;
        }
      | undefined;

    const cleanupDrag = () => {
      drag?.box.remove();
      drag = undefined;
    };
    signal.addEventListener("abort", cleanupDrag);

    frame.addEventListener(
      "keydown",
      (event) => {
        if (event.key === "Escape" && activeMode !== "off") {
          event.preventDefault();
          cleanupDrag();
          setMode("off");
        }
      },
      { signal },
    );

    root.addEventListener(
      "pointerdown",
      (event) => {
        if (activeMode === "off" || event.button !== 0) return;
        const target = event.target instanceof Element ? event.target : null;
        if (!target) return;
        event.preventDefault();
        const startCell = target.closest<HTMLElement>("td[data-node-id],th[data-node-id]");
        const startTable = startCell?.closest<HTMLElement>("table[data-node-id]") ?? null;
        const startParagraph =
          target.closest<HTMLElement>("p[data-node-id],h1[data-node-id],h2[data-node-id],h3[data-node-id],h4[data-node-id],h5[data-node-id],h6[data-node-id]") ??
          root.querySelector<HTMLElement>("p[data-node-id],h1[data-node-id],h2[data-node-id]");
        const box = frame.createElement("div");
        box.className = "sdk-draw-table-box";
        box.dataset.mode = activeMode;
        box.style.left = event.clientX + "px";
        box.style.top = event.clientY + "px";
        box.style.width = "0px";
        box.style.height = "0px";
        frame.body.append(box);
        drag = {
          startX: event.clientX,
          startY: event.clientY,
          startCell,
          startTable,
          startParagraph,
          box,
        };
      },
      { signal, capture: true },
    );

    frame.addEventListener(
      "pointermove",
      (event) => {
        if (!drag) return;
        const left = Math.min(drag.startX, event.clientX);
        const top = Math.min(drag.startY, event.clientY);
        const width = Math.abs(event.clientX - drag.startX);
        const height = Math.abs(event.clientY - drag.startY);
        drag.box.style.left = left + "px";
        drag.box.style.top = top + "px";
        drag.box.style.width = width + "px";
        drag.box.style.height = height + "px";
      },
      { signal },
    );

    frame.addEventListener(
      "pointerup",
      (event) => {
        if (!drag) return;
        const state = drag;
        cleanupDrag();
        const dx = Math.abs(event.clientX - state.startX);
        const dy = Math.abs(event.clientY - state.startY);
        try {
          if (activeMode === "draw") {
            if (state.startCell && state.startTable) {
              const tableId = state.startTable.dataset.nodeId!;
              const cellId = state.startCell.dataset.nodeId!;
              const grid = locateCellGrid(session, tableId, cellId);
              if (!grid) return;
              if (dx > 16 && dx >= dy) {
                // Horizontal draw stroke inside a cell splits it vertically into 2 rows (Word Draw Table line)
                // Wait: in Word, drawing a horizontal line across a cell divides it top/bottom (2 rows, 1 col);
                // drawing a vertical line divides it left/right (1 row, 2 cols)!
                commit([
                  {
                    op: "editWordTable",
                    nodeId: tableId,
                    change: { kind: "splitCell", row: grid.row, column: grid.column, rows: 2, columns: 1 },
                  },
                ]);
              } else if (dy > 16 && dy > dx) {
                commit([
                  {
                    op: "editWordTable",
                    nodeId: tableId,
                    change: { kind: "splitCell", row: grid.row, column: grid.column, rows: 1, columns: 2 },
                  },
                ]);
              } else {
                commit([
                  {
                    op: "formatWordCell",
                    nodeId: cellId,
                    properties: {
                      "border-top": "1pt solid #000000",
                      "border-right": "1pt solid #000000",
                      "border-bottom": "1pt solid #000000",
                      "border-left": "1pt solid #000000",
                    },
                  },
                ]);
              }
            } else if (state.startParagraph?.dataset.nodeId) {
              const rows = Math.max(1, Math.min(8, Math.round(dy / 36) || 2));
              const columns = Math.max(1, Math.min(8, Math.round(dx / 90) || 2));
              commit([
                {
                  op: "insertWordTable",
                  nodeId: state.startParagraph.dataset.nodeId,
                  position: "after",
                  rows,
                  columns,
                },
              ]);
            } else {
              throw new DocumentError("InvalidOperation", "请在文档正文或表格区域拖动以绘制表格。");
            }
          } else if (activeMode === "erase") {
            const endEl = frame.elementFromPoint(event.clientX, event.clientY);
            const endCell = endEl?.closest<HTMLElement>("td[data-node-id],th[data-node-id]");
            const endTable = endCell?.closest<HTMLElement>("table[data-node-id]");
            if (
              state.startCell &&
              state.startTable &&
              endCell &&
              endTable === state.startTable &&
              endCell !== state.startCell
            ) {
              const tableId = state.startTable.dataset.nodeId!;
              const a = locateCellGrid(session, tableId, state.startCell.dataset.nodeId!);
              const b = locateCellGrid(session, tableId, endCell.dataset.nodeId!);
              if (a && b) {
                const topRow = Math.min(a.row, b.row);
                const leftCol = Math.min(a.column, b.column);
                const bottomRow = Math.max(a.row + a.rows, b.row + b.rows);
                const rightCol = Math.max(a.column + a.columns, b.column + b.columns);
                commit([
                  {
                    op: "editWordTable",
                    nodeId: tableId,
                    change: {
                      kind: "mergeCells",
                      row: topRow,
                      column: leftCol,
                      rows: bottomRow - topRow,
                      columns: rightCol - leftCol,
                    },
                  },
                ]);
              }
            } else if (state.startCell) {
              const rect = state.startCell.getBoundingClientRect();
              const edgeThreshold = 10;
              const props: Record<string, string> = {};
              if (Math.abs(state.startY - rect.top) <= edgeThreshold) props["border-top"] = "0pt none #000000";
              else if (Math.abs(state.startY - rect.bottom) <= edgeThreshold) props["border-bottom"] = "0pt none #000000";
              else if (Math.abs(state.startX - rect.left) <= edgeThreshold) props["border-left"] = "0pt none #000000";
              else if (Math.abs(state.startX - rect.right) <= edgeThreshold) props["border-right"] = "0pt none #000000";
              else {
                props["border-top"] = "0pt none #000000";
                props["border-right"] = "0pt none #000000";
                props["border-bottom"] = "0pt none #000000";
                props["border-left"] = "0pt none #000000";
              }
              commit([
                {
                  op: "formatWordCell",
                  nodeId: state.startCell.dataset.nodeId!,
                  properties: props,
                },
              ]);
            }
          }
        } catch (error) {
          callbacks.onError(error);
        }
      },
      { signal },
    );
  }

  return {
    mode: () => activeMode,
    setMode,
    bind,
  };
}

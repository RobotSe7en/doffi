import { readFileSync } from "node:fs";
import { expect, test } from "vitest";
import { createDocumentSession, DEFAULT_WORKBOOK_RULE_LIMITS, openDocument } from "@doffi/core";
import { importXlsx } from "@doffi/office";
import { modelWindowSource } from "../packages/render/src/workbook/window-source.js";
import { modelAssetSource } from "../packages/render/src/workbook/assets.js";
import { DEFAULT_CHART_READ_LIMITS, modelChartSource } from "../packages/render/src/workbook/chart-source.js";
import { modelRuleSource } from "../packages/render/src/workbook/rule-source.js";

test.each(["06-rules-validation", "07-three-charts", "08-images-shapes"])("%s read adapters retain one immutable session revision across edits and disposal", async name => {
  const model = openDocument(await importXlsx(readFileSync("examples/xlsx/" + name + ".xlsx"), { documentId: "read-snapshot" }), { kind: "workbook" }), session = createDocumentSession(model);
  const windows = modelWindowSource(session), assets = modelAssetSource(session), charts = modelChartSource(session), rules = modelRuleSource(session);
  const snapshot = session.snapshot(), sheetId = snapshot.metadata.workbook.sheets[0]!.sheetId, signal = new AbortController().signal;
  const identity = { documentId: "read-snapshot", revision: 0 };
  const windowRequest = { ...identity, windows: [{ sheetId, row: 0, column: 0, rows: 10, columns: 10 }], maxCells: 100, maxCodeUnits: 65536 };
  const chartRequest = { ...identity, ...DEFAULT_CHART_READ_LIMITS, chartIds: snapshot.metadata.workbook.charts?.map(chart => chart.chartId) ?? [] };
  const ruleRequest = { ...identity, ...DEFAULT_WORKBOOK_RULE_LIMITS, sheetId, coordinates: [{ row: 0, column: 0 }, { row: 1, column: 0 }] };
  const beforeWindow = await windows.read(windowRequest, signal), beforeCharts = await charts.read(chartRequest, signal), beforeRules = await rules.read(ruleRequest, signal);
  const asset = snapshot.manifest.assets[0], assetRequest = asset ? { ...identity, path: asset.path, maxBytes: 8 * 1024 * 1024 } : undefined;
  const beforeAsset = assetRequest ? await assets.readAsset(assetRequest, signal) : undefined;
  session.commit(session.prepare({ protocolVersion: 1, documentId: identity.documentId, baseRevision: 0, operations: [{ op: "updateWorkbookSheet", sheetId, changes: { name: "Changed" } }] }));
  session.dispose();
  expect(await windows.read(windowRequest, signal)).toEqual(beforeWindow);
  expect(await charts.read(chartRequest, signal)).toEqual(beforeCharts);
  expect(await rules.read(ruleRequest, signal)).toEqual(beforeRules);
  if (assetRequest) expect(await assets.readAsset(assetRequest, signal)).toEqual(beforeAsset);
  expect(windows.snapshot.manifest.revision).toBe(0);
});

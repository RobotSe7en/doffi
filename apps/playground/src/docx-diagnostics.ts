import { DocumentError } from "@doffi/core";

const route = "/__doffi/docx-diagnostics";
const headers = { "x-doffi-docx-diagnostics": "1" };
export const docxDiagnosticsEnabled = Boolean(document.querySelector('meta[name="doffi-docx-diagnostics"]'));
export const docxDiagnosticsNotice = "DOCX 诊断已开启。测试服务器会将选择的本地 DOCX 原件临时保留最多 1 小时。";

export async function retainDocx(bytes: Uint8Array<ArrayBuffer>, signal: AbortSignal): Promise<string | undefined> {
  try {
    const response = await fetch(route, {
      method: "POST",
      headers: { ...headers, "Content-Type": "application/octet-stream" },
      body: bytes,
      signal,
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const { id } = await response.json();
    if (typeof id === "string" && /^[a-f0-9-]{36}$/.test(id)) return id;
  } catch (error) {
    if (!signal.aborted) console.warn("Unable to save the DOCX diagnostic copy.", error);
  }
  return undefined;
}

export function reportDocx(id: string, phase: "import" | "open-document" | "preview", started: number,
  result: { annotations: number } | { error: unknown }): void {
  let outcome: object = result;
  if ("error" in result) {
    const { error } = result;
    const stack = error instanceof Error ? error.stack?.split("\n").slice(1)
      .filter((line) => line.includes(location.origin + "/")).slice(0, 8).join("\n").slice(0, 4000) : undefined;
    outcome = error instanceof DocumentError
      ? { code: error.code, message: error.message, stack }
      : { code: "UnexpectedError", stack };
  }
  void fetch(`${route}/${id}`, {
    method: "POST",
    headers: { ...headers, "Content-Type": "application/json" },
    body: JSON.stringify({ phase, elapsedMs: Math.round(performance.now() - started), ...outcome }),
    keepalive: true,
  }).catch(() => {});
}

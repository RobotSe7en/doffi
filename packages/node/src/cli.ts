#!/usr/bin/env node
import { createDocumentConverter } from "./render.js";
import type { ConverterOptions, DocumentConverter } from "./render.js";
import { serveConversionProtocol } from "./protocol.js";
import type { ProtocolLimits } from "./protocol.js";

async function run(): Promise<void> {
  let converter: DocumentConverter | undefined;
  const controller = new AbortController();
  const stop = (): void => { controller.abort(); process.stdin.destroy(); };
  try {
    const raw = process.env.DOFFI_CONVERSION_CONFIG;
    const config = raw ? JSON.parse(raw) as { converter?: ConverterOptions; protocol?: ProtocolLimits } : {};
    converter = createDocumentConverter({ ...config.converter,
      ...(process.env.DOFFI_CHROMIUM_EXECUTABLE ? { browserExecutable: process.env.DOFFI_CHROMIUM_EXECUTABLE } : {}),
      ...(process.env.DOFFI_CONVERSION_TIMEOUT ? { timeoutMs: Number(process.env.DOFFI_CONVERSION_TIMEOUT) } : {}),
    });
    process.once("SIGTERM", stop); process.once("SIGINT", stop);
    await serveConversionProtocol(process.stdin, async (line) => {
      await new Promise<void>((resolve, reject) => process.stdout.write(line, (error) => error ? reject(error) : resolve()));
    }, converter, config.protocol, controller.signal);
  } catch {
    if (!controller.signal.aborted) { process.stderr.write("Conversion process failed or exceeded its transport budget.\n"); process.exitCode = 1; }
  } finally { process.off("SIGTERM", stop); process.off("SIGINT", stop); await converter?.close(); }
}
if (process.argv.slice(2).includes("--help")) process.stdout.write("doffi: convert DOCX, PPTX, XLSX or document-package ZIP bytes to PDF or PNG. Send one V1 JSON request per stdin line and read byte results on stdout.\n");
else await run();

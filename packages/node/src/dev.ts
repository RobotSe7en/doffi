import { spawn } from "node:child_process";
import type { IncomingMessage, ServerResponse } from "node:http";
import { fileURLToPath } from "node:url";
import { conversionEnvironment } from "./environment.js";
import { ConversionError } from "./render.js";
import type { ConverterOptions } from "./render.js";
import { parseConversionRequest } from "./protocol.js";
import type { ConversionRequestV1, ConversionResponseV1 } from "./protocol.js";

export interface ConversionProcessOptions extends ConverterOptions {
  readonly signal?: AbortSignal;
  readonly browserCacheDirectory?: string;
  readonly maxMessageBytes?: number;
}
/** A host-owned one-request child process. Only public runtime configuration is inherited. */
export async function runConversionProcess(input: ConversionRequestV1, options: ConversionProcessOptions = {}): Promise<ConversionResponseV1> {
  const request = parseConversionRequest(input), maximum = options.maxMessageBytes ?? 96 * 1024 * 1024, timeout = options.timeoutMs ?? 30_000;
  if (!Number.isSafeInteger(maximum) || maximum <= 0 || !Number.isSafeInteger(timeout) || timeout <= 0) throw new ConversionError("InvalidRequest", "Conversion process budgets are invalid.");
  const serialized = JSON.stringify(request) + "\n";
  if (Buffer.byteLength(serialized) > maximum) throw new ConversionError("ResourceLimitExceeded", "Conversion request exceeds its process frame budget.");
  if (options.signal?.aborted) throw new ConversionError("Cancelled", "Conversion was cancelled before startup.");
  const env = conversionEnvironment();
  if (options.browserExecutable) env.DOFFI_CHROMIUM_EXECUTABLE = options.browserExecutable;
  if (options.browserCacheDirectory) env.PLAYWRIGHT_BROWSERS_PATH = options.browserCacheDirectory;
  env.DOFFI_CONVERSION_TIMEOUT = String(timeout);
  const serializableFonts = options.webFonts
    ?.filter((font) => typeof font.path === "string" && font.path.length > 0)
    .map(({ family, aliases, weight, style, equivalentAliases, fallbackOnly, path }) => ({
      family,
      ...(aliases ? { aliases } : {}),
      ...(weight !== undefined ? { weight } : {}),
      ...(style !== undefined ? { style } : {}),
      ...(equivalentAliases !== undefined ? { equivalentAliases } : {}),
      ...(fallbackOnly !== undefined ? { fallbackOnly } : {}),
      path,
    }));
  env.DOFFI_CONVERSION_CONFIG = JSON.stringify({ converter: {
    ...(options.browserChannel ? { browserChannel: options.browserChannel } : {}),
    ...(serializableFonts?.length ? { webFonts: serializableFonts } : {}),
    ...(options.maxInputBytes !== undefined ? { maxInputBytes: options.maxInputBytes } : {}),
    ...(options.maxOutputBytes !== undefined ? { maxOutputBytes: options.maxOutputBytes } : {}),
    ...(options.maxPixels !== undefined ? { maxPixels: options.maxPixels } : {}),
    ...(options.limits ? { limits: options.limits } : {}),
  }, protocol: { maxInputBytes: options.maxInputBytes, maxMessageBytes: maximum, maxOutputMessageBytes: maximum } });
  return new Promise<ConversionResponseV1>((resolve, reject) => {
    const child = spawn(process.execPath, [fileURLToPath(new URL("./cli.js", import.meta.url))], { env, stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
    const chunks: Buffer[] = []; let size = 0, stopped: ConversionError | undefined, force: ReturnType<typeof setTimeout> | undefined;
    const stop = (reason: ConversionError): void => {
      if (stopped) return; stopped = reason; child.kill("SIGTERM"); force = setTimeout(() => child.kill("SIGKILL"), 3000); force.unref();
    };
    const timer = setTimeout(() => stop(new ConversionError("Timeout", "Conversion process timed out.")), timeout);
    const abort = (): void => stop(options.signal?.reason instanceof ConversionError ? options.signal.reason : new ConversionError("Cancelled", "Conversion process was cancelled."));
    options.signal?.addEventListener("abort", abort, { once: true });
    const cleanup = (): void => { clearTimeout(timer); if (force) clearTimeout(force); options.signal?.removeEventListener("abort", abort); };
    child.stdout.on("data", (chunk: Buffer) => {
      size += chunk.length; if (size > maximum) { stop(new ConversionError("ResourceLimitExceeded", "Conversion response exceeds its process frame budget.")); return; } chunks.push(chunk);
    });
    child.stderr.on("data", () => { /* Child diagnostics are bounded fixed messages; no host logging is implicit. */ });
    child.stdin.on("error", () => { /* Exit handling supplies one typed result. */ });
    child.once("error", () => { cleanup(); reject(new ConversionError("ExportFailed", "Unable to start the conversion process.")); });
    child.once("close", (code) => {
      cleanup(); if (stopped) { reject(stopped); return; }
      if (code !== 0) { reject(new ConversionError("ExportFailed", "Conversion process did not complete successfully.")); return; }
      try {
        const response = JSON.parse(Buffer.concat(chunks).toString("utf8")) as ConversionResponseV1;
        if (response.protocolVersion !== 1 || response.taskId !== request.taskId || typeof response.ok !== "boolean") throw new Error();
        resolve(response);
      } catch { reject(new ConversionError("ExportFailed", "Conversion process returned an invalid response.")); }
    });
    child.stdin.end(serialized);
  });
}

export interface DevelopmentConverterOptions extends Omit<ConversionProcessOptions, "signal"> {
  readonly route?: string;
  /** Exact host-owned development origins. Empty keeps the loopback default. */
  readonly allowedOrigins?: readonly string[];
}
export type DevelopmentConversionMiddleware = (request: IncomingMessage, response: ServerResponse, next: () => void) => void;
function loopback(address: string | undefined): boolean { return address === "127.0.0.1" || address === "::1" || address === "::ffff:127.0.0.1"; }
function readBody(request: IncomingMessage, maximum: number, signal: AbortSignal): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = []; let size = 0;
    const cleanup = (): void => { request.off("data", data); request.off("end", end); request.off("error", error); request.off("aborted", error); signal.removeEventListener("abort", cancel); };
    const cancel = (): void => { cleanup(); request.resume(); reject(signal.reason instanceof ConversionError ? signal.reason : new ConversionError("Cancelled", "Conversion upload was cancelled.")); };
    const error = (): void => { cleanup(); reject(new ConversionError("InvalidRequest", "Conversion request body could not be read.")); };
    const data = (chunk: Buffer): void => {
      size += chunk.length;
      if (size > maximum) { cleanup(); request.resume(); reject(new ConversionError("ResourceLimitExceeded", "Conversion request exceeds its HTTP byte budget.")); return; }
      chunks.push(chunk);
    };
    const end = (): void => { cleanup(); resolve(Buffer.concat(chunks)); };
    request.on("data", data); request.once("end", end); request.once("error", error); request.once("aborted", error);
    signal.addEventListener("abort", cancel, { once: true }); if (signal.aborted) cancel();
  });
}
/** Development adapter; loopback by default, with explicit same-origin opt-in.
 * The SDK opens no listener and this gate does not replace host authentication. */
export function createDevelopmentConversionMiddleware(options: DevelopmentConverterOptions = {}): DevelopmentConversionMiddleware {
  const path = options.route ?? "/__doffi/convert", maximum = options.maxMessageBytes ?? 96 * 1024 * 1024, concurrency = options.maxConcurrency ?? 2;
  if (!Number.isSafeInteger(concurrency) || concurrency <= 0 || !Number.isSafeInteger(maximum) || maximum <= 0
    || options.timeoutMs !== undefined && (!Number.isSafeInteger(options.timeoutMs) || options.timeoutMs <= 0)) throw new ConversionError("InvalidRequest", "Development conversion quotas are invalid.");
  if (options.allowedOrigins !== undefined && (!Array.isArray(options.allowedOrigins) || options.allowedOrigins.length > 32)) throw new ConversionError("InvalidRequest", "Development origins must be an explicit bounded list.");
  const allowed = new Set<string>();
  for (const source of options.allowedOrigins ?? []) {
    try {
      if (typeof source !== "string" || source.length > 2048) throw new Error();
      const origin = new URL(source);
      if (!["http:", "https:"].includes(origin.protocol) || origin.hostname.includes("*") || origin.username || origin.password || origin.pathname !== "/" || origin.search || origin.hash) throw new Error();
      allowed.add(origin.origin);
    } catch { throw new ConversionError("InvalidRequest", "Development origins must be explicit HTTP origins without credentials, paths or wildcards."); }
  }
  let active = 0;
  return (request, response, next): void => {
    if (request.url?.split("?")[0] !== path) { next(); return; }
    const send = (status: number, value: unknown): void => {
      if (response.destroyed) return; response.statusCode = status; response.setHeader("Content-Type", "application/json; charset=utf-8"); response.setHeader("Cache-Control", "no-store");
      if (status >= 400) response.setHeader("Connection", "close");
      response.end(JSON.stringify(value), () => { if (status >= 400 && !request.complete) request.destroy(); });
    };
    const protocol = (request.socket as { encrypted?: boolean }).encrypted ? "https:" : "http:";
    let sameOrigin = true, localHost = false, validAuthority = false, authority: URL | undefined;
    try {
      authority = new URL(protocol + "//" + request.headers.host);
      validAuthority = !authority.username && !authority.password
        && authority.pathname === "/" && !authority.search && !authority.hash && Number(authority.port || (protocol === "https:" ? 443 : 80)) === request.socket.localPort;
      localHost = ["localhost", "127.0.0.1", "[::1]"].includes(authority.hostname) && validAuthority;
    } catch { /* Invalid Host is rejected below. */ }
    if (request.headers.origin) {
      try { const origin = new URL(request.headers.origin); sameOrigin = origin.host === authority?.host && origin.protocol === protocol; } catch { sameOrigin = false; }
    }
    const permitted = loopback(request.socket.remoteAddress) && localHost || Boolean(request.headers.origin && authority && allowed.has(authority.origin));
    if (!validAuthority || !permitted || !sameOrigin || request.headers["x-doffi-dev"] !== "1") { send(403, { error: { code: "InvalidRequest", message: "An authorized same-origin development conversion is required." } }); return; }
    if (request.method !== "POST") { send(405, { error: { code: "InvalidRequest", message: "Use a POST conversion request." } }); return; }
    if (!/^application\/json(?:\s*;|$)/i.test(request.headers["content-type"] ?? "")) { send(415, { error: { code: "InvalidRequest", message: "Conversion requires application/json." } }); return; }
    if (active >= concurrency) { send(429, { error: { code: "Busy", message: "Development conversion capacity is in use." } }); return; }
    const length = Number(request.headers["content-length"] ?? 0);
    if (length > maximum) { request.resume(); send(413, { error: { code: "ResourceLimitExceeded", message: "Conversion request exceeds its HTTP byte budget." } }); return; }
    active++; const controller = new AbortController(); let finished = false;
    const cancel = (): void => { if (!finished) controller.abort(); };
    const timer = setTimeout(() => controller.abort(new ConversionError("Timeout", "Development conversion timed out.")), options.timeoutMs ?? 30_000);
    request.once("aborted", cancel); response.once("close", cancel);
    void (async () => {
      try {
        const bytes = await readBody(request, maximum, controller.signal), text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
        let value: unknown; try { value = JSON.parse(text); } catch { throw new ConversionError("InvalidRequest", "Invalid conversion JSON."); }
        const result = await runConversionProcess(parseConversionRequest(value), { ...options, signal: controller.signal }); finished = true; send(200, result);
      } catch (reason) {
        finished = true;
        const error = reason instanceof ConversionError ? reason : new ConversionError("ExportFailed", "Development conversion failed.");
        send(error.code === "ResourceLimitExceeded" ? 413 : ["InvalidRequest", "UnsupportedVersion", "UnsupportedFormat"].includes(error.code) ? 400 : error.code === "Timeout" ? 504 : 500, { error: error.toJSON() });
      } finally { clearTimeout(timer); active--; request.off("aborted", cancel); response.off("close", cancel); }
    })();
  };
}

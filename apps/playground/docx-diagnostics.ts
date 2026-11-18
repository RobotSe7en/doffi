import { randomUUID } from "node:crypto";
import { createWriteStream, mkdtempSync } from "node:fs";
import { rm } from "node:fs/promises";
import type { IncomingMessage, ServerResponse } from "node:http";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { addAbortSignal } from "node:stream";
import { pipeline } from "node:stream/promises";
import type { Plugin } from "vite";

const route = "/__doffi/docx-diagnostics";
const maxUploadBytes = 64 * 1024 * 1024;
const retentionMs = 60 * 60 * 1000;

function sameOrigin(request: IncomingMessage): boolean {
  const { host, origin, "sec-fetch-site": site } = request.headers;
  // Browsers may omit Fetch Metadata on an HTTP LAN address.
  if (site) return site === "same-origin";
  return Boolean(host && (origin === `http://${host}` || origin === `https://${host}`));
}

function contentLength(request: IncomingMessage, maximum: number): number | undefined {
  const value = Number(request.headers["content-length"]);
  return Number.isSafeInteger(value) && value > 0 && value <= maximum ? value : undefined;
}

function reply(request: IncomingMessage, response: ServerResponse, status: number, data?: object): void {
  if (response.destroyed) return;
  response.statusCode = status;
  response.setHeader("Cache-Control", "no-store");
  if (data) response.setHeader("Content-Type", "application/json; charset=utf-8");
  if (status >= 400) response.setHeader("Connection", "close");
  response.end(data && JSON.stringify(data), () => {
    if (status >= 400 && !request.complete) request.destroy();
  });
}

/** Opt-in development host capture. Original bytes have no download endpoint. */
export function docxDiagnostics(): Plugin {
  return {
    name: "docx-upload-diagnostics",
    apply: "serve",
    transformIndexHtml() {
      return [{ tag: "meta", attrs: { name: "doffi-docx-diagnostics", content: "1" } }];
    },
    configureServer(server) {
      const directory = mkdtempSync(join(tmpdir(), "doffi-docx-"));
      const captures = new Map<string, { timer: ReturnType<typeof setTimeout>; reports: number }>();
      const log = (entry: object): void => server.config.logger.info(`[docx-diagnostics] ${JSON.stringify(entry)}`);
      let active = 0;
      log({ event: "enabled", directory, retentionMs });
      server.httpServer?.once("close", () => {
        for (const { timer } of captures.values()) clearTimeout(timer);
        void rm(directory, { recursive: true, force: true }).catch(() => log({ event: "cleanup-failed", directory }));
      });

      server.middlewares.use((request, response, next) => {
        const path = request.url?.split("?")[0];
        if (path !== route && !path?.startsWith(route + "/")) { next(); return; }
        const send = (status: number, data?: object): void => reply(request, response, status, data);
        if (!sameOrigin(request) || request.headers["x-doffi-docx-diagnostics"] !== "1") { send(403); return; }
        if (request.method !== "POST") { send(405); return; }

        if (path === route) {
          if (request.headers["content-type"] !== "application/octet-stream") { send(415); return; }
          const bytes = contentLength(request, maxUploadBytes);
          if (bytes === undefined) { send(413); return; }
          if (active >= 2 || captures.size + active >= 8) { send(429); return; }
          active++;
          const id = randomUUID(), file = join(directory, `${id}.docx`);
          void (async () => {
            try {
              await pipeline(request, createWriteStream(file, { flags: "wx", mode: 0o600 }), {
                signal: AbortSignal.timeout(60_000),
              });
              const timer = setTimeout(() => {
                void rm(file, { force: true }).then(() => captures.delete(id), () => log({ event: "cleanup-failed", id }));
              }, retentionMs);
              timer.unref();
              captures.set(id, { timer, reports: 0 });
              log({ event: "saved", id, bytes, file });
              send(201, { id });
            } catch {
              await rm(file, { force: true }).catch(() => log({ event: "cleanup-failed", id }));
              log({ event: "upload-failed", id, bytes });
              send(500);
            } finally {
              active--;
            }
          })();
          return;
        }

        const id = path!.slice(route.length + 1), capture = captures.get(id);
        if (!capture) { send(404); return; }
        if (request.headers["content-type"] !== "application/json") { send(415); return; }
        if (contentLength(request, 8 * 1024) === undefined) { send(413); return; }
        if (capture.reports >= 12) { send(429); return; }
        capture.reports++;
        void (async () => {
          try {
            const chunks: Buffer[] = [];
            for await (const chunk of addAbortSignal(AbortSignal.timeout(5000), request)) chunks.push(chunk);
            const { phase, elapsedMs, annotations, code, message, stack } = JSON.parse(Buffer.concat(chunks).toString("utf8"));
            if (!["import", "open-document", "preview"].includes(phase) || !Number.isFinite(elapsedMs) || elapsedMs < 0) { send(400); return; }
            log({ event: "result", id, phase, elapsedMs,
              ...(Number.isSafeInteger(annotations) && annotations >= 0 ? { annotations } : {}),
              ...(typeof code === "string" ? { code: code.slice(0, 64) } : {}),
              ...(typeof message === "string" ? { message: message.slice(0, 256) } : {}),
              ...(typeof stack === "string" ? { stack: stack.slice(0, 4000) } : {}),
            });
            send(204);
          } catch {
            send(400);
          }
        })();
      });
    },
  };
}

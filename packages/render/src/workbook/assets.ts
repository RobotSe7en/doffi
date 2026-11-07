import { captureWorkbookModel, DocumentError } from "@doffi/core";
import type { WorkbookModel, WorkbookSnapshot } from "@doffi/core";
import { checkWindowCancellation, checkWindowIdentity } from "./window-source.js";
import { workbookImageFootprint } from "./image-footprint.js";

export interface WorkbookAssetRequest { readonly documentId: string; readonly revision: number; readonly path: string; readonly maxBytes: number }
export interface WorkbookAssetResponse { readonly documentId: string; readonly revision: number; readonly path: string; readonly bytes: Uint8Array }
/** This source owns an already validated, authorized immutable workbook. */
export interface WorkbookAssetSource {
  readonly snapshot: WorkbookSnapshot;
  readAsset(request: WorkbookAssetRequest, signal: AbortSignal): Promise<WorkbookAssetResponse>;
}
export interface WorkbookAssetLimits {
  readonly maxPaths: number; readonly maxAssets: number; readonly maxAssetBytes: number; readonly maxBytes: number;
  readonly maxImagePixels: number; readonly maxPixels: number; readonly maxImageFrames: number;
}
const defaults: WorkbookAssetLimits = { maxPaths: 5000, maxAssets: 256, maxAssetBytes: 8 * 1024 * 1024, maxBytes: 64 * 1024 * 1024,
  maxImagePixels: 16 * 1024 * 1024, maxPixels: 64 * 1024 * 1024, maxImageFrames: 256 };
export interface WorkbookPreparedAssets {
  readonly count: number; readonly byteLength: number; readonly pixels: number;
  url(path: string): string | undefined;
  dispose(): void;
}

export function modelAssetSource(model: WorkbookModel): WorkbookAssetSource {
  model = captureWorkbookModel(model);
  const snapshot = model.snapshot();
  return Object.freeze({ snapshot, async readAsset(request: WorkbookAssetRequest, signal: AbortSignal): Promise<WorkbookAssetResponse> {
    checkWindowCancellation(signal); checkWindowIdentity(request, snapshot);
    if (!Number.isSafeInteger(request.maxBytes) || request.maxBytes < 1) throw new DocumentError("InvalidOperation", "Workbook resource requests require a positive byte budget.");
    const bytes = model.getAsset(request.path, request.maxBytes);
    if (!bytes) throw new DocumentError("AssetUnavailable", "The requested resource is not declared by this workbook.");
    checkWindowCancellation(signal);
    return Object.freeze({ documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, path: request.path, bytes });
  } });
}

/** Prepare one candidate window's resource lease. Duplicate references share a
 * URL in this lease only. The previous visible window keeps its own lease until
 * commit; failed/cancelled candidates release every URL, even when a provider
 * ignores cancellation. A successful lease is explicitly disposed by its owner.
 * No URL fetch or whole-workbook serialization takes place here. */
export async function prepareWorkbookAssets(source: WorkbookAssetSource, paths: readonly string[], signal: AbortSignal, input: Partial<WorkbookAssetLimits> = {}): Promise<WorkbookPreparedAssets> {
  const limits = { ...defaults, ...input };
  if (Object.keys(input).some(key => !Object.hasOwn(defaults, key)) || !Object.values(limits).every(value => Number.isSafeInteger(value) && value > 0)) throw new DocumentError("InvalidOperation", "Workbook asset limits must be supported positive safe integers.");
  checkWindowCancellation(signal);
  if (paths.length > limits.maxPaths) throw new DocumentError("ResourceLimitExceeded", "Workbook resource references exceed their request budget.");
  const unique = [...new Set(paths)];
  if (unique.length > limits.maxAssets) throw new DocumentError("ResourceLimitExceeded", "Workbook resources exceed their asset-count budget.");
  const snapshot = source.snapshot, descriptors = new Map(snapshot.manifest.assets.map(asset => [asset.path, asset]));
  for (const path of unique) {
    const descriptor = descriptors.get(path);
    if (!descriptor) throw new DocumentError("AssetUnavailable", "A workbook display requested an undeclared resource.");
    if (!["image/png", "image/jpeg", "image/webp"].includes(descriptor.mimeType)) throw new DocumentError("UnsupportedFeature", "Workbook display resources must be supported raster images.");
  }
  const urls = new Map<string, string>(); let total = 0, pixels = 0, disposed = false;
  const dispose = (): void => { if (!disposed) { disposed = true; for (const url of urls.values()) URL.revokeObjectURL(url); urls.clear(); } };
  let abort: () => void = () => {};
  const cancelled = new Promise<never>((_, reject) => { abort = () => { dispose(); reject(new DocumentError("OperationCancelled", "Workbook resource preparation was cancelled.")); }; });
  signal.addEventListener("abort", abort, { once: true });
  try {
    checkWindowCancellation(signal);
    const { documentId, revision } = snapshot.manifest;
    for (const path of unique) {
      const maxBytes = Math.min(limits.maxAssetBytes, limits.maxBytes - total);
      if (maxBytes <= 0) throw new DocumentError("ResourceLimitExceeded", "Workbook resources exceed their cumulative byte budget.");
      const response = await Promise.race([source.readAsset(Object.freeze({ documentId, revision, path, maxBytes }), signal), cancelled]);
      checkWindowCancellation(signal); checkWindowIdentity(response, snapshot);
      if (response.path !== path || !(response.bytes instanceof Uint8Array)) throw new DocumentError("InvalidDocument", "A workbook resource response does not match its request.");
      if (response.bytes.byteLength > maxBytes) throw new DocumentError("ResourceLimitExceeded", "Workbook resource bytes exceed their requested budget.");
      // Isolate the provider's mutable buffer before asynchronous hashing and
      // Blob creation. A valid digest refers to bytes already validated by core.
      const bytes = new Uint8Array(response.bytes), descriptor = descriptors.get(path)!;
      const digest = await Promise.race([crypto.subtle.digest("SHA-256", bytes), cancelled]);
      checkWindowCancellation(signal);
      const actual = [...new Uint8Array(digest)].map(byte => byte.toString(16).padStart(2, "0")).join("");
      if (actual !== descriptor.sha256) throw new DocumentError("InvalidDocument", "Workbook display resource bytes do not match the source digest.");
      const maxPixels = Math.min(limits.maxImagePixels, limits.maxPixels - pixels);
      if (maxPixels <= 0) throw new DocumentError("ResourceLimitExceeded", "Workbook images exceed their cumulative pixel budget.");
      const footprint = workbookImageFootprint(bytes, descriptor.mimeType, maxPixels, limits.maxImageFrames); pixels += footprint.pixels;
      const url = URL.createObjectURL(new Blob([bytes], { type: descriptor.mimeType })); urls.set(path, url); total += bytes.byteLength;
    }
    return Object.freeze({ count: urls.size, byteLength: total, pixels, dispose, url(path: string) {
      if (disposed) throw new DocumentError("InvalidOperation", "These workbook display resources have been released.");
      return urls.get(path);
    } });
  } catch (error) { dispose(); throw error; }
  finally { signal.removeEventListener("abort", abort); }
}

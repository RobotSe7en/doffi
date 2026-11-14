import { DocumentError, createOpaqueId } from '@doffi/core';
import type { TransactionPrepareOptions } from '@doffi/core';
export async function wordImageResource(document: Document, file: File): Promise<{
    path: string;
    width: number;
    height: number;
    options: TransactionPrepareOptions;
}> { if (file.size > 64 * 1024 * 1024)
    throw new DocumentError('ResourceLimitExceeded', '图片超过 64 MiB。'); const bytes = new Uint8Array(await file.arrayBuffer()), digest = await document.defaultView!.crypto.subtle.digest('SHA-256', bytes), sha256 = Array.from(new Uint8Array(digest), b => b.toString(16).padStart(2, '0')).join(''), mime = file.type || (/\.jpe?g$/i.test(file.name) ? 'image/jpeg' : /\.webp$/i.test(file.name) ? 'image/webp' : 'image/png'), ext = mime === 'image/jpeg' ? 'jpg' : mime === 'image/webp' ? 'webp' : 'png', path = 'assets/' + createOpaqueId() + '.' + ext, bitmap = await createImageBitmap(file), width = Math.min(360, bitmap.width * .75), height = width * bitmap.height / bitmap.width; bitmap.close(); return { path, width, height, options: { assets: [{ descriptor: { path, mimeType: mime, sha256 }, bytes }] } }; }

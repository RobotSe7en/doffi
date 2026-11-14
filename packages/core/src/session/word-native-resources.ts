import { WordNativeFragmentCollectionSchema } from '../contracts/word.js';
import type { WordNativeContent, WordMetadata } from '../contracts/word.js';
import type { ParsedHtml } from '../document/html.js';
import type { DeepReadonly, DocumentReadLimits } from '../runtime.js';
import { assetDigest } from '../document/assets.js';
import { readJson, encode, resolveReference } from '../document/package.js';
import { parseContract } from '../schema-registry.js';
import { checkLimit, invalid } from '../document/validation.js';
/** The native graph owns opaque resources. Core follows its declared edges,
 * without parsing OOXML or fetching relationships. Older unindexed fragments
 * conservatively retain their source-part dependencies. */
export function pruneWordNativeResources(word: WordMetadata, previous: DeepReadonly<WordNativeContent> | undefined, files: Map<string, Uint8Array>, source: ParsedHtml, entrypoint: string, limits: DocumentReadLimits): void {
    const native = word.nativeContent;
    if (!native || JSON.stringify(native) === JSON.stringify(previous))
        return;
    const owned = new Set([...(previous?.fragments ?? []).flatMap(fragment => [fragment.assetPath, ...(fragment.preview ? [fragment.preview.assetPath] : [])]), ...(previous?.parts ?? []).map(part => part.assetPath), ...native.fragments.flatMap(fragment => [fragment.assetPath, ...(fragment.preview ? [fragment.preview.assetPath] : [])]), ...native.parts.map(part => part.assetPath)]);
    const selected = new Set<string>(), keepParts = new Set<string>(), parts = new Map(native.parts.map(part => [part.sourcePart, part])), relations = new Map<string, typeof native.relationships>();
    for (const relation of native.relationships) {
        const entries = relations.get(relation.sourcePart) ?? [];
        entries.push(relation);
        relations.set(relation.sourcePart, entries);
    }
    const byId = new Map(native.relationships.map(relation => [relation.sourcePart + '\n' + relation.id, relation])), wildcards = new Set<string>();
    const pending: string[] = [];
    const select = (relation: typeof native.relationships[number]): void => { const key = relation.sourcePart + '\n' + relation.id; if (selected.has(key))
        return; selected.add(key); if (!relation.external && parts.has(relation.target) && !keepParts.has(relation.target)) {
        keepParts.add(relation.target);
        pending.push(relation.target);
    } };
    for (const relation of native.relationships)
        if (relation.root)
            select(relation);
    for (const fragment of native.fragments) {
        if (parts.has(fragment.sourcePart) && !keepParts.has(fragment.sourcePart)) {
            keepParts.add(fragment.sourcePart);
            pending.push(fragment.sourcePart);
        }
        if (fragment.relationshipIds === undefined)
            wildcards.add(fragment.sourcePart);
        else
            for (const id of fragment.relationshipIds) {
                const relation = byId.get(fragment.sourcePart + '\n' + id);
                if (!relation)
                    invalid('The native fragment dependency is missing.');
                select(relation);
            }
    }
    for (const part of wildcards)
        for (const relation of relations.get(part) ?? [])
            select(relation);
    for (let index = 0; index < pending.length; index++)
        for (const relation of relations.get(pending[index]!) ?? [])
            select(relation);
    native.parts = native.parts.filter(part => keepParts.has(part.sourcePart));
    native.relationships = native.relationships.filter(relation => selected.has(relation.sourcePart + '\n' + relation.id));
    const manifest = parseContract('document-manifest', readJson(files, 'manifest.json')), collections = new Map<string, Set<string>>();
    for (const fragment of native.fragments)
        if (fragment.fragmentFormat === 'json') {
            if (!fragment.fragmentKey)
                invalid('A native JSON fragment requires its collection key.');
            const keys = collections.get(fragment.assetPath) ?? new Set<string>();
            keys.add(fragment.fragmentKey);
            collections.set(fragment.assetPath, keys);
        }
    for (const [path, keys] of collections) {
        const parsed = WordNativeFragmentCollectionSchema.safeParse(readJson(files, path));
        if (!parsed.success)
            invalid('The native fragment collection is invalid.');
        checkLimit(Object.keys(parsed.data.fragments).length, limits.maxNodes);
        const entries = [...keys].map(key => { const value = Object.hasOwn(parsed.data.fragments, key) ? parsed.data.fragments[key] : undefined; if (value === undefined)
            invalid('The native fragment collection entry is missing.'); return [key, value] as const; });
        if (entries.length === Object.keys(parsed.data.fragments).length)
            continue;
        const bytes = encode(JSON.stringify({ ...parsed.data, fragments: Object.fromEntries(entries) }));
        checkLimit(bytes.byteLength, limits.maxFileBytes);
        const sha256 = assetDigest(bytes), next = 'assets/native-fragments-' + sha256 + '.json', existing = files.get(next);
        if (existing && assetDigest(existing) !== sha256)
            invalid('A native collection resource path is already occupied.');
        files.set(next, bytes);
        if (!manifest.assets.some(asset => asset.path === next))
            manifest.assets.push({ path: next, mimeType: 'application/json', sha256 });
        for (const fragment of native.fragments)
            if (fragment.assetPath === path && fragment.fragmentFormat === 'json')
                fragment.assetPath = next;
        owned.add(path);
    }
    const used = new Set(native.fragments.flatMap(fragment => [fragment.assetPath, ...(fragment.preview ? [fragment.preview.assetPath] : [])]).concat(native.parts.map(part => part.assetPath)));
    for (const id of source.nodeIds()) {
        const node = source.peekNode(id)!;
        if (node.tagName === 'img')
            used.add(resolveReference(node.attributes.src!, entrypoint));
    }
    const visit = (value: unknown): void => { if (typeof value === 'string') {
        if (owned.has(value))
            used.add(value);
    }
    else if (Array.isArray(value))
        for (const item of value)
            visit(item);
    else if (value && typeof value === 'object')
        for (const [key, item] of Object.entries(value))
            if (key !== 'nativeContent')
                visit(item); };
    visit(word);
    const removed = new Set([...owned].filter(path => !used.has(path)));
    for (const path of removed)
        files.delete(path);
    manifest.assets = manifest.assets.filter(asset => !removed.has(asset.path));
    if (!native.fragments.length && !native.parts.length && !native.relationships.length)
        delete word.nativeContent;
    files.set('manifest.json', encode(JSON.stringify(manifest)));
}

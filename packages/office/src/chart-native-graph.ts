import { DocumentError } from '@doffi/core';
import type { DeepReadonly, DocumentPackage, DocumentReadLimits, WordChart } from '@doffi/core';
import type { OpcPackage } from './ooxml/package.js';
import { R, xml } from './ooxml/xml.js';
import type { XmlNode, XmlBudget } from './ooxml/xml.js';
import { serializeXml } from './ooxml/serialize.js';
import type { OutputOpc } from './export/opc.js';
type Graph = NonNullable<WordChart['native']>;
export function readChartGraph(opc: OpcPackage, chartPart: string, asset: (part: string, type: string) => string): Pick<Graph, 'parts' | 'relationships'> {
    const parts: Graph['parts'] = [], relationships: Graph['relationships'] = [], seen = new Set<string>();
    const visit = (part: string): void => {
        if (seen.has(part))
            return;
        seen.add(part);
        const contentType = opc.contentType(part) ?? 'application/octet-stream';
        if (!opc.files.has(part))
            throw new DocumentError('InvalidDocument', 'A native chart dependency is missing.');
        parts.push({ sourcePart: part, assetPath: asset(part, contentType), contentType });
        const relations = opc.relationships(part);
        for (const relation of relations.values()) {
            relationships.push({ sourcePart: part, ...relation });
            if (!relation.external)
                visit(relation.target);
        }
        if (/(?:\/|\+)xml(?:;|$)/i.test(contentType) || part.endsWith('.xml')) {
            const inspect = (node: XmlNode): void => {
                for (const [key, id] of node.attrs)
                    if (key.startsWith(R + ' ') && !relations.has(id))
                        throw new DocumentError('InvalidDocument', 'A native chart XML relationship is missing.');
                for (const child of node.content)
                    if (typeof child !== 'string')
                        inspect(child);
            };
            inspect(opc.read(part));
        }
    };
    visit(chartPart);
    return { parts, relationships };
}
/** Copy one immutable native part graph per chart instance. Current chart XML
 * and workbook bytes replace their templates before relationship IDs are
 * remapped, so copied charts can be edited independently. */
export function writeChartGraph(opc: OutputOpc, graph: DeepReadonly<Graph>, source: DocumentPackage, chartPart: string, overrides: ReadonlyMap<string, Uint8Array>, limits: DocumentReadLimits, budget: XmlBudget): void {
    const prefix = chartPart.slice(0, -4) + '-parts/', paths = new Map<string, string>([[graph.chartPart, chartPart]]);
    let serial = 0;
    for (const part of graph.parts)
        if (!paths.has(part.sourcePart)) {
            const ext = /\.([a-zA-Z0-9]{1,8})$/.exec(part.sourcePart)?.[1]?.toLowerCase() ?? 'bin';
            paths.set(part.sourcePart, prefix + (++serial) + '.' + ext);
        }
    const mappings = new Map<string, Map<string, string>>();
    for (const relation of graph.relationships) {
        const part = paths.get(relation.sourcePart), target = relation.external ? relation.target : paths.get(relation.target);
        if (!part || !target)
            throw new DocumentError('ExportFailed', 'A chart dependency cannot be remapped.');
        const id = opc.relate(part, relation.type, relation.external ? target : '/' + target, relation.external), mapping = mappings.get(relation.sourcePart) ?? new Map<string, string>();
        mapping.set(relation.id, id);
        mappings.set(relation.sourcePart, mapping);
    }
    const codecs = globalThis as unknown as {
        TextEncoder: new () => {
            encode(text: string): Uint8Array;
        };
    };
    for (const part of graph.parts) {
        let bytes = overrides.get(part.sourcePart) ?? source.get(part.assetPath);
        if (!bytes)
            throw new DocumentError('AssetUnavailable', 'A native chart template is unavailable.');
        if (/(?:\/|\+)xml(?:;|$)/i.test(part.contentType) || part.sourcePart.endsWith('.xml')) {
            const root = xml(bytes, part.sourcePart, limits, budget), mapping = mappings.get(part.sourcePart);
            const rewrite = (node: XmlNode): void => {
                for (const [key, value] of node.attrs)
                    if (key.startsWith(R + ' ')) {
                        const mapped = mapping?.get(value);
                        if (!mapped)
                            throw new DocumentError('ExportFailed', 'A native chart relationship is unavailable.');
                        (node.attrs as Map<string, string>).set(key, mapped);
                    }
                for (const child of node.content)
                    if (typeof child !== 'string')
                        rewrite(child);
            };
            rewrite(root);
            bytes = new codecs.TextEncoder().encode(serializeXml(root));
        }
        opc.binary(paths.get(part.sourcePart)!, part.contentType, bytes);
    }
}

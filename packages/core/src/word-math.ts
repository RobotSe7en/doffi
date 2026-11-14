import { DocumentError } from './errors.js';
import type { WordMathNode, WordMathTree } from './contracts/word-math.js';
import type { DeepReadonly } from './runtime.js';
function fail(message: string): never {
    throw new DocumentError('InvalidDocument', message);
}
export function validateWordMath(math: DeepReadonly<WordMathTree>): void {
    const nodes = new Map(math.nodes.map(node => [node.mathId, node]));
    if (nodes.size !== math.nodes.length || !nodes.has(math.rootMathId))
        fail('Math requires unique local identities and a root.');
    const visited = new Set<string>(), active = new Set<string>();
    const visit = (id: string, depth: number, parent?: DeepReadonly<WordMathNode>): void => {
        if (depth > 128)
            fail('Math exceeds the supported structural depth.');
        if (active.has(id) || visited.has(id))
            fail('Math must be one acyclic tree without shared child ownership.');
        const node = nodes.get(id);
        if (!node)
            fail('A math argument is missing.');
        if (node.kind === 'text' && /[\ud800-\udfff\u0000]/u.test(node.text))
            fail('Math text must preserve well-formed Unicode.');
        if(node.kind==='matrixRow'&&(parent?.kind!=='matrix'||node.style))fail('A matrix row requires its matrix and cannot carry independent formatting.');
        if(node.kind==='row'&&node.argumentSize!==undefined&&(!parent||parent.kind==='row'))fail('Relative math sizes require a structural math argument.');
        for(const value of node.kind==='fence'?[node.open,node.close,node.separator]:node.kind==='nary'?[node.operator]:node.kind==='accent'?[node.symbol]:[])if(value!==undefined&&(Array.from(value).length>1||/[\ud800-\udfff\u0000]/u.test(value)))fail('A math delimiter or operator must be a single Unicode character.');
        if (node.kind === 'script' && node.children.length !== (['sub', 'sup'].includes(node.mode) ? 2 : 3))
            fail('The script arguments do not match their mode.');
        if (node.kind === 'matrix') {
            const rows = node.children.map(id => nodes.get(id));
            if (rows.some(row => row?.kind !== 'matrixRow') || rows.some(row => row?.kind === 'matrixRow' && row.children.length !== (rows[0] as Extract<WordMathNode, {
                kind: 'matrixRow';
            }>).children.length))
                fail('Math matrix rows must form a rectangle.');
            if(node.columnAlignments&&node.columnAlignments.length>(rows[0] as Extract<WordMathNode,{kind:'matrixRow'}>).children.length)fail('Matrix column alignment exceeds its column count.');
        }
        active.add(id);
        for (const child of 'children' in node ? node.children : [])
            visit(child, depth + 1,node);
        active.delete(id);
        visited.add(id);
    };
    visit(math.rootMathId, 0);
    if (visited.size !== nodes.size)
        fail('Math cannot retain disconnected nodes.');
}
/** Small first-party linear input: structural templates and literal Unicode.
 * Native OMML can represent additional trees through the same public contract. */
export function parseWordMath(source: string): WordMathTree {
    if (source.length > 1000000)
        throw new DocumentError('ResourceLimitExceeded', 'Linear math exceeds its character budget.');
    const nodes: WordMathNode[] = [];
    let cursor = 0, serial = 0;
    type Input<T> = T extends WordMathNode ? Omit<T, 'mathId'> : never;
    const add = (node: Input<WordMathNode>): string => {
        if(serial>=50_000)throw new DocumentError('ResourceLimitExceeded','Linear math exceeds its node budget.');
        const id = 'm' + ++serial;
        nodes.push({ ...node, mathId: id } as WordMathNode);
        return id;
    };
    const text = (value: string, variant?: 'normal'): string => add({ kind: 'text', text: value, ...(variant ? { variant } : {}) });
    const symbols: Readonly<Record<string, string>> = { alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', theta: 'θ', lambda: 'λ', mu: 'μ', pi: 'π', sigma: 'σ', phi: 'φ', omega: 'ω', Delta: 'Δ', Sigma: 'Σ', Omega: 'Ω', infty: '∞', times: '×', cdot: '⋅', pm: '±', le: '≤', leq: '≤', ge: '≥', geq: '≥', ne: '≠', neq: '≠', to: '→', rightarrow: '→', partial: '∂', nabla: '∇', forall: '∀', exists: '∃', in: '∈', notin: '∉', subset: '⊂', cup: '∪', cap: '∩' };
    const error = (message: string): never => {
        throw new DocumentError('InvalidOperation', message + ' At ' + cursor + '.');
    };
    const group = (depth: number, close = '}'): string => {
        while (cursor < source.length && /\s/.test(source[cursor]!))
            cursor++;
        if (source[cursor] !== '{')
            return atom(depth + 1);
        cursor++;
        const value = row(depth + 1, close);
        if (source[cursor] !== close)
            error('A math argument is not closed');
        cursor++;
        return value;
    };
    const atom = (depth: number): string => {
        if (depth > 128)
            error('Math nesting exceeds its depth budget');
        const character = source[cursor];
        if (character === undefined)
            return text('');
        if (character === '{')
            return group(depth);
        if (character === '\\') {
            cursor++;
            const command = /^[A-Za-z]+/.exec(source.slice(cursor))?.[0] ?? source[cursor++] ?? '';
            if (/[A-Za-z]/.test(command))
                cursor += command.length;
            if (command === 'matrix') {
                if (source[cursor++] !== '{')
                    error('A matrix requires braces');
                const rows: string[] = [], cells: string[] = [];
                while (cursor < source.length) {
                    cells.push(row(depth + 1, '}', true));
                    if (source[cursor] === '&') {
                        cursor++;
                        continue;
                    }
                    rows.push(add({ kind: 'matrixRow', children: cells.splice(0) }));
                    if (source.startsWith('\\\\', cursor)) {
                        cursor += 2;
                        continue;
                    }
                    if (source[cursor++] === '}')
                        return add({ kind: 'matrix', children: rows });
                    error('The matrix is not closed');
                }
                return error('The matrix is not closed');
            }
            if (['mathbf', 'mathit', 'mathsf', 'mathbb', 'mathcal', 'mathfrak', 'mathtt'].includes(command)) {
                const id = group(depth), variant = ({ mathbf: 'bold', mathit: 'italic', mathsf: 'sansSerif', mathbb: 'doubleStruck', mathcal: 'script', mathfrak: 'fraktur', mathtt: 'monospace' } as const)[command as 'mathbf'], visit = (id: string): void => {
                    const node = nodes.find(node => node.mathId === id)!;
                    if (node.kind === 'text')
                        node.variant = variant;
                    else
                        for (const child of node.children)
                            visit(child);
                };
                visit(id);
                return id;
            }
            if (command === 'overset' || command === 'underset') {
                const limit = group(depth), base = group(depth);
                return add({ kind: 'limit', position: command === 'overset' ? 'upper' : 'lower', children: [base, limit] });
            }
            if (command === 'frac' || command === 'binom') {
                const a = group(depth), b = group(depth), fraction = add({ kind: 'fraction', children: [a, b], mode: command === 'binom' ? 'noBar' : 'bar' });
                return command === 'binom' ? add({ kind: 'fence', children: [fraction], open: '(', close: ')' }) : fraction;
            }
            if (command === 'sqrt') {
                let degree: string | undefined;
                if (source[cursor] === '[') {
                    cursor++;
                    degree = row(depth + 1, ']');
                    if (source[cursor++] !== ']')
                        error('The radical degree is not closed');
                }
                const base = group(depth);
                return add({ kind: 'radical', children: degree ? [base, degree] : [base] });
            }
            if (['hat', 'vec', 'tilde', 'dot', 'ddot', 'overline', 'underline'].includes(command)) {
                const body = group(depth);
                return command === 'overline' || command === 'underline' ? add({ kind: 'bar', children: [body], position: command === 'overline' ? 'top' : 'bottom' }) : add({ kind: 'accent', children: [body], symbol: ({ hat: '̂', vec: '⃗', tilde: '̃', dot: '̇', ddot: '̈' } as Record<string, string>)[command]! });
            }
            if (command === 'text' || command === 'mathrm') {
                if (source[cursor++] !== '{')
                    error('Text requires braces');
                const end = source.indexOf('}', cursor);
                if (end < 0)
                    error('Text is not closed');
                const value = source.slice(cursor, end);
                cursor = end + 1;
                return text(value, 'normal');
            }
            if (command === 'left') {
                const delimiter = (): string => {
                    while (cursor < source.length && /\s/.test(source[cursor]!))
                        cursor++;
                    if (source[cursor] === '\\') {
                        cursor++;
                        const name = /^[A-Za-z]+/.exec(source.slice(cursor))?.[0] ?? source[cursor++] ?? '';
                        if (/[A-Za-z]/.test(name))
                            cursor += name.length;
                        return ({ langle: '⟨', rangle: '⟩', lbrace: '{', rbrace: '}', lvert: '|', rvert: '|', Vert: '‖' } as Record<string, string>)[name] ?? name;
                    }
                    const value = source[cursor++] ?? '';
                    return value === '.' ? '' : value;
                };
                const open = delimiter(), body = row(depth + 1, '\\right');
                if (!source.startsWith('\\right', cursor))
                    error('A delimiter is not closed');
                cursor += 6;
                const close = delimiter();
                return add({ kind: 'fence', children: [body], open, close, grow: true });
            }
            if (command === 'right')
                return error('Unexpected right delimiter');
            if (['sin', 'cos', 'tan', 'log', 'ln', 'exp', 'lim', 'max', 'min','sup','inf'].includes(command))
                return text(command, 'normal');
            if (Object.hasOwn(symbols,command))
                return text(symbols[command]!);
            if (command === 'sum' || command === 'prod' || command === 'int')
                return add({ kind: 'nary', operator: command === 'sum' ? '∑' : command === 'prod' ? '∏' : '∫', children: [add({ kind: 'row', children: [] }), add({ kind: 'row', children: [] }), add({ kind: 'row', children: [] })], limits: command === 'int' ? 'subSup' : 'aboveBelow', hideLower: true, hideUpper: true });
            if (command.length === 1)
                return text(command);
            return error('Unknown math command ' + command);
        }
        const literal = /^(?:\P{M}\p{M}*)(?:\u200d\P{M}\p{M}*)*/u.exec(source.slice(cursor))?.[0] ?? String.fromCodePoint(source.codePointAt(cursor)!);
        cursor += literal.length;
        return text(literal);
    };
    const row = (depth: number, close?: string, matrix = false): string => {
        const children: string[] = [];
        while (cursor < source.length && (!close || !source.startsWith(close, cursor) || close.startsWith('\\') && /[A-Za-z]/.test(source[cursor + close.length] ?? '')) && !(matrix && (source[cursor] === '&' || source.startsWith('\\\\', cursor)))) {
            if (source[cursor] === '}')
                error('An unexpected closing brace was found');
            if ((source[cursor] === '^' || source[cursor] === '_') && children.length) {
                const mode = source[cursor++] === '^' ? 'sup' : 'sub', base = children.pop()!, script = group(depth), previous = nodes.find(node => node.mathId === base)!;
                if (previous.kind === 'nary') {
                    previous.children[mode === 'sub' ? 1 : 2] = script;
                    if (mode === 'sub')
                        previous.hideLower = false;
                    else
                        previous.hideUpper = false;
                    children.push(base);
                }
                else if(previous.kind==='text'&&previous.variant==='normal'&&['lim','max','min','sup','inf'].includes(previous.text)){children.push(add({kind:'limit',position:mode==='sub'?'lower':'upper',children:[base,script]}));}
                else if (previous.kind === 'script' && previous.mode !== mode && previous.mode !== 'subsup') {
                    nodes.splice(nodes.indexOf(previous), 1);
                    children.push(add({ kind: 'script', mode: 'subsup', children: [previous.children[0]!, mode === 'sub' ? script : previous.children[1]!, mode === 'sup' ? script : previous.children[1]!] }));
                }
                else
                    children.push(add({ kind: 'script', mode, children: [base, script] }));
            }
            else if (source[cursor] === '{' && children.length && nodes.find(node => node.mathId === children.at(-1))?.kind === 'nary') {
                const previous = nodes.find(node => node.mathId === children.at(-1))!;
                if (previous.kind === 'nary')
                    previous.children[0] = group(depth);
            }
            else
                children.push(atom(depth + 1));
        }
        return add({ kind: 'row', children });
    };
    const rootMathId = row(0), byId = new Map(nodes.map(node => [node.mathId, node])), reachable = new Set<string>(), visit = (id: string): void => {
        reachable.add(id);
        const node = byId.get(id)!;
        for (const child of 'children' in node ? node.children : [])
            visit(child);
    };
    visit(rootMathId);
    const math = { rootMathId, nodes: nodes.filter(node => reachable.has(node.mathId)) };
    validateWordMath(math);
    return math;
}

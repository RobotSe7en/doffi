import { DocumentError, parseWordMath } from '@doffi/core';
import type { DocumentEditingSession, DocumentOperation, WordMathTree, WordMathNode } from '@doffi/core';
export function createWordMathControls(parent: ShadowRoot, before: HTMLElement, session: DocumentEditingSession, selection: () => {
    paragraphId: string;
    offset: number;
} | undefined, onError: (error: unknown) => void): {
    open(insert: boolean): void;
    bind(root: HTMLElement, signal: AbortSignal): void;
    dispose(): void;
} {
    const document = before.ownerDocument;
    let panel: HTMLElement | undefined, controller: AbortController | undefined, selected: string | undefined;
    const close = (): void => {
        controller?.abort();
        panel?.remove();
        panel = undefined;
    };
    function open(insert: boolean): void {
        close();
        const location = selection(), equations = session.snapshot().metadata.word?.equations ?? [], equation = equations.find(equation => equation.nodeId === selected) ?? equations[0];
        if (insert && !location)
            throw new DocumentError('InvalidOperation', '请先在段落中放置光标。');
        if (!insert && !equation)
            throw new DocumentError('InvalidOperation', '文档中还没有可编辑公式。');
        if (!insert)
            selected = equation!.nodeId;
        controller = new AbortController();
        const signal = controller.signal;
        panel = document.createElement('section');
        panel.setAttribute('role', 'region');
        panel.setAttribute('aria-label', '公式编辑');
        panel.style.cssText = 'padding:12px;background:white;border:1px solid #dbe3ef;max-height:330px;overflow:auto;display:flex;flex-wrap:wrap;align-items:center;gap:8px';
        parent.insertBefore(panel, before);
        const apply = (operations: DocumentOperation[]): void => {
            const snapshot = session.snapshot();
            session.commit(session.prepare({ protocolVersion: 1, documentId: snapshot.manifest.documentId, baseRevision: snapshot.manifest.revision, operations }));
        };
        const button = (name: string, run: () => void): void => {
            const node = document.createElement('button');
            node.type = 'button';
            node.textContent = name;
            node.addEventListener('click', () => {
                try {
                    run();
                }
                catch (error) {
                    onError(error);
                }
            }, { signal });
            panel!.append(node);
        };
        const field = (name: string, value: string): HTMLInputElement => {
            const label = document.createElement('label'), input = document.createElement('input');
            label.textContent = name;
            input.setAttribute('aria-label', name);
            input.value = value;
            label.append(input);
            panel!.append(label);
            return input;
        };
        const choose = (name: string, options: readonly (readonly [
            string,
            string
        ])[], value: string): HTMLSelectElement => {
            const label = document.createElement('label'), input = document.createElement('select');
            label.textContent = name;
            input.setAttribute('aria-label', name);
            for (const [key, text] of options) {
                const option = document.createElement('option');
                option.value = key;
                option.textContent = text;
                input.append(option);
            }
            input.value = value;
            label.append(input);
            panel!.append(label);
            return input;
        };
        button('关闭公式编辑', close);
        if (!insert) {
            const target = choose('当前公式', equations.map((equation, index) => [equation.nodeId, '公式 ' + (index + 1)]), equation!.nodeId);
            target.addEventListener('change', () => {
                selected = target.value;
                open(false);
            }, { signal });
        }
        const display = choose('公式显示', [['inline', '行内'], ['block', '独立行']], equation?.display ?? 'inline'), alignment = choose('公式对齐', [['left', '左对齐'], ['centerGroup', '居中'], ['right', '右对齐']], equation?.alignment ?? 'centerGroup'), font = field('公式字体', equation?.fontFamily ?? 'Cambria Math'), inputLabel = document.createElement('label'), input = document.createElement('textarea');
        inputLabel.textContent = '线性公式';
        input.setAttribute('aria-label', '线性公式');
        input.value = insert ? '\\frac{x^2+1}{2}' : '';
        input.placeholder = '输入公式以替换当前结构';
        input.rows = 2;
        input.cols = 45;
        inputLabel.append(input);
        panel.append(inputLabel);
        for (const [name, value] of [['分式模板', '\\frac{a}{b}'], ['根式模板', '\\sqrt[3]{x}'], ['上下标模板', 'x_1^2'], ['重音模板', '\\vec{x}+\\hat{y}'], ['矩阵模板', '\\matrix{a&b\\\\c&d}'], ['求和模板', '\\sum_{i=1}^{n}{i}']])
            button(name!, () => {
                input.value = value!;
                input.focus();
            });
        const applyMath = (math: WordMathTree): void => {
            const settings = { math, display: display.value as 'inline' | 'block', alignment: alignment.value as 'left' | 'right' | 'centerGroup', fontFamily: font.value };
            if (insert) {
                const previous = new Set(equations.map(equation => equation.nodeId));
                apply([{ op: 'insertWordEquation', nodeId: location!.paragraphId, offset: location!.offset, ...settings }]);
                selected = session.snapshot().metadata.word!.equations!.find(equation => !previous.has(equation.nodeId))!.nodeId;
                open(false);
            }
            else {
                apply([{ op: 'setWordEquation', nodeId: equation!.nodeId, ...settings }]);
                open(false);
            }
        };
        button(insert ? '插入公式' : '按线性输入重建公式', () => {
            if (!input.value.trim())
                throw new DocumentError('InvalidOperation', '请输入公式。');
            applyMath(parseWordMath(input.value));
        });
        if (!insert) {
            button('应用公式显示设置', () => applyMath({ rootMathId: equation!.rootMathId, nodes: JSON.parse(JSON.stringify(equation!.nodes)) as WordMathNode[] }));
            const textNodes = equation!.nodes.filter(node => node.kind === 'text');
            if (textNodes.length) {
                const target = choose('公式文字组件', textNodes.map(node => [node.mathId, node.kind === 'text' ? node.text || '空白文字' : node.kind]), textNodes[0]!.mathId), text = field('公式组件文字', textNodes[0]!.kind === 'text' ? textNodes[0]!.text : '');
                target.addEventListener('change', () => {
                    const node = textNodes.find(node => node.mathId === target.value)!;
                    text.value = node.kind === 'text' ? node.text : '';
                }, { signal });
                button('更新公式文字组件', () => {
                    const nodes = JSON.parse(JSON.stringify(equation!.nodes)) as WordMathNode[], node = nodes.find(node => node.mathId === target.value)!;
                    if (node.kind === 'text')
                        node.text = text.value;
                    applyMath({ rootMathId: equation!.rootMathId, nodes });
                });
            }
            button('删除公式', () => {
                apply([{ op: 'removeWordEquation', nodeId: equation!.nodeId }]);
                selected = undefined;
                close();
            });
        }
    }
    return { open, bind(root, signal) {
            root.addEventListener('dblclick', event => {
                const equation = (event.target as Element).closest<HTMLElement>('[data-word-equation]');
                if (!equation)
                    return;
                event.preventDefault();
                event.stopImmediatePropagation();
                selected = equation.dataset.nodeId;
                try {
                    open(false);
                }
                catch (error) {
                    onError(error);
                }
            }, { signal, capture: true });
        }, dispose: close };
}

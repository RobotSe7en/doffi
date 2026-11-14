import type { TextFormatState } from './text-format-state.js';
/** Font roles and language are authored values; installed preview substitutes
 * are never used to rewrite the source faces. */
export function createWordFontControls(parent: ShadowRoot, before: HTMLElement, read: () => TextFormatState, apply: (properties: ReadonlyMap<string, string>) => void, onError: (error: unknown) => void): {
    open(): void;
    dispose(): void;
} {
    const document = before.ownerDocument;
    let panel: HTMLElement | undefined, controller: AbortController | undefined;
    const close = (): void => { controller?.abort(); panel?.remove(); panel = undefined; };
    function open(): void {
        close();
        controller = new AbortController();
        const signal = controller.signal, state = read();
        panel = document.createElement('section');
        panel.setAttribute('role', 'region');
        panel.setAttribute('aria-label', '字符格式');
        panel.style.cssText = 'display:flex;flex-wrap:wrap;gap:8px;padding:12px;background:white;border:1px solid #dbe3ef';
        parent.insertBefore(panel, before);
        const button = (label: string, run: () => void): void => { const button = document.createElement('button'); button.type = 'button'; button.textContent = label; button.addEventListener('click', () => { try {
            run();
        }
        catch (error) {
            onError(error);
        } }, { signal }); panel!.append(button); };
        const field = (name: string, type = 'text'): HTMLInputElement => { const label = document.createElement('label'), input = document.createElement('input'); label.textContent = name; input.type = type; input.setAttribute('aria-label', name); input.placeholder = '保留'; input.style.width = '110px'; if (type === 'number')
            input.step = '.5'; label.append(input); panel!.append(label); return input; };
        const choose = (name: string, options: readonly (readonly [
            string,
            string
        ])[]): HTMLSelectElement => { const label = document.createElement('label'), input = document.createElement('select'); label.textContent = name; input.setAttribute('aria-label', name); for (const [value, text] of [['', '保留'], ...options]) {
            const option = document.createElement('option');
            option.value = value!;
            option.textContent = text!;
            input.append(option);
        } label.append(input); panel!.append(label); return input; };
        button('关闭字符格式', close);
        const hint = document.createElement('span');
        hint.textContent = '留空保留原设置；未选中文字时用于后续输入。';
        panel.append(hint);
        const family = field('统一字体'), size = field('字符字号（磅）', 'number'), roles = (['ascii', 'hAnsi', 'eastAsia', 'cs'] as const).map((role, index) => ({ role, input: field(['ASCII 字体', '高位 ANSI 字体', '东亚字体', '复杂文字字体'][index]!) }));
        size.min = '.5';
        size.max = '1638';
        const weight = choose('字符字重', [['400', '常规'], ['700', '加粗']]), italic = choose('字符倾斜', [['normal', '正体'], ['italic', '斜体']]), underline = choose('字符下划线', [['none', '无'], ['solid', '单线'], ['double', '双线'], ['dotted', '点线'], ['dashed', '虚线'], ['wavy', '波浪线']]), strike = choose('字符删除线', [['none', '无'], ['line-through', '删除线']]), vertical = choose('字符位置', [['baseline', '基线'], ['super', '上标'], ['sub', '下标']]), caps = choose('字符大小写', [['none', '正常'], ['uppercase', '全部大写'], ['small-caps', '小型大写']]);
        const spacing = field('字符间距（磅）', 'number'), color = field('字符颜色'), background = field('字符底纹'), complexSize = field('复杂文字字号（磅）', 'number'), complexBold = choose('复杂文字字重', [['0', '常规'], ['1', '加粗']]), complexItalic = choose('复杂文字倾斜', [['0', '正体'], ['1', '斜体']]), kern = field('字距调整起始字号（磅）', 'number');
        complexSize.min = '.5';
        kern.min = '0';
        const languages = (['val', 'eastAsia', 'bidi'] as const).map((role, index) => ({ role, input: field(['西文语言', '东亚语言', '复杂文字语言'][index]!) }));
        button('应用字符格式', () => {
            const properties = new Map<string, string>();
            if (family.value.trim()) {
                const font = JSON.stringify(family.value.trim());
                properties.set('font-family', font);
                for (const role of ['ascii', 'hAnsi', 'eastAsia', 'cs'])
                    properties.set('--word-font-' + role, font);
            }
            if (size.value)
                properties.set('font-size', size.value + 'pt');
            for (const { role, input } of roles)
                if (input.value.trim())
                    properties.set('--word-font-' + role, JSON.stringify(input.value.trim()));
            if (weight.value)
                properties.set('font-weight', weight.value);
            if (italic.value)
                properties.set('font-style', italic.value);
            if (vertical.value)
                properties.set('vertical-align', vertical.value);
            if (spacing.value)
                properties.set('letter-spacing', spacing.value + 'pt');
            if (color.value.trim())
                properties.set('color', color.value.trim());
            if (background.value.trim())
                properties.set('background-color', background.value.trim());
            if (underline.value || strike.value) {
                const lines = [...(underline.value ? underline.value === 'none' ? [] : ['underline'] : state.underline ? ['underline'] : []), ...(strike.value ? strike.value === 'none' ? [] : ['line-through'] : state.strikethrough ? ['line-through'] : [])];
                properties.set('text-decoration-line', lines.join(' ') || 'none');
                if (underline.value && underline.value !== 'none')
                    properties.set('text-decoration-style', underline.value);
            }
            if (caps.value) {
                properties.set('font-variant', caps.value === 'small-caps' ? 'small-caps' : 'normal');
                properties.set('text-transform', caps.value === 'uppercase' ? 'uppercase' : 'none');
            }
            if (complexSize.value)
                properties.set('--word-font-size-cs', complexSize.value + 'pt');
            if (complexBold.value)
                properties.set('--word-bold-cs', complexBold.value);
            if (complexItalic.value)
                properties.set('--word-italic-cs', complexItalic.value);
            if (kern.value)
                properties.set('--word-kerning-threshold', kern.value + 'pt');
            for (const { role, input } of languages)
                if (input.value.trim())
                    properties.set('--word-language-' + role, JSON.stringify(input.value.trim()));
            if (properties.size)
                apply(properties);
        });
    }
    return { open, dispose: close };
}

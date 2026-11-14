import { WordControlSchema } from '@doffi/core';
import type { WordControl, DeepReadonly } from '@doffi/core';
import { W, attr, child, elements, check, checkAttributes, number, fail, unsupported } from './ooxml/xml.js';
import type { XmlNode } from './ooxml/xml.js';
import { attributes, empty } from './export/xml.js';
export const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const on = (node: XmlNode | undefined, namespace = W): boolean => { if (!node)
    return false; const value = attr(node, 'val', namespace); if (value === undefined)
    return true; if (['1', 'true', 'on'].includes(value))
    return true; if (['0', 'false', 'off'].includes(value))
    return false; return fail(node, 'A content control Boolean is invalid.'); };
export function readWordControl(source: XmlNode, level: 'inline' | 'block'): Omit<WordControl, 'nodeId'> {
    const properties = child(source, 'sdtPr');
    if(properties)checkAttributes(properties, [], W);
    if (child(properties, 'dataBinding') || child(properties, 'docPartObj') || child(properties, 'docPartList'))
        unsupported(source, 'This data-bound or building-block content control is preserved natively.');
    for (const node of properties?.content ?? []) {
        if (typeof node === 'string') {
            if (node.trim())
                fail(properties!, 'Unexpected control property text.');
            continue;
        }
        if (!(node.uri === W && ['rPr', 'alias', 'tag', 'id', 'lock', 'placeholder', 'temporary', 'showingPlcHdr', 'label', 'tabIndex', 'richText', 'text', 'comboBox', 'dropDownList', 'date', 'picture', 'equation', 'group', 'citation', 'bibliography'].includes(node.name) || node.uri === W14 && node.name === 'checkbox'))
            unsupported(node, 'This content control property is preserved natively.');
    }
    const kinds = elements(properties, undefined, W).filter(node => ['richText', 'text', 'comboBox', 'dropDownList', 'date', 'picture', 'equation', 'group', 'citation', 'bibliography'].includes(node.name)), checkbox = child(properties, 'checkbox', W14);
    if (kinds.length + (checkbox ? 1 : 0) > 1)
        fail(source, 'A content control cannot declare multiple types.');
    const type = checkbox ?? kinds[0], kind = checkbox ? 'checkBox' : type?.name === 'dropDownList' ? 'dropDown' : type?.name ?? 'richText';
    for (const item of elements(properties, undefined, W))
        if (['alias', 'tag', 'id', 'lock', 'label', 'tabIndex', 'temporary', 'showingPlcHdr'].includes(item.name)) {
            checkAttributes(item, ['val'], W);
            check(item, []);
        }
    const control: Record<string, unknown> = { nodeId: 'control', kind, level };
    for (const name of ['alias', 'tag', 'lock']) {
        const value = attr(child(properties, name), 'val');
        if (value !== undefined)
            control[name] = value;
    }
    const id = attr(child(properties, 'id'), 'val');
    if (id !== undefined) {
        number(id, properties!);
        control.sourceId = id;
    }
    if (child(properties, 'temporary'))
        control.temporary = on(child(properties, 'temporary'));
    if (child(properties, 'showingPlcHdr'))
        control.showingPlaceholder = on(child(properties, 'showingPlcHdr'));
    const placeholder = attr(child(child(properties, 'placeholder'), 'docPart'), 'val');
    if (placeholder !== undefined)
        control.placeholder = placeholder;
    if (kind === 'text') {
        if(type)checkAttributes(type, ['multiLine'], W);
        check(type, []);
        if (type && attr(type, 'multiLine') !== undefined) {
            const value = attr(type, 'multiLine')!;
            if (!['1', '0', 'true', 'false', 'on', 'off'].includes(value))
                fail(type, 'The multiline flag is invalid.');
            control.multiline = ['1', 'true', 'on'].includes(value);
        }
    }
    if (checkbox) {
        checkAttributes(checkbox, [], W14);
        const checked=child(checkbox,'checked',W14);if(checked)checkAttributes(checked,['val'],W14);
        check(checkbox, ['checked', 'checkedState', 'uncheckedState'], W14);
        control.checked = on(child(checkbox, 'checked', W14), W14);
        for (const side of ['checked', 'unchecked']) {
            const state = child(checkbox, side + 'State', W14);
            if (!state)
                continue;
            checkAttributes(state, ['val', 'font'], W14);
            const code = attr(state, 'val', W14), font = attr(state, 'font', W14);
            if (code !== undefined) {
                if (!/^[0-9a-f]{4}$/i.test(code))
                    fail(state, 'A checkbox symbol requires four hexadecimal digits.');
                control[side + 'Symbol'] = parseInt(code, 16);
            }
            if (font !== undefined)
                control[side + 'Font'] = font;
        }
    }
    if (kind === 'dropDown' || kind === 'comboBox') {
        if(type)checkAttributes(type, ['lastValue'], W);
        check(type, ['listItem']);
        control.items = elements(type, 'listItem').map(item => { checkAttributes(item, ['displayText', 'value'], W); return { displayText: attr(item, 'displayText') ?? attr(item, 'value') ?? '', value: attr(item, 'value') ?? attr(item, 'displayText') ?? '' }; });
        const value = attr(type, 'lastValue');
        if (value !== undefined)
            control.value = value;
    }
    if (kind === 'date') {
        if(type)checkAttributes(type, ['fullDate'], W);
        check(type, ['dateFormat', 'lid', 'storeMappedDataAs', 'calendar']);
        const fullDate = attr(type, 'fullDate');
        if (fullDate !== undefined)
            control.fullDate = fullDate;
        for (const [sourceName, name] of [['dateFormat', 'dateFormat'], ['lid', 'locale'], ['storeMappedDataAs', 'dateMapping'], ['calendar', 'calendar']] as const) {
            const value = attr(child(type, sourceName), 'val');
            if (value !== undefined)
                control[name] = value;
        }
    }
    // End properties and property-only flags are residuals bound to this wrapper.
    if (child(properties, 'label'))
        unsupported(child(properties, 'label')!, 'The native control label is retained.');
    if (child(properties, 'tabIndex'))
        unsupported(child(properties, 'tabIndex')!, 'The native control tab order is retained.');
    if (child(source, 'sdtEndPr'))
        unsupported(child(source, 'sdtEndPr')!, 'Native end-run properties are retained.');
    const parsed = WordControlSchema.safeParse(control);
    if (!parsed.success)
        unsupported(source, 'This content control is outside the editable model.');
    const { nodeId: _id, ...result } = parsed.data;
    return result;
}
export function writeWordControlProperties(control: DeepReadonly<WordControl>, id: number, runProperties = ''): string {
    const val = (name: string, value: string | number | undefined): string => value === undefined ? '' : empty('w:' + name, { 'w:val': value });
    let result = '<w:sdtPr>' + runProperties + val('alias', control.alias) + val('tag', control.tag) + val('id', id) + val('lock', control.lock) + (control.placeholder ? '<w:placeholder>' + val('docPart', control.placeholder) + '</w:placeholder>' : '') + (control.temporary === undefined ? '' : val('temporary', control.temporary ? 1 : 0)) + (control.showingPlaceholder === undefined ? '' : val('showingPlcHdr', control.showingPlaceholder ? 1 : 0));
    if (control.kind === 'checkBox')
        result += '<w14:checkbox>' + empty('w14:checked', { 'w14:val': control.checked ? 1 : 0 }) + (['checked', 'unchecked'] as const).map(side => { const code = side === 'checked' ? control.checkedSymbol : control.uncheckedSymbol, font = side === 'checked' ? control.checkedFont : control.uncheckedFont; return code !== undefined || font !== undefined ? empty('w14:' + side + 'State', { 'w14:val': (code ?? (side === 'checked' ? 0x2612 : 0x2610)).toString(16).toUpperCase().padStart(4, '0'), 'w14:font': font }) : ''; }).join('') + '</w14:checkbox>';
    else if (control.kind === 'dropDown' || control.kind === 'comboBox') {
        const name = control.kind === 'dropDown' ? 'dropDownList' : 'comboBox';
        result += '<w:' + name + attributes({ 'w:lastValue': control.value }) + '>' + (control.items ?? []).map(item => empty('w:listItem', { 'w:displayText': item.displayText, 'w:value': item.value })).join('') + '</w:' + name + '>';
    }
    else if (control.kind === 'date')
        result += '<w:date' + attributes({ 'w:fullDate': control.fullDate }) + '>' + val('dateFormat', control.dateFormat) + val('lid', control.locale) + val('storeMappedDataAs', control.dateMapping) + val('calendar', control.calendar) + '</w:date>';
    else
        result += empty('w:' + control.kind, control.kind === 'text' && control.multiline !== undefined ? { 'w:multiLine': control.multiline ? 1 : 0 } : {});
    return result + '</w:sdtPr>';
}

import { wordControlDate } from '../word-control.js';
import { DocumentError } from '../errors.js';
import type { WordMetadata } from '../contracts/word.js';
import type { WordControl, WordControlOperation } from '../contracts/word-control.js';
import type { DocumentEdit } from '../document/edit.js';
import type { ParsedHtml } from '../document/html.js';
function fail(message: string): never { throw new DocumentError('InvalidOperation', message); }
const escape = (text: string): string => text.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replace(/\r\n?/g, '\n').replaceAll('\n', '<br>');
const symbol = (control: WordControl): string => String.fromCharCode(control.checked ? (control.checkedSymbol ?? 0x2612) : (control.uncheckedSymbol ?? 0x2610));
export function isWordControlOperation(operation: {
    readonly op: string;
}): operation is WordControlOperation { return ['insertWordControl', 'setWordControl', 'setWordControlValue', 'setWordControlPicture', 'removeWordControl'].includes(operation.op); }
export function applyWordControl(word: WordMetadata, edit: DocumentEdit, parse: () => ParsedHtml, operation: WordControlOperation): void {
    const source = parse(), node = source.peekNode(operation.nodeId);
    if (!node)
        fail('The content control source is missing.');
    const controls = word.controls ??= [];
    const writeValue = (control: WordControl, value: string | boolean): void => {
        if (control.kind === 'checkBox') {
            if (typeof value !== 'boolean')
                fail('A checkbox value must be Boolean.');
            control.checked = value;
            edit.replaceSourceContent(control.nodeId, escape(symbol(control)));
            const font = control.checked ? control.checkedFont : control.uncheckedFont;
            if (font) {
                edit.sourceStyle(control.nodeId, 'font-family', JSON.stringify(font));
                for(const role of ['ascii','hAnsi','eastAsia','cs'])edit.sourceStyle(control.nodeId,'--word-font-'+role,JSON.stringify(font));
            }
        }
        else {
            if (typeof value !== 'string')
                fail('This content control requires text.');
            if (control.kind === 'dropDown') {
                const item = control.items?.find(item => item.value === value);
                if (!item)
                    fail('Choose a declared dropdown value.');
                control.value = value;
                value = item.displayText;
            }
            else if (control.kind === 'comboBox') {
                control.value = value;
                value = control.items?.find(item => item.value === value)?.displayText ?? value;
            }
            else if (control.kind === 'date') {
                const date = wordControlDate(value, control);
                control.fullDate = date.fullDate;
                value = date.text;
            }
            else if (!['text', 'richText', 'citation', 'bibliography', 'group', 'equation'].includes(control.kind))
                fail('Use the picture control operation to replace an image.');
            edit.replaceSourceContent(control.nodeId, control.level === 'block' ? '<p>' + escape(value) + '</p>' : escape(value));
        }
        control.showingPlaceholder = false;
    };
    if (operation.op === 'insertWordControl') {
        if (!/^p$|^h[1-6]$/.test(node.tagName))
            fail('Insert a content control beside a source paragraph or at its text offset.');
        const control = { ...operation.control, nodeId: '' }, tag = control.level === 'block' ? 'div' : 'span', content = operation.html ?? (control.level === 'block' ? '<p>' + escape(operation.text ?? '') + '</p>' : escape(operation.text ?? ''));
        const html = '<' + tag + ' data-word-content-control="' + control.kind + '">' + content + '</' + tag + '>';
        let id: string;
        if (control.level === 'inline')
            id = edit.insertSourceAt(operation.nodeId, operation.offset ?? 0, html).rootNodeIds[0]!;
        else {
            const parent = node.parentNodeId!, siblings = source.peekNode(parent)!.childNodeIds;
            id = edit.insertSource(parent, html, siblings[siblings.indexOf(node.nodeId) + 1]).rootNodeIds[0]!;
        }
        control.nodeId = id;
        controls.push(control);
        if (control.kind === 'checkBox')
            writeValue(control, control.checked ?? false);
        else if (control.kind === 'dropDown' && control.items?.length)
            writeValue(control, control.value ?? control.items[0]!.value);
        else if (control.kind === 'date' && control.fullDate)
            writeValue(control, control.fullDate);
        return;
    }
    const control = controls.find(control => control.nodeId === operation.nodeId);
    if (!control)
        fail('The content control definition is missing.');
    if (operation.op === 'removeWordControl') {
        if (operation.preserveContent !== false)
            edit.unwrapSource(control.nodeId);
        else
            edit.removeSource(control.nodeId);
        word.controls = controls.filter(item => item !== control);
        return;
    }
    if (operation.op === 'setWordControlValue') {
        writeValue(control, operation.value);
        return;
    }
    if (operation.op === 'setWordControlPicture') {
        if (control.kind !== 'picture')
            fail('Image selection requires a picture content control.');
        edit.replaceSourceContent(control.nodeId, (control.level === 'block' ? '<p>' : '') + '<img data-word-drawing="image" src="' + escape(edit.resourceReference(operation.assetPath)) + '" alt="' + escape(operation.alt ?? '') + '">' + (control.level === 'block' ? '</p>' : ''));
        const container = parse().peekNode(control.nodeId)!, id = control.level === 'block' ? parse().peekNode(container.childNodeIds[0]!)!.childNodeIds[0]! : container.childNodeIds[0]!;
        (word.drawings ??= []).push({ nodeId: id, kind: 'image', geometry: operation.geometry ?? { x: 0, y: 0, width: 180, height: 120, rotation: 0, flipH: false, flipV: false }, layout: { mode: 'inline' } });
        control.showingPlaceholder = false;
        return;
    }
    const previous={...control};
    Object.assign(control, operation.changes);
    edit.attributes(control.nodeId, { 'data-word-content-control': control.kind });
    if (control.kind === 'checkBox' && ['checked', 'checkedSymbol', 'uncheckedSymbol', 'checkedFont', 'uncheckedFont'].some(key => Reflect.has(operation.changes,key)&&Reflect.get(previous,key)!==Reflect.get(control,key)))
        writeValue(control, control.checked ?? false);
    else if (control.kind === 'date' && control.fullDate && (previous.dateFormat!==control.dateFormat||previous.locale!==control.locale))
        writeValue(control, control.fullDate);
    else if (operation.changes.items && JSON.stringify(previous.items)!==JSON.stringify(control.items) && (control.kind === 'dropDown' || control.kind === 'comboBox') && control.value !== undefined)
        writeValue(control, control.items?.some(item => item.value === control.value) ? control.value : control.items?.[0]?.value ?? '');
    else if (operation.changes.value !== undefined)
        writeValue(control, operation.changes.value);
    else if (operation.changes.fullDate !== undefined)
        writeValue(control, operation.changes.fullDate);
}

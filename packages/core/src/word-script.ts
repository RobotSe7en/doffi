export type WordScriptRole = 'ascii' | 'hAnsi' | 'eastAsia' | 'cs';
const complex = /[\p{Script=Arabic}\p{Script=Hebrew}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Devanagari}\p{Script=Bengali}\p{Script=Gurmukhi}\p{Script=Gujarati}\p{Script=Oriya}\p{Script=Tamil}\p{Script=Telugu}\p{Script=Kannada}\p{Script=Malayalam}\p{Script=Sinhala}\p{Script=Thai}\p{Script=Lao}\p{Script=Tibetan}\p{Script=Myanmar}\p{Script=Khmer}\p{Script=Mongolian}]/u;
/** Combining marks, joiners and variation sequences stay with their base role. */
export function wordScriptRole(character: string, previous?: WordScriptRole): WordScriptRole {
    if (previous && (/[\p{Mark}\u200c\u200d]/u.test(character)))
        return previous;
    const code = character.codePointAt(0)!;
    if (/[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Bopomofo}]/u.test(character) || code >= 0x2e80 && code <= 0xa4cf || code >= 0xf900 && code <= 0xfaff || code >= 0xfe10 && code <= 0xfe6f || code >= 0xff00 && code <= 0xffef || code >= 0x20000 && code <= 0x3ffff)
        return 'eastAsia';
    return code <= 0x7f ? 'ascii' : complex.test(character) ? 'cs' : 'hAnsi';
}

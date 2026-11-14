import { DocumentError } from './errors.js';
import type { DeepReadonly } from './runtime.js';
import type { WordControl } from './contracts/word-control.js';
export function wordControlDate(value: string, control: DeepReadonly<WordControl>): {
    fullDate: string;
    text: string;
} {
    const match = /^(\d{4})-(\d\d)-(\d\d)(?:T\d\d:\d\d:\d\d(?:\.\d+)?(?:Z|[+-]\d\d:\d\d))?$/.exec(value);
    const fail = (): never => { throw new DocumentError('InvalidOperation', 'A date control requires a valid ISO date and a supported calendar.'); };
    if (!match || !Number.isFinite(Date.parse(value)))
        fail();
    const date = new Date(match![1] + '-' + match![2] + '-' + match![3] + 'T00:00:00Z');
    if (date.getUTCFullYear() !== Number(match![1]) || date.getUTCMonth() + 1 !== Number(match![2]) || date.getUTCDate() !== Number(match![3]))
        fail();
    if (control.calendar && !['gregorian', 'gregorianUs', 'gregorianMeFrench', 'gregorianArabic', 'gregorianXlitEnglish', 'gregorianXlitFrench'].includes(control.calendar))
        fail();
    const locales: Readonly<Record<string, string>> = { '1033': 'en-US', '2052': 'zh-CN', '1028': 'zh-TW', '1041': 'ja-JP', '1031': 'de-DE', '1036': 'fr-FR' }, locale = locales[control.locale ?? ''] ?? control.locale ?? 'zh-CN';
    const part = (options: Intl.DateTimeFormatOptions): string => { try {
        return new Intl.DateTimeFormat(locale, { ...options, timeZone: 'UTC', calendar: 'gregory' }).format(date);
    }
    catch {
        return new Intl.DateTimeFormat('zh-CN', { ...options, timeZone: 'UTC', calendar: 'gregory' }).format(date);
    } };
    const tokens: Record<string, string> = { yyyy: match![1]!, yy: match![1]!.slice(-2), MMMM: part({ month: 'long' }), MMM: part({ month: 'short' }), MM: match![2]!, M: String(Number(match![2])), dddd: part({ weekday: 'long' }), ddd: part({ weekday: 'short' }), dd: match![3]!, d: String(Number(match![3])), gg: part({ era: 'short', year: 'numeric' }).replace(/[\d\s.,]/g, '') };
    let text = '', format = control.dateFormat ?? 'yyyy-MM-dd';
    for (let i = 0; i < format.length;) {
        const char = format[i]!;
        if (char === '"' || char === "'") {
            const quote = char;
            i++;
            while (i < format.length) {
                if (format[i] === quote) {
                    i++;
                    break;
                }
                text += format[i++];
            }
        }
        else if (char === '\\') {
            text += format[++i] ?? '';
            i++;
        }
        else {
            const token = /^(yyyy|dddd|MMMM|ddd|MMM|yy|MM|dd|gg|M|d)/.exec(format.slice(i))?.[0];
            if (token) {
                text += tokens[token];
                i += token.length;
            }
            else
                text += format[i++];
        }
    }
    return { fullDate: value.includes('T') ? value : value + 'T00:00:00Z', text };
}

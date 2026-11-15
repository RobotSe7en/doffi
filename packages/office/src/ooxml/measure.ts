import {fail,number} from './xml.js';
import type {XmlNode} from './xml.js';
/** OOXML integer units and UniversalMeasure literals share one point value. */
export function measurePoints(raw:string|undefined,node:XmlNode,units=20,fallback?:number):number{
    if(raw===undefined&&fallback!==undefined)return fallback;
    if(raw!==undefined&&/^-?\d+$/.test(raw))return number(raw,node)/units;
    const match=raw&&/^([+-]?(?:\d+(?:\.\d+)?|\.\d+))(mm|cm|in|pt|pc|pi)$/.exec(raw);
    if(!match)fail(node,'An OOXML measurement is invalid.');
    const value=Number(match[1])*({mm:72/25.4,cm:72/2.54,in:72,pt:1,pc:12,pi:12}as Record<string,number>)[match[2]!]!;
    if(!Number.isFinite(value)||Math.abs(value)>Number.MAX_SAFE_INTEGER)fail(node,'An OOXML measurement exceeds its numeric range.');
    return value;
}

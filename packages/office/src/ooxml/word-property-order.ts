import { W } from './xml.js';
import type { XmlNode } from './xml.js';

/** ECMA-376 Transitional CT_PPrBase followed by CT_PPr's own children. */
export const wordParagraphPropertyOrder = ['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange'] as const;
const paragraphPositions=new Map<string,number>(wordParagraphPropertyOrder.map((name,index)=>[name,index]));
const rowPositions=new Map(['cnfStyle','divId','gridBefore','gridAfter','wBefore','wAfter','cantSplit','trHeight','tblHeader','tblCellSpacing','jc','hidden','ins','del','trPrChange'].map((name,index)=>[name,index]));
const runPositions=new Map(['ins','del','moveFrom','moveTo','rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline','shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden','color','spacing','w','kern','position','sz','szCs','highlight','u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em','lang','eastAsianLayout','specVanish','oMath','rPrChange'].map((name,index)=>[name,index]));
/** Native preservation may add a property absent from the compiled output.
 * Reorder standard children in their existing slots; extension nodes stay put. */
export function sortWordProperties(root:XmlNode):void {
    const positions=root.uri===W?(root.name==='pPr'?paragraphPositions:root.name==='trPr'?rowPositions:root.name==='rPr'?runPositions:undefined):undefined;
    if(positions){
        const known=root.content.filter((node):node is XmlNode=>typeof node!=='string'&&node.uri===W&&positions.has(node.name)).sort((a,b)=>positions.get(a.name)!-positions.get(b.name)!);
        let index=0;for(let slot=0;slot<root.content.length;slot++){const node=root.content[slot]!;if(typeof node!=='string'&&node.uri===W&&positions.has(node.name))root.content[slot]=known[index++]!;}
    }
    for(const child of root.content)if(typeof child!=='string')sortWordProperties(child);
}

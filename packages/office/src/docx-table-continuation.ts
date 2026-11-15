import type { DeepReadonly, WordTableEmptyParagraph, SourceStyleNode } from '@doffi/core';
import type { ExportCss } from './export/css.js';
import type { DocxStyles, Css } from './ooxml/styles.js';
import { mergeTabs } from './ooxml/styles.js';
import { W, attr, child, elements, number, text, unsupported } from './ooxml/xml.js';
import type { XmlNode } from './ooxml/xml.js';
import { writeWordConditionalFormat, readWordConditionalFormat } from './docx-conditional-format.js';
import { paragraphProperties, runProperties } from './export/properties.js';
import { empty, orderedProperties, value } from './export/xml.js';
import { nativeInteger } from './export/values.js';

/** Preserve empty continuation formatting without inventing source text,
 * editable cells or a second native document tree. Nonempty content must still
 * use the normal compatibility recovery path. */
export function readWordEmptyContinuation(cell: XmlNode, styles: DocxStyles, compat: (node: XmlNode) => XmlNode[], register: (css: Css) => string | undefined, defaultTabInterval=36): WordTableEmptyParagraph[] {
    const result: WordTableEmptyParagraph[] = [];
    for (const paragraph of compat(cell)) {
        if (paragraph.uri === W && paragraph.name === 'tcPr') continue;
        if (paragraph.uri !== W || paragraph.name !== 'p') unsupported(paragraph, 'A vertical merge continuation has independent content.');
        const properties = child(paragraph, 'pPr');
        if (paragraph.attrs.size || properties?.attrs.size) unsupported(paragraph, 'An empty continuation paragraph has unmapped attributes.');
        for (const name of ['pPrChange','sectPr']) if (child(properties, name)) unsupported(child(properties,name)!, 'An empty continuation paragraph has independent structure or history.');
        const css = styles.paragraph(properties), run = styles.run(child(properties,'rPr'));
        for (const marker of elements(child(properties,'rPr'))) if (['ins','del','moveFrom','moveTo','rPrChange'].includes(marker.name)) unsupported(marker,'An empty continuation paragraph has independent mark history.');
        const definition: WordTableEmptyParagraph = { runs: [], defaultTabInterval }, styleId = attr(child(properties,'pStyle'),'val');
        const inherited=styles.get(styleId ?? styles.defaultParagraph);
        if (styleId) definition.styleId = styleId;
        const paragraphClass = register(css), runClass = register(run);
        if (paragraphClass) definition.paragraphCssClass = paragraphClass;
        if (runClass) definition.runCssClass = runClass;
        const format = readWordConditionalFormat(properties); if(format)definition.format=format;
        const outline = child(properties,'outlineLvl'); if(outline)definition.outlineLevel=number(attr(outline,'val'),outline);
        const tabs = child(properties,'tabs'); if(tabs){
            mergeTabs([],tabs); // Validate each native declaration, including CLEAR.
            definition.tabs=elements(tabs,'tab').map(tab=>({position:number(attr(tab,'pos'),tab)/20,alignment:(attr(tab,'val')??'left')as NonNullable<WordTableEmptyParagraph['tabs']>[number]['alignment'],leader:(attr(tab,'leader')??'none')as NonNullable<WordTableEmptyParagraph['tabs']>[number]['leader']}));
        }
        const resolved=mergeTabs(inherited.tabs??styles.defaults.tabs,tabs);
        if(resolved.length||tabs){definition.resolvedTabs=resolved;}
        const numPr=child(properties,'numPr'),numId=attr(child(numPr,'numId'),'val');
        if(numId==='0')definition.numberingDisabled=true;
        else if(numId){const level=number(attr(child(numPr,'ilvl'),'val'),numPr!,0);if(!styles.numbering.some(item=>item.numberingId===numId&&item.levels.some(item=>item.level===level)))unsupported(numPr!,'An empty continuation paragraph has missing numbering.');definition.numbering={numberingId:numId,level};}
        for (const source of compat(paragraph)) {
            if (source.uri === W && source.name === 'pPr') continue;
            if (source.uri !== W || source.name !== 'r') unsupported(source, 'A vertical merge continuation has independent content.');
            const runProperties=child(source,'rPr');
            if(source.attrs.size||runProperties?.attrs.size)unsupported(source,'An empty continuation run has unmapped attributes.');
            if(child(runProperties,'rPrChange'))unsupported(child(runProperties,'rPrChange')!,'An empty continuation run has independent history.');
            const current:WordTableEmptyParagraph['runs'][number]={},cssClass=register(styles.run(runProperties)),character=attr(child(runProperties,'rStyle'),'val');
            if(cssClass)current.cssClass=cssClass;
            if(character){styles.get(character);current.styleId=character;}
            for(const content of compat(source)){
                if(content.uri===W&&content.name==='rPr')continue;
                if(content.uri!==W||content.name!=='t'||content.content.some(item=>typeof item!=='string')||text(content)!=='')unsupported(content,'A vertical merge continuation has independent content.');
            }
            definition.runs.push(current);
        }
        result.push(definition);
    }
    return result;
}

export function writeWordEmptyContinuation(paragraph: DeepReadonly<WordTableEmptyParagraph>, resolver: ExportCss, parent: SourceStyleNode, styles: ReadonlyMap<string,string>, numbering: ReadonlyMap<string,string>): string {
    const css=(name:string):Readonly<Css>=>{
        const keys=Object.keys(resolver.specifiedClass(name)),computed=resolver.classStyle(name,resolver.get(parent));
        return Object.fromEntries(Object.entries(computed).filter(([key])=>keys.some(name=>key===name||/^(border(?:-(?:top|right|bottom|left))?|margin|padding|text-decoration)$/.test(name)&&key.startsWith(name+'-'))));
    };
    const direct = paragraph.paragraphCssClass ? paragraphProperties(css(paragraph.paragraphCssClass)) : '', run = paragraph.runCssClass ? runProperties(css(paragraph.runCssClass)) : '';
    const tabs = paragraph.tabs?.map(tab=>empty('w:tab',{'w:val':tab.alignment,'w:leader':tab.leader,'w:pos':nativeInteger(tab.position)})).join('');
    const num=paragraph.numbering, list=num||paragraph.numberingDisabled?'<w:numPr>'+ (num?value('ilvl',num.level):'')+value('numId',num?numbering.get(num.numberingId)!:0)+'</w:numPr>':'';
    const properties = (paragraph.styleId?value('pStyle',styles.get(paragraph.styleId)!):'')+direct+writeWordConditionalFormat(paragraph.format)+(tabs?'<w:tabs>'+tabs+'</w:tabs>':'')+list+(paragraph.outlineLevel!==undefined?value('outlineLvl',paragraph.outlineLevel):'')+(run?'<w:rPr>'+run+'</w:rPr>':'');
    return '<w:p>'+(properties?'<w:pPr>'+orderedProperties('pPr',properties)+'</w:pPr>':'')+paragraph.runs.map(source=>{const props=(source.styleId?value('rStyle',styles.get(source.styleId)!):'')+(source.cssClass?runProperties(css(source.cssClass)):'');return '<w:r>'+(props?'<w:rPr>'+orderedProperties('rPr',props)+'</w:rPr>':'')+'</w:r>';}).join('')+'</w:p>';
}

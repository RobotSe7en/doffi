import { z } from 'zod';
import { IdentifierSchema } from './contracts/primitives.js';
import { PageSetupSchema } from './contracts/document.js';
import { DocumentError } from './errors.js';
import { openDocument } from './document-model.js';
import { encode } from './document/package.js';
import { checkLimit, readLimits } from './document/validation.js';
import type { DocumentPackage, OpenDocumentOptions } from './runtime.js';
import type { PageSetup, WordDocumentMetadata } from './contracts/document.js';
const optionsSchema=z.strictObject({documentId:IdentifierSchema,title:z.string().optional(),locale:IdentifierSchema.optional(),pageSetup:PageSetupSchema.optional()});
export interface CreateWordDocumentOptions { readonly documentId:string; readonly title?:string; readonly locale?:string; readonly pageSetup?:PageSetup }
const escape=(value:string):string=>value.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
/** Create an isolated Word package; all subsequent edits use the shared session. */
export function createWordDocument(input:CreateWordDocumentOptions,readOptions:OpenDocumentOptions={}):DocumentPackage{
  const parsed=optionsSchema.safeParse(input);if(!parsed.success)throw new DocumentError('InvalidDocument','Invalid new Word document options.');
  const options=parsed.data,limits=readLimits(readOptions.limits),title=options.title??'未命名文档',locale=options.locale??'zh-CN';
  checkLimit((title.length+locale.length+options.documentId.length)*6,limits.maxFileBytes);
  for(const value of [title,locale])if(/[\u0000\ud800-\udfff]/u.test(value))throw new DocumentError('InvalidDocument','Word document text must be well-formed Unicode without null characters.');
  const metadata:WordDocumentMetadata={title,locale,pageSetup:options.pageSetup??{unit:'pt',width:595.2755905511812,height:841.8897637795276,margins:{top:72,right:72,bottom:72,left:72}},sections:[{sectionId:'section-1',rootNodeId:'section-root-1',breakType:'nextPage'}],word:{stories:[{storyId:'main',kind:'main',rootNodeId:'main-story'}],styles:[{styleId:'Normal',default:true,kind:'paragraph',cssClass:'word-normal'},{styleId:'Heading1',outlineLevel:0,kind:'paragraph',cssClass:'word-heading-1',basedOn:'Normal',nextStyleId:'Normal'},{styleId:'Heading2',outlineLevel:1,kind:'paragraph',cssClass:'word-heading-2',basedOn:'Normal',nextStyleId:'Normal'}],numbering:[],hyperlinks:[],bookmarks:[],fields:[],tabs:[],revisions:[],tables:[]}};
  const files:DocumentPackage=new Map([
    ['manifest.json',encode(JSON.stringify({format:'doffi',formatVersion:1,documentId:options.documentId,revision:0,kind:'document',entrypoint:'content.html',styles:['styles.css'],metadata:'metadata.json',annotations:'annotations.json',assets:[]}))],
    ['metadata.json',encode(JSON.stringify(metadata))],['annotations.json',encode('[]')],
    ['content.html',encode('<!doctype html><html lang="'+escape(locale)+'"><head><meta charset="utf-8"><title>'+escape(title)+'</title><link rel="stylesheet" href="styles.css"></head><body><article data-node-id="document-root" data-office-kind="document"><div data-node-id="main-story" data-office-story="main"><section data-node-id="section-root-1"><p data-node-id="paragraph-1" data-word-style="Normal" class="word-normal"></p></section></div></article></body></html>')],
    ['styles.css',encode('.word-normal{font-family:"Noto Serif CJK SC",serif;--word-font-ascii:"Liberation Serif";--word-font-hAnsi:"Liberation Serif";--word-font-eastAsia:"Noto Serif CJK SC";font-size:12pt;line-height:1.5;white-space:pre-wrap;margin:0}.word-heading-1{font-size:18pt;font-weight:700;break-after:avoid}.word-heading-2{font-size:16pt;font-weight:700;break-after:avoid}')],
  ]);
  return openDocument(files,{...readOptions,kind:'document'}).serialize();
}

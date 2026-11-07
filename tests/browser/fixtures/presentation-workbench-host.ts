import { openContentFixture as openDocument } from "../../content-fixture.js";
import { decodeTestPackage, type TestPackage } from './package-transport.js';
import { createDocumentSession } from "@doffi/core";
import type { AssetDescriptor,DeepReadonly,DocumentSession,PresentationEditProposal,PresentationInteractionRequest } from "@doffi/core";
import { mountDocumentWorkbench } from "@doffi/editor";
import type { PresentationWorkbenchView } from "@doffi/editor";
import { createElement } from "react";
import { createRoot } from "react-dom/client";
import type { Root } from "react-dom/client";
import { PresentationWorkbench } from "@doffi/editor/react";
export interface PresentationWorkbenchHost {
  session:DocumentSession|undefined;
  view:PresentationWorkbenchView|undefined;
  errors:string[];
  submission:"fixed"|"deferred";
  complete:(()=>void)|undefined;
  mount(files:TestPackage,mode:"js"|"react",resources?:{descriptor:AssetDescriptor;bytes:number[]}[]):Promise<void>;
  dispose():void;
}
declare global{interface Window{presentationWorkbenchHost:PresentationWorkbenchHost}}
const container=document.createElement("div");container.style.width="1500px";document.body.replaceChildren(container);let root:Root|undefined;
function fixed(request:DeepReadonly<PresentationInteractionRequest>):PresentationEditProposal{return{protocolVersion:1,requestId:request.requestId,proposalId:"fixed",documentId:request.documentId,baseRevision:request.baseRevision,annotationIds:request.annotations.map(a=>a.annotationId),summary:"修改幻灯片标题",operations:[{op:"replaceText",nodeId:"heading-p",from:0,to:4,expectedText:"项目进展",text:"提案标题"}]};}
window.presentationWorkbenchHost={session:undefined,view:undefined,errors:[],submission:"fixed",complete:undefined,
  dispose(){root?.unmount();root=undefined;this.view?.dispose();this.view=undefined;this.session?.dispose();this.session=undefined;},
  async mount(files,mode,resources=[]){
    this.dispose();this.errors=[];this.submission="fixed";this.complete=undefined;this.session=createDocumentSession(openDocument(decodeTestPackage(files)));
    const onError=(error:Error):void=>{this.errors.push(error.message);},onSubmitAnnotations=async(request:DeepReadonly<PresentationInteractionRequest>):Promise<PresentationEditProposal>=>this.submission==="fixed"?fixed(request):new Promise(resolve=>{this.complete=()=>resolve(fixed(request));});
    const options={onError,onSubmitAnnotations,...(resources.length?{proposalResources:{assets:resources.map(r=>({descriptor:r.descriptor,bytes:new Uint8Array(r.bytes)}))}}:{})};
    if(mode==="js")this.view=mountDocumentWorkbench(container,this.session,{...options,format:"presentation"});
    else{root=createRoot(container);await new Promise<void>(resolve=>{root!.render(createElement(PresentationWorkbench,{session:this.session!,...options,onViewChange:view=>{this.view=view;if(view)resolve();}}));});}
    await this.view!.ready;
  },
};

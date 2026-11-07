import { openContentFixture as openDocument } from "../../content-fixture.js";
import { decodeTestPackage, type TestPackage } from './package-transport.js';
import { createDocumentSession } from "@doffi/core";
import type { DocumentSession } from "@doffi/core";
import { mountPresentationEditor } from "@doffi/editor";
import type { PresentationEditorView, PresentationEditorOptions } from "@doffi/editor";
export interface PresentationEditorHost {
  session:DocumentSession|undefined;
  view:PresentationEditorView|undefined;
  errors:string[];
  mount(files:TestPackage,options?:PresentationEditorOptions):Promise<void>;
}
declare global {interface Window{presentationEditorHost:PresentationEditorHost}}
const container=document.createElement("div");container.style.width="1500px";document.body.replaceChildren(container);
window.presentationEditorHost={
  session:undefined,view:undefined,errors:[],
  async mount(files,options={}) {
    this.view?.dispose();this.session?.dispose();this.errors=[];
    this.session=createDocumentSession(openDocument(decodeTestPackage(files)));
    this.view=mountPresentationEditor(container,this.session,{...options,onError:error=>{this.errors.push(error.message);}});await this.view.ready;
  },
};

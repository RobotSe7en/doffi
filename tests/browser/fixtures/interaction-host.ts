import { openContentFixture as openDocument } from "../../content-fixture.js";
import { decodeTestPackage, type TestPackage } from './package-transport.js';
import { createDocumentSession } from "@doffi/core";
import type { DocumentEditingSession, DocumentSelection } from "@doffi/core";
import { mountDocumentInteraction, mountDocumentPreview } from "@doffi/editor";
import type { DocumentInteractionView, DocumentPreview } from "@doffi/editor";

/** A separate plain JavaScript host for public API/lifecycle browser acceptance. */
export interface InteractionHost {
  session: DocumentEditingSession | undefined;
  view: DocumentInteractionView | undefined;
  preview: DocumentPreview | undefined;
  selections: (DocumentSelection | undefined)[];
  errors: string[];
  nodeReads: string[];
  mount(files: TestPackage, readOnly?: boolean, appearance?: "default" | "workspace"): Promise<void>;
  settle(): Promise<void>;
}
declare global { interface Window { interactionHost: InteractionHost } }
const container = document.createElement("div");
container.style.cssText = "width:1000px;margin:20px";
document.body.replaceChildren(container);
window.interactionHost = {
  session: undefined, view: undefined, preview: undefined, selections: [], errors: [], nodeReads: [],
  async mount(files, readOnly = false, appearance = "default") {
    this.view?.dispose();
    this.view = undefined;
    this.preview?.dispose();
    this.preview = undefined;
    this.session?.dispose();
    const session = createDocumentSession(openDocument(decodeTestPackage(files)));
    this.session = { ...session, getNode: (id) => { this.nodeReads.push(id); return session.getNode(id); } };
    if (readOnly) this.preview = mountDocumentPreview(container, this.session);
    else this.view = mountDocumentInteraction(container, this.session, {
      appearance,
      onSelectionChange: (selection) => this.selections.push(selection), onError: (error) => this.errors.push(error.message),
    });
    await this.settle();
  },
  async settle() {
    await Promise.resolve();
    await this.view?.ready;
    await this.preview?.ready;
    await new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
  },
};

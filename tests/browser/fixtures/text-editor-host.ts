import { openContentFixture as openDocument } from "../../content-fixture.js";
import { decodeTestPackage, type TestPackage } from './package-transport.js';
import { createDocumentSession } from "@doffi/core";
import type { DocumentSession } from "@doffi/core";
import { mountDocumentTextEditor } from "@doffi/editor";
import type { DocumentTextEditor } from "@doffi/editor";
export interface TextEditorHost {
  session: DocumentSession | undefined;
  view: DocumentTextEditor | undefined;
  errors: string[];
  mount(files: TestPackage, beforeReadyText?: string): Promise<void>;
}
declare global { interface Window { textEditorHost: TextEditorHost } }
const container = document.createElement("div"); container.style.cssText = "width:1050px;margin:20px"; document.body.replaceChildren(container);
window.textEditorHost = {
  session: undefined, view: undefined, errors: [],
  async mount(files, beforeReadyText) {
    this.view?.dispose(); this.session?.dispose();
    this.session = createDocumentSession(openDocument(decodeTestPackage(files)));
    this.view = mountDocumentTextEditor(container, this.session, { onError: (error) => this.errors.push(error.message) });
    if (beforeReadyText) this.session.commit(this.session.prepare({ protocolVersion: 1,
      documentId: this.session.snapshot().manifest.documentId, baseRevision: this.session.snapshot().manifest.revision,
      operations: [{ op: "replaceText", nodeId: "a", from: 0, to: 0, expectedText: "", text: beforeReadyText }] }));
    await this.view.ready;
  },
};

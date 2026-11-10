import { openContentFixture as openDocument } from "../../content-fixture.js";
import { decodeTestPackage, type TestPackage } from './package-transport.js';
import { createDocumentSession } from "@doffi/core";
import type { DocumentInteractionRequest, DocumentSession, EditProposal } from "@doffi/core";
import { mountDocumentWorkbench } from "@doffi/editor";
import type { DocumentWorkbenchView } from "@doffi/editor";

export interface WorkbenchHost {
  session: DocumentSession | undefined;
  view: DocumentWorkbenchView | undefined;
  requests: DocumentInteractionRequest[];
  events: string[];
  errors: string[];
  failSubmission: boolean;
  deferSubmission: boolean;
  release: (() => void) | undefined;
  submissionProposal: EditProposal | undefined;
  mount(files: TestPackage, maxCandidates?: number): Promise<void>;
  settle(): Promise<void>;
}
declare global { interface Window { workbenchHost: WorkbenchHost } }
const container = document.createElement("div"); container.style.cssText = "width:1000px;margin:20px";
document.body.replaceChildren(container);
window.workbenchHost = {
  session: undefined, view: undefined, requests: [], events: [], errors: [], failSubmission: false, deferSubmission: false, release: undefined, submissionProposal: undefined,
  async mount(files, maxCandidates = 8) {
    this.view?.dispose(); this.session?.dispose();
    this.session = createDocumentSession(openDocument(decodeTestPackage(files)), { sessionLimits: { maxCandidates } });
    this.view = mountDocumentWorkbench(container, this.session, {
      onSubmitAnnotations: async (request) => {
        if (this.failSubmission) throw new Error("Private host failure");
        this.requests.push(request as DocumentInteractionRequest);
        if (this.submissionProposal) this.view!.loadProposal(request as DocumentInteractionRequest,
          { ...this.submissionProposal, requestId: request.requestId, documentId: request.documentId, baseRevision: request.baseRevision });
        if (this.deferSubmission) await new Promise<void>((resolve) => { this.release = resolve; });
      },
      onEvent: (event) => this.events.push(event.type), onError: (error) => this.errors.push(error.message),
    });
    await this.settle();
  },
  async settle() { await Promise.resolve(); await this.view?.ready; await new Promise<void>((resolve) => requestAnimationFrame(() => resolve())); },
};

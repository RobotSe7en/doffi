import { decodeTestPackage, type TestPackage } from './package-transport.js';
import { openContentFixture as openDocument } from "../../content-fixture.js";
import { createDocumentSession } from "@doffi/core";
import type { ContentDocumentModel as DocumentModel, DocumentSession } from "@doffi/core";
import { createLayoutFieldTransaction, mountDocumentInteraction, mountDocumentPreview, mountDocumentTextEditor } from "@doffi/editor";
import type { DocumentInteractionView, DocumentPreview, DocumentPreviewOptions, DocumentRenderReport, DocumentTextEditor } from "@doffi/editor";
import { SourceProjection as ProjectionUnderTest } from '../../../packages/render/src/projection.js';
import { markViewNode } from '../../../packages/render/src/view-nodes.js';

export interface PaginationHost {
  model: DocumentModel | undefined;
  session: DocumentSession | undefined;
  view: DocumentPreview | DocumentInteractionView | DocumentTextEditor | undefined;
  interaction: DocumentInteractionView | undefined;
  errors: string[];
  mount(files: TestPackage, mode?: "preview" | "edit" | "interaction", options?: DocumentPreviewOptions): Promise<void>;
  layout(): DocumentRenderReport;
  updateFields(): Promise<void>;
  lateProjection(): { original: string; html: string; copy: string; shellMissing: boolean; unowned: string; instance: string };
}
declare global { interface Window { paginationHost: PaginationHost } }
const container = document.createElement("div"); container.style.cssText = "width:1050px;margin:20px"; document.body.replaceChildren(container);
window.paginationHost = {
  model: undefined, session: undefined, view: undefined, interaction: undefined, errors: [],
  async mount(files, mode = "preview", options = {}) {
    this.view?.dispose(); this.session?.dispose(); this.session = undefined; this.interaction = undefined; this.errors = [];
    this.model = openDocument(decodeTestPackage(files));
    if (mode === "preview") this.view = mountDocumentPreview(container, this.model, options);
    else {
      this.session = createDocumentSession(this.model);
      const input = { ...options, onError: (error: Error) => this.errors.push(error.message) };
      this.view = mode === "edit" ? mountDocumentTextEditor(container, this.session, input) : this.interaction = mountDocumentInteraction(container, this.session, input);
    }
    await this.view.ready;
  },
  layout() { const result = this.view?.layout(); if (!result) throw new Error("Layout is not ready."); return result; },
  async updateFields() {
    const session = this.session!;
    session.commit(session.prepare(createLayoutFieldTransaction(session, this.layout())));
    await Promise.resolve(); await this.view!.ready;
  },
  lateProjection() {
    const source = document.createElement('p'); source.dataset.nodeId = 'p'; source.innerHTML = '<span data-node-id="run" style="color:#123456">原文</span>';
    const root = source.cloneNode(true) as HTMLElement, projection = new ProjectionUnderTest(); projection.capture(root, source);
    const wrapper = markViewNode(document.createElement('span'), 'unwrap'); wrapper.className = 'sdk-inline-flow'; wrapper.style.display = 'block'; wrapper.append(...root.childNodes); root.append(wrapper);
    const html = projection.serialize(root).outerHTML, range = document.createRange(); range.selectNodeContents(wrapper);
    const holder = document.createElement('div'); holder.append(projection.serializeRange(range));
    const unowned = document.createElement('span'); unowned.className = 'sdk-inline-flow'; unowned.style.display = 'block'; unowned.textContent = '源内容'; root.append(unowned);
    const instance = markViewNode(source.cloneNode(true) as HTMLElement, 'unwrap'), bound = new ProjectionUnderTest(); bound.captureInstance(instance, source);
    return { original: source.outerHTML, html, copy: holder.innerHTML, shellMissing: projection.shell(wrapper) === undefined, unowned: projection.serialize(root).outerHTML, instance: bound.serialize(instance).outerHTML };
  },
};

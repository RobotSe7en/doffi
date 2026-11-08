import { createElement, useEffect, useRef } from "react";
import type { ReactElement } from "react";
import type { DocumentEditingSession, DocumentSession, WorkbookSession } from "@doffi/core";
import { mountWorkbookWorkbench } from "./workbook/workbench.js";
import type { WorkbookWorkbenchOptions, WorkbookWorkbenchView } from "./workbook/workbench.js";
import type { WorkbookSessionConnection } from "./workbook/session-transport.js";

export interface WorkbookWorkbenchProps extends WorkbookWorkbenchOptions {
  readonly session: WorkbookSession | WorkbookSessionConnection;
  readonly className?: string;
  readonly onViewChange?: (view: WorkbookWorkbenchView | undefined) => void;
}
/** React owns the container; the ordinary JS workbench owns view cleanup. */
export function WorkbookWorkbench(props: WorkbookWorkbenchProps): ReactElement {
  const container = useRef<HTMLDivElement>(null), latest = useRef(props); latest.current = props;
  const hasSubmission = Boolean(props.onSubmitAnnotations);
  useEffect(() => {
    if (!container.current) return;
    const view = mountWorkbookWorkbench(container.current, props.session, {
      ...(props.title !== undefined ? { title: props.title } : {}), ...(props.widthPt !== undefined ? { widthPt: props.widthPt } : {}), ...(props.heightPt !== undefined ? { heightPt: props.heightPt } : {}),
      ...(props.locale !== undefined ? { locale: props.locale } : {}), ...(props.limits ? { limits: props.limits } : {}), ...(props.scope ? { scope: props.scope } : {}), ...(props.allowGlobalScope !== undefined ? { allowGlobalScope: props.allowGlobalScope } : {}),
      ...(props.readRanges ? { readRanges: props.readRanges } : {}), ...(props.supportedOperations ? { supportedOperations: props.supportedOperations } : {}), ...(props.contextLimits ? { contextLimits: props.contextLimits } : {}),
      ...(props.prepareOptions ? { prepareOptions: () => latest.current.prepareOptions!() } : {}),
      ...(props.confirmValidationWarnings ? { confirmValidationWarnings: confirmation => latest.current.confirmValidationWarnings!(confirmation) } : {}),
      ...(hasSubmission ? { onSubmitAnnotations: (request, context) => latest.current.onSubmitAnnotations!(request, context) } : {}),
      onError: error => latest.current.onError?.(error), onSelectionChange: selection => latest.current.onSelectionChange?.(selection), onPendingChange: pending => latest.current.onPendingChange?.(pending), onEvent: event => latest.current.onEvent?.(event),
    });
    try { latest.current.onViewChange?.(view); } catch (error) { view.dispose(); throw error; }
    void view.ready.catch(() => { /* The ordinary workbench reports errors. */ });
    return () => { view.dispose(); latest.current.onViewChange?.(undefined); };
  }, [props.session, props.title, props.widthPt, props.heightPt, props.locale, props.limits, props.scope, props.allowGlobalScope, props.readRanges, props.supportedOperations, props.contextLimits, Boolean(props.prepareOptions), Boolean(props.confirmValidationWarnings), hasSubmission]);
  return createElement("div", { ref: container, className: props.className });
}
import { mountDocumentWorkbench } from "./workbench.js";
import type { DocumentWorkbenchOptions, DocumentWorkbenchView } from "./workbench.js";
import { mountPresentationEditor } from "./presentation-editor.js";
import type { PresentationEditorOptions, PresentationEditorView } from "./presentation-editor.js";
import { mountPresentationWorkbench } from "./presentation-workbench.js";
import type { PresentationWorkbenchOptions,PresentationWorkbenchView } from "./presentation-workbench.js";

export interface PresentationWorkbenchProps extends PresentationWorkbenchOptions {
  readonly session:DocumentSession;
  readonly className?:string;
  readonly onViewChange?:(view:PresentationWorkbenchView|undefined)=>void;
}
/** Optional React wrapper delegates to the ordinary JS PPT workbench. */
export function PresentationWorkbench(props:PresentationWorkbenchProps):ReactElement {
  const container=useRef<HTMLDivElement>(null),latest=useRef(props);latest.current=props;const hasSubmission=!!props.onSubmitAnnotations;
  useEffect(()=>{
    if(!container.current)return;
    const view=mountPresentationWorkbench(container.current,props.session,{
      ...(props.title?{title:props.title}:{}),...(props.layout?{layout:props.layout}:{}),...(props.limits?{limits:props.limits}:{}),...(props.presentationSurfaceId?{presentationSurfaceId:props.presentationSurfaceId}:{}),...(props.appearance?{appearance:props.appearance}:{}),
      ...(props.supportedOperations?{supportedOperations:props.supportedOperations}:{}),...(props.contextLimits?{contextLimits:props.contextLimits}:{}),...(props.proposalResources?{proposalResources:props.proposalResources}:{}),
      ...(hasSubmission?{onSubmitAnnotations:request=>latest.current.onSubmitAnnotations!(request)}:{}),
      onSelectionChange:selection=>latest.current.onSelectionChange?.(selection),onError:error=>latest.current.onError?.(error),onInputRecovery:text=>latest.current.onInputRecovery?.(text),onEvent:event=>latest.current.onEvent?.(event),
    });
    try{latest.current.onViewChange?.(view);}catch(error){view.dispose();throw error;}
    void view.ready.catch(()=>{/* The JS workbench reports rendering failures. */});
    return()=>{view.dispose();latest.current.onViewChange?.(undefined);};
  },[props.session,props.title,props.layout,props.limits,props.presentationSurfaceId,props.supportedOperations,props.contextLimits,props.proposalResources,props.appearance,hasSubmission]);
  return createElement("div",{ref:container,className:props.className});
}

export interface PresentationEditorProps extends PresentationEditorOptions {
  readonly session:DocumentEditingSession;
  readonly className?:string;
  readonly onViewChange?:(view:PresentationEditorView|undefined)=>void;
}
export function PresentationEditor(props:PresentationEditorProps):ReactElement {
  const container=useRef<HTMLDivElement>(null),latest=useRef(props);latest.current=props;
  useEffect(()=>{
    if(!container.current)return;
    const view=mountPresentationEditor(container.current,props.session,{
      ...(props.title?{title:props.title}:{}),...(props.layout?{layout:props.layout}:{}),...(props.limits?{limits:props.limits}:{}),
      onSelectionChange:selection=>latest.current.onSelectionChange?.(selection),onError:error=>latest.current.onError?.(error),onInputRecovery:text=>latest.current.onInputRecovery?.(text),
      ...(props.presentationSurfaceId?{presentationSurfaceId:props.presentationSurfaceId}:{}),
      ...(props.appearance?{appearance:props.appearance}:{}),
    });
    try{latest.current.onViewChange?.(view);}catch(error){view.dispose();throw error;}
    void view.ready.catch(()=>{/* The mounted editor forwards rendering errors through onError. */});
    return()=>{view.dispose();latest.current.onViewChange?.(undefined);};
  },[props.session,props.title,props.layout,props.limits,props.presentationSurfaceId,props.appearance]);
  return createElement("div",{ref:container,className:props.className});
}

export interface DocumentWorkbenchProps extends DocumentWorkbenchOptions {
  readonly session: DocumentSession;
  readonly className?: string;
  readonly onViewChange?: (view: DocumentWorkbenchView | undefined) => void;
}

/** Optional React entry; ordinary JavaScript consumers need no React installation. */
export function DocumentWorkbench(props: DocumentWorkbenchProps): ReactElement {
  const container = useRef<HTMLDivElement>(null), latest = useRef(props);
  latest.current = props;
  const hasSubmission = Boolean(props.onSubmitAnnotations);
  useEffect(() => {
    if (!container.current) return;
    const workbenchOptions:DocumentWorkbenchOptions={
      ...(props.editable !== undefined ? { editable: props.editable } : {}),
      ...(props.appearance ? { appearance: props.appearance } : {}),
      ...(props.title ? { title: props.title } : {}),
      ...(props.layout ? { layout: props.layout } : {}), ...(props.limits ? { limits: props.limits } : {}),
      ...(props.supportedOperations ? { supportedOperations: props.supportedOperations } : {}),
      ...(props.contextLimits ? { contextLimits: props.contextLimits } : {}),
      ...(hasSubmission ? { onSubmitAnnotations: (request) => latest.current.onSubmitAnnotations?.(request) } : {}),
      onSelectionChange: (selection) => latest.current.onSelectionChange?.(selection),
      onError: (error) => latest.current.onError?.(error),
      onEvent: (event) => latest.current.onEvent?.(event),
    };
    const view=mountDocumentWorkbench(container.current,props.session,workbenchOptions);
    latest.current.onViewChange?.(view);
    let active = true;
    void view.ready.catch((error: unknown) => { if (active) latest.current.onError?.(error instanceof Error ? error : new Error("Preview failed.")); });
    return () => { active = false; view.dispose(); latest.current.onViewChange?.(undefined); };
  }, [props.session, props.title, props.layout, props.limits, props.supportedOperations, props.contextLimits, props.editable, props.appearance, hasSubmission]);
  return createElement("div", { ref: container, className: props.className });
}

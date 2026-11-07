import { openContentFixture as openDocument } from "../../../tests/content-fixture.js";
import {expect,test} from "vitest";
import { createDocumentSession, createPresentation } from "@doffi/core";
import type {DocumentOperation,DocumentSession} from "@doffi/core";
const blank=()=>createDocumentSession(openDocument(createPresentation({documentId:"comments"})));
const p=(s:DocumentSession)=>s.snapshot().metadata.presentation!;
function commit(s:DocumentSession,operations:DocumentOperation[]){const{documentId,revision}=s.snapshot().manifest;return s.commit(s.prepare({protocolVersion:1,documentId,baseRevision:revision,operations}));}
test.each(["classic","modern"]as const)("%s comments preserve rich bodies, replies, format, status and one-step history",format=>{
  const s=blank();commit(s,[{op:"addPresentationComment",anchor:{kind:"node",nodeId:"slide-root-1"},html:'<p><span style="font-weight:700">中文批注</span></p>',format,author:{name:"甲"},date:"2026-11-01T00:00:00Z"}]);const annotation=s.getAnnotations()[0]!,body=annotation.bodyNodeId!;expect(s.getNode(body)!.text).toBe("中文批注");expect(p(s).comments[0]!.format).toBe(format);
  commit(s,[{op:"addPresentationCommentReply",annotationId:annotation.annotationId,html:"<p>回复𠮷󠄀</p>",author:{name:"乙"}},{op:"updatePresentationComment",annotationId:annotation.annotationId,changes:{status:"resolved"}}]);const reply=s.getAnnotations()[0]!.replies![0]!;expect(s.getNode(reply.bodyNodeId!)!.text).toBe("回复𠮷󠄀");expect(s.getAnnotations()[0]!.status).toBe("resolved");expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
  commit(s,[{op:"setPresentationCommentBody",annotationId:annotation.annotationId,replyId:reply.replyId,html:"<p>更新回复</p>"}]);expect(s.getNode(reply.bodyNodeId!)!.text).toBe("更新回复");s.undo();expect(s.getNode(reply.bodyNodeId!)!.text).toBe("回复𠮷󠄀");
  commit(s,[{op:"removePresentationCommentReply",annotationId:annotation.annotationId,replyId:reply.replyId}]);expect(s.getNode(reply.bodyNodeId!)).toBeUndefined();expect(s.getAnnotations()[0]!.replies).toBeUndefined();s.undo();expect(s.getNode(reply.bodyNodeId!)!.text).toBe("回复𠮷󠄀");
  commit(s,[{op:"removeAnnotation",annotationId:annotation.annotationId}]);expect(s.getNode(body)).toBeUndefined();expect(s.getNode(reply.bodyNodeId!)).toBeUndefined();expect(p(s).comments).toEqual([]);s.undo();expect(s.getAnnotations()[0]!.status).toBe("resolved");
});
test("comment recovery retains workflow and rich text, requires an explicit new anchor and cleans newly removed bodies",()=>{
  const s=blank();commit(s,[{op:"insertSlide",layoutId:"layout-body"}]);const slide=p(s).slides[1]!,node=s.getNode(slide.rootNodeId)!.childNodeIds[0]!;
  commit(s,[{op:"addPresentationComment",anchor:{kind:"node",nodeId:node},html:"<p>正文</p>"}]);const id=s.getAnnotations()[0]!.annotationId;commit(s,[{op:"updatePresentationComment",annotationId:id,changes:{status:"resolved"}},{op:"removeSlide",slideId:slide.slideId}]);expect(s.getAnnotations()[0]!.status).toBe("needs_reanchor");expect(p(s).comments[0]!.recoveryStatus).toBe("resolved");
  const before=s.snapshot();expect(()=>commit(s,[{op:"updatePresentationComment",annotationId:id,changes:{status:"open"}}])).toThrow();expect(s.snapshot()).toEqual(before);
  commit(s,[{op:"updatePresentationComment",annotationId:id,changes:{anchor:{kind:"node",nodeId:"slide-root-1"},status:"resolved"}}]);expect(s.getAnnotations()[0]!.status).toBe("resolved");expect(p(s).comments[0]!.recoveryStatus).toBeUndefined();
  commit(s,[{op:"addPresentationCommentReply",annotationId:id,html:"<p>同批删除</p>"},{op:"removeAnnotation",annotationId:id}]);expect(p(s).stories.filter(s=>s.kind==="comment")).toEqual([]);expect(openDocument(s.serialize()).snapshot()).toEqual(s.snapshot());
});

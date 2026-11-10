import { useState } from "react";
import type { DragEvent } from "react";
import { Icon } from "./icons.js";
import { brandLogo } from "./brand.js";

const samples = [
  { id: "docx", format: "DOCX", title: "文字与编号", detail: "段落样式、多级编号与 Word 文档预览", label: "打开 DOCX 样例", kind: "word" },
  { id: "xlsx", format: "XLSX", title: "一万行工作簿", detail: "10,000 行 · 20 万单元格 · 公式与网格编辑", label: "打开一万行 XLSX 样例", kind: "workbook" },
  { id: "pptx", format: "PPTX", title: "多页演示文稿", detail: "固定幻灯片画布、对象与原生文件编辑", label: "打开 PPTX 样例", kind: "presentation" },
] as const;

/** Decorative miniatures describe the formats; document content is rendered by the SDK. */
function Miniature({ kind }: { kind: "word" | "workbook" | "presentation" }) {
  return <div className={"sample-art " + kind} aria-hidden="true">
    <div className="mini-chrome"><i /><i /><i /><span /></div>
    {kind === "word" ? <div className="mini-paper"><b>实施报告</b><i /><i /><i /><strong>项目进展</strong><i /><i /><div className="mini-table" /><i /><i /></div>
      : kind === "workbook" ? <div className="mini-sheet"><div className="mini-cell" /><div className="mini-chart">{[35, 54, 43, 72, 61, 87, 76].map((height, index) => <i key={index} style={{ height: height + "%" }} />)}</div></div>
      : <div className="mini-slide"><span>PROJECT UPDATE</span><b>让想法<br />清晰呈现。</b><i /><small>项目进展演示</small></div>}
  </div>;
}

export function WelcomeScreen({ loading, onOpen, onChooseFile, onFile, onNewPresentation }: {
  loading: boolean; onOpen(id: string): void; onChooseFile(): void; onFile(file: File): void; onNewPresentation(): void;
}) {
  const [dragging, setDragging] = useState(false);
  function drop(event: DragEvent<HTMLDivElement>) {
    event.preventDefault(); setDragging(false);
    if (!loading && event.dataTransfer.files[0]) onFile(event.dataTransfer.files[0]);
  }
  return <div className="welcome">
    <header className="welcome-header"><a href="./" aria-label="doffi Playground"><img src={brandLogo} alt="" /><span>Playground</span></a><a className="source-link" href="https://github.com/RobotSe7en/doffi" target="_blank" rel="noreferrer">GitHub <Icon name="external" size={14} /></a></header>
    <section className="welcome-hero">
      <span className="welcome-eyebrow">DOFFI PLAYGROUND</span>
      <h1>Word、Excel、PPT，<br className="mobile-break" />在浏览器里打开。</h1>
      <p>预览文档，直接编辑，与智能体一起完善每个想法。</p>
      <div className="suite-preview">{samples.map(item => <Miniature key={item.id} kind={item.kind} />)}</div>
      <p className="hero-description">从一份样例开始，或打开自己的 Office 文件。<br />预览、编辑、批注与修改提案，在同一个工作台完成。</p>
      <div className="hero-actions"><a className="primary-button" href="#sample-documents">试用样例文档 <Icon name="arrow" size={16} /></a><button className="subtle-button" disabled={loading} onClick={onChooseFile}><Icon name="folder" size={17} />打开自己的文件</button></div>
    </section>
    <section className="welcome-samples" id="sample-documents"><div className="welcome-container">
      <div className="section-heading"><div><h2>样例文档</h2><p>选择一种格式，体验完整的文档工作台。</p></div><span>三种格式，一个工作空间</span></div>
      <div className="sample-cards">{samples.map(item => <button key={item.id} className={"sample-card " + item.kind} aria-label={item.label} disabled={loading} onClick={() => onOpen(item.id)}><Miniature kind={item.kind} /><div className="sample-card-body"><span className="format-badge">{item.format}</span><h3>{item.title}</h3><p>{item.detail}</p><span className="sample-open">打开文档 <Icon name="arrow" size={16} /></span></div></button>)}</div>
      <div className="collaboration-samples"><Icon name="comment" size={19} /><span>体验人和智能体的文档协作</span><button disabled={loading} aria-label="打开报告协作样例" onClick={() => onOpen("report")}>报告与批注 <Icon name="arrow" size={14} /></button><button disabled={loading} aria-label="打开 PPT 协作样例" onClick={() => onOpen("presentation")}>演示与提案 <Icon name="arrow" size={14} /></button><button disabled={loading} onClick={onNewPresentation}><Icon name="plus" size={14} />新建 PPT</button></div>
    </div></section>
    <section className="welcome-upload welcome-container"><h2>打开自己的文件</h2><p>支持 .docx、.xlsx、.pptx 和 doffi 文档包 .zip。</p><div className={"upload-zone" + (dragging ? " dragging" : "")} onDragOver={event => { event.preventDefault(); if (!loading) setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={drop}>
      <Icon name="upload" size={28} /><strong>将文件拖放到这里</strong><span>或从电脑中选择一份文档</span><button className="primary-button" disabled={loading} onClick={onChooseFile}>选择文件</button><small>最大 64 MiB</small>
    </div><div className="local-file-note"><Icon name="shield" size={20} /><p><strong>在浏览器中打开与编辑</strong><span>文档与修改保留在当前页面；刷新前请下载保存。PDF / PNG 导出使用本地转换服务。</span></p></div></section>
    <footer className="welcome-footer"><span>doffi · 人与智能体，共同编辑同一份文档。</span><span>AGPL-3.0-only / 商业许可</span></footer>
  </div>;
}

import { createElement as h, useState } from 'react';
import type { ReactElement } from 'react';
import { createRoot } from 'react-dom/client';
import { OfficeRibbon } from '@doffi/editor/react';
import type { TextFormatState } from '@doffi/editor';

function Example({ name }: { name: string }): ReactElement {
  const [mode, setMode] = useState<'edit' | 'preview'>('edit'), [tab, setTab] = useState('开始'), [opened, setOpened] = useState(0);
  const [format, setFormat] = useState<TextFormatState>({ enabled: true, fontFamily: 'Arial', fontSize: 12 }), [rulers, setRulers] = useState(true), [zoom, setZoom] = useState(1), [layout, setLayout] = useState<'pages' | 'continuous'>('pages');
  return h('section', { 'aria-label': name }, h('button', { onClick: () => { setTab('开始'); setMode(mode === 'edit' ? 'preview' : 'edit'); } }, mode === 'edit' ? '切换预览' : '切换编辑'),
    h(OfficeRibbon, { kind: 'document', mode, tab, onTab: setTab, busy: false, commands: [], onCommand() {}, actions: { open: { run: () => setOpened(value => value + 1) } }, formatState: format,
      onTextCommand(command) { if (command.type === 'format') setFormat(value => ({ ...value, ...(command.property === 'font-family' ? { fontFamily: command.value } : {}) })); },
      rulers, onRulers: setRulers, layout, onLayout: setLayout, zoom, onZoom: setZoom }), h('output', { 'aria-label': '打开次数' }, String(opened)));
}
const style = document.createElement('style'); style.textContent = 'button{background:rgb(255,0,0);font-size:21px}body{margin:0}section{margin-bottom:25px}';
document.head.append(style); document.body.replaceChildren();
const outside = document.createElement('button'); outside.id = 'outside'; outside.textContent = '宿主按钮'; document.body.append(outside);
const container = document.createElement('div'); document.body.append(container); createRoot(container).render(h('main', null, h(Example, { name: '文档一' }), h(Example, { name: '文档二' })));

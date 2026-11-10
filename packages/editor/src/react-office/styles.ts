export const officeRibbonStyles: string = `
.doffi-office-ribbon{font:13px/1.4 system-ui,sans-serif;color:#202938;min-width:0}
:where(.doffi-office-ribbon) button,:where(.doffi-office-ribbon) input,:where(.doffi-office-ribbon) select{font:inherit}
:where(.doffi-office-ribbon) button{display:inline-flex;align-items:center;justify-content:center;gap:7px;border:1px solid transparent;border-radius:5px;padding:7px 10px;background:transparent;color:inherit;cursor:pointer}
:where(.doffi-office-ribbon) button:disabled{cursor:default}
:where(.doffi-office-ribbon) input,:where(.doffi-office-ribbon) select{color:inherit}
:where(.doffi-office-ribbon) button:focus-visible,:where(.doffi-office-ribbon) input:focus-visible,:where(.doffi-office-ribbon) select:focus-visible{outline:2px solid #185abd;outline-offset:2px}
:where(.doffi-office-ribbon) *{box-sizing:border-box}
:where(.doffi-office-ribbon) [hidden]{display:none!important}
.doffi-office-ribbon{ min-width: 0; background: #f5f6f8; }
:where(.doffi-office-ribbon) .office-tabs{ display: flex; align-items: center; gap: 2px; min-height: 36px; padding: 0 13px; background: #f5f7fa; overflow-x: auto; scrollbar-width: thin; border-bottom: 1px solid #e2e6ed; }
:where(.doffi-office-ribbon) .office-tabs>button{ align-self: stretch; position: relative; padding: 8px 13px; border-radius: 0; border: 0; color: #3e4959; font-size: 12px; flex-shrink: 0; }
:where(.doffi-office-ribbon) .office-tabs>button:hover{ background: #e9edf4; }
:where(.doffi-office-ribbon) .office-tabs>button[aria-selected=true]{ color: #185abd; font-weight: 600; }
:where(.doffi-office-ribbon) .office-tabs>button[aria-selected=true]:after{ content: ""; position: absolute; bottom: 0; left: 13px; right: 13px; height: 2px; background: #185abd; border-radius: 2px; }
:where(.doffi-office-ribbon) .office-tabs>button:first-child{ color: #185abd; }
:where(.doffi-office-ribbon) .office-ribbon-note{ margin-left: auto; padding-left: 20px; padding-right: 7px; color: #9aa6b8; font-size: 10px; }
:where(.doffi-office-ribbon) .office-ribbon-panel{ display: flex; min-height: 107px; padding: 6px 8px 4px; margin: 0 7px 5px; border: 1px solid #e1e5ec; border-top: 0; border-radius: 0 0 8px 8px; background: #fff; box-shadow: 0 2px 3px #182d4d05; overflow-x: auto; scrollbar-width: thin; }
:where(.doffi-office-ribbon) .ribbon-group{ display: flex; position: relative; flex-direction: column; flex-shrink: 0; padding: 0 9px; border-right: 1px solid #e3e6ed; }
:where(.doffi-office-ribbon) .ribbon-group:last-child{ border-right: 0; }
:where(.doffi-office-ribbon) .ribbon-group-content{ display: flex; align-items: flex-start; justify-content: center; gap: 3px; min-height: 70px; flex: 1; }
:where(.doffi-office-ribbon) .ribbon-group-caption{ display: block; padding-top: 3px; text-align: center; color: #7f8897; font-size: 10px; line-height: 17px; }
:where(.doffi-office-ribbon) .ribbon-tool{ display: inline-flex; align-items: center; justify-content: flex-start; gap: 5px; padding: 4px; min-height: 23px; border: 1px solid transparent; border-radius: 3px; font-size: 11px; line-height: 1.3; color: #384656; }
:where(.doffi-office-ribbon) .ribbon-tool:hover:not(:disabled){ border-color: #c7d5ec; background: #eaf1fc; }
:where(.doffi-office-ribbon) .ribbon-tool:disabled{ opacity: .37; }
:where(.doffi-office-ribbon) .ribbon-tool>svg{ color: #506d94; }
:where(.doffi-office-ribbon) .ribbon-tool.large{ flex-direction: column; justify-content: flex-start; min-width: 46px; padding: 5px 6px 3px; gap: 7px; min-height: 67px; }
:where(.doffi-office-ribbon) .ribbon-tool.large>svg{ color: #5a7aab; }
:where(.doffi-office-ribbon) .ribbon-tool.large>span:last-child{ white-space: nowrap; font-size: 11px; }
:where(.doffi-office-ribbon) .ribbon-tool-stack{ display: flex; align-items: flex-start; flex-direction: column; gap: 0; }
:where(.doffi-office-ribbon) .clipboard-group>.ribbon-group-content{ gap: 4px; }
:where(.doffi-office-ribbon) .clipboard-group .large>svg{ color: #c38e4f; }
:where(.doffi-office-ribbon) .font-group>.ribbon-group-content,:where(.doffi-office-ribbon) .paragraph-group>.ribbon-group-content{ flex-direction: column; justify-content: flex-start; gap: 6px; }
:where(.doffi-office-ribbon) .ribbon-font-selects{ display: flex; align-items: center; gap: 3px; }
:where(.doffi-office-ribbon) .ribbon-font-selects select{ height: 25px; border: 1px solid #cdd4df; border-radius: 2px; background: #fff; padding: 2px 5px; color: #334056; font-size: 11px; }
:where(.doffi-office-ribbon) .ribbon-font-selects select:first-child{ width: 111px; }
:where(.doffi-office-ribbon) .ribbon-font-selects select:nth-child(2){ width: 46px; }
:where(.doffi-office-ribbon) .ribbon-font-selects select:disabled{ color: #929baa; border-color: #e0e5ed; background: #fafbfc; }
:where(.doffi-office-ribbon) .ribbon-icon-row{ display: flex; align-items: center; gap: 2px; }
:where(.doffi-office-ribbon) .ribbon-icon-row>.ribbon-tool,:where(.doffi-office-ribbon) .ribbon-font-selects>.ribbon-tool{ width: 27px; height: 25px; min-height: 25px; justify-content: center; padding: 3px; }
:where(.doffi-office-ribbon) .ribbon-icon-row>.ribbon-tool>span:last-child,:where(.doffi-office-ribbon) .ribbon-font-selects>.ribbon-tool>span:last-child{ display: none; }
:where(.doffi-office-ribbon) .ribbon-glyph{ display: inline-flex; align-items: center; justify-content: center; min-width: 17px; height: 20px; font-family: Georgia, serif; font-size: 16px; line-height: 20px; }
:where(.doffi-office-ribbon) .ribbon-glyph sup,:where(.doffi-office-ribbon) .ribbon-glyph sub{ font-size: 9px; }
:where(.doffi-office-ribbon) .ribbon-glyph .font-color{ border-bottom: 3px solid #1f4b99; height: 21px; }
:where(.doffi-office-ribbon) .ribbon-glyph .normal-weight{ font-size: 13px; font-family: system-ui; }
:where(.doffi-office-ribbon) .ribbon-divider{ display: inline-block; width: 1px; height: 19px; background: #e4e8f0; margin: 0 3px; }
:where(.doffi-office-ribbon) .style-gallery{ display: flex; gap: 4px; align-items: center; height: 68px; }
:where(.doffi-office-ribbon) .style-gallery>button{ display: flex; flex-direction: column; justify-content: center; gap: 6px; height: 62px; min-width: 65px; padding: 6px 9px; border: 1px solid #e3e7ef; border-radius: 2px; background: #fff; opacity: .5; }
:where(.doffi-office-ribbon) .style-gallery small{ color: #5d6f87; font-size: 9px; }
:where(.doffi-office-ribbon) .styles-group .ribbon-group-content{ align-items: center; }
:where(.doffi-office-ribbon) .ribbon-group-content>select{ min-width: 65px; align-self: center; border: 1px solid #d4dce8; border-radius: 3px; font-size: 11px; background: #fff; padding: 5px; }
:where(.doffi-office-ribbon) .ribbon-tool-stack>label{ display: flex; align-items: center; gap: 5px; margin: 3px 0; color: #7c8797; font-size: 11px; }
:where(.doffi-office-ribbon) .ribbon-tool-stack input[type=checkbox]{ width: 12px; height: 12px; margin: 1px; }
:where(.doffi-office-ribbon) .view-description{ color: #8e9caf; align-self: center; font-size: 10px; margin-left: 5px; }
:where(.doffi-office-ribbon) .pen-gallery{ display: flex; align-items: center; gap: 10px; }
:where(.doffi-office-ribbon) .pen-gallery>button{ border-bottom: 3px solid currentColor; margin-inline: 5px; padding: 6px 7px; }
:where(.doffi-office-ribbon) .ribbon-tool[aria-pressed=true]{ background: #e3edfc; border-color: #b7cdec; color: #174e9e; }
:where(.doffi-office-ribbon) .office-combo{ display:flex; position:relative; border:1px solid #cdd4df; border-radius:2px; background:#fff; height:25px; }
:where(.doffi-office-ribbon) .office-combo input{ border:0; border-radius:2px; min-width:0; padding:2px 5px; font-size:11px; width:100%; background:transparent; }
:where(.doffi-office-ribbon) .office-combo>button{ border:0; padding:0 4px; font-size:9px; color:#536171; background:transparent; border-radius:0; }
:where(.doffi-office-ribbon) .office-combo:has(input:disabled){ opacity:.5; background:#fafbfc; }
:where(.doffi-office-ribbon) .font-family-field{ width:132px; }
:where(.doffi-office-ribbon) .font-size-field{ width:58px; }
:where(.doffi-office-ribbon) .office-combo-options{ position:absolute; z-index:70; left:-1px; top:24px; min-width:0; max-height:260px; overflow:auto; border:1px solid #bec9d6; background:#fff; box-shadow:0 5px 16px #243f6928; }
:where(.doffi-office-ribbon) .office-combo-options>button{ display:block; width:100%; text-align:left; padding:5px 9px; white-space:nowrap; font-size:12px; border:0; border-radius:0; }
:where(.doffi-office-ribbon) .office-combo-options>button:hover,:where(.doffi-office-ribbon) .office-combo-options>button[aria-selected=true]{ background:#dceafd; }
:where(.doffi-office-ribbon) .office-combo-options>span{ display:block; padding:7px; font-size:11px; white-space:nowrap; color:#66768a; }
:where(.doffi-office-ribbon) .office-combo-error{ position:absolute; z-index:71; left:0; top:26px; width:220px; padding:7px; color:#9c3131; background:#fff1f0; border:1px solid #e6b7b2; font-size:11px; }
:where(.doffi-office-ribbon) .office-font-hint{ align-self:center; border-radius:50%; color:#8b6007; background:#fff2c9; font:bold 10px system-ui; padding:0 4px; margin-right:1px; cursor:help; }
:where(.doffi-office-ribbon) .office-combo-options{ max-width:min(280px,calc(100vw - 24px)); overflow-x:hidden; }
:where(.doffi-office-ribbon) .office-combo-options>button{ overflow:hidden; text-overflow:ellipsis; }

@media(max-width:1100px){:where(.doffi-office-ribbon) .office-tabs>button{padding-inline:10px}
:where(.doffi-office-ribbon) .office-tabs>button[aria-selected=true]:after{left:10px;right:10px}
}

@media(max-width:680px){:where(.doffi-office-ribbon) .office-tabs{padding-inline:6px;min-height:34px}
:where(.doffi-office-ribbon) .office-tabs>button{padding:8px 11px;font-size:11px}
:where(.doffi-office-ribbon) .office-ribbon-note{display:none}
:where(.doffi-office-ribbon) .office-ribbon-panel{margin-inline:4px}
:where(.doffi-office-ribbon) .ribbon-group{padding-inline:7px}
}

`;

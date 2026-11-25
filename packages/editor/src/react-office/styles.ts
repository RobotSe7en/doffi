/** Scoped chrome follows the compact macOS Office ribbon. */
export const officeRibbonStyles: string = `
.doffi-office-ribbon{--office-accent:#185abd;min-width:0;background:#f5f5f5;color:#292929;font:12px/1.35 -apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif}
.doffi-office-ribbon[data-kind=presentation]{--office-accent:#b7472a}.doffi-office-ribbon[data-kind=workbook]{--office-accent:#107c41}
:where(.doffi-office-ribbon) *{box-sizing:border-box}:where(.doffi-office-ribbon) [hidden]{display:none!important}
:where(.doffi-office-ribbon) button,:where(.doffi-office-ribbon) input,:where(.doffi-office-ribbon) select{font:inherit;color:inherit}
:where(.doffi-office-ribbon) button{display:inline-flex;align-items:center;justify-content:center;border:1px solid transparent;border-radius:4px;background:transparent;padding:3px;color:inherit;cursor:pointer;gap:3px;white-space:nowrap;transition:none}
:where(.doffi-office-ribbon) button:disabled{opacity:.33;cursor:default}:where(.doffi-office-ribbon) button:hover:not(:disabled){background:#e7e7e7}
:where(.doffi-office-ribbon) button:focus-visible,:where(.doffi-office-ribbon) input:focus-visible,:where(.doffi-office-ribbon) select:focus-visible{outline:2px solid #8badde;outline-offset:1px}
:where(.doffi-office-ribbon) .office-tabs{display:flex;align-items:center;gap:0;height:32px;min-height:32px;padding:0 10px;background:#f5f5f5;overflow-x:auto;scrollbar-width:none;border:0}
:where(.doffi-office-ribbon) .office-tabs>button{position:relative;align-self:stretch;flex-shrink:0;padding:5px 12px;border:0;border-radius:0;font-size:12px;color:#292929}
:where(.doffi-office-ribbon) .office-tabs>button[aria-selected=true]{font-weight:600;color:#292929}
:where(.doffi-office-ribbon) .office-tabs>button[aria-selected=true]:after{content:"";position:absolute;left:12px;right:12px;bottom:3px;height:3px;border-radius:2px;background:var(--office-accent)}
:where(.doffi-office-ribbon) .office-ribbon-panel{display:flex;align-items:stretch;height:74px;min-height:74px;padding:5px 8px;margin:0 7px 6px;border:1px solid #e8e8e8;border-radius:13px;background:#fff;box-shadow:0 2px 4px #00000012;overflow-x:auto;overflow-y:hidden;scrollbar-width:none}
:where(.doffi-office-ribbon) .ribbon-group{position:relative;display:flex;flex-direction:column;flex-shrink:0;padding:0 9px;border-right:1px solid #d9d9d9}
:where(.doffi-office-ribbon) .ribbon-group:last-child{border-right:0}
:where(.doffi-office-ribbon) .ribbon-group-content{display:flex;align-items:flex-start;justify-content:center;gap:4px;height:62px;min-height:0;flex:1}
:where(.doffi-office-ribbon) .ribbon-group-caption{display:none}
:where(.doffi-office-ribbon) .ribbon-tool{display:inline-flex;align-items:center;justify-content:flex-start;gap:4px;padding:3px;min-height:22px;border:1px solid transparent;border-radius:4px;font-size:11px;line-height:1.2;color:#333}
:where(.doffi-office-ribbon) .ribbon-tool>svg{color:#595959;flex-shrink:0}
:where(.doffi-office-ribbon) .ribbon-tool.large{flex-direction:column;justify-content:flex-start;min-width:43px;height:62px;padding:4px 3px 0;gap:4px}
:where(.doffi-office-ribbon) .ribbon-tool.large>span:last-child{font-size:11px;white-space:normal;line-height:1.15;max-width:60px;text-align:center}
:where(.doffi-office-ribbon) .ribbon-tool[aria-pressed=true]{background:#e5e5e5;border-color:transparent;color:#292929}
:where(.doffi-office-ribbon) .ribbon-tool-stack{display:flex;flex-direction:column;align-items:flex-start;gap:0;height:62px;justify-content:center}
:where(.doffi-office-ribbon) .ribbon-tool-stack>.ribbon-tool{min-height:20px;padding-block:1px}
:where(.doffi-office-ribbon) .clipboard-group{padding-inline:2px 9px}
:where(.doffi-office-ribbon) .clipboard-group .ribbon-tool-stack>.ribbon-tool{width:22px;height:20px;padding:1px}
:where(.doffi-office-ribbon) .clipboard-group .ribbon-tool-stack>.ribbon-tool>span:last-child{display:none}
:where(.doffi-office-ribbon) [data-icon=paste]>svg,:where(.doffi-office-ribbon) [data-icon=brush]>svg{color:#d07a25}
:where(.doffi-office-ribbon) [data-icon=mic]>svg{color:#0089d6}:where(.doffi-office-ribbon) [data-icon=addins]>svg{color:#ff6c25}:where(.doffi-office-ribbon) [data-icon=editor]>svg{color:#185abd}
:where(.doffi-office-ribbon) .font-group>.ribbon-group-content,:where(.doffi-office-ribbon) .paragraph-group>.ribbon-group-content{flex-direction:column;justify-content:center;gap:6px}
:where(.doffi-office-ribbon[data-kind=document]) .font-group{width:375px;padding-inline:13px}
:where(.doffi-office-ribbon[data-kind=document]) .paragraph-group{width:338px;padding-inline:13px}
:where(.doffi-office-ribbon[data-kind=document]) :is(.font-group,.paragraph-group) :is(.ribbon-font-selects,.ribbon-icon-row){width:100%;justify-content:space-between}
:where(.doffi-office-ribbon) .ribbon-font-selects,:where(.doffi-office-ribbon) .ribbon-icon-row{display:flex;align-items:center;gap:2px;height:24px;flex-shrink:0}
:where(.doffi-office-ribbon) .ribbon-icon-row>.ribbon-tool,:where(.doffi-office-ribbon) .ribbon-font-selects>.ribbon-tool{width:23px;height:24px;min-height:24px;padding:2px;justify-content:center}
:where(.doffi-office-ribbon) .ribbon-icon-row>.ribbon-tool>span:last-child,:where(.doffi-office-ribbon) .ribbon-font-selects>.ribbon-tool>span:last-child{display:none}
:where(.doffi-office-ribbon) .ribbon-glyph{display:inline-flex;align-items:center;justify-content:center;min-width:16px;height:20px;font-family:Georgia,serif;font-size:15px;line-height:20px}
:where(.doffi-office-ribbon) .ribbon-glyph sup,:where(.doffi-office-ribbon) .ribbon-glyph sub{font-size:9px}
:where(.doffi-office-ribbon) .ribbon-divider{display:inline-block;width:1px;height:19px;background:#dfdfdf;margin-inline:3px}
:where(.doffi-office-ribbon) .ribbon-dialog-launcher{position:absolute;right:0;bottom:0;z-index:1;width:10px;height:10px;min-height:0;padding:0}
:where(.doffi-office-ribbon) .ribbon-dialog-launcher>svg{width:9px;height:9px}:where(.doffi-office-ribbon) .ribbon-dialog-launcher>span{display:none}
:where(.doffi-office-ribbon) .office-combo{display:flex;position:relative;border:1px solid #d2d2d2;border-radius:3px;background:#fff;height:24px;flex-shrink:0}
:where(.doffi-office-ribbon) .office-combo input{border:0;border-radius:3px;min-width:0;padding:2px 5px;font-size:11px;width:100%;background:transparent;outline-offset:0}
:where(.doffi-office-ribbon) .office-combo>button{border:0;padding:0 3px;font-size:9px;background:transparent;border-radius:0}
:where(.doffi-office-ribbon) .office-combo:has(input:disabled){opacity:.5}
:where(.doffi-office-ribbon) .font-family-field{width:100px}:where(.doffi-office-ribbon) .font-size-field{width:54px}
:where(.doffi-office-ribbon) .office-combo-options{position:fixed;z-index:70;overflow:auto;border:1px solid #b7b7b7;border-radius:5px;padding:4px 0;background:#fff;box-shadow:0 5px 20px #0003;max-width:calc(100vw - 16px)}
:where(.doffi-office-ribbon) .office-combo-options>button{display:block;width:100%;text-align:left;padding:3px 12px;min-height:25px;font-size:12px;border:0;border-radius:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
:where(.doffi-office-ribbon) .font-options>button{font-size:14px;min-height:27px}
:where(.doffi-office-ribbon) .office-combo-options>button:hover,:where(.doffi-office-ribbon) .office-combo-options>button[aria-selected=true]{background:#dbe7f7}
:where(.doffi-office-ribbon) .office-option-section,:where(.doffi-office-ribbon) .office-menu-heading{padding:5px 10px;color:#666;font-size:11px;text-align:left;font-weight:500;background:#f3f3f3}
:where(.doffi-office-ribbon) .office-combo-options>span{display:block;padding:7px;color:#666;font-size:11px}
:where(.doffi-office-ribbon) .office-combo-error{position:absolute;z-index:80;left:0;top:26px;width:220px;padding:7px;color:#9c3131;background:#fff1f0;border:1px solid #e6b7b2;font-size:11px}
:where(.doffi-office-ribbon) .styles-group{flex:1;min-width:180px;overflow:hidden;padding-inline:9px 3px}
:where(.doffi-office-ribbon) .styles-group>.ribbon-group-content{align-items:center;min-width:0;gap:0}
:where(.doffi-office-ribbon) .style-gallery{display:flex;align-items:center;gap:4px;height:58px;overflow:hidden;min-width:0;flex:1}
:where(.doffi-office-ribbon) .style-gallery>button{display:flex;flex-direction:column;justify-content:center;gap:2px;height:52px;flex:0 0 76px;padding:4px 3px;border:1px solid #d0d0d0;border-radius:4px;background:linear-gradient(white,#f6f6f6);opacity:1;overflow:hidden}
:where(.doffi-office-ribbon) .style-gallery>button[aria-pressed=true]{outline:2px solid #b1c7e6;outline-offset:1px;background:#edf2f9}
:where(.doffi-office-ribbon) .style-gallery>button>span{display:block;width:100%;height:25px;border-bottom:1px solid #e3e3e3;overflow:hidden}
:where(.doffi-office-ribbon) .style-gallery small{color:#333;font-size:10px;line-height:16px}
:where(.doffi-office-ribbon) .style-gallery-menu>.ribbon-tool{width:16px;height:56px;padding:0;border:1px solid #ddd;border-radius:0 3px 3px 0}
:where(.doffi-office-ribbon) .style-gallery-menu>.ribbon-tool>span:not(:first-child),:where(.doffi-office-ribbon) .style-gallery-menu .menu-arrow{display:none}
:where(.doffi-office-ribbon) .office-menu{position:relative;display:inline-flex;flex-shrink:0;align-items:stretch;height:24px}
:where(.doffi-office-ribbon) .office-menu>.ribbon-tool{justify-content:center;padding:2px;gap:2px;height:24px;min-height:24px}
:where(.doffi-office-ribbon) .ribbon-icon-row .office-menu>.ribbon-tool>span:last-of-type,:where(.doffi-office-ribbon) .ribbon-font-selects .office-menu>.ribbon-tool>span:last-of-type{display:none}
:where(.doffi-office-ribbon) .office-menu.large{height:62px}:where(.doffi-office-ribbon) .office-menu.large>.ribbon-tool{height:62px;min-height:62px}
:where(.doffi-office-ribbon) .menu-arrow{font:10px system-ui;width:8px;line-height:16px}
:where(.doffi-office-ribbon) .menu-split-arrow{width:9px;padding:0;border:0;font-size:10px;line-height:24px}
:where(.doffi-office-ribbon) .office-menu:has(button[aria-expanded=true]){background:#e8e8e8;border-radius:4px}
:where(.doffi-office-ribbon) .office-menu-popup{position:fixed;z-index:75;min-width:180px;width:max-content;max-width:calc(100vw - 16px);max-height:calc(100dvh - 155px);overflow:auto;padding:5px 0;border:1px solid #c0c0c0;border-radius:6px;background:#fff;box-shadow:0 7px 24px #0003;color:#222;font:12px/1.4 -apple-system,BlinkMacSystemFont,"PingFang SC",sans-serif}
:where(.doffi-office-ribbon) .office-menu-popup>button{display:flex;justify-content:flex-start;width:100%;padding:5px 16px;min-height:25px;border:0;border-radius:0;text-align:left}
:where(.doffi-office-ribbon) .office-menu-popup>button:hover:not(:disabled){background:#dce8f8}
:where(.doffi-office-ribbon) .office-color-symbol{display:inline-flex;height:19px;width:18px;align-items:center;justify-content:center;border-bottom:3px solid;font:16px Georgia,serif}
:where(.doffi-office-ribbon) .highlight-symbol{transform:rotate(-12deg);font-size:20px}:where(.doffi-office-ribbon) .fill-symbol{font-size:22px}
:where(.doffi-office-ribbon) .color-menu .office-menu-popup{width:230px}
:where(.doffi-office-ribbon) .office-color-grid{display:grid;grid-template-columns:repeat(10,17px);gap:3px;padding:7px 14px}
:where(.doffi-office-ribbon) .office-color-grid>button{width:17px;height:17px;padding:0;border:1px solid #0002;border-radius:0}
:where(.doffi-office-ribbon) .office-color-grid>button:hover{outline:2px solid #aaa;outline-offset:1px}
:where(.doffi-office-ribbon) .office-custom-color{display:flex;align-items:center;justify-content:space-between;padding:5px 14px;border-top:1px solid #ddd}
:where(.doffi-office-ribbon) .office-custom-color input{width:28px;height:24px;padding:0;opacity:0;position:absolute;inset:auto 0 0 0;width:100%;cursor:pointer}
:where(.doffi-office-ribbon) .office-list-gallery{display:grid;grid-template-columns:repeat(3,1fr);gap:5px;padding:8px}
:where(.doffi-office-ribbon) .office-list-gallery button{width:76px;height:64px;border:1px solid #ddd;display:flex;gap:4px;padding:5px;white-space:normal}
:where(.doffi-office-ribbon) .office-list-gallery button>i{font-size:9px;font-style:normal;line-height:16px;color:#777}
:where(.doffi-office-ribbon) .ribbon-group-content>select{align-self:center;border:1px solid #d0d0d0;border-radius:4px;background:#fff;font-size:11px;padding:4px}
:where(.doffi-office-ribbon) .ribbon-tool-stack>label{display:flex;align-items:center;gap:5px;font-size:11px;margin:2px 0}
:where(.doffi-office-ribbon) .ribbon-tool-stack input[type=checkbox]{width:12px;height:12px;margin:0}
:where(.doffi-office-ribbon) .view-description{align-self:center;font-size:11px;color:#777}
:where(.doffi-office-ribbon) .pen-gallery{display:flex;align-items:center;gap:10px}:where(.doffi-office-ribbon) .pen-gallery>button{border-bottom:3px solid currentColor;padding:6px}
:where(.doffi-office-ribbon) .workbook-alignment{width:274px}
:where(.doffi-office-ribbon) .workbook-alignment>.ribbon-group-content{align-items:center;gap:12px}
:where(.doffi-office-ribbon) .workbook-alignment-left{display:flex;flex-direction:column;justify-content:center;gap:6px}
:where(.doffi-office-ribbon) .workbook-alignment .ribbon-tool-stack{gap:6px}
:where(.doffi-office-ribbon) .number-group{width:177px}
:where(.doffi-office-ribbon) .number-group .ribbon-tool-stack{gap:6px;align-items:stretch;width:100%}
:where(.doffi-office-ribbon) .number-format-field{width:155px}
:where(.doffi-office-ribbon) .number-group .ribbon-icon-row{justify-content:space-between}
:where(.doffi-office-ribbon[data-kind=workbook]) .font-group{width:307px}
:where(.doffi-office-ribbon[data-kind=workbook]) .font-family-field{width:177px}
:where(.doffi-office-ribbon[data-kind=workbook]) .font-size-field{width:52px}
:where(.doffi-office-ribbon[data-kind=workbook]) .font-group :is(.ribbon-font-selects,.ribbon-icon-row){width:100%;justify-content:space-between}
:where(.doffi-office-ribbon) .style-gallery-menu>button>svg{display:none}
:where(.doffi-office-ribbon) .office-combo-more{border-top:1px solid #ddd!important;margin-top:4px}
:where(.doffi-office-ribbon) .native-command-gallery{display:flex;align-items:center;gap:5px;height:60px;max-width:680px;overflow:hidden}
:where(.doffi-office-ribbon) .native-command-gallery>button{display:flex;flex-direction:column;width:78px;flex:0 0 78px;height:60px;padding:2px;gap:2px;opacity:.6}
:where(.doffi-office-ribbon) .native-command-gallery>button>span{display:flex;align-items:center;justify-content:center;position:relative;width:72px;height:44px;background:#fff;border:1px solid #ccc;border-top:5px solid #4472c4;font:22px Georgia,serif}
:where(.doffi-office-ribbon) .native-command-gallery>button>span>i{position:absolute;bottom:6px;left:9px;width:45px;height:3px;background:#c8c8c8}
:where(.doffi-office-ribbon) .native-command-gallery small{font-size:10px;white-space:nowrap;max-width:78px;overflow:hidden;text-overflow:ellipsis}
:where(.doffi-office-ribbon) .native-command-gallery.transition>button>span{background:#e9f2f9;border-top:1px solid #ccc}
:where(.doffi-office-ribbon) .native-pen-gallery{display:flex;align-items:center;gap:14px;height:60px;padding:0 8px}
:where(.doffi-office-ribbon) .native-pen-gallery button{height:60px;align-items:flex-end;padding-bottom:2px;border-bottom:3px solid currentColor;opacity:.65}
:where(.doffi-office-ribbon[data-compact=true]) .styles-group{min-width:55px;flex:0 0 55px;padding-inline:5px}
:where(.doffi-office-ribbon[data-compact=true]) .style-gallery{display:none}
:where(.doffi-office-ribbon[data-compact=true]) .style-gallery-menu>.ribbon-tool{width:44px;height:62px;border:0;flex-direction:column;gap:3px}
:where(.doffi-office-ribbon[data-compact=true]) .style-gallery-menu>button>svg{display:block}
:where(.doffi-office-ribbon[data-compact=true]) .style-gallery-menu .gallery-chevron{display:none}
:where(.doffi-office-ribbon[data-compact=true]) .style-gallery-menu>button:after{content:"样式";font-size:11px}
@media(max-width:680px){:where(.doffi-office-ribbon) .office-tabs{padding-inline:4px}:where(.doffi-office-ribbon) .office-tabs>button{padding-inline:10px}:where(.doffi-office-ribbon) .ribbon-group{padding-inline:7px}}
`;

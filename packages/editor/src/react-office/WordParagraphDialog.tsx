import { useState } from 'react';
import type { ReactElement } from 'react';
/** Transient form values become one source transaction when applied. */
export function WordParagraphDialog({apply,close}:{apply(properties:Record<string,string>):void;close():void}):ReactElement{
 const[numbers,setNumbers]=useState<Record<string,string>>({}),[choices,setChoices]=useState<Record<string,string>>({}),[color,setColor]=useState('#000000'),[fill,setFill]=useState('#ffffff'),[borderWidth,setBorderWidth]=useState('0.5');
 const choice=(name:string,key:string,options:readonly(readonly[string,string])[])=> <label style={{display:'block',marginBottom:8}}>{name}<select aria-label={name} value={choices[key]??''} onChange={event=>setChoices({...choices,[key]:event.currentTarget.value})}><option value="">保持不变</option>{options.map(([value,label])=><option key={value} value={value}>{label}</option>)}</select></label>;
 const flags=[['--word-contextual-spacing','相同样式的段落之间不添加间距','1','0'],['break-before','段前分页','page','auto'],['break-after','与下段同页','avoid','auto'],['break-inside','段中不分页','avoid','auto'],['widows','孤行控制','2','1'],['--word-paragraph-snap-to-grid','对齐到文档网格','1','0'],['--word-kinsoku','中文行首行尾禁则','1','0'],['--word-word-wrap','英文单词整体换行','1','0'],['--word-auto-space-de','自动调整中文与西文间距','1','0'],['--word-auto-space-dn','自动调整中文与数字间距','1','0']] as const;
 return <div role="dialog" aria-label="段落设置" style={{position:'absolute',zIndex:30,background:'white',padding:16,border:'1px solid #c8d0dd',boxShadow:'0 4px 20px #0002',maxHeight:'70vh',overflow:'auto'}}><form onSubmit={event=>{event.preventDefault();const properties:Record<string,string>={};
  for(const[key,value]of Object.entries(numbers))if(value!=='')properties[key]=key==='line-height'&&(!choices.rule||choices.rule==='auto')?String(Number(value)):Number(value)+'pt';
  if(numbers['line-height'])properties['--word-line-rule']=choices.rule||'auto';
  for(const[key]of flags)if(choices[key]){properties[key]=choices[key]!;if(key==='widows')properties.orphans=choices[key]!;}
  if(choices.border)for(const side of ['top','right','bottom','left'])properties['border-'+side]=choices.border==='none'?'0pt none '+color:borderWidth+'pt '+choices.border+' '+color;
  if(choices.fill)properties['background-color']=choices.fill==='none'?'transparent':fill;
  apply(properties);close();}}>
  {choice('行距类型','rule',[['auto','多倍行距'],['exact','固定行距'],['atLeast','最小行距']])}
  {([['line-height',!choices.rule||choices.rule==='auto'?'行距（倍）':'行距（磅）'],['margin-top','段前（磅）'],['margin-bottom','段后（磅）'],['margin-left','左缩进（磅）'],['margin-right','右缩进（磅）'],['text-indent','首行缩进（负值为悬挂，磅）']] as const).map(([key,label])=><label key={key} style={{display:'block',marginBottom:8}}>{label}<input aria-label={label} type="number" step="0.1" min={key==='text-indent'?undefined:key==='line-height'?'.1':0} placeholder="保持不变" required={key==='line-height'&&Boolean(choices.rule)} value={numbers[key]??''} onChange={event=>setNumbers({...numbers,[key]:event.currentTarget.value})}/></label>)}
  {flags.map(([key,label,yes,no])=><span key={key}>{choice(label,key,[[yes,'启用'],[no,'关闭']])}</span>)}
  {choice('段落边框','border',[['none','无边框'],['solid','实线'],['double','双线'],['dotted','点线'],['dashed','虚线']])}<label>边框颜色<input type="color" aria-label="段落边框颜色" value={color} onChange={event=>setColor(event.currentTarget.value)}/></label><label>边框宽度（磅）<input type="number" aria-label="段落边框宽度（磅）" min="0" step=".25" value={borderWidth} onChange={event=>setBorderWidth(event.currentTarget.value)}/></label>
  {choice('段落底纹','fill',[['set','设置颜色'],['none','无底纹']])}<input type="color" aria-label="段落底纹颜色" value={fill} onChange={event=>setFill(event.currentTarget.value)}/>
  <div><button type="submit">应用段落设置</button><button type="button" onClick={close}>取消</button></div>
 </form></div>;
}

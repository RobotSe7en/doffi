import { parseArgs } from 'node:util';
import { resolve as outputResolve } from 'node:path';
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {readFileSync,readdirSync,writeFileSync} from 'node:fs';
import {join} from 'node:path';
import {build} from 'vite';
const { values: outputOptions } = parseArgs({ options: { record: { type: 'boolean' }, output: { type: 'string' } } });
if (Boolean(outputOptions.record) === Boolean(outputOptions.output)) throw Error('Use exactly one of --record or --output <file> to save verified evidence.');
const sha=value=>createHash('sha256').update(value).digest('hex');
const expected=JSON.parse(readFileSync('examples/xlsx/expected.json','utf8')).cases;
const built=await build({configFile:false,build:{write:false,minify:false,lib:{entry:join(process.cwd(),'tests/browser/fixtures/workbook-preview.ts'),formats:['iife'],name:'WorkbookPreview'}}});
const outputs=Array.isArray(built)?built.flatMap(part=>part.output):built.output;
const bundle=sha(outputs.filter(part=>part.type==='chunk').map(part=>part.code).join('\n'));
const runs=[];
for(const browser of ['chromium','firefox','webkit','legacy-chromium']){
  const base=join('.cache/workbook-reading-matrix', browser);
  assert.equal(JSON.parse(readFileSync(join(base,'.last-run.json'),'utf8')).status,'passed','The full matrix run must pass, including cleanup');
  const files=readdirSync(base,{recursive:true}).filter(path=>String(path).endsWith('/public-reading.json')&&String(path).split('/')[0].endsWith('-'+(browser==='legacy-chromium'?'chromium':browser)));
  assert.equal(files.length,32,browser+' must cover all 16 source and 16 actual producer files');
  const reports=files.map(path=>JSON.parse(readFileSync(join(base,path),'utf8'))).sort((a,b)=>Number(a.producer)-Number(b.producer)||a.file.localeCompare(b.file));
  for(const producer of [false,true])for(const sample of expected){
    const matches=reports.filter(report=>report.file===sample.file&&report.producer===producer);assert.equal(matches.length,1);
    const report=matches[0];assert.equal(report.hostBundleSha256,bundle,'Matrix must run the current public preview bundle');
    assert.equal(report.sourceSha256,sha(readFileSync('examples/xlsx/'+(producer?'producer/':'')+sample.file)));
    assert.ok(report.observations.length>=2&&report.observations.length%2===0);
    for(const observation of report.observations){
      const view=observation.report;assert.equal(view.kind,'workbook');assert.equal(view.layout,'grid');assert.equal(view.revision,0);
      assert.ok(view.cellFragmentCount>0&&view.cellFragmentCount<=5000);assert.ok(view.visibleCellCount>0&&view.fonts.length>0);
      assert.ok(sample.sheets.includes(observation.sheet));
    }
  }
  runs.push({browser,cases:reports.map(({file,producer,sourceSha256,observations})=>({file,producer,sourceSha256,windows:observations.map(({sheet,last,report:{fonts,...report}})=>({sheet,last,...report,fontResolutions:fonts.length}))}))});
}
const evidence={scope:'R3-06–08 public reading matrix: all selected sources and unchanged actual Calc producers. Feature-specific fixed-font and native semantic evidence remains separate; this is not a pixel-identity claim.',
  hostBundleSha256:bundle,testSha256:sha(readFileSync('tests/browser/workbook-preview.spec.ts')),runs};
writeFileSync(outputOptions.output ? outputResolve(outputOptions.output) : 'docs/evidence/workbook-reading-matrix.json',JSON.stringify(evidence,null,2)+'\n');
process.stdout.write('Recorded 128 complete public source/producer cases on the current bundle.\n');

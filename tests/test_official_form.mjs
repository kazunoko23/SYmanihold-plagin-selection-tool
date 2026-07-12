// SMC公式マニホールド仕様書 自動記入テスト
// 実テンプレート（SS5Y3-10S6-B.xlsx）に EX600 構成を記入し、出力xlsxを生成する。
// 実行: node tests/test_official_form.mjs <テンプレxlsx> <出力xlsx>
import { readFileSync, writeFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const here = dirname(fileURLToPath(import.meta.url));
const html = readFileSync(join(here, '..', 'smc_sy_plugin_v10.html'), 'utf8');
const blocks = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const appJs = blocks[blocks.length - 1];

const templatePath = process.argv[2] || 'C:/Users/mtyf2/Downloads/SS5Y3-10S6-B.xlsx';
const outPath = process.argv[3] || join(here, 'out_filled.xlsx');

// ─── DOMスタブ ───
const elements = {};
function makeEl(id) {
  return {
    id, value: 'U', textContent: '', innerHTML: '', style: {}, dataset: {}, className: '',
    classList: { toggle(){}, add(){}, remove(){}, contains(){ return false; } },
    appendChild(){}, removeChild(){}, addEventListener(){}, setAttribute(){}, click(){},
    getBoundingClientRect(){ return { left:0, top:0 }; }, scrollIntoView(){}, parentElement: null,
  };
}
globalThis.document = {
  getElementById(id) { if (!elements[id]) elements[id] = makeEl(id); return elements[id]; },
  querySelectorAll() { return []; }, querySelector() { return makeEl('q'); },
  createElement() { return makeEl('c'); }, addEventListener() {},
  body: makeEl('body'), documentElement: makeEl('root'),
};
globalThis.window = new Proxy({}, { get(){ return function(){}; } });
const alerts = [];
globalThis.alert = (m) => alerts.push(m);
globalThis.localStorage = { getItem(){ return null; }, setItem(){}, removeItem(){} };
globalThis.XLSX = { utils: { book_new(){return{}}, aoa_to_sheet(){return{}}, book_append_sheet(){} }, writeFile(){} };
globalThis.FileReader = function(){};
let capturedBlob = null;
globalThis.URL = { createObjectURL(b){ capturedBlob = b; return 'blob:x'; }, revokeObjectURL(){} };

const ex = {};
const runner = new Function('document','window','alert','localStorage','XLSX','FileReader','URL','__ex',
  appJs + '\n;__ex.S=S; __ex.buildPN=buildPN; __ex.fillOfficialSpecForm=fillOfficialSpecForm; __ex.getValveSpec=getValveSpec;');
runner(globalThis.document, globalThis.window, globalThis.alert, globalThis.localStorage,
       globalThis.XLSX, globalThis.FileReader, globalThis.URL, ex);
const { S, buildPN, fillOfficialSpecForm, getValveSpec } = ex;

// ─── EX600構成: SS5Y3-10S6Q72-05B-C6（5連: S,S,D,3CS,ブランキング）───
Object.assign(S, {
  series:'3', base:'conn', pipe:'横配管', wiring:'ex600',
  ex600Si:'Q', ex600Io:'7', ex600IoUnits:['EX600-DXPD','EX600-DYPB'], eportPilot:'int',
  valveCount:5, abCode:'C6', mountMethod:'direct', mountPlate:false, mountPrint:false,
  valveTypes:['single','single','double','3cs','single'],
});
elements['sel-eport'] = makeEl('sel-eport'); elements['sel-eport'].value = 'B';
// バルブ個別: 4連目=メタルシール+節電、5連目=ブランキングプレート
for (let i = 0; i < 5; i++) getValveSpec(i);
S.valveSpecs[3].seal = '1'; S.valveSpecs[3].coil = 'T';
S.valveSpecs[4].manifoldOpt = { slot: 'blanking' };
S.valveSpecs[1].manifoldOpt = { supSpacerPipe: '1', supSpacerSize: 'C6' };

const pn = buildPN();
console.log('選定品番:', pn.full, '| complete:', pn.complete);

const bytes = readFileSync(templatePath);
const fakeFile = {
  name: templatePath.split(/[\\/]/).pop(),
  arrayBuffer: async () => bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
};

await fillOfficialSpecForm(fakeFile);
console.log('alerts:', JSON.stringify(alerts, null, 1));
if (!capturedBlob) { console.log('✗ 出力Blobが生成されませんでした'); process.exit(1); }
writeFileSync(outPath, new Uint8Array(await capturedBlob.arrayBuffer()));
console.log('出力:', outPath, (await capturedBlob.arrayBuffer()).byteLength, 'bytes');

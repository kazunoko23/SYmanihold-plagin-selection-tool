// カタログ手配例テスト — smc_sy_plugin_v11.html の buildPN()/getValvePN() を
// SMC SYカタログ（7-1-2-p0387-0722）に実在する手配例品番と突合する。
// 実行: node tests/test_catalog_examples.mjs
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const here = dirname(fileURLToPath(import.meta.url));
const html = readFileSync(join(here, '..', 'smc_sy_plugin_v11.html'), 'utf8');

// 最後の<script>ブロック＝アプリ本体（先頭側はXLSXライブラリ）
const blocks = [...html.matchAll(/<script>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const appJs = blocks.find(b => b.includes("function buildPN")); // アプリ本体（XLSXライブラリ・オンボーディング等を除外）

// ─── DOMスタブ ───
const elements = {};
function makeEl(id) {
  return {
    id, value: 'U', textContent: '', innerHTML: '', style: {}, dataset: {},
    className: '', options: [],
    classList: { toggle(){}, add(){}, remove(){}, contains(){ return false; } },
    appendChild(){}, removeChild(){}, addEventListener(){}, setAttribute(){},
    getBoundingClientRect(){ return { left:0, top:0, width:0, height:0 }; },
    scrollIntoView(){},
    parentElement: null,
  };
}
globalThis.document = {
  getElementById(id) { if (!elements[id]) elements[id] = makeEl(id); return elements[id]; },
  querySelectorAll() { return []; },
  querySelector() { return makeEl('q'); },
  createElement() { return makeEl('c'); },
  addEventListener() {},
  body: makeEl('body'),
  documentElement: makeEl('root'),
};
globalThis.window = new Proxy({}, { get(){ return function(){}; } });
globalThis.alert = () => {};
globalThis.localStorage = { getItem(){ return null; }, setItem(){}, removeItem(){} };
globalThis.XLSX = { utils: { book_new(){return{}}, aoa_to_sheet(){return{}}, book_append_sheet(){} }, writeFile(){} };
globalThis.FileReader = function(){};
globalThis.Blob = function(){};
globalThis.URL = { createObjectURL(){ return ''; }, revokeObjectURL(){} };

const exported = {};
// 直接evalだと外側スコープと衝突するため、関数化して必要シンボルだけ取り出す
const runner = new Function(
  'document','window','alert','localStorage','navigator','XLSX','FileReader','Blob','URL','__ex',
  appJs + '\n;__ex.S=S; __ex.buildPN=buildPN; __ex.getValvePN=getValvePN; __ex.getWiringLimits=getWiringLimits; __ex.getSolenoidUsage=getSolenoidUsage; __ex.getMissingItems=getMissingItems;'
);
runner(globalThis.document, globalThis.window, globalThis.alert, globalThis.localStorage,
       globalThis.navigator, globalThis.XLSX, globalThis.FileReader, globalThis.Blob, globalThis.URL, exported);

const { S, buildPN, getValvePN } = exported;
const S0 = JSON.parse(JSON.stringify(S));

function reset() {
  Object.keys(S).forEach(k => { delete S[k]; });
  Object.assign(S, JSON.parse(JSON.stringify(S0)));
  globalThis.document.getElementById('sel-eport').value = 'U';
}

function setEport(v) { globalThis.document.getElementById('sel-eport').value = v; }

// 共通: 直接取付
function directMount() { S.mountMethod = 'direct'; S.mountPlate = false; S.mountPrint = false; }

// ─── マニホールドベース品番テスト（カタログ手配例）───
const manifoldCases = [
  ['SS5Y3-10F1-05D-C6', () => { // p.430 手配例
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'dsub', connType:'F', connDir:'1', valveCount:5, abCode:'C6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-10FW2-05D-C6', () => {
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'dsub', connType:'FW', connDir:'2', valveCount:5, abCode:'C6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-10P2-05D-C6', () => {
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'flat', connType:'P', connDir:'2', valveCount:5, abCode:'C6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-10L11-05D-C6', () => {
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'lead', leadNum:'1', leadLen:'1', valveCount:5, abCode:'C6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-10M-05D-C6', () => { // p.476 マルチコネクタ手配例
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'multi', valveCount:5, abCode:'C6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-10T-05D-C6', () => {
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'term', valveCount:5, abCode:'C6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-10TC-05D-C6', () => {
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'term_sp', valveCount:5, abCode:'C6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-10SW2-04D-C6', () => { // 小型無線リモート
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'wireless', wirelessType:'W2', valveCount:4, abCode:'C6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-10S6Q72-05B-C6', () => { // EX600 DeviceNet+極性7+I/O2連
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'ex600', ex600Si:'Q', ex600Io:'7',
      ex600IoUnits:['EX600-DXPD','EX600-DXPD'], eportPilot:'int', valveCount:5, abCode:'C6' });
    setEport('B'); directMount();
  }],
  ['SS5Y3-10S6EP-05D-C6', () => { // EX600 CIP Safety
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'ex600', ex600Si:'EP', ex600IoUnits:[], eportPilot:'int', valveCount:5, abCode:'C6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-10SQ1C-05D-C6', () => { // EX250 p.522 手配例
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'ex250', ex250Si:'Q', ex250Io:'1', ex250Spec:'C', eportPilot:'int', valveCount:5, abCode:'C6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-10SNAN-04D-C6', () => { // EX260 p.530 手配例
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'ex260', ex260Si:'NAN', ex260Pts:'32', valveCount:4, abCode:'C6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-10SAANY2-05D-C6', () => { // EX245
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'ex245', ex245Si:'AAN', ex245IoMod:'Y', ex245Io:'2', valveCount:5, abCode:'C6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-10SA3N-05D-C6', () => { // EX500
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'ex500', ex500Si:'A3N', valveCount:5, abCode:'C6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-10S3V-05D-C6', () => { // EX120
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'ex120', ex120Si:'V', valveCount:5, abCode:'C6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-10S4V-05D-C6', () => { // EX126
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'ex126', ex126Si:'V', valveCount:5, abCode:'C6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-10S8V2-05D-C6', () => { // EX180 CC-Link 32点(2)
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'ex180', ex180Si:'V', ex180Pts:'32', ex180PolarSel:'2', ex180Conn:'', eportPilot:'int', valveCount:5, abCode:'C6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-10F1-05UR-C6AA', () => { // 外部パイロットR + 直接取付・銘板+連数印字(AA)
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'dsub', connType:'F', connDir:'1', eportAssy:'R', valveCount:5, abCode:'C6',
      mountMethod:'direct', mountPlate:true, mountPrint:true });
    setEport('U');
  }],
  ['SS5Y3-10F2-05DR-L6', () => { // エルボ上向きø6
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'dsub', connType:'F', connDir:'2', eportAssy:'R', valveCount:5, abCode:'L6' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-12F1-05D', () => { // 12型 上配管（p.442 手配例）
    Object.assign(S, { series:'3', base:'conn', pipe:'上配管', wiring:'dsub', connType:'F', connDir:'1', connUpperPePort:'', valveCount:5 });
    setEport('D'); directMount();
  }],
  ['SS5Y3-12SNAN-04D', () => { // 12型 EX260
    Object.assign(S, { series:'3', base:'conn', pipe:'上配管', wiring:'ex260', ex260Si:'NAN', ex260Pts:'32', connUpperPePort:'', valveCount:4 });
    setEport('D'); directMount();
  }],
  ['SS5Y3-50F2-05B-01', () => { // 金属ベース ねじ配管01
    Object.assign(S, { series:'3', base:'metal', pipe:'横配管', wiring:'dsub', connDir:'2', metalFlatS:'', valveCount:5, abCode:'01', subScrew:'' });
    setEport('B'); directMount();
  }],
  ['SS5Y3-50F2-05D-C6', () => { // 金属ベース ワンタッチC6（p.636 手配例）
    Object.assign(S, { series:'3', base:'metal', pipe:'横配管', wiring:'dsub', connDir:'2', metalFlatS:'', valveCount:5, abCode:'C6', subScrew:'' });
    setEport('D'); directMount();
  }],
  ['SS5Y3-50S5-05B-C6', () => { // 金属ベース EX510（プラスコモン・ダブル）
    Object.assign(S, { series:'3', base:'metal', pipe:'横配管', wiring:'ex510', ex510N:'', ex510S:'', valveCount:5, abCode:'C6', subScrew:'' });
    setEport('B'); directMount();
  }],
  ['SS5Y5-M10F1-05D-C86', () => { // 混合取付 3000/5000（p.574 手配例）
    Object.assign(S, { series:'5', base:'mixed', pipe:'横配管', wiring:'dsub', connType:'F', connDir:'1', valveCount:5, abCode:'C86', mixBig:'8', mixSmall:'6' });
    setEport('D'); directMount();
  }],
  ['SS5Y7-M10F1-05D-C128', () => { // 混合取付 5000/7000
    Object.assign(S, { series:'7', base:'mixed', pipe:'横配管', wiring:'dsub', connType:'F', connDir:'1', valveCount:5, abCode:'C128', mixBig:'12', mixSmall:'8' });
    setEport('D'); directMount();
  }],
  ['SS5Y5-M12F1-05D', () => { // 混合取付 上配管
    Object.assign(S, { series:'5', base:'mixed', pipe:'上配管', wiring:'dsub', connType:'F', connDir:'1', connUpperPePort:'', valveCount:5 });
    setEport('D'); directMount();
  }],
];

// ─── バルブ品番テスト（カタログ手配例）───
const valveCases = [
  ['SY3100-5U1', () => { // 2位置シングル
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'dsub', valveCount:1, valveTypes:['single'] });
    S.valveSpecs = [{ type:'single', seal:'0', pilot:'0', backpress:'', pilotOpt:'', coil:'', voltage:'5', lamp:'U', manual:'', screw:'', portSize:'', portScrew:'', special:null, vacBport:'', sensorCount:'1' }];
  }],
  ['SY3200-5U1', () => { // 2位置ダブル
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'dsub', valveCount:1, valveTypes:['double'] });
    S.valveSpecs = [{ type:'double', seal:'0', pilot:'0', backpress:'', pilotOpt:'', coil:'', voltage:'5', lamp:'U', manual:'', screw:'', portSize:'', portScrew:'', special:null, vacBport:'', sensorCount:'1' }];
  }],
  ['SY3201-5U1', () => { // メタルシール
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'dsub', valveCount:1, valveTypes:['double'] });
    S.valveSpecs = [{ type:'double', seal:'1', pilot:'0', backpress:'', pilotOpt:'', coil:'', voltage:'5', lamp:'U', manual:'', screw:'', portSize:'', portScrew:'', special:null, vacBport:'', sensorCount:'1' }];
  }],
  ['SY3200R-5U1', () => { // 外部パイロット
    Object.assign(S, { series:'3', base:'conn', pipe:'横配管', wiring:'dsub', valveCount:1, valveTypes:['double'] });
    S.valveSpecs = [{ type:'double', seal:'0', pilot:'R', backpress:'', pilotOpt:'', coil:'', voltage:'5', lamp:'U', manual:'', screw:'', portSize:'', portScrew:'', special:null, vacBport:'', sensorCount:'1' }];
  }],
  ['SY3130-5U1-C6', () => { // 上配管形（12型搭載）
    Object.assign(S, { series:'3', base:'conn', pipe:'上配管', wiring:'dsub', valveCount:1, valveTypes:['single'] });
    S.valveSpecs = [{ type:'single', seal:'0', pilot:'0', backpress:'', pilotOpt:'', coil:'', voltage:'5', lamp:'U', manual:'', screw:'', portSize:'C6', portScrew:'', special:null, vacBport:'', sensorCount:'1' }];
  }],
  ['SY5130-5U1-C6-P2', () => { // 圧力センサ付き（センサ2個）
    Object.assign(S, { series:'5', base:'conn', pipe:'横配管', wiring:'dsub', valveCount:1, valveTypes:['single'] });
    S.valveSpecs = [{ type:'single', seal:'0', pilot:'0', backpress:'', pilotOpt:'', coil:'', voltage:'5', lamp:'U', manual:'', screw:'', portSize:'', portScrew:'', special:'pressure', vacBport:'', sensorCount:'2' }];
  }],
  ['SY5300-5U1E', () => { // 残圧排気弁付（3位置クローズドセンタ）
    Object.assign(S, { series:'5', base:'conn', pipe:'横配管', wiring:'dsub', valveCount:1, valveTypes:['3cs'] });
    S.valveSpecs = [{ type:'3cs', seal:'0', pilot:'0', backpress:'', pilotOpt:'', coil:'', voltage:'5', lamp:'U', manual:'', screw:'', portSize:'', portScrew:'', special:'residual', vacBport:'', sensorCount:'1' }];
  }],
];

let pass = 0, fail = 0;
console.log('════ マニホールドベース品番（カタログ手配例突合）════');
for (const [expect, setup] of manifoldCases) {
  reset(); setup();
  let got = '(error)';
  try { const r = buildPN(); got = r ? r.full : '(null)'; } catch (e) { got = '(throw) ' + e.message; }
  const ok = got === expect;
  ok ? pass++ : fail++;
  console.log((ok ? '  OK ' : '✗ NG ') + expect + (ok ? '' : '   ← 生成: ' + got));
}
console.log('════ バルブ品番（カタログ手配例突合）════');
for (const [expect, setup] of valveCases) {
  reset(); setup();
  let got = '(error)';
  try { got = getValvePN(0); } catch (e) { got = '(throw) ' + e.message; }
  const ok = got === expect;
  ok ? pass++ : fail++;
  console.log((ok ? '  OK ' : '✗ NG ') + expect + (ok ? '' : '   ← 生成: ' + got));
}
console.log('──────────────────────────────');
console.log(`結果: ${pass} passed / ${fail} failed`);
process.exit(fail > 0 ? 1 : 0);

// 制約テスト — SMC公式WEBカタログ「品番確認」との突合(2026-07-25)で判明した
// カタログ制約が、ツール側で強制されていることを検証する。
// 根拠: SY総合カタログ（EX600 / CIP Safety / EX245 / EX250 各型式表示方法ページ）
// 実行: node tests/test_sy_constraints.mjs
import { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const here = dirname(fileURLToPath(import.meta.url));
const html = readFileSync(join(here, '..', 'smc_sy_plugin_v10.html'), 'utf8');
// scriptブロックが増えても壊れないよう buildPN を含むブロックを探す
const blocks = [...html.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)].map(m => m[1]);
const appJs = blocks.find(b => /function buildPN/.test(b));
if (!appJs) { console.error('本体スクリプト（buildPNを含むブロック）が見つかりません'); process.exit(1); }

// ─── DOMスタブ ───
const elements = {};
function makeEl(id) {
  return {
    id, value: 'U', textContent: '', innerHTML: '', style: {}, dataset: {},
    className: '', options: [],
    classList: { toggle(){}, add(){}, remove(){}, contains(){ return false; } },
    appendChild(){}, removeChild(){}, addEventListener(){}, setAttribute(){},
    getAttribute(){ return null; },
    querySelector(){ return makeEl('q'); }, querySelectorAll(){ return []; },
    getBoundingClientRect(){ return { left:0, top:0, width:0, height:0 }; },
    scrollIntoView(){}, click(){},
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

const ex = {};
new Function(
  'document','window','alert','localStorage','navigator','XLSX','FileReader','Blob','URL','__ex',
  appJs + `
;__ex.S=S; __ex.buildPN=buildPN; __ex.sanitizeAllValves=sanitizeAllValves;
 __ex.updateMountOptions=updateMountOptions; __ex.getBlockingDiscPN=getBlockingDiscPN;
 __ex.getEx250MaxInputBlocks=getEx250MaxInputBlocks;`
)(globalThis.document, globalThis.window, globalThis.alert, globalThis.localStorage,
  globalThis.navigator, globalThis.XLSX, globalThis.FileReader, globalThis.Blob, globalThis.URL, ex);

const { S, buildPN, sanitizeAllValves, updateMountOptions, getBlockingDiscPN, getEx250MaxInputBlocks } = ex;
const S0 = JSON.parse(JSON.stringify(S));

let pass = 0, fail = 0;
function ck(name, got, want) {
  if (got === want) { pass++; console.log('  OK ' + name); }
  else { fail++; console.log('  NG ' + name + '\n     got =' + JSON.stringify(got) + '\n     want=' + JSON.stringify(want)); }
}
function reset() {
  Object.keys(S).forEach(k => { delete S[k]; });
  Object.assign(S, JSON.parse(JSON.stringify(S0)));
  globalThis.document.getElementById('sel-eport').value = 'D';
}
function pn() { sanitizeAllValves(); updateMountOptions(); return buildPN().full; }
function setCount(n) {
  S.valveCount = n;
  S.valveTypes = Array.from({ length: n }, () => 'single');
  S.valveSpecs = [];
}
// SY3000・コネクタベース・10型横配管・5連・D側・ø6ストレート・直接取付
function base(wiring) {
  reset();
  S.series = '3'; S.base = 'conn'; S.pipe = '横配管'; S.wiring = wiring;
  setCount(5);
  S.abPipeType = 'push'; S.abDir = 'ST'; S.abCode = 'C6';
  S.mountMethod = 'direct'; S.mountPlate = false; S.mountPrint = false;
}

console.log('════ DINレール長さ（□はバルブ連数より長いこと）════');
base('dsub'); S.connType = 'F'; S.connDir = '1'; S.mountMethod = 'din'; S.mountDinLen = '6';
ck('5連+D6 は有効', pn(), 'SS5Y3-10F1-05D-C6D6');
setCount(10);
ck('10連に増連するとD6は解除される', pn(), 'SS5Y3-10F1-10D-C6');
base('dsub'); S.connType = 'F'; S.connDir = '1'; S.mountMethod = 'din'; S.mountDinLen = '12';
ck('5連+D12 は維持', pn(), 'SS5Y3-10F1-05D-C6D12');
base('dsub'); S.connType = 'F'; S.connDir = '1'; S.mountMethod = 'din'; S.mountDinLen = '0';
ck('D0（レールなし）は連数チェックの対象外', pn(), 'SS5Y3-10F1-05D-C6D0');

console.log('════ SIユニットなし時のDINレール制約 ════');
base('ex250'); S.ex250Si = '0'; S.mountMethod = 'din'; S.mountDinLen = '';
ck('EX250 SIなしはDIN不可→直接取付', pn(), 'SS5Y3-10S0-05D-C6');
base('ex600'); S.ex600Si = '0'; S.ex600Io = ''; S.ex600IoPolar = 'none';
S.mountMethod = 'din'; S.mountDinLen = '';
ck('EX600 SIなし+DINはD0に補正', pn(), 'SS5Y3-10S60-05D-C6D0');

console.log('════ EX245 SIユニットなし ════');
base('ex245'); S.ex245Si = '0'; S.ex245IoMod = 'Y'; S.ex245Io = '3';
ck('SIなしは入出力モジュール④⑤とも無記号', pn(), 'SS5Y3-10S0-05D-C6');

console.log('════ EX600 出力極性 ════');
base('ex600'); S.ex600Si = 'FB'; S.ex600IoPolar = 'plus'; S.ex600Io = '2';
sanitizeAllValves();
ck('FBはマイナスコモンのみ→プラス系コネクタを解除', S.ex600Io, null);
base('ex600'); S.ex600Si = 'FB'; S.ex600IoPolar = 'minus'; S.ex600Io = '4';
ck('FB+マイナスコモンは有効', pn(), 'SS5Y3-10S6FB4-05D-C6');
base('ex600'); S.ex600Si = 'DA'; S.ex600IoPolar = 'plus'; S.ex600Io = '2';
ck('DAはプラスコモン可', pn(), 'SS5Y3-10S6DA2-05D-C6');
base('ex600'); S.ex600Si = 'DA'; S.ex600IoPolar = null; S.ex600Io = '';
sanitizeAllValves();
ck('SIユニットありで④無記号は未選択に戻る', S.ex600Io, null);
base('ex600'); S.ex600Si = 'EP'; S.ex600Io = '';
ck('CIP Safety(EP)は④なしで確定', pn(), 'SS5Y3-10S6EP-05D-C6');

console.log('════ EX250 入力ブロック連数の上限 ════');
// SMC公式WEBカタログ実測: TA/TC=4連 / TB/TD=2連 / Q・ZE=8連（入力ブロック仕様に依存しない）
function ex250(si, io, spec) {
  base('ex250'); S.ex250Si = si; S.ex250Io = io; S.ex250Spec = spec;
  return pn();
}
function maxIo(si) { S.ex250Si = si; return getEx250MaxInputBlocks(); }
ck('上限 TA(8in/8out)', maxIo('TA'), 4);
ck('上限 TC(8in/8out)', maxIo('TC'), 4);
ck('上限 TB(4in/4out)', maxIo('TB'), 2);
ck('上限 TD(4in/4out)', maxIo('TD'), 2);
ck('上限 Q(制限なし)',  maxIo('Q'),  8);
ck('TA+5連→4連にクランプ',   ex250('TA', '5', 'A'), 'SS5Y3-10STA4A-05D-C6');
ck('TA+4連×4点入力は有効',   ex250('TA', '4', 'B'), 'SS5Y3-10STA4B-05D-C6');
ck('TC+2連×4点入力は有効',   ex250('TC', '2', 'C'), 'SS5Y3-10STC2C-05D-C6');
ck('Q(制限なし)+8連は有効',  ex250('Q',  '8', 'B'), 'SS5Y3-10SQ8B-05D-C6');
ck('ZE(制限なし)+8連は有効', ex250('ZE', '8', 'E'), 'SS5Y3-10SZE8E-05D-C6');

console.log('════ SUP./EXH.ブロッキングディスク品番 ════');
ck('SY3000 SUP', getBlockingDiscPN('3', 'sup'), 'SY30M-40-1A');
ck('SY3000 EXH（専用品）', getBlockingDiscPN('3', 'exh'), 'SY30M-40-2A');
ck('SY5000 SUP', getBlockingDiscPN('5', 'sup'), 'SY50M-40-1A');
ck('SY5000 EXH（SUPと同一・2個使用）', getBlockingDiscPN('5', 'exh'), 'SY50M-40-1A');
ck('SY7000 EXH（SUPと同一・2個使用）', getBlockingDiscPN('7', 'exh'), 'SY70M-40-1A');

console.log('──────────────────────────────');
console.log(`結果: ${pass} passed / ${fail} failed`);
process.exit(fail ? 1 : 0);

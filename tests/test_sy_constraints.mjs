// 制約テスト — SMC公式WEBカタログ「品番確認」との突合(2026-07-25)で判明した
// カタログ制約が、ツール側で強制されていることを検証する。
// 根拠: SY総合カタログ（EX600 / CIP Safety / EX245 / EX250 各型式表示方法ページ）
// 実行: node tests/test_sy_constraints.mjs
import fs, { readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const here = dirname(fileURLToPath(import.meta.url));
const html = readFileSync(join(here, '..', 'smc_sy_plugin_v13.html'), 'utf8');
const HTML_PATH_FOR_DISCON_TEST = join(here, '..', 'smc_sy_plugin_v13.html');
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
 __ex.getEx250MaxInputBlocks=getEx250MaxInputBlocks;
 __ex.getValvePN=getValvePN; __ex.backpressAllowed=backpressAllowed;
 __ex.getEportStationMax=getEportStationMax; __ex.getWiringLimits=getWiringLimits;
 __ex.getMountCode=getMountCode; __ex.getAbActualCode=getAbActualCode; __ex.AB_COMPAT=AB_COMPAT;
 __ex.getValveSpec=getValveSpec; __ex.veSet=veSet; __ex.renderAbPortBody=renderAbPortBody;
 __ex.updateMetalPeWrap=updateMetalPeWrap;`
)(globalThis.document, globalThis.window, globalThis.alert, globalThis.localStorage,
  globalThis.navigator, globalThis.XLSX, globalThis.FileReader, globalThis.Blob, globalThis.URL, ex);

const { S, buildPN, sanitizeAllValves, updateMountOptions, getBlockingDiscPN, getEx250MaxInputBlocks, getValvePN, backpressAllowed, getEportStationMax, getWiringLimits, getMountCode, getAbActualCode, AB_COMPAT, getValveSpec, veSet, renderAbPortBody, updateMetalPeWrap } = ex;
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

console.log('════ 生産終了品の告知（監査 2026-08-10） ════');
// 根拠: EX600-W 旧カタログ（生産終了品ページ EX600-W-2-old.pdf）の手配例
//   「SS5Y3-10S6WS72-05B-C6（10型5連マニホールドベース、リモート）… EX600-WSV1」
//   → マニホールド型式記号 WS の実体は EX600-WSV1/WSV2。
//   公式生産終了情報: EX600-WSV1/WSV2 は2026年8月終了（→EXW1-RD##E#/G#・RAX#・RDY#M5C／取付互換なし）、
//   無線ベース EX600-WEN1/WEN2/WPN1/WPN2 は2025年3月終了（→EXW1-BENAC1/BPNAC1）。
{
  const src = fs.readFileSync(HTML_PATH_FOR_DISCON_TEST, 'utf8');
  ck('EX600 WS のプロトコル表に生産終了を明記',
     /無線リモート（EX600-WSV1\/2）/.test(src) && /2026年8月生産終了/.test(src), true);
  ck('EX600 WS のSIユニット選択行に終了バッジ',
     /⚠ 生産終了品（EX600-WSV1\/2・2026年8月）/.test(src), true);
  ck('EX600 WS 選択時に品番の注記で警告',
     /EX600-WSV1\/WSV2 は2026年8月生産終了です/.test(src), true);
  ck('後継 EXW1（小型無線リモート）へ誘導している',
     /小型無線リモート EXW1/.test(src), true);
  // 生産終了記号を「無警告で」出していないことの負例: WS を選ぶと必ず警告noteが付く
  base('ex600'); S.ex600Si = 'WS'; S.ex600IoPolar = 'minus'; S.ex600Io = '4';
  const r = buildPN();
  ck('WS選択時のnotesに生産終了警告が入る',
     (r.notes || []).some(function(n){ return /2026年8月生産終了/.test(String(n)); }), true);
  ck('WS選択時の品番自体は従来どおり生成される', r.full.indexOf('SS5Y3-10S6WS') === 0, true);
  // 現行の小型無線リモート（EXW1）は警告なしで通ること
  base('wireless'); S.wirelessType = 'W2';
  const r2 = buildPN();
  ck('小型無線リモート(EXW1)には終了警告が付かない',
     (r2.notes || []).every(function(n){ return !/生産終了/.test(String(n)); }), true);
}

console.log('──────────────────────────────');

/* v12: 2026-08-17 くうあつーる横断監査（品番確認）で確定した制約 */
{
  const setValve = (series, pipe, spec) => {
    Object.assign(S, { series, base:'conn', pipe, wiring:'dsub', valveCount:1, valveTypes:[spec.type||'single'] });
    S.valveSpecs = [Object.assign({ type:'single', seal:'0', pilot:'0', backpress:'', pilotOpt:'', coil:'',
      voltage:'5', lamp:'U', manual:'', screw:'', portSize:'', portScrew:'', special:'', vacBport:'' }, spec)];
  };
  // 残圧排気弁付は「-E」
  setValve('5','横配管',{ type:'3cs', special:'residual' });
  ck('v12 残圧排気弁付は -E（ハイフン必須）', /-5U1-E$/.test(getValvePN(0)), true);
  // 背圧防止弁H: SY7000不可・3位置不可・メタル不可
  ck('v12 H: SY3000 2位置は可', backpressAllowed('3',{seal:'0',type:'double'}), true);
  ck('v12 H: SY5000 2位置は可', backpressAllowed('5',{seal:'0',type:'single'}), true);
  ck('v12 H: SY7000は不可', backpressAllowed('7',{seal:'0',type:'double'}), false);
  ck('v12 H: 3位置クローズドは不可', backpressAllowed('5',{seal:'0',type:'3cs'}), false);
  ck('v12 H: 3位置排気/プレッシャも不可', backpressAllowed('5',{seal:'0',type:'3es'}) || backpressAllowed('5',{seal:'0',type:'3ps'}), false);
  ck('v12 H: メタルシールは不可', backpressAllowed('5',{seal:'1',type:'single'}), false);
  // 絞り弁付真空切換弁: 上配管形かつSY3000/5000のみ
  setValve('5','上配管',{ special:'vacuum' });
  ck('v12 真空: SY5000上配管は品番が出る', /^SY5A3R/.test(getValvePN(0)), true);
  setValve('7','上配管',{ special:'vacuum' });
  ck('v12 真空: SY7000は警告', /⚠/.test(getValvePN(0)), true);
  setValve('5','横配管',{ special:'vacuum' });
  ck('v12 真空: 横配管形は口径記号なしで成立', /^SY5A0R-5U1$/.test(getValvePN(0)), true);
  /* v12訂正: 連数上限は「P,Eポート取出位置」で決まる（配線方式に関わらず共通）。
     実証: SS5Y3-10TC-10U/-10D=OK、-12U/-12D=NG、-12B/-16B/-24B=OK、SS5Y3-10F2-16B-C6=OK／-12D=NG。
     一時 TC だけ一律10連にしていたが、両側Bの11〜24連を塞いでしまう誤りだった */
  ck('v12 P,E取出位置 U/D は10連まで', getEportStationMax(), 10);
  globalThis.document.getElementById('sel-eport').value = 'B';
  ck('v12 P,E取出位置 両側B は24連まで', getEportStationMax(), 24);
  ck('v12 TCの配線方式上限はカタログどおり24', (function(){ S.wiring='term_sp'; return getWiringLimits().st; })(), 24);
  globalThis.document.getElementById('sel-eport').value = 'U';
  // sanitize が成立しない組合せを落とす
  setValve('7','横配管',{ type:'3cs', backpress:'H' });
  sanitizeAllValves();
  ck('v12 sanitize: 成立しないHを落とす', (S.valveSpecs[0].backpress||''), '');
}

console.log('══════════════════════════════════════');
console.log('v13: くうあつーる再検証(2026-08-27〜29)で確定したパターンA〜I ＋副産物 の修正確認');
console.log('══════════════════════════════════════');

function mkValve(opts) {
  reset();
  Object.assign(S, { series: opts.series, base: opts.base || 'conn', pipe: opts.pipe, wiring: opts.wiring || 'dsub' });
  S.eportAssy = opts.eportAssy !== undefined ? opts.eportAssy : '';
  if (S.wiring === 'dsub' || S.wiring === 'flat') {
    S.connDir = opts.connDir || '1';
    if (S.base === 'metal') S.metalFlatS = opts.metalFlatS !== undefined ? opts.metalFlatS : '';
  }
  setCount(opts.count || 5);
  if (opts.abCode !== undefined) S.abCode = opts.abCode;
  if (opts.abPipeType) S.abPipeType = opts.abPipeType;
  if (opts.abDir) S.abDir = opts.abDir;
  S.mountMethod = opts.mountMethod || 'direct';
  if (S.mountMethod === 'din') S.mountDinLen = opts.mountDinLen !== undefined ? opts.mountDinLen : '';
  if (opts.connUpperPePort !== undefined) S.connUpperPePort = opts.connUpperPePort;
  globalThis.document.getElementById('sel-eport').value = opts.eport || 'D';
  const spec = Object.assign({ type:'single', seal:'0', pilot:'0', backpress:'', pilotOpt:'', coil:'',
    voltage:'5', lamp:'U', manual:'', screw:'', portSize:'', portScrew:'', special:null, vacBport:'' }, opts.spec||{});
  S.valveSpecs = [spec];
  for (let i=1;i<S.valveCount;i++) getValveSpec(i);
  return spec;
}

console.log('════ v13 A) ワンタッチ継手径・M5にはねじ種類記号(F/N/T)を付けない ════');
{
  mkValve({ series:'3', pipe:'上配管', connUpperPePort:'', spec:{ portSize:'C6', portScrew:'' } });
  ck('v13 A: SY3130-5U1-C6（ねじ種類なし）は成立', getValvePN(0), 'SY3130-5U1-C6');
  veSet(0, 'portSize', '01');
  veSet(0, 'portScrew', 'N');
  ck('v13 A: portSize=01(ねじ口径)のときportScrewは保持される', getValveSpec(0).portScrew, 'N');
  veSet(0, 'portSize', 'C6');
  ck('v13 A: portSizeをC6(ワンタッチ)に変えるとportScrewは自動で外れる', getValveSpec(0).portScrew, '');

  mkValve({ series:'5', pipe:'上配管', connUpperPePort:'', spec:{ portSize:'01', portScrew:'' } });
  veSet(0, 'portScrew', 'N');
  ck('v13 A: SY5130-5U1-01N（ねじ口径+NPT）は成立', getValvePN(0), 'SY5130-5U1-01N');
  veSet(0, 'portSize', 'C6');
  ck('v13 A: series5でC6(ワンタッチ)へ変更するとportScrewは自動で外れる', getValveSpec(0).portScrew, '');
  ck('v13 A: SY5130-5U1-C6（ねじ種類なし）が生成される', getValvePN(0), 'SY5130-5U1-C6');

  mkValve({ series:'5', pipe:'上配管', connUpperPePort:'', spec:{ portSize:'M5', portScrew:'N' } });
  sanitizeAllValves();
  ck('v13 A: sanitizeAllValvesでM5+portScrew残留を封鎖', getValveSpec(0).portScrew, '');
}

console.log('════ v13 B) 口径混合(CM/LM)を横配管/裏配管で選んだ時は個別バルブ品番に口径記号を付けない ════');
{
  mkValve({ series:'5', pipe:'横配管', abCode:'CM' });
  ck('v13 B: 横配管+CM混合の個別バルブ品番は口径記号なし', getValvePN(0), 'SY5100-5U1');
  getValveSpec(0).portSize = 'C8';
  ck('v13 B: portSizeを設定しても横配管+CM品番には出ない', getValvePN(0), 'SY5100-5U1');

  mkValve({ series:'5', pipe:'裏配管', abCode:'LM', spec:{ portSize:'C8' } });
  ck('v13 B: 裏配管+LM混合でも品番に口径記号は出ない', getValvePN(0), 'SY5100-5U1');

  mkValve({ series:'5', pipe:'上配管', connUpperPePort:'', spec:{ portSize:'C8' } });
  ck('v13 B対照: 上配管形は従来通り口径記号が出る', getValvePN(0), 'SY5130-5U1-C8');
}

console.log('════ v13 C) SY5000上配管ポート選択肢からC10を除外 ════');
{
  const m5 = html.match(/else if\(series==='5'\) portOpts=\[[\s\S]*?\];/);
  ck('v13 C: series5のportOpts配列が見つかる', !!m5, true);
  if (m5) {
    ck('v13 C: series5のportOpts配列にC10は含まれない', /'C10'/.test(m5[0]), false);
    ck('v13 C: series5のportOpts配列にC8は含まれる(存置)', /'C8'/.test(m5[0]), true);
  }
  const m7 = html.match(/else portOpts=\[[\s\S]*?\];/);
  if (m7) ck('v13 C: series7のportOpts配列にC10は含まれる(変更なし)', /'C10'/.test(m7[0]), true);
}

console.log('════ v13 F) 金属ベース裏配管(51型)はP,Eポート取出位置Dのみ ════');
{
  mkValve({ series:'5', base:'metal', pipe:'裏配管', abCode:'C6', eport:'U', count:6 });
  const full1 = pn();
  ck('v13 F: eport=Uで開始しても51型はDに補正される', globalThis.document.getElementById('sel-eport').value, 'D');
  ck('v13 F: 51型はD位置の品番になる', full1, 'SS5Y5-51F1-06D-C6');

  mkValve({ series:'7', base:'metal', pipe:'裏配管', abCode:'C8', eport:'B', count:5 });
  pn();
  ck('v13 F: series7でもeport=BからDに補正される', globalThis.document.getElementById('sel-eport').value, 'D');

  mkValve({ series:'5', base:'conn', pipe:'裏配管', abCode:'C6', eport:'U', count:5 });
  pn();
  ck('v13 F対照: 11型(コネクタベース裏配管)はU/D/B制約なし(Uのまま)', globalThis.document.getElementById('sel-eport').value, 'U');
}

console.log('════ v13 G) 上配管形(12/52型)はDINレール取付非対応 ════');
{
  mkValve({ series:'3', base:'conn', pipe:'上配管', wiring:'term_sp', connUpperPePort:'', mountMethod:'direct', count:3, eport:'U' });
  ck('v13 G: 直接取付はOK(選定完了)', pn(), 'SS5Y3-12TC-03U');

  mkValve({ series:'3', base:'conn', pipe:'上配管', wiring:'term_sp', connUpperPePort:'', mountMethod:'din', mountDinLen:'', count:3, eport:'U' });
  ck('v13 G: DINレール(標準長)はgetMountCodeがnull', getMountCode(), null);
  ck('v13 G: DIN指定時は品番が未完成(complete=false)', buildPN().complete, false);

  mkValve({ series:'5', base:'metal', pipe:'上配管', wiring:'dsub', metalFlatS:'', mountMethod:'din', mountDinLen:'8', count:4, eport:'U', spec:{ portSize:'C8' } });
  for (let i=0;i<4;i++) getValveSpec(i).portSize='C8';
  ck('v13 G: 金属ベース上配管(52型)でもDINは未完成扱い', buildPN().complete, false);
}

console.log('════ v13 H) SY7000はエルボ継手にインチ表記(LN#/BN#)が実在しない ════');
{
  S.series = '7';
  ck('v13 H: SY7000エルボ上向きミリ(L6)は成立', getAbActualCode('C6','EL_UP','mm'), 'L6');
  ck('v13 H: SY7000エルボ上向きにインチ指定してもミリ表記に強制される', getAbActualCode('C6','EL_UP','inch'), 'L6');
  ck('v13 H: SY7000エルボ下向きインチ指定もミリに強制される', getAbActualCode('C8','EL_DN','inch'), 'B8');
  ck('v13 H: SY7000ストレートはインチ可(N7・変更なし)', getAbActualCode('C6','ST','inch'), 'N7');
  S.series = '5';
  ck('v13 H対照: SY5000エルボはインチ表記が成立する(既存動作を維持)', getAbActualCode('C6','EL_UP','inch'), 'LN7');
  S.series = '3';
  ck('v13 H対照: SY3000エルボもインチ表記が成立する(既存動作を維持)', getAbActualCode('C4','EL_UP','inch'), 'LN3');
}

console.log('════ v13 I) 金属ベースSY7000のA,BポートはC8/C10のみ ════');
{
  ck('v13 I: metal_横配管_7のSTがC8,C10のみに限定', JSON.stringify(AB_COMPAT['metal_横配管_7'].ST), JSON.stringify(['C8','C10']));
  ck('v13 I: metal_裏配管_7のSTがC8,C10のみに限定', JSON.stringify(AB_COMPAT['metal_裏配管_7'].ST), JSON.stringify(['C8','C10']));

  mkValve({ series:'7', base:'metal', pipe:'横配管', abCode:'C6', count:5 });
  sanitizeAllValves();
  ck('v13 I: 金属ベースSY7000でC6を指定するとsanitizeで無効化される', S.abCode, null);

  mkValve({ series:'7', base:'metal', pipe:'横配管', abCode:'C12', count:5 });
  sanitizeAllValves();
  ck('v13 I: 金属ベースSY7000でC12を指定するとsanitizeで無効化される', S.abCode, null);

  mkValve({ series:'7', base:'metal', pipe:'横配管', abCode:'C8', count:5 });
  sanitizeAllValves();
  ck('v13 I: C8はそのまま有効', S.abCode, 'C8');
}

console.log('════ v13 副産物) SY5000金属ベース×EX510はA,BポートC4が不成立 ════');
{
  mkValve({ series:'5', base:'metal', pipe:'横配管', wiring:'ex510', abCode:'C4', count:5 });
  sanitizeAllValves();
  ck('v13 副産物: EX510配線でC4はsanitizeで無効化される', S.abCode, null);

  mkValve({ series:'5', base:'metal', pipe:'横配管', wiring:'dsub', metalFlatS:'', abCode:'C4', count:5 });
  sanitizeAllValves();
  ck('v13 副産物対照: dsub配線ではC4は維持される(既存動作)', S.abCode, 'C4');
}

console.log(`結果: ${pass} passed / ${fail} failed`);
process.exit(fail ? 1 : 0);

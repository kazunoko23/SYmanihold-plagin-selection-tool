// 公式仕様書「マニホールドオプション」欄の記入位置テスト（v14〜）
//
// 各スペーサ区画は  [搭載順番(灰色)] → [品名ラベル行] → [品番行]  の順に並ぶ。
// 「搭載順番」はスペーサ2段重ね時の順番欄なので、ツールは何も書かない。
//   ・ブランキングプレート: 品名ラベル行の該当連に ○
//   ・単独SUP/EXHスペーサ : 品名ラベル行＝継手仕様(1/2/3)、次の行＝口径
// v13までは搭載順番行に○を書いて1段ずれていた（SMC公式の記入例と照合して判明）。
//
//   node tests/test_official_form_options.mjs <テンプレ展開フォルダ>
import { readFileSync, existsSync, readdirSync, statSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const here = dirname(fileURLToPath(import.meta.url));
const templateDir = process.argv[2]
  || join(process.env.USERPROFILE || process.env.HOME || '', 'Downloads', 'smc_forms');

if (!existsSync(templateDir)) {
  console.log('SKIP: テンプレートフォルダがありません → ' + templateDir);
  console.log('      SMC配布の仕様書xlsxを展開したフォルダを引数で指定してください。');
  process.exit(0);
}

import { CASES, latestHtml, loadChromium } from './official_form_cases.mjs';


// 再帰的にテンプレを探す
function findTemplates(dir) {
  const out = [];
  for (const e of readdirSync(dir)) {
    const p = join(dir, e);
    if (statSync(p).isDirectory()) out.push(...findTemplates(p));
    else if (e.endsWith('.xlsx') && CASES[e]) out.push({ name: e, path: p });
  }
  return out;
}
const found = findTemplates(templateDir);
if (!found.length) {
  console.log('SKIP: 対象テンプレートが見つかりません → ' + templateDir);
  process.exit(0);
}

// playwright-core は開発時のみ使う（未導入ならスキップ）
//   npm i playwright-core   ※ブラウザDLは不要。インストール済みChromeを使う
const chromium = await loadChromium();
if (!chromium) { console.log('SKIP: playwright-core が未導入です（npm i playwright-core）'); process.exit(0); }
const CHROME = process.env.CHROME_PATH || 'C:/Program Files/Google/Chrome/Application/chrome.exe';
if (!existsSync(CHROME)) { console.log('SKIP: Chromeが見つかりません → ' + CHROME); process.exit(0); }

const url = 'file:///' + latestHtml(join(here, '..')).replace(/\\/g, '/').replace(/ /g, '%20').replace(/[^\x00-\x7F]/g, c => encodeURIComponent(c));
const browser = await chromium.launch({ executablePath: CHROME, headless: true });
const page = await (await browser.newContext()).newPage();
const errs = [];
page.on('pageerror', e => errs.push(e.message));
page.on('dialog', d => d.accept());
await page.goto(url, { waitUntil: 'load' });
await page.waitForTimeout(400);
await page.evaluate(() => {
  window.__captured = null;
  const o = URL.createObjectURL.bind(URL);
  URL.createObjectURL = (b) => { window.__captured = b; return o(b); };
});

let pass = 0, fail = 0;
for (const { name, path } of found) {
  const c = CASES[name];
  const setup = await page.evaluate((c) => {
    resetAll();
    setSeries(c.series); setBase(c.base); setPipe(c.pipe); setWiring(c.wiring);
    const s = c.sub || {};
    if (s.connType) { setConnType(s.connType); setConnDir(s.connDir || '1'); }
    if (s.metalFlatS !== undefined && typeof setMetalFlatS === 'function') setMetalFlatS(s.metalFlatS);
    if (s.leadNum)  { setLeadNum(s.leadNum); setLeadLen(s.leadLen); }
    if (s.ex600Si)  { setEx600Si(s.ex600Si); if (s.ex600IoPolar) setEx600IoPolar(s.ex600IoPolar); if (s.ex600Io) setEx600Io(s.ex600Io); }
    if (s.ex250Si)  { setEx250Si(s.ex250Si); setEx250Io(s.ex250Io); setEx250Spec(s.ex250Spec); }
    if (s.ex260Si)  { setEx260Pts(s.ex260Pts); setEx260Si(s.ex260Si); }
    if (s.ex120Si)  { setEx120Proto(s.ex120Proto); setEx120Si(s.ex120Si); }
    changeValve(1); changeValve(1); changeValve(1);          // 5連
    document.getElementById('sel-eport').value = 'D';
    if (c.ab) { setAbPipeType('push'); setAbDir('ST'); setAbCode(c.ab); }
    if (c.upperPe !== undefined && typeof setConnUpperPePort === 'function') setConnUpperPePort(c.upperPe);
    if (c.mix) { S.mixBig = c.mix[0]; S.mixSmall = c.mix[1]; S.abCode = 'C' + c.mix[0] + c.mix[1]; }
    setMountMethod('direct');
    const kinds = ['single', 'double', '3cs', 'single', 'single'];
    for (let i = 0; i < S.valveCount; i++) {
      const k = kinds[i % kinds.length];
      S.valveTypes[i] = k; getValveSpec(i).type = k;
      // 上配管・口径混合はバルブ連別のA,B口径が必要（未設定だとcomplete=false）
      if (k !== 'blank' && (S.pipe === '上配管' || S.abCode === 'CM' || S.abCode === 'LM')) {
        getValveSpec(i).portSize = (S.series === '3' ? 'C6' : 'C8');
      }
    }
    const sz = { '3': 'C6', '5': 'C8', '7': 'C10' }[S.series];
    getValveSpec(1).manifoldOpt = { supSpacerPipe: '1A', supSpacerSize: sz };
    // v15: 1連目にパイロット弁オプション・背圧防止弁Ass'y・SUPブロッキングディスク
    getValveSpec(0).pilotOpt = 'B';
    getValveSpec(0).manifoldOpt = Object.assign({}, getValveSpec(0).manifoldOpt || {}, { backpressAssy: true, blockingDiscSup: true });
    const szEl = { '3': 'L4', '5': 'L6', '7': 'L10' }[S.series];   // エルボ(2A/3A)はL系（v15〜 実在組合せのみ）
    getValveSpec(2).manifoldOpt = { exhSpacerPipe: '3A', exhSpacerSize: szEl };
    getValveSpec(S.valveCount - 1).manifoldOpt = { slot: 'blanking' };   // 最終連（金属ベースは4連になる）
    updateAll();
    const pn = buildPN();
    pn.sz = sz; pn.szEl = szEl; pn.last = S.valveCount;
    return { pn: pn.full, complete: !!pn.complete, sz: pn.sz, szEl: pn.szEl, last: pn.last, metal: S.base === 'metal', ab: (S.pipe !== '上配管' && S.abCode && S.abCode !== 'CM' && S.abCode !== 'LM') ? S.abCode : '' };
  }, c);

  const b64 = readFileSync(path).toString('base64');
  const res = await page.evaluate(async ({ b64, name }) => {
    window.__captured = null;
    const bin = atob(b64); const buf = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
    try { await fillOfficialSpecForm({ name, arrayBuffer: async () => buf.buffer }); }
    catch (e) { return { error: e.message }; }
    if (!window.__captured) return { error: '出力Blobなし' };
    // 記入後のシートXMLから ◆マニホールド型式 行を読み直して左から連結する
    return { ok: true };
  }, { b64, name });

  if (res.error) { fail++; console.log(`  NG ${name} — 記入失敗: ${res.error}`); continue; }
  if (!setup.complete) { fail++; console.log(`  NG ${name} — 選定が未完了: ${setup.pn}`); continue; }

  // 記入後のxlsxを、ツール自身のZIP/シートパーサで読み直して検査する
  const got = await page.evaluate(async (last) => {
    const ab = await window.__captured.arrayBuffer();
    const entries = _zipParse(new Uint8Array(ab));
    const sheet = entries.filter(e => /^xl\/worksheets\/sheet\d+\.xml$/.test(e.name))
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }))[0];
    const sst = _parseSharedStrings(await _zipEntryText(entries, 'xl/sharedStrings.xml'));
    const cells = _sheetCellMap(await _zipEntryText(entries, sheet.name), sst);
    const rowNum = r => parseInt(r.match(/\d+$/)[0], 10);
    const colOf  = r => r.match(/^[A-Z]+/)[0];
    const keys = Object.keys(cells);
    const hdr = keys.filter(r => cells[r] === 'Ｄ側').map(rowNum).sort((a, b) => a - b)[0];
    const st = {};
    keys.filter(r => rowNum(r) === hdr && /^\d+$/.test(String(cells[r]).trim()))
      .forEach(r => { st[+String(cells[r]).trim()] = colOf(r); });
    const mount = keys.filter(r => cells[r] === '搭載順番').map(rowNum);
    const lab = sub => keys.filter(r => String(cells[r]).indexOf(sub) >= 0).map(rowNum)
      .sort((a, b) => a - b).find(lr => mount.some(m => Math.abs(m - lr) <= 3)) || 0;
    const v = (n, row) => (cells[st[n] + row] === undefined ? '' : String(cells[st[n] + row]));
    const rB = lab("ブランキングプレートAss'y"), rS = lab('単独SUP.スペーサ'), rE = lab('単独EXH.スペーサ');
    const nz = x => String(x).replace(/[\uFF01-\uFF5E]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)).replace(/[\s\u3000]/g, '');
    const firstRow = f => keys.filter(r => f(nz(cells[r]), r)).map(rowNum).sort((a, b) => a - b)[0] || 0;
    const dCol = keys.filter(r => nz(cells[r]) === 'D側').map(r => _colToNum(colOf(r))).sort((a, b) => a - b)[0] || 99;
    const left = r => _colToNum(colOf(r)) < dCol;
    const rPO = firstRow((t, r) => left(r) && /パイロット弁?オプション/.test(t));
    const rBP = firstRow((t, r) => left(r) && t.indexOf("背圧防止弁Ass'y") >= 0);
    const rDS = firstRow((t, r) => left(r) && t.indexOf('SUP.ブロッキングディスク') >= 0);
    const rG  = firstRow((t, r) => left(r) && t.indexOf('A,Bポート管接続口径') >= 0
      && keys.some(k => rowNum(k) === rowNum(r) && left(k) && nz(cells[k]) === 'Aポート'));
    const discRow = rDS ? keys.filter(r => rowNum(r) === rDS && String(cells[r]) === '○').length : -1;
    const listText = keys.filter(r => /SY\d0M-26-/.test(String(cells[r]))).map(r => String(cells[r])).join(' ');
    // 搭載順番行に何か書かれていないか（全連）
    const mountDirty = [];
    mount.forEach(m => Object.keys(st).forEach(n => { if (v(n, m)) mountDirty.push(n + '連@' + m + '=' + v(n, m)); }));
    return { rDS, rPO, po1: rPO ? v(1, rPO) : '-', rBP, bp1: rBP ? v(1, rBP) : '-', discRow, rG, g1: rG ? v(1, rG) + '/' + v(1, rG + 1) : '-', listText,
             rB, rS, rE, blankLast: v(last, rB), supV: v(2, rS), supW: v(2, rS + 1), exhV: v(3, rE), exhW: v(3, rE + 1),
             blankOther: Object.keys(st).filter(n => +n !== last).map(n => v(n, rB)).join(''), mountDirty };
  }, setup.last);
  const errs1 = [];
  if (!got.rB || !got.rS || !got.rE) errs1.push(`ラベル行が見つからない B${got.rB} S${got.rS} E${got.rE}`);
  if (got.blankLast !== '○') errs1.push(`ブランキング${setup.last}連目="${got.blankLast}"`);
  if (got.blankOther) errs1.push(`ブランキング他連="${got.blankOther}"`);
  if (got.supV !== '1' || got.supW !== setup.sz) errs1.push(`SUPスペーサ2連目 V="${got.supV}" W="${got.supW}"`);
  if (got.exhV !== '3' || got.exhW !== setup.szEl) errs1.push(`EXHスペーサ3連目 V="${got.exhV}" W="${got.exhW}"`);
  if (got.mountDirty.length) errs1.push('搭載順番行に記入: ' + got.mountDirty.join(','));
  // v15 追加分
  if (got.rPO && got.po1 !== 'B') errs1.push(`パイロット弁オプション1連目="${got.po1}"`);
  if (got.rBP && got.bp1 !== '○') errs1.push(`背圧防止弁Ass'y1連目="${got.bp1}"`);
  if (got.rDS && got.discRow !== 1) errs1.push(`SUPブロッキングディスク○=${got.discRow}個`);
  if (setup.ab && got.rG && got.g1 !== setup.ab + '/' + setup.ab) errs1.push(`A,B口径1連目="${got.g1}"（期待 ${setup.ab}）`);
  const wantBlk = setup.metal ? '-26-2A' : '-26-1A';
  if (got.listText.indexOf(wantBlk) < 0) errs1.push(`構成製品リストのブランキング="${got.listText}"（期待 ${wantBlk}）`);
  if (!errs1.length) { pass++; console.log(`  OK ${name.padEnd(28)} ブランキング行${got.rB} SUP行${got.rS} EXH行${got.rE} ディスク行${got.rDS}(○${got.discRow}) 口径${got.g1}`); }
  else { fail++; console.log(`  NG ${name.padEnd(28)} ${errs1.join(' / ')}  (行: B${got.rB} S${got.rS} E${got.rE} / 品番 ${setup.pn})`); }
}

console.log('──────────────────────────────');
console.log(`結果: ${pass} passed / ${fail} failed`);
if (errs.length) console.log('page errors:', errs.slice(0, 3).join(' | '));
await browser.close();
process.exit(fail ? 1 : 0);

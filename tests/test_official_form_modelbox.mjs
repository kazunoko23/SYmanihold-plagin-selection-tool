// 公式仕様書「◆マニホールド型式」ボックスの記入位置テスト
//
// テンプレの型式ボックスは  ― SS5Y□ ― [固定表記] [枠][枠] ― [枠][枠] ― [枠]  の形をしており、
// 枠に入れる値は「―」で区切られた3グループに配る。基準となる固定表記の列を取り違えると
// グループが丸ごとずれて、連数がコネクタ種類の枠に入る等の誤記入になる。
// ここでは記入後のボックスを左から連結し、選定品番と一致することを確認する。
//
// テンプレート（SMC配布のxlsx）はリポジトリに含めないため、展開済みフォルダを引数で渡す。
//   node tests/test_official_form_modelbox.mjs <テンプレ展開フォルダ>
// 省略時は既定パスを見に行き、無ければSKIPする。
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

// テンプレ名 → 選定構成（ツール側の設定値）
const CASES = {
  'SS5Y3-10S6-B.xlsx':       { series:'3', base:'conn',  pipe:'横配管', wiring:'ex600', sub:{ ex600Si:'Q', ex600IoPolar:'plus', ex600Io:'2' }, ab:'C6'  },
  'SS5Y3-12S-B(EX250).xlsx': { series:'3', base:'conn',  pipe:'上配管', wiring:'ex250', sub:{ ex250Si:'Q', ex250Io:'1', ex250Spec:'A' }, upperPe:'' },
  'SS5Y5-10F-B.xlsx':        { series:'5', base:'conn',  pipe:'横配管', wiring:'dsub',  sub:{ connType:'F',  connDir:'1' }, ab:'C8'  },
  'SS5Y5-10PH-B.xlsx':       { series:'5', base:'conn',  pipe:'横配管', wiring:'flat',  sub:{ connType:'PH', connDir:'1' }, ab:'C8'  },
  'SS5Y5-10S(EX250)-B.xlsx': { series:'5', base:'conn',  pipe:'横配管', wiring:'ex250', sub:{ ex250Si:'Q', ex250Io:'1', ex250Spec:'C' }, ab:'C8' },
  'SS5Y5-10S3-B.xlsx':       { series:'5', base:'conn',  pipe:'横配管', wiring:'ex120', sub:{ ex120Proto:'DeviceNet', ex120Si:'Q' }, ab:'C8' },
  'SS5Y5-11M-A.xlsx':        { series:'5', base:'conn',  pipe:'裏配管', wiring:'multi', sub:{}, ab:'C8' },
  'SS5Y5-12L+-A.xlsx':       { series:'5', base:'conn',  pipe:'上配管', wiring:'lead',  sub:{ leadNum:'1', leadLen:'1' }, upperPe:'' },
  'SS5Y5-M10M-A.xlsx':       { series:'5', base:'mixed', pipe:'横配管', wiring:'multi', sub:{}, mix:['8','6'] },
  'SS5Y5-M11M.xlsx':         { series:'5', base:'mixed', pipe:'裏配管', wiring:'multi', sub:{}, mix:['8','6'] },
  'SS5Y7-10S(EX260)-A.xlsx': { series:'7', base:'conn',  pipe:'横配管', wiring:'ex260', sub:{ ex260Pts:'32', ex260Si:'QA' }, ab:'C10' },
  // 2026-07-29 追加DL分
  'SS5Y3-10S-EX260-B.xlsx':  { series:'3', base:'conn',  pipe:'横配管', wiring:'ex260', sub:{ ex260Pts:'32', ex260Si:'QA' }, ab:'C6' },
  'SS5Y5-11S(EX260)-A.xlsx': { series:'5', base:'conn',  pipe:'裏配管', wiring:'ex260', sub:{ ex260Pts:'32', ex260Si:'QA' }, ab:'C8' },
  'SS5Y5-12M.xlsx':          { series:'5', base:'conn',  pipe:'上配管', wiring:'multi', sub:{}, upperPe:'' },
  'SS5Y5-M12M.xlsx':         { series:'5', base:'mixed', pipe:'上配管', wiring:'multi', sub:{}, upperPe:'' },
  'SS5Y5-M12S(EX250).xlsx':  { series:'5', base:'mixed', pipe:'上配管', wiring:'ex250', sub:{ ex250Si:'Q', ex250Io:'1', ex250Spec:'A' }, upperPe:'' },
  'SS5Y5-M12T.xlsx':         { series:'5', base:'mixed', pipe:'上配管', wiring:'term',  sub:{}, upperPe:'' },
  // 2026-07-29 金属ベース50/51/52型（型式ボックスに印字済みセルが挟まる構造）
  'SS5Y3-50P-A.xlsx':        { series:'3', base:'metal', pipe:'横配管', wiring:'flat', sub:{ connType:'P',  connDir:'2', metalFlatS:'S' }, ab:'C6' },
  'SS5Y5-50PH-A.xlsx':       { series:'5', base:'metal', pipe:'横配管', wiring:'flat', sub:{ connType:'PH', connDir:'1', metalFlatS:'' }, ab:'C8' },
  'SS5Y3-51P-A.xlsx':        { series:'3', base:'metal', pipe:'裏配管', wiring:'flat', sub:{ connType:'P',  connDir:'2', metalFlatS:'S' }, ab:'C6' },
  'SS5Y5-51P-A.xlsx':        { series:'5', base:'metal', pipe:'裏配管', wiring:'flat', sub:{ connType:'P',  connDir:'2', metalFlatS:'' }, ab:'C8' },
  'SS5Y3-52PH-A.xlsx':       { series:'3', base:'metal', pipe:'上配管', wiring:'flat', sub:{ connType:'PH', connDir:'1', metalFlatS:'S' } },
  'SS5Y5-52F-A.xlsx':        { series:'5', base:'metal', pipe:'上配管', wiring:'dsub', sub:{ connType:'F',  connDir:'2', metalFlatS:'' } },
};

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
let chromium;
try { ({ chromium } = await import('playwright-core')); }
catch { console.log('SKIP: playwright-core が未導入です（npm i playwright-core）'); process.exit(0); }
const CHROME = process.env.CHROME_PATH || 'C:/Program Files/Google/Chrome/Application/chrome.exe';
if (!existsSync(CHROME)) { console.log('SKIP: Chromeが見つかりません → ' + CHROME); process.exit(0); }

const url = 'file:///' + join(here, '..', 'smc_sy_plugin_v11.html').replace(/\\/g, '/').replace(/ /g, '%20').replace(/[^\x00-\x7F]/g, c => encodeURIComponent(c));
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
    const kinds = ['single', 'double', '3cs', 'single', 'blank'];
    for (let i = 0; i < S.valveCount; i++) {
      const k = kinds[i % kinds.length];
      S.valveTypes[i] = k; getValveSpec(i).type = k;
      // 上配管・口径混合はバルブ連別のA,B口径が必要（未設定だとcomplete=false）
      if (k !== 'blank' && (S.pipe === '上配管' || S.abCode === 'CM' || S.abCode === 'LM')) {
        getValveSpec(i).portSize = (S.series === '3' ? 'C6' : 'C8');
      }
    }
    updateAll();
    const pn = buildPN();
    return { pn: pn.full, complete: !!pn.complete };
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

  // 記入後のxlsxは、ツール自身のZIP/シートパーサをページ内で再利用して読む
  const assembled = await page.evaluate(async () => {
    const ab = await window.__captured.arrayBuffer();
    const entries = _zipParse(new Uint8Array(ab));
    const sheet = entries.filter(e => /^xl\/worksheets\/sheet\d+\.xml$/.test(e.name))
      .sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true }))[0];
    const sst = _parseSharedStrings(await _zipEntryText(entries, 'xl/sharedStrings.xml'));
    const cells = _sheetCellMap(await _zipEntryText(entries, sheet.name), sst);
    const rowNum = r => parseInt(r.match(/\d+$/)[0], 10);
    const colOf  = r => r.match(/^[A-Z]+/)[0];
    const anchor = Object.keys(cells).find(r =>
      String(cells[r]).indexOf('マニホールド型式') >= 0 && String(cells[r]).indexOf('◆') >= 0);
    if (!anchor) return '(型式行なし)';
    const row = rowNum(anchor);
    return Object.keys(cells).filter(r => rowNum(r) === row)
      .sort((a, b) => _colToNum(colOf(a)) - _colToNum(colOf(b)))
      .map(r => String(cells[r]).trim())
      .filter(v => v && v.indexOf('◆') < 0 && v !== '⇒' && v.indexOf('右頁') < 0)
      .join('')
      .replace(/―/g, '-')
      .replace(/[Ａ-Ｚ０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
      .replace(/^-+|-+$/g, '');
  });
  const expect = setup.pn;
  if (assembled === expect) { pass++; console.log(`  OK ${name.padEnd(28)} ${assembled}`); }
  else { fail++; console.log(`  NG ${name.padEnd(28)} 型式ボックス="${assembled}" 期待="${expect}"`); }
}

console.log('──────────────────────────────');
console.log(`結果: ${pass} passed / ${fail} failed`);
if (errs.length) console.log('page errors:', errs.slice(0, 3).join(' | '));
await browser.close();
process.exit(fail ? 1 : 0);

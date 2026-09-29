// 公式仕様書テスト共通: テンプレ名 → 選定構成、最新版HTML、playwright-core の読み込み
import { readdirSync } from 'fs';
import { createRequire } from 'module';
import { join } from 'path';

// テンプレ名 → 選定構成（ツール側の設定値）
export const CASES = {
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

// リポジトリ内で最も新しい smc_sy_plugin_v*.html
export function latestHtml(root) {
  const v = f => +(f.match(/_v(\d+)\.html$/) || [0, 0])[1];
  const fs = readdirSync(root).filter(f => /^smc_sy_plugin_v\d+\.html$/.test(f)).sort((a, b) => v(a) - v(b));
  return join(root, fs[fs.length - 1]);
}

// playwright-core: ローカル導入 → 環境変数 PLAYWRIGHT_CORE_PATH → グローバル導入 の順に探す
export async function loadChromium() {
  try { return (await import('playwright-core')).chromium; } catch {}
  const req = createRequire(import.meta.url);
  for (const p of [process.env.PLAYWRIGHT_CORE_PATH,
                   join(process.env.APPDATA || '', 'npm', 'node_modules', 'playwright-core')]) {
    if (!p) continue;
    try { return req(p).chromium; } catch {}
  }
  return null;
}

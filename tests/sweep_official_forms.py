# -*- coding: utf-8 -*-
"""SMC従来仕様書（SY・約145種）への自動記入スイープ（v15〜）
各テンプレの型式に合う実在マニホールド品番を kuatsu-pn-db から選び、「品番から逆引き→公式仕様書に記入」を実行して
  ・型式ボックスの連結＝選定品番 ・印刷済みセルの上書きなし ・搭載順番行への記入なし ・仕様欄の連別記入あり
を確認する。混合取付(M型)とPCワイヤリング(J/G型)は品番DBに無いため対象外。
  python tests/sweep_official_forms.py <テンプレ展開フォルダ> [<kuatsu_smc_pn.sqlite>] [<html>]
要: pip install playwright（chromium）
"""
import sys, os, json, re, base64, pathlib, sqlite3
HERE = os.path.dirname(os.path.abspath(__file__))
TPL = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.expanduser('~'), 'Downloads', 'smc_forms_sy')
DB = sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, '..', '..', 'kuatsu-pn-db', 'kuatsu_smc_pn.sqlite')
if len(sys.argv) > 3: HTML = sys.argv[3]
else:
    hs = sorted([f for f in os.listdir(os.path.join(HERE, '..')) if re.match(r'smc_sy_plugin_v\d+\.html$', f)], key=lambda f: int(re.search(r'v(\d+)', f).group(1)))
    HTML = os.path.join(HERE, '..', hs[-1])
if not os.path.isdir(TPL) or not os.path.exists(DB):
    print('SKIP: テンプレフォルダまたは品番DBがありません'); sys.exit(0)
c=sqlite3.connect(DB)
rows=[(pn,json.loads(a)) for pn,a in c.execute("select pn,attrs from pn where tool='sy-plugin' and pn like 'SS5Y%'")]
def want(f):
    m=re.match(r'SS5Y(\d)-(M?\d\d)(.*)\.xlsx',f)
    if not m: return None
    s,typ,rest=m.groups()
    r=rest.upper()
    if 'EX250' in r: w='ex250'
    elif 'EX260' in r: w='ex260'
    elif 'EX500' in r: w='ex500'
    elif r.startswith('S3'): w='ex120'
    elif r.startswith('S4'): w='ex126'
    elif r.startswith('S6'): w='ex600'
    elif r.startswith('S5'): w='ex510'
    elif r.startswith('S8'): w='ex180'
    elif r.startswith('SW'): w='wireless'
    elif r.startswith('S'): w='ex250'
    elif r.startswith('F'): w='dsub'
    elif r.startswith('PG'): w='flat:PG'
    elif r.startswith('PH'): w='flat:PH'
    elif r.startswith('P'): w='flat:P'
    elif r.startswith('L'): w='lead'
    elif r.startswith('M'): w='multi'
    elif r.startswith('TC'): w='term_sp'
    elif r.startswith('T') or r.startswith('Ｔ'): w='term'
    else: w='?'+r[:2]
    return s,typ,w
out={}
for f in sorted(os.listdir(TPL)):
    if f.startswith('25A'): continue
    wv=want(f)
    if not wv: out[f]=None; continue
    s,typ,w=wv; wk,_,ct=w.partition(':')
    cands=[pn for pn,a in rows if pn.startswith(f'SS5Y{s}-{typ}') and a.get('wiring')==wk and (not ct or re.match(rf'SS5Y{s}-{typ}{ct}\d',pn)) and re.search(r'-05[A-Z]',pn)]
    if not cands: cands=[pn for pn,a in rows if pn.startswith(f'SS5Y{s}-{typ}') and a.get('wiring')==wk and (not ct or re.match(rf'SS5Y{s}-{typ}{ct}\d',pn))]
    # prefer C-size straight & mid
    cands.sort(key=lambda p:(('-C' not in p), len(p)))
    out[f]={'want':w,'pn':cands[0] if cands else None}
miss=[f for f,v in out.items() if not v or not v['pn']]
print('対象', len(out)-len(miss), '/ 品番なし', len(miss))
from playwright.sync_api import sync_playwright
pick=out
res={}
JS_CHECK = r"""async ([b64o, b64n]) => {
  const load = async (b64) => { const s=atob(b64); const u=new Uint8Array(s.length); for(let i=0;i<s.length;i++)u[i]=s.charCodeAt(i);
    const e=_zipParse(u); const sh=e.filter(x=>/^xl\/worksheets\/sheet\d+\.xml$/.test(x.name)).sort((a,b)=>a.name.localeCompare(b.name,undefined,{numeric:true}))[0];
    const sst=_parseSharedStrings(await _zipEntryText(e,'xl/sharedStrings.xml')); return _sheetCellMap(await _zipEntryText(e,sh.name), sst); };
  const A = await load(b64o), B = await load(b64n);
  const over=[], added=[];
  Object.keys(B).forEach(r=>{ if (A[r]===undefined || A[r]==='') added.push(r+'='+B[r]); else if (String(A[r])!==String(B[r])) over.push(r+':'+A[r]+'→'+B[r]); });
  const rowNum=r=>parseInt(r.match(/\d+$/)[0],10), colOf=r=>r.match(/^[A-Z]+/)[0];
  const anchor=Object.keys(B).find(r=>String(B[r]).indexOf('マニホールド型式')>=0&&(String(B[r]).indexOf('◆')>=0||String(B[r]).indexOf('□')>=0));
  let box='';
  if(anchor){const row=rowNum(anchor); box=Object.keys(B).filter(r=>rowNum(r)===row).sort((a,b)=>_colToNum(colOf(a))-_colToNum(colOf(b)))
    .map(r=>String(B[r]).trim()).filter(v=>v&&v.indexOf('◆')<0&&v.indexOf('□')<0&&v!=='⇒'&&v.indexOf('右頁')<0).join('').replace(/―/g,'-')
    .replace(/[Ａ-Ｚ０-９]/g,ch=>String.fromCharCode(ch.charCodeAt(0)-0xFEE0)).replace(/^-+|-+$/g,'');}
  const mount=Object.keys(A).filter(r=>A[r]==='搭載順番').map(rowNum);
  const inMount=added.filter(x=>mount.indexOf(rowNum(x.split('=')[0]))>=0);
  return {over, nAdded:added.length, box, inMount, added};
}"""
with sync_playwright() as pw:
    br=pw.chromium.launch(); ctx=br.new_context(viewport={'width':1600,'height':1000},accept_downloads=True)
    p=ctx.new_page(); alerts=[]
    p.on('dialog', lambda d:(alerts.append(d.message), d.accept()))
    p.on('pageerror', lambda e: alerts.append('PAGEERR '+str(e)))
    p.goto(pathlib.Path(HTML).as_uri()); p.wait_for_timeout(800)
    p.evaluate("()=>{const o=URL.createObjectURL.bind(URL);URL.createObjectURL=(b)=>{window.__cap=b;return o(b)};}")
    for f,v in pick.items():
        if not v or not v['pn']: continue
        pn=v['pn']; s=pn[4]; alerts.clear()
        r=p.evaluate("""([pn,s])=>{resetAll(); showReversePNModal();
            document.getElementById('rev-manifold-input').value=pn;
            const n=+document.getElementById('rev-row-count').textContent; revSetRows(3-n);
            ['SY'+s+'100-5U1','SY'+s+'200-5U1','SY'+s+'300-5U1'].forEach((x,i)=>{const c=document.getElementById('rev-cell-'+i); if(c) c.value=x;});
            runReversePN(); closeReversePNModal(); const b=buildPN(); return {full:b.full, complete:!!b.complete, st:document.getElementById('rev-status').innerText.slice(0,300)};}""",[pn,s])
        b64=base64.b64encode(open(os.path.join(TPL,f),'rb').read()).decode()
        p.evaluate("""async ([b64,name])=>{window.__cap=null; const s=atob(b64);const u=new Uint8Array(s.length);for(let i=0;i<s.length;i++)u[i]=s.charCodeAt(i);
            await fillOfficialSpecForm({name, arrayBuffer: async()=>u.buffer});}""",[b64,f])
        has=p.evaluate("()=>!!window.__cap")
        chk=None
        if has:
            nb=p.evaluate("async()=>{const ab=await window.__cap.arrayBuffer();let s='';const u=new Uint8Array(ab);for(let i=0;i<u.length;i+=8192)s+=String.fromCharCode.apply(null,u.subarray(i,i+8192));return btoa(s);}")
            chk=p.evaluate(JS_CHECK,[b64,nb])
        res[f]=dict(pn=pn, built=r['full'], complete=r['complete'], rev=r['st'], alerts=list(alerts), out=chk)
        ok = has and chk['box']==r['full'] and not chk['over'] and not chk['inMount'] and r['full']==pn
        if False: print(f, pn, '' if ok else ('built='+r['full']+' box='+(chk['box'] if chk else '-')+' over='+str(chk['over'][:3] if chk else '-')+' alerts='+str([a[:80] for a in alerts][:2])), flush=True)
    br.close()
r=res; ok=0; tot=0; addsum=0
for f,v in r.items():
    tot+=1; o=v['out']; al=' '.join(v['alerts'])
    m=re.search(r'仕様欄（連別）: (\d+)連分',al); st=int(m.group(1)) if m else -1
    prob=[]
    if not o: prob.append('出力なし')
    else:
        addsum+=o['nAdded']
        if o['box']!=v['built']: prob.append(f"box={o['box']} built={v['built']}")
        if o['over']: prob.append('上書き'+str(o['over'][:3]))
        if o['inMount']: prob.append('搭載順番'+str(o['inMount'][:3]))
    if st!=3: prob.append(f'仕様欄{st}連')
    if 'PAGEERR' in al: prob.append('ERR')
    if prob: print(' ',f, ' / '.join(prob))
    else: ok+=1
print(os.path.basename(HTML),'OK',ok,'/',tot,' 記入セル合計',addsum)
sys.exit(0 if ok==tot else 1)

// ===== 全件検算レポート(読み取り専用) =====
// 商品コードごとに EMSリスト×消込台帳(出荷済み)×受注明細(取り寄せ)×Yahoo自由在庫 を突き合わせ、
// 「発送済み消費・現在受注・過去便の対応」が数量として説明できるかを1枚で見る。
// 受注明細・EMSリスト・台帳・履歴には一切書き込まない(書くのは「全件検算」シートだけ)。
// 設計: docs/superpowers/specs/2026-07-10-full-allocation-rebuild-v2-design.md

const ZENKEN_CFG = { シート: '全件検算' };

// 純粋集計(Nodeテスト対象)。src:
//   ems:    [{code, p, st, qty, arrival, ems}]    EMSリストの行(素の値。実EMS番号だけ供給扱い)
//   出荷済: [{ban, code, sku, qty, 入荷日}]        消込台帳_出荷済み行_()の戻り(重複排除前)
//   受注:   [{ban, code, sku, qty, 選択肢, 商品名, 入荷日}] 受注明細の行(素の値)
//   a在庫:  {基底コード:数量} YahooCSV集計_の結果。null=Yahoo照合なし
//   商品名: {基底コード:商品名} Yahoo側の商品名(任意)
// 戻り: {rows:[{code,名,判定,到着済,反映済,未着,出荷到着,出荷過去,出荷不明,確保到着,確保過去,確保ズレ,待ち,残,yahooA}], counts:{判定:件数}}
function 全件検算_受注基底マップ_(src){
  const map={};
  (src.出荷済||[]).concat(src.受注||[]).forEach(r=>{
    const ban=String(r.ban||'').trim(), base=受注基底コード_(r.sku,r.code);
    if(ban && base){
      if(!map[ban]) map[ban]=new Set();
      map[ban].add(base);
    }
  });
  return map;
}

function 全件検算_EMS集計コード_(r, baseByBan){
  const raw=normCode_(r.code); if(!raw) return '';
  if(/^\d{5,}$/.test(raw)) return raw; // 中古・名指し商品の注文専用キー
  const bans=String(r.p||'').match(/\d{5,}/g)||[];
  const bases=[]; let unresolved=false;
  bans.forEach(b=>{
    const set=baseByBan[b]; if(!set){ unresolved=true; return; }
    set.forEach(base=>{ if(bases.indexOf(base)<0) bases.push(base); });
  });
  return !unresolved && bases.length===1? bases[0] : raw;
}

function 全件検算_需要集計コード_(r, orderSupply){
  const ban=String(r.ban||'').trim();
  if(ban && orderSupply[ban]) return ban;
  return 受注基底コード_(r.sku,r.code);
}

function 全件検算_集計_(src){
  // --- EMSリストを 供給(コード -> 数量とステータス別の到着日集合) へ ---
  const 供給={}, orderSupply={};
  const 供給を=c=>供給[c]||(供給[c]={到着済:0,反映済:0,未着:0,到着日:new Set(),反映日:new Set()});
  const baseByBan=全件検算_受注基底マップ_(src);
  (src.ems||[]).forEach(r=>{
    if(r.ems!==undefined && !実EMS番号_(r.ems)) return;
    const code=全件検算_EMS集計コード_(r,baseByBan); if(!code) return;
    if(/^\d{5,}$/.test(code)) orderSupply[code]=1;
    const qty=Number(r.qty)||0, st=String(r.st||'').trim(), d=ymd_(r.arrival);
    const e=供給を(code);
    if(st==='到着済'){ e.到着済+=qty; if(d) e.到着日.add(d); }
    else if(st==='在庫反映済み'){ e.反映済+=qty; if(d) e.反映日.add(d); }
    else e.未着+=qty;
  });

  const 集={}, 名={};
  const 集を=c=>集[c]||(集[c]={出荷到着:0,出荷過去:0,出荷不明:0,確保到着:0,確保過去:0,確保ズレ:0,待ち:0});

  // --- 台帳の出荷済みを箱の到着日で分類(重複排除してから) ---
  出荷済み重複排除_(src.出荷済||[]).forEach(r=>{
    const base=全件検算_需要集計コード_(r,orderSupply); if(!base) return;
    const qty=Number(r.qty)||0; if(qty<=0) return;
    const e=供給[base], d=ymd_(r.入荷日), a=集を(base);
    if(e && d && e.到着日.has(d)) a.出荷到着+=qty;
    else if(e && d && e.反映日.has(d)) a.出荷過去+=qty;
    else a.出荷不明+=qty;
  });

  // --- 受注明細の未出荷取り寄せ(台湾/中国ルートと即納は対象外) ---
  (src.受注||[]).forEach(r=>{
    if(区分_(r.選択肢)!=='取り寄せ') return;
    if(/台湾|中国/.test(String(r.選択肢||'')) || /台湾|中国/.test(String(r.商品名||''))) return;
    const qty=Number(r.qty)||0; if(qty<=0) return;
    const base=全件検算_需要集計コード_(r,orderSupply); if(!base) return;
    if(!名[base] && String(r.商品名||'').trim()) 名[base]=String(r.商品名).trim();
    const a=集を(base);
    if(String(r.入荷日==null?'':r.入荷日).trim()===''){ a.待ち+=qty; return; }
    const e=供給[base], d=ymd_(r.入荷日);
    if(e && d && e.到着日.has(d)) a.確保到着+=qty;
    else if(e && d && e.反映日.has(d)) a.確保過去+=qty;
    else a.確保ズレ+=qty;
  });

  // --- コードごとに判定(上から先勝ち) ---
  const yahooあり=src.a在庫!=null;
  const 順={'⚠️超過消費':0,'⚠️入荷日ズレ':1,'📦供給不足':2,'ℹ️箱残>Yahoo':3,'ℹ️EMS外在庫':4,'OK':5};
  const rows=[], counts={};
  const codes={};
  Object.keys(供給).forEach(c=>codes[c]=1);
  Object.keys(集).forEach(c=>codes[c]=1);
  Object.keys(codes).sort().forEach(c=>{
    const e=供給[c]||{到着済:0,反映済:0,未着:0};
    const a=集[c]||{出荷到着:0,出荷過去:0,出荷不明:0,確保到着:0,確保過去:0,確保ズレ:0,待ち:0};
    const 残=e.到着済-a.出荷到着-a.確保到着;
    const ya=yahooあり? (Number((src.a在庫||{})[c])||0) : null;
    let 判定;
    if(残<0) 判定='⚠️超過消費';
    else if(a.確保ズレ>0) 判定='⚠️入荷日ズレ';
    else if(a.待ち > e.未着+Math.max(0,残)) 判定='📦供給不足';
    else if(yahooあり && 残>ya) 判定='ℹ️箱残>Yahoo';
    else if(yahooあり && 残<ya) 判定='ℹ️EMS外在庫';
    else 判定='OK';
    counts[判定]=(counts[判定]||0)+1;
    rows.push({code:c, 名:名[c]||((src.商品名||{})[c]||''), 判定:判定,
      到着済:e.到着済, 反映済:e.反映済, 未着:e.未着,
      出荷到着:a.出荷到着, 出荷過去:a.出荷過去, 出荷不明:a.出荷不明,
      確保到着:a.確保到着, 確保過去:a.確保過去, 確保ズレ:a.確保ズレ,
      待ち:a.待ち, 残:残, yahooA:ya});
  });
  rows.sort((x,y)=> (順[x.判定]-順[y.判定]) || (x.code<y.code? -1 : x.code>y.code? 1 : 0));
  return {rows:rows, counts:counts};
}

// ===== Task9: 取り置き台帳式の箱別検算(純粋・Nodeテスト対象) =====
// src:
//   supply:    [{ems, code, sourceCode, matchCode, qty, status}] EMSリスト全状態の行(素の値)
//   ledger:    取り置き台帳の行 / movements: EMS在庫移動台帳の行
//   yahoo:     {基底コード:数量} | null(照合なし)
// 戻り: {rows, pendingRows, productRows, counts, errors}
//   rows: 箱別(EMS番号×元コードidentity)。供給=到着済+在庫反映済みの実数量。
//         不変式: 供給 = 取り置き中+発送済み+戻し未処理+在庫なし確定+Yahoo移動済み+余り
//         供給に無い台帳/移動の使用も行として出す(余りが負=⚠️超過消費)。
//   pendingRows: 未着の箱(検算対象外の見える化)
//   productRows: 基底コード単位に余りを合算しYahoo a在庫と比較(比較のみ。箱別判定には使わない)
function 全件検算_台帳集計_(src){
  const summary=取り置き_集計_(src.ledger||[], src.movements||[]);
  const errors=summary.errors.slice();
  const boxes={}, meta={}, pending={};
  const boxFor=(key,ems,sourceCode,matchCode)=>{
    if(!boxes[key]){
      boxes[key]={EMS番号:ems,商品コード:sourceCode,判定:'OK',供給:0,取り置き中:0,発送済み:0,戻し未処理:0,在庫なし確定:0,Yahoo移動済み:0,余り:0};
      meta[key]={matchCode:matchCode||normCode_(sourceCode)};
    }
    return boxes[key];
  };
  (src.supply||[]).forEach(r=>{
    const ems=String(r.ems||'').trim(), sourceCode=String(r.sourceCode||r.code||'').trim();
    const qty=Number(r.qty)||0, status=String(r.status||'').trim();
    if(!実EMS番号_(ems) || !sourceCode || qty<=0) return;
    const key=取り置き_供給キー_(ems,sourceCode);
    const matchCode=normCode_(r.matchCode||r.code||sourceCode);
    if(status==='到着済' || status==='在庫反映済み') boxFor(key,ems,sourceCode,matchCode).供給+=qty;
    else {
      const p=pending[key]||(pending[key]={EMS番号:ems,商品コード:sourceCode,未着数量:0});
      p.未着数量+=qty;
    }
  });
  // 供給に無い使用(締め済み過去箱への追記漏れ・台帳の誤記)も行に出して負の余りで見せる
  Object.keys(summary.usageBySupply).forEach(key=>{
    if(boxes[key]) return;
    const sep=key.indexOf('|'), ems=sep>=0?key.slice(0,sep):key, sourceCode=sep>=0?key.slice(sep+1):'';
    if(!実EMS番号_(ems) || !sourceCode) return;
    boxFor(key,ems,sourceCode,normCode_(sourceCode));
  });
  Object.keys(boxes).forEach(key=>{
    const row=boxes[key], u=summary.usageBySupply[key];
    row.取り置き中=u?u.取り置き中:0; row.発送済み=u?u.発送済み:0; row.戻し未処理=u?u.戻し未処理:0;
    row.在庫なし確定=u?u.在庫なし確定:0; row.Yahoo移動済み=u?u.Yahoo移動済み:0;
    row.余り=row.供給-取り置き_使用合計_(u);
    row.判定=row.余り<0?'⚠️超過消費':'OK';
  });
  const rank={'⚠️超過消費':0,'OK':1};
  const rows=Object.keys(boxes).map(k=>boxes[k]).sort((a,b)=> (rank[a.判定]-rank[b.判定])
    || (a.商品コード<b.商品コード?-1:a.商品コード>b.商品コード?1:0)
    || (a.EMS番号<b.EMS番号?-1:a.EMS番号>b.EMS番号?1:0));
  const pendingRows=Object.keys(pending).sort().map(k=>pending[k]);
  const yahooあり=src.yahoo!=null;
  const products={};
  Object.keys(boxes).forEach(key=>{
    const mc=meta[key].matchCode; if(!mc) return;
    const p=products[mc]||(products[mc]={商品コード:mc,余り:0,'Yahoo a在庫':yahooあり?(Number((src.yahoo||{})[mc])||0):'',差:''});
    p.余り+=boxes[key].余り;
  });
  const productRows=Object.keys(products).sort().map(mc=>{
    const p=products[mc];
    if(yahooあり) p.差=(Number(p['Yahoo a在庫'])||0)-p.余り;
    return p;
  });
  const counts={}; rows.forEach(r=>{ counts[r.判定]=(counts[r.判定]||0)+1; });
  return {rows,pendingRows,productRows,counts,errors};
}

function 全件検算レポート(){ 直列_(全件検算レポート本体_); }
function 全件検算レポート本体_(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();

  // --- 発注共有EMSリスト(読み取りのみ) ---
  let sh;
  try{ sh=発注共有を開く_().getSheetByName(P_KAKUTEI_CFG.シート); }
  catch(e){ ui.alert('発注共有ファイルが開けません:\n'+e.message); return; }
  if(!sh){ ui.alert('発注共有に「'+P_KAKUTEI_CFG.シート+'」がありません'); return; }
  const hr=P_KAKUTEI_CFG.ヘッダー行, last=sh.getLastRow();
  if(last<=hr){ ui.alert('EMSリストにデータがありません'); return; }
  const head=sh.getRange(hr,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
  const f=(...names)=>{ for(const n of names){ const i=head.indexOf(n); if(i>=0) return i; } return -1; };
  const cSt=f('ステータス列'), cC=f('商品コード'), cQ=f('数量','個数'), cE=f('EMS番号'), cP=f('注文番号');
  let cA=f('EMS到着日','到着日','到着'); if(cA<0) cA=4; // 既定E列(🔎整合チェックと同じ)
  if(cC<0){ ui.alert('EMSリストの'+hr+'行目に「商品コード」見出しがありません'); return; }
  const ems=sh.getRange(hr+1,1,last-hr,sh.getLastColumn()).getValues().map(r=>({
    code:r[cC], p:cP>=0? r[cP]:'', st:cSt>=0? r[cSt]:'到着済', qty:cQ>=0? r[cQ]:1, arrival:r[cA], ems:cE>=0?r[cE]:''}));

  // --- 受注明細(読み取りのみ) ---
  const recv=ss.getSheetByName(HIKIATE_CFG.受注);
  if(!recv){ ui.alert('「'+HIKIATE_CFG.受注+'」タブがありません'); return; }
  const M=列マップ_(recv);
  const R=recv.getDataRange().getValues();
  const 受注=[];
  for(let i=M.hr;i<R.length;i++){
    const row=R[i];
    if(!String(row[M.番号]||'').trim()) continue;
    受注.push({ban:row[M.番号], code:row[M.コード], sku:M.SKU>=0? row[M.SKU]:'', qty:row[M.個数],
      選択肢:row[M.選択肢], 商品名:M.商品名>=0? row[M.商品名]:'', 入荷日:M.入荷>=0? row[M.入荷]:''});
  }

  // --- 取り置き台帳・EMS在庫移動台帳(見出し不足は検算不能として中止) ---
  let ledgerRows, movementRows;
  try{ ledgerRows=取り置き台帳_読む_(); movementRows=EMS在庫移動台帳_読む_(); }
  catch(e){ ui.alert('全件検算を中止しました','取り置き台帳またはEMS在庫移動台帳を読み込めません:\n'+e.message,ui.ButtonSet.OK); return; }

  // --- 消込台帳の出荷済み(参考情報用。読めない時は0件で続行) ---
  let 出荷済=[]; try{ 出荷済=消込台帳_出荷済み行_(); }catch(e){}

  // --- Yahoo在庫CSV(読めない時はYahoo照合なしで続行) ---
  const y=Yahoo在庫を読む_();

  // --- 本判定: 取り置き台帳式の箱別検算 ---
  const t=全件検算_台帳集計_({
    supply: ems.map(r=>({ems:r.ems, code:r.code, sourceCode:String(r.code==null?'':r.code).trim(),
      matchCode:normCode_(r.code), qty:r.qty, status:String(r.st||'').trim()||'到着済'})),
    ledger: ledgerRows, movements: movementRows, yahoo: y.error? null : y.a在庫
  });

  // --- 参考情報: 旧・日付ベース集計(判定には使わない) ---
  const r=全件検算_集計_({ems:ems, 出荷済:出荷済, 受注:受注,
    a在庫:y.error? null : y.a在庫, 商品名:y.error? null : y.商品名});

  // --- 描画(書くのはこのシートだけ) ---
  let rep=ss.getSheetByName(ZENKEN_CFG.シート); if(!rep) rep=ss.insertSheet(ZENKEN_CFG.シート);
  rep.clear();
  const cnt=['⚠️超過消費','OK'].filter(k=>t.counts[k]).map(k=>k+' '+t.counts[k]).join(' ／ ');
  rep.getRange(1,1).setValue('全件検算(読み取り専用): '+Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm')
    +(y.error? ' ／ ⚠️Yahoo照合なし('+y.error+')' : ' ／ Yahoo CSV: '+y.fileName)
    +' ／ '+(cnt||'対象箱なし')
    +(t.errors.length? ' ／ ⚠️台帳エラー '+t.errors.length+'件(下部参照)':''));
  const HDR=['EMS番号','商品コード','判定','供給','取り置き中','発送済み','戻し未処理','在庫なし確定','Yahoo移動済み','余り'];
  rep.getRange(2,1,1,HDR.length).setValues([HDR]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff').setFontSize(HIKIATE_CFG.字);
  rep.setFrozenRows(2);
  const out=t.rows.map(x=>[x.EMS番号,x.商品コード,x.判定,x.供給,x.取り置き中,x.発送済み,x.戻し未処理,x.在庫なし確定,x.Yahoo移動済み,x.余り]);
  if(out.length){
    rep.getRange(3,1,out.length,HDR.length).setValues(out).setFontSize(HIKIATE_CFG.字);
    rep.getRange(3,1,out.length,HDR.length).setBackgrounds(t.rows.map(x=>new Array(HDR.length).fill(x.判定==='⚠️超過消費'? HIKIATE_CFG.色_赤 : null)));
  } else {
    rep.getRange(3,1).setValue('(対象箱なし: 実EMS番号の供給も台帳の使用もありません)');
  }
  let cur=3+Math.max(out.length,1)+1;
  const 表を=(title,hdr,rows)=>{
    rep.getRange(cur,1).setValue(title).setFontWeight('bold').setFontSize(HIKIATE_CFG.字); cur++;
    rep.getRange(cur,1,1,hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#d9d9d9').setFontSize(HIKIATE_CFG.字); cur++;
    if(rows.length){ rep.getRange(cur,1,rows.length,hdr.length).setValues(rows).setFontSize(HIKIATE_CFG.字); cur+=rows.length; }
    else { rep.getRange(cur,1).setValue('(なし)').setFontSize(HIKIATE_CFG.字); cur++; }
    cur++;
  };
  表を('■ 商品別Yahoo比較(比較のみ・箱別判定には使わない)',['商品コード','余り(箱合算)','Yahoo a在庫','差'],
    t.productRows.map(p=>[p.商品コード,p.余り,p['Yahoo a在庫']===''?'—':p['Yahoo a在庫'],p.差===''?'—':p.差]));
  表を('■ 未着の箱(検算対象外)',['EMS番号','商品コード','未着数量'],
    t.pendingRows.map(p=>[p.EMS番号,p.商品コード,p.未着数量]));
  if(t.errors.length) 表を('■ ⚠️台帳エラー(取置ID・数量の不備。台帳を直してから再実行)',['内容'],t.errors.map(e=>[e]));
  const 旧HDR=['商品コード','商品名','判定','EMS到着済','EMS反映済','EMS未着','出荷済(到着箱)','出荷済(過去便)','出荷済(箱不明)','確保済(到着箱)','確保済(過去便)','確保済(ズレ)','待ち(入荷日なし)','到着済残','Yahoo a在庫'];
  表を('■ 参考情報(旧・日付ベース集計。判定には使わない)',旧HDR,
    r.rows.map(x=>[x.code, x.名, x.判定, x.到着済, x.反映済, x.未着, x.出荷到着, x.出荷過去, x.出荷不明,
      x.確保到着, x.確保過去, x.確保ズレ, x.待ち, x.残, x.yahooA==null? '—' : x.yahooA]));
  const 凡例=[
    '⚠️超過消費: 箱の実数量以上を台帳・Yahoo移動が消費扱い(記録漏れ/重複の疑い) → 取り置き台帳の該当EMS番号×商品コードを確認',
    '余り: 供給−(取り置き中+発送済み+戻し未処理+在庫なし確定+Yahoo移動済み)。締め前の箱の余りは⑤でYahooへ移す',
    '商品別Yahoo比較の差: Yahoo a在庫−余り合算。プラス=EMS外在庫(免税店買付などがあれば正常)/マイナス=Yahoo未反映の疑い',
    '※ このレポートは何も書き換えない。棚卸・EMS番号空欄は供給対象外。Yahoo即納数は比較だけで引当には使わない'
  ];
  rep.getRange(cur,1,凡例.length,1).setValues(凡例.map(s=>[s])).setFontSize(HIKIATE_CFG.字);
  ss.setActiveSheet(rep);
  ss.toast('全件検算: '+t.rows.length+'箱'+(cnt? ' ／ '+cnt : ''),'🧮全件検算',8);
}

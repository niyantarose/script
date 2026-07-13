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

  // --- 消込台帳の出荷済み(読めない時は0件で続行) ---
  let 出荷済=[]; try{ 出荷済=消込台帳_出荷済み行_(); }catch(e){}

  // --- Yahoo在庫CSV(読めない時はYahoo照合なしで続行) ---
  const y=Yahoo在庫を読む_();

  const r=全件検算_集計_({ems:ems, 出荷済:出荷済, 受注:受注,
    a在庫:y.error? null : y.a在庫, 商品名:y.error? null : y.商品名});

  // --- 描画(書くのはこのシートだけ) ---
  let rep=ss.getSheetByName(ZENKEN_CFG.シート); if(!rep) rep=ss.insertSheet(ZENKEN_CFG.シート);
  rep.clear();
  const 判定順=['⚠️超過消費','⚠️入荷日ズレ','📦供給不足','ℹ️箱残>Yahoo','ℹ️EMS外在庫','OK'];
  const cnt=判定順.filter(k=>r.counts[k]).map(k=>k+' '+r.counts[k]).join(' ／ ');
  rep.getRange(1,1).setValue('全件検算(読み取り専用): '+Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm')
    +(y.error? ' ／ ⚠️Yahoo照合なし('+y.error+')' : ' ／ Yahoo CSV: '+y.fileName)
    +' ／ '+(cnt||'対象コードなし'));
  const HDR=['商品コード','商品名','判定','EMS到着済','EMS反映済','EMS未着','出荷済(到着箱)','出荷済(過去便)','出荷済(箱不明)','確保済(到着箱)','確保済(過去便)','確保済(ズレ)','待ち(入荷日なし)','到着済残','Yahoo a在庫'];
  rep.getRange(2,1,1,HDR.length).setValues([HDR]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff').setFontSize(HIKIATE_CFG.字);
  rep.setFrozenRows(2);
  const out=r.rows.map(x=>[x.code, x.名, x.判定, x.到着済, x.反映済, x.未着, x.出荷到着, x.出荷過去, x.出荷不明,
    x.確保到着, x.確保過去, x.確保ズレ, x.待ち, x.残, x.yahooA==null? '—' : x.yahooA]);
  if(out.length){
    rep.getRange(3,1,out.length,HDR.length).setValues(out).setFontSize(HIKIATE_CFG.字);
    const bg=r.rows.map(x=>{
      const c= x.判定.indexOf('⚠️')===0? HIKIATE_CFG.色_赤 : x.判定==='📦供給不足'? HIKIATE_CFG.色_黄 : null;
      return new Array(HDR.length).fill(c);
    });
    rep.getRange(3,1,out.length,HDR.length).setBackgrounds(bg);
  } else {
    rep.getRange(3,1).setValue('(対象コードなし: EMSリスト・取り寄せ注文・出荷済みのいずれにも商品がありません)');
  }
  let cur=3+Math.max(out.length,1)+1;
  const 凡例=[
    '⚠️超過消費: 到着済の箱の個数以上に消費扱い(幽霊スタンプ/記録漏れの疑い) → 🔎整合チェックで行単位を確認',
    '⚠️入荷日ズレ: 入荷日がどの箱の到着日とも一致しない → 🔎整合チェックで確認(手入力の正しい日付もある)',
    '📦供給不足: 待ち注文が見えている実EMS供給(EMS未着+到着済残)を超える。棚卸数量やYahoo即納数で補わず、EMSリストを確認',
    'ℹ️箱残>Yahoo: 締め前の便なら正常(箱の余りは締めでYahooへ反映される)。締めたはずの商品なら未記録出荷の疑い',
    'ℹ️EMS外在庫: Yahoo自由在庫がEMS由来より多い(免税店買付などがあれば正常)',
    '※ このレポートは何も書き換えない。棚卸・EMS番号空欄は供給対象外。Yahoo即納数は比較だけで引当には使わない'
  ];
  rep.getRange(cur,1,凡例.length,1).setValues(凡例.map(s=>[s])).setFontSize(HIKIATE_CFG.字);
  ss.setActiveSheet(rep);
  ss.toast('全件検算: '+r.rows.length+'コード'+(cnt? ' ／ '+cnt : ''),'🧮全件検算',8);
}

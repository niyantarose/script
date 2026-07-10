// ===== 棚卸箱の下書きを生成(Yahoo在庫CSV × 未出荷注文) =====
// 「机上棚卸」: 現物を数える代わりに、待ち注文の確保分だけを持つ棚卸箱を作る。
//   棚卸箱数量(基底コード) = max(0, 未出荷の取り寄せ数 - 未着箱の入荷予定数)
// 自由在庫(即納a)は Yahoo在庫DB が唯一の正なので、引当ファイル側へは持ち込まない
// (写すと「正が2つ」になり再び乖離が生まれる。②に必要なのは待ち注文の確保分だけ)。
// Yahoo在庫CSV は「その注文の商品が本当に届いているか」の突き合わせ材料(内訳列)として使う。
//   sub-code 末尾 a=即納(quantity=物理の自由在庫) / b=お取り寄せ(販売枠なので無視)。
// 使い方: 🔄全リセット → これで下書き生成 → 内容確認 → EMSリストへ棚卸箱として貼る → ② → 🔎
const TANAOROSHI_CFG = {
  フォルダ名: 'Yahooの全在庫(商品名付き)', // Googleドライブのフォルダ名(最新のcsvを読む)
  出力シート: '棚卸箱下書き',
};

// ===== 棚卸箱をEMSリストへ追記(手貼りの代行) =====
// EMSリストは見出しが6行目で列順も下書きと異なるため、手貼りは事故りやすい。
// 見出し名(ステータス列/商品コード/数量/EMS到着日/EMS番号)で列を探して末尾に追記する。
// 他の列(No./照合キー/注文番号/数式列など)には一切書き込まない。
function 棚卸箱をEMSリストへ追記(){ 直列_(棚卸箱をEMSリストへ追記本体_); }
function 棚卸箱をEMSリストへ追記本体_(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  const rep=ss.getSheetByName(TANAOROSHI_CFG.出力シート);
  if(!rep || rep.getLastRow()<3){ ui.alert('「'+TANAOROSHI_CFG.出力シート+'」がありません。先に 📦 棚卸箱の下書きを生成 を実行してください。'); return; }
  const vals=rep.getRange(3,1,rep.getLastRow()-2,5).getDisplayValues()
    .filter(r=>String(r[2]||'').trim() && (Number(r[3])||0)>0); // 商品コードあり・数量>0の行だけ
  if(!vals.length){ ui.alert('下書きに追記できる行がありません'); return; }
  const 箱番号=String(vals[0][4]||'').trim();

  let sh;
  try{ sh=発注共有を開く_().getSheetByName(P_KAKUTEI_CFG.シート); }
  catch(e){ ui.alert('発注共有ファイルが開けません:\n'+e.message); return; }
  if(!sh){ ui.alert('発注共有に「'+P_KAKUTEI_CFG.シート+'」がありません'); return; }
  const hr=P_KAKUTEI_CFG.ヘッダー行;
  const head=sh.getRange(hr,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
  const f=n=>head.indexOf(n)+1; // 1始まり。0=見出し無し
  const cSt=f('ステータス列'), cCode=f('商品コード'), cQty=f('数量'), cEms=f('EMS番号');
  let cArr=f('EMS到着日'); if(!cArr) cArr=5; // 既定E列(整合チェックのスキャンと同じ既定)
  if(!cSt||!cCode||!cQty){ ui.alert('EMSリストの'+hr+'行目に ステータス列/商品コード/数量 の見出しが見つかりません'); return; }

  // 二重貼り防止: 同じ箱番号が既にあれば中止
  const last=sh.getLastRow();
  if(cEms && last>hr){
    const emsVals=sh.getRange(hr+1,cEms,last-hr,1).getDisplayValues();
    if(emsVals.some(r=>String(r[0]||'').trim()===箱番号)){
      ui.alert('EMSリストに既に「'+箱番号+'」の行があります。\n貼り直す場合は先にEMSリストからその行を削除してください(二重計上防止)。');
      return;
    }
  }

  const a=ui.alert('棚卸箱をEMSリストへ追記',
    vals.length+'行を「'+P_KAKUTEI_CFG.シート+'」の末尾('+(last+1)+'行目〜)に追記します。\n'+
    'EMS番号: '+箱番号+' ／ ステータス: 到着済\n\n続行しますか？',
    ui.ButtonSet.OK_CANCEL);
  if(a!==ui.Button.OK) return;

  const start=last+1;
  // 列ごとに書く(他の列の数式・チェックボックスに触らない)
  sh.getRange(start,cSt,vals.length,1).setValues(vals.map(function(r){return ['到着済'];}));
  sh.getRange(start,cArr,vals.length,1).setValues(vals.map(function(r){return [r[1]];}));
  sh.getRange(start,cCode,vals.length,1).setValues(vals.map(function(r){return [r[2]];}));
  sh.getRange(start,cQty,vals.length,1).setValues(vals.map(function(r){return [Number(r[3])||0];}));
  if(cEms) sh.getRange(start,cEms,vals.length,1).setValues(vals.map(function(r){return [r[4]];}));

  ui.alert('追記完了',
    vals.length+'行を追記しました('+start+'行目〜)。\n\n次: ②引き当て実行 → 🔎整合チェックで⚠️超過ゼロを確認',
    ui.ButtonSet.OK);
}

// ===== Yahoo在庫CSVの読み込み(棚卸と全件検算で共用) =====
// "a","b","c" 形式の1行をフィールド配列へ(引用符の "" エスケープ対応)
function CSV行分解_(line){
  const out=[]; let cur=''; let inQ=false;
  for(let i=0;i<line.length;i++){
    const ch=line[i];
    if(inQ){
      if(ch==='"'){ if(line[i+1]==='"'){ cur+='"'; i++; } else { inQ=false; } }
      else cur+=ch;
    } else {
      if(ch==='"') inQ=true;
      else if(ch===','){ out.push(cur); cur=''; }
      else cur+=ch;
    }
  }
  out.push(cur);
  return out;
}

// Yahoo在庫CSVのテキストを集計する(純粋関数)。列: code,name,sub-code,quantity,...(先頭行がヘッダー)
// sub-code末尾 a=即納(quantity=物理の自由在庫) / b=お取り寄せ(販売枠なので無視)
// 戻り: {a在庫:{基底コード:数量}, 商品名:{基底コード:商品名}, subなし:[…], 解析スキップ:n} または {error}
function YahooCSV集計_(text){
  const lines=String(text||'').split(/\r?\n/);
  if(lines.length<2) return {error:'CSVが空です'};
  const a在庫={};   // 基底コード -> 即納aの在庫数(自由在庫)
  const 商品名={};  // 基底コード -> 商品名
  const subなし=[]; // sub-code無しで在庫>0(a/b運用外。要確認)
  let 解析スキップ=0;
  for(let i=1;i<lines.length;i++){
    const line=lines[i]; if(!line || !line.trim()) continue;
    let r;
    try{ r=CSV行分解_(line); }catch(e){ 解析スキップ++; continue; }
    if(!r || r.length<4){ 解析スキップ++; continue; }
    const code=String(r[0]||'').trim(), name=String(r[1]||'').trim();
    const sub=String(r[2]||'').trim(), qty=Number(r[3])||0;
    if(!sub){ if(qty>0 && subなし.length<50) subなし.push(code+' x'+qty); continue; }
    const m=sub.match(/^(.+)([abAB])$/);
    if(!m) continue;
    if(m[2].toLowerCase()!=='a') continue; // b=お取り寄せ販売枠は在庫ではない
    const base=normCode_(m[1]);
    if(!base) continue;
    if(qty>0) a在庫[base]=(a在庫[base]||0)+qty;
    if(!商品名[base]) 商品名[base]=name;
  }
  return {a在庫, 商品名, subなし, 解析スキップ};
}

// Googleドライブの最新Yahoo在庫CSVを読んで集計する。
// Utilities.parseCsv は商品名内の不正な引用符などで全体が「テキストを解析できませんでした」に
// なるため使わない。1行ずつ寛容に分解し、壊れた行はスキップして続行する。
// 戻り: {a在庫,商品名,subなし,解析スキップ,fileName} または {error:'…'}
function Yahoo在庫を読む_(){
  const folders=DriveApp.getFoldersByName(TANAOROSHI_CFG.フォルダ名);
  if(!folders.hasNext()) return {error:'ドライブにフォルダ「'+TANAOROSHI_CFG.フォルダ名+'」が見つかりません'};
  const folder=folders.next();
  let newest=null;
  const it=folder.getFiles();
  while(it.hasNext()){
    const f=it.next();
    if(!/\.csv$/i.test(f.getName())) continue;
    if(!newest || f.getLastUpdated().getTime()>newest.getLastUpdated().getTime()) newest=f;
  }
  if(!newest) return {error:'フォルダ「'+TANAOROSHI_CFG.フォルダ名+'」にCSVがありません'};
  const r=YahooCSV集計_(newest.getBlob().getDataAsString('Shift_JIS'));
  if(r.error) return {error:r.error+': '+newest.getName()};
  r.fileName=newest.getName();
  return r;
}

// 棚卸箱の自動数量。EMSリストに一度も載ったことが無い商品(実績なし)は、予約・未発注なのに
// 「棚にある」ことにされてゴースト在庫を生む(実例: KRSJCM20-01S, JMEE-SPKZ系 2026-07-10)ため、
// 自動計上せず数量0で要確認に回す。現物が本当にあれば人が数量を手入力して貼る。
function 棚卸自動数量_(需要数, 未着数, EMS実績あり){
  const qty=Math.max(0, (Number(需要数)||0)-(Number(未着数)||0));
  if(qty>0 && !EMS実績あり) return {qty:0, 注意:'EMS到着実績なし(予約/未発注の可能性)。現物があれば数量を手入力'};
  return {qty:qty, 注意:''};
}

function 棚卸箱の下書きを生成(){ 直列_(棚卸箱の下書きを生成本体_); }
function 棚卸箱の下書きを生成本体_(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();

  // --- 1) Yahoo在庫CSV(フォルダ内で最新のもの)を読む ---
  const y=Yahoo在庫を読む_();
  if(y.error){ ui.alert(y.error); return; }
  const a在庫=y.a在庫, 商品名=y.商品名, subなし=y.subなし, 解析スキップ=y.解析スキップ;

  // --- 2) 受注明細: 未出荷の取り寄せ数(台湾/中国ルートは韓国便対象外なので除外) ---
  const recv=ss.getSheetByName(HIKIATE_CFG.受注);
  if(!recv){ ui.alert('「'+HIKIATE_CFG.受注+'」タブがありません'); return; }
  const M=列マップ_(recv);
  const R=recv.getDataRange().getValues();
  const 需要={};
  for(let i=M.hr;i<R.length;i++){
    const row=R[i];
    const ban=String(row[M.番号]||'').trim(); if(!ban) continue;
    if(区分_(row[M.選択肢])!=='取り寄せ') continue;
    const qty=Number(row[M.個数])||0; if(qty<=0) continue;
    const 別ルート=/台湾|中国/.test(String(row[M.選択肢]||'')) || (M.商品名>=0 && /台湾|中国/.test(String(row[M.商品名]||'')));
    if(別ルート) continue;
    const sku=M.SKU>=0? String(row[M.SKU]||'').trim():'';
    let base=normCode_(sku || String(row[M.コード]||''));
    if(/[AB]$/.test(base) && base.length>2) base=base.slice(0,-1); // 末尾のa/b枝番を落として基底へ
    if(!base) continue;
    需要[base]=(需要[base]||0)+qty;
    if(!商品名[base] && M.商品名>=0) 商品名[base]=String(row[M.商品名]||'').trim();
  }

  // --- 3) 未着箱の入荷予定数(発注共有EMSリスト): まだ棚に無い分は棚卸箱に入れない ---
  //        あわせて「EMSに載ったことがある商品」の集合を作る(過去の棚卸箱は実績に数えない)
  const 未着供給={};
  const EMS実績=new Set();
  let 実績確認可=false;
  try{
    const sh=発注共有を開く_().getSheetByName(P_KAKUTEI_CFG.シート);
    if(sh){
      const hr=P_KAKUTEI_CFG.ヘッダー行, last=sh.getLastRow();
      if(last>hr){
        const head=sh.getRange(hr,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
        const f=n=>head.indexOf(n);
        const cSt=f('ステータス列'), cCode=f('商品コード'), cQty=f('数量'), cEms=f('EMS番号');
        if(cCode>=0){
          実績確認可=true;
          sh.getRange(hr+1,1,last-hr,sh.getLastColumn()).getDisplayValues().forEach(r=>{
            const base=normCode_(r[cCode]); if(!base) return;
            const ems=cEms>=0? String(r[cEms]||'').trim():'';
            if(!/^棚卸/.test(ems)) EMS実績.add(base); // 実績=本物のEMS行のみ(棚卸箱は数えない)
            const st=cSt>=0? String(r[cSt]||'').trim():'';
            if(st==='到着済'||st==='在庫反映済み') return; // 着いている分は対象外(未着だけ)
            const q=cQty>=0? (Number(r[cQty])||0):1;
            未着供給[base]=(未着供給[base]||0)+q;
          });
        }
      }
    }
  }catch(e){ /* 発注共有が読めない時は未着差引きなし(棚卸箱が大きめに出る=安全側ではないので注意書きを出す) */ }

  // --- 4) 棚卸箱数量 = max(0, 需要 - 未着供給)。待ち注文の確保分だけ(自由在庫はYahooが正) ---
  //        EMSに載ったことが無い商品は自動計上しない(予約/未発注のゴースト防止。数量0で要確認)
  const today=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd');
  const 箱番号='棚卸'+Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMdd');
  const out=[];
  let 要確認数=0;
  Object.keys(需要).sort().forEach(base=>{
    const a=a在庫[base]||0, d=需要[base]||0, m=未着供給[base]||0;
    const g=棚卸自動数量_(d, m, 実績確認可? EMS実績.has(base) : true);
    if(g.qty<=0 && !g.注意) return;
    if(g.注意) 要確認数++;
    out.push(['到着済', today, base, g.qty, 箱番号, a, d, m, 商品名[base]||'', g.注意]);
  });

  // --- 5) 出力 ---
  let rep=ss.getSheetByName(TANAOROSHI_CFG.出力シート); if(!rep) rep=ss.insertSheet(TANAOROSHI_CFG.出力シート);
  rep.clearContents();
  rep.getRange(1,1).setValue('棚卸箱下書き: '+Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm')
    +' / CSV: '+y.fileName+' / '+out.length+'コード'
    +(要確認数? ' ／ ⚠️EMS到着実績なし(数量0) '+要確認数+'件=現物があれば数量を手入力':'')
    +' ｜ 数量 = max(0, 未出荷取り寄せ - 未着入荷予定)＝待ち注文の確保分のみ(自由在庫はYahooが正)。'
    +'内容を確認して発注共有「EMSリスト」へ貼る(②はEMSリストしか読まない)');
  const HDR=['ステータス列','EMS到着日','商品コード','数量','EMS番号','(内訳)即納a在庫','(内訳)未出荷取り寄せ','(内訳)未着入荷予定','商品名','(注意)'];
  rep.getRange(2,1,1,HDR.length).setValues([HDR]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff').setFontSize(HIKIATE_CFG.字);
  rep.setFrozenRows(2);
  if(out.length){
    rep.getRange(3,1,out.length,HDR.length).setValues(out).setFontSize(HIKIATE_CFG.字);
  } else {
    rep.getRange(3,1).setValue('(対象なし)');
  }
  let cur=3+Math.max(out.length,1)+1;
  if(subなし.length){
    rep.getRange(cur,1).setValue('■ 要確認: sub-code無しで在庫>0(a/b運用外のため棚卸箱に入れていない) '+subなし.length+'件: '+subなし.slice(0,20).join('、')).setFontSize(HIKIATE_CFG.字);
  }
  ss.setActiveSheet(rep);
  ui.alert('棚卸箱下書き 完成',
    out.length+'コードを出力しました（CSV: '+y.fileName+'）。\n'+
    (解析スキップ? '※解析できず読み飛ばした行: '+解析スキップ+'行\n':'')+
    (要確認数? '⚠️EMS到着実績なしのため数量0にした商品: '+要確認数+'件\n'+
      '　(予約/未発注の可能性。現物が本当にあるものだけ数量を手入力してください)\n':'')+'\n'+
    '数量は「待ち注文の確保分」だけです（自由在庫はYahoo在庫DBが正なので持ち込みません）。\n\n'+
    '次の手順:\n'+
    '1) 一覧を確認（⚠️整合チェックの数量超過に出た商品は現物と突き合わせて数量を直す）\n'+
    '2) 📦 棚卸箱をEMSリストへ追記（数量0の行は貼られません）\n'+
    '3) ②引き当て実行 → 🔎整合チェックで検証',
    ui.ButtonSet.OK);
}

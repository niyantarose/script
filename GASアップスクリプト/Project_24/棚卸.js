// ===== 棚卸箱の下書きを生成(Yahoo在庫CSV × 未出荷注文) =====
// 「机上棚卸」: 現物を数える代わりに、
//   棚卸箱数量(基底コード) = Yahoo即納a在庫(自由在庫の正) + max(0, 未出荷の取り寄せ数 - 未着箱の入荷予定数)
// を計算して、発注共有EMSリストに貼る「棚卸箱」の下書きを作る。
// Yahoo在庫CSV: ストアクリエイターProの在庫CSV(code,name,sub-code,quantity,… / CP932)。
//   sub-code 末尾 a=即納(quantity=物理の自由在庫) / b=お取り寄せ(販売枠なので無視)。
// 使い方: 🔄全リセット → これで下書き生成 → 内容確認 → EMSリストへ棚卸箱として貼る → ② → 🔎
const TANAOROSHI_CFG = {
  フォルダ名: 'Yahooの全在庫(商品名付き)', // Googleドライブのフォルダ名(最新のcsvを読む)
  出力シート: '棚卸箱下書き',
};

function 棚卸箱の下書きを生成(){ 直列_(棚卸箱の下書きを生成本体_); }
function 棚卸箱の下書きを生成本体_(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();

  // --- 1) Yahoo在庫CSV(フォルダ内で最新のもの)を読む ---
  const folders=DriveApp.getFoldersByName(TANAOROSHI_CFG.フォルダ名);
  if(!folders.hasNext()){ ui.alert('ドライブにフォルダ「'+TANAOROSHI_CFG.フォルダ名+'」が見つかりません'); return; }
  const folder=folders.next();
  let newest=null;
  const it=folder.getFiles();
  while(it.hasNext()){
    const f=it.next();
    if(!/\.csv$/i.test(f.getName())) continue;
    if(!newest || f.getLastUpdated().getTime()>newest.getLastUpdated().getTime()) newest=f;
  }
  if(!newest){ ui.alert('フォルダ「'+TANAOROSHI_CFG.フォルダ名+'」にCSVがありません'); return; }

  const text=newest.getBlob().getDataAsString('Shift_JIS');
  const rows=Utilities.parseCsv(text);
  if(!rows.length){ ui.alert('CSVが空です: '+newest.getName()); return; }

  // 列: code,name,sub-code,quantity,...(先頭行がヘッダー)
  const a在庫={};   // 基底コード -> 即納aの在庫数(自由在庫)
  const 商品名={};  // 基底コード -> 商品名
  const subなし=[]; // sub-code無しで在庫>0(a/b運用外。要確認)
  for(let i=1;i<rows.length;i++){
    const r=rows[i]; if(!r || r.length<4) continue;
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
  const 未着供給={};
  try{
    const sh=発注共有を開く_().getSheetByName(P_KAKUTEI_CFG.シート);
    if(sh){
      const hr=P_KAKUTEI_CFG.ヘッダー行, last=sh.getLastRow();
      if(last>hr){
        const head=sh.getRange(hr,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
        const f=n=>head.indexOf(n);
        const cSt=f('ステータス列'), cCode=f('商品コード'), cQty=f('数量');
        if(cCode>=0){
          sh.getRange(hr+1,1,last-hr,sh.getLastColumn()).getDisplayValues().forEach(r=>{
            const st=cSt>=0? String(r[cSt]||'').trim():'';
            if(st==='到着済'||st==='在庫反映済み') return; // 着いている分は対象外(未着だけ)
            const base=normCode_(r[cCode]); if(!base) return;
            const q=cQty>=0? (Number(r[cQty])||0):1;
            未着供給[base]=(未着供給[base]||0)+q;
          });
        }
      }
    }
  }catch(e){ /* 発注共有が読めない時は未着差引きなし(棚卸箱が大きめに出る=安全側ではないので注意書きを出す) */ }

  // --- 4) 棚卸箱数量 = a在庫 + max(0, 需要 - 未着供給) ---
  const 全コード={};
  Object.keys(a在庫).forEach(k=>全コード[k]=1);
  Object.keys(需要).forEach(k=>全コード[k]=1);
  const today=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd');
  const 箱番号='棚卸'+Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMdd');
  const out=[];
  Object.keys(全コード).sort().forEach(base=>{
    const a=a在庫[base]||0, d=需要[base]||0, m=未着供給[base]||0;
    const 確保=Math.max(0, d-m);
    const qty=a+確保;
    if(qty<=0) return;
    out.push(['到着済', today, base, qty, 箱番号, a, d, m, 商品名[base]||'']);
  });

  // --- 5) 出力 ---
  let rep=ss.getSheetByName(TANAOROSHI_CFG.出力シート); if(!rep) rep=ss.insertSheet(TANAOROSHI_CFG.出力シート);
  rep.clearContents();
  rep.getRange(1,1).setValue('棚卸箱下書き: '+Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm')
    +' / CSV: '+newest.getName()+' / '+out.length+'コード'
    +' ｜ 数量 = Yahoo即納a在庫 + max(0, 未出荷取り寄せ - 未着入荷予定)。内容を確認して発注共有EMSリストへ貼る');
  const HDR=['ステータス列','EMS到着日','商品コード','数量','EMS番号','(内訳)即納a在庫','(内訳)未出荷取り寄せ','(内訳)未着入荷予定','商品名'];
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
    out.length+'コードを出力しました（CSV: '+newest.getName()+'）。\n\n'+
    '次の手順:\n'+
    '1) 一覧を確認（⚠️整合チェックの数量超過に出た商品は現物と突き合わせて数量を直す）\n'+
    '2) A〜E列を発注共有のEMSリストへ「棚卸箱」として貼る\n'+
    '3) ②引き当て実行 → 🔎整合チェックで検証',
    ui.ButtonSet.OK);
}

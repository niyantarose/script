// ===== 日本在庫の余りを ★Yahoo在庫変更(code/sub-code/quantity/mode) 形式で出力 =====
// 運用: この出力 → Z:\業務\■在庫管理\Yahoo在庫作業\★Yahoo在庫変更.xlsm へ貼ってYahoo反映 → 📦⑤で便を締める。
// (締めた後は日本在庫シートからその便の余りが消えるため、必ず締める前に出力する)
// mode「+」=加算のため、同じ便を2回出力して二重貼りしないこと。出力シートは毎回置換。

const YAHOO_EXPORT_CFG = Object.freeze({ 出力シート: 'Yahoo在庫変更出力' });

// 日本在庫の行を対象/除外へ振り分ける純粋関数。emsSetが空なら全便を対象にする。
// 除外: PromotionalItem(贈呈品)・受注番号形式コード(例10117508)・★コピペ・マスタ除外コード(人名等)は
// Yahoo在庫に足さない(承認済み仕様)
function Yahoo変更_対象行_(rows, emsSet, excludeSet){
  const 対象=[], 除外=[];
  (rows||[]).forEach(r=>{
    const code=String(r&&r.商品コード||'').trim(), ems=String(r&&r.EMS番号||'').trim();
    const qty=Number(r&&r.余り数)||0;
    if(!code || qty<=0) return;
    if(emsSet && emsSet.size && !emsSet.has(ems)) return;
    // 未着(先行)の余りは現物ではないためYahooへ足せない(箱が着いて到着済になってから出力する)
    const status=String(r&&r.状態||'').trim();
    if(status && status!=='到着済'){ 除外.push({商品コード:code,余り数:qty,EMS番号:ems,理由:'未着(先行)の余りはYahooへ足せない'}); return; }
    if(code==='★コピペ'){ 除外.push({商品コード:code,余り数:qty,EMS番号:ems,理由:'付属ポスター印(★コピペ)'}); return; }
    if(/promotional/i.test(code)){ 除外.push({商品コード:code,余り数:qty,EMS番号:ems,理由:'PromotionalItem(贈呈品)'}); return; }
    if(/^\d{7,}$/.test(code)){ 除外.push({商品コード:code,余り数:qty,EMS番号:ems,理由:'受注番号形式コード'}); return; }
    if(excludeSet && excludeSet.size && excludeSet.has(normCode_(code))){ 除外.push({商品コード:code,余り数:qty,EMS番号:ems,理由:'マスタ除外コード'}); return; }
    対象.push({商品コード:code,余り数:qty,EMS番号:ems});
  });
  return {対象,除外};
}

// Yahoo全在庫CSVテキストから 正規化sub-code→{code,sub} の逆引きを作る(壊れた行はスキップ)。
// xlsmのcode列は親コードで、箱コードから機械的に導けないため、YahooのCSVを正とする
function Yahoo変更_サブコード逆引き_(text){
  const lines=String(text||'').split(/\r?\n/);
  if(lines.length<2) return {};
  const header=全件再計算_CSV行_(lines[0]).map((v,i)=>(i===0?String(v||'').replace(/^﻿/,''):String(v||'')).trim());
  const cCode=header.indexOf('code'), cSub=header.indexOf('sub-code');
  if(cCode<0 || cSub<0) return {};
  const map={};
  for(let i=1;i<lines.length;i++){
    if(!String(lines[i]||'').trim()) continue;
    const row=全件再計算_CSV行_(lines[i]);
    const code=String(row[cCode]||'').trim(), sub=String(row[cSub]||'').trim();
    if(!code || !sub) continue;
    const key=normCode_(sub);
    if(!(key in map)) map[key]={code:code,sub:sub};
  }
  return map;
}

// 対象行をxlsm形式へ変換。同一コードは合算し、sub-code=箱コード+a をYahoo CSVで逆引き。
// 逆引きできないコードは出力せず要確認へ(未出品・a/b運用外の可能性)
function Yahoo変更_行変換_(対象, 逆引き){
  const byCode={}, order=[];
  (対象||[]).forEach(r=>{
    const key=normCode_(r.商品コード);
    if(!(key in byCode)){ byCode[key]={箱コード:String(r.商品コード).trim(),qty:0}; order.push(key); }
    byCode[key].qty+=Number(r.余り数)||0;
  });
  const 行=[], 要確認=[];
  order.forEach(key=>{
    const e=byCode[key], hit=(逆引き||{})[normCode_(e.箱コード+'a')];
    if(hit) 行.push({code:hit.code,'sub-code':hit.sub,quantity:e.qty,mode:'+'});
    else 要確認.push({商品コード:e.箱コード,余り数:e.qty,理由:'Yahoo全在庫CSVに '+e.箱コード+'a が無い(未出品/a-b運用外?)'});
  });
  return {行,要確認};
}

function Yahoo在庫変更を出力(){ 直列_(Yahoo在庫変更を出力本体_); } // 書き込み系は直列_で排他

// preTargets: ⑤(便の締め)から対象EMSを引き継いで呼ばれる場合はプロンプトを出さない
function Yahoo在庫変更を出力本体_(preTargets){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  const sh=ss.getSheetByName(HIKIATE_CFG.純在庫);
  if(!sh || sh.getLastRow()<3){ ui.alert('「'+HIKIATE_CFG.純在庫+'」にデータがありません。②を実行してから使ってください。'); return; }
  let tokens;
  if(preTargets && preTargets.length){ tokens=preTargets.slice(); }
  else {
    const resp=ui.prompt('📤 Yahoo在庫変更を出力(締める便の余り)',
      '出力する便のEMS番号を入力してください(カンマ/スペース区切りで複数可)。\n'
      +'例: EG050152967KR\n\n'
      +'⚠️ mode「+」は加算です。同じ便を2回貼ると二重加算になります。\n'
      +'順序: この出力 → ★Yahoo在庫変更.xlsmへ貼ってYahoo反映 → 📦⑤で便を締める',
      ui.ButtonSet.OK_CANCEL);
    if(resp.getSelectedButton()!==ui.Button.OK) return;
    tokens=String(resp.getResponseText()||'').split(/[,、\s]+/).map(s=>s.trim()).filter(Boolean);
    if(!tokens.length){ ui.alert('EMS番号が空です'); return; }
  }
  const emsSet=new Set(tokens);

  // 日本在庫(1行目=最終引当メタ / 2行目=見出し / 3行目〜=データ)
  const values=sh.getRange(3,1,sh.getLastRow()-2,5).getDisplayValues();
  const japanRows=values.map(v=>({状態:v[0],到着日:v[1],商品コード:v[2],余り数:v[3],EMS番号:v[4]}));
  const 未知=tokens.filter(t=>!japanRows.some(r=>String(r.EMS番号||'').trim()===t));
  // 全部が一覧に無い時だけ中止(打ち間違い・未着便の防止)。複数便の一部に余りが無いのは正常
  if(未知.length===tokens.length){ ui.alert('出力を中止しました','日本在庫にこのEMS番号の行がありません: '+未知.join(', ')+'\n(未着の便は⑤前のYahoo出力対象になりません)',ui.ButtonSet.OK); return; }
  const 振り分け=Yahoo変更_対象行_(japanRows,emsSet,全件再計算_マスタ除外集合_());

  // 最新のYahoo全在庫CSVで親コードを逆引き
  let 逆引き={}, csvName='';
  const folders=DriveApp.getFoldersByName(TANAOROSHI_CFG.フォルダ名);
  if(folders.hasNext()){
    const it=folders.next().getFiles(); let newest=null;
    while(it.hasNext()){ const f=it.next(); if(!/\.csv$/i.test(f.getName())) continue;
      if(!newest || f.getLastUpdated().getTime()>newest.getLastUpdated().getTime()) newest=f; }
    if(newest){ 逆引き=Yahoo変更_サブコード逆引き_(newest.getBlob().getDataAsString('Shift_JIS')); csvName=newest.getName(); }
  }
  const 変換=Yahoo変更_行変換_(振り分け.対象,逆引き);

  // 出力シート(毎回置換)
  let out=ss.getSheetByName(YAHOO_EXPORT_CFG.出力シート); if(!out) out=ss.insertSheet(YAHOO_EXPORT_CFG.出力シート);
  out.clearContents();
  const data=[['code','sub-code','quantity','mode']].concat(変換.行.map(r=>[r.code,r['sub-code'],r.quantity,r.mode]));
  out.getRange(1,1,data.length,4).setValues(data);
  out.getRange(1,1,1,4).setFontWeight('bold');
  let cur=data.length+2;
  out.getRange(cur,1).setValue('作成: '+Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd HH:mm')
    +' / 便: '+tokens.join(', ')+' / 照合CSV: '+(csvName||'なし(全件が要確認になります)')); cur++;
  out.getRange(cur,1).setValue('手順: 上のA:D列を ★Yahoo在庫変更.xlsm(Z:\\業務\\■在庫管理\\Yahoo在庫作業)へ貼り付け → Yahoo反映 → 📦⑤で便を締める。mode「+」=加算のため二重貼り禁止'); cur+=2;
  if(変換.要確認.length){
    out.getRange(cur,1).setValue('■ 要確認(出力していません)'); out.getRange(cur,1).setFontWeight('bold'); cur++;
    const rows=変換.要確認.map(r=>[r.商品コード,r.余り数,r.理由]);
    out.getRange(cur,1,rows.length,3).setValues(rows); cur+=rows.length+1;
  }
  if(振り分け.除外.length){
    out.getRange(cur,1).setValue('■ 除外(仕様どおり出力しません)'); out.getRange(cur,1).setFontWeight('bold'); cur++;
    const rows=振り分け.除外.map(r=>[r.商品コード,r.余り数,r.EMS番号,r.理由]);
    out.getRange(cur,1,rows.length,4).setValues(rows);
  }
  ss.setActiveSheet(out);
  // ⑤が「出力済みか」を自動判定するための記録(対象行の内容署名)。内容が変われば再出力になる
  try{ PropertiesService.getDocumentProperties().setProperty('YAHOO出力記録',
    JSON.stringify({sig:Yahoo変更_内容署名_(振り分け.対象),at:String(new Date())})); }catch(e){}
  ui.alert('Yahoo在庫変更出力を作成しました',
    '出力: '+変換.行.length+'行 / 要確認: '+変換.要確認.length+'件 / 除外: '+振り分け.除外.length+'件'
    +(未知.length?'\n⚠️ 日本在庫にこのEMS番号の行がありません: '+未知.join(', '):'')
    +'\n\nA:D列をxlsmへ貼り付け→Yahoo反映→📦⑤で締めてください。',ui.ButtonSet.OK);
}

// 出力対象行の内容署名(EMS|コード|数量を整列連結)。⑤が「同じ内容を出力済みか」の判定に使う
function Yahoo変更_内容署名_(対象){
  return (対象||[]).map(r=>String(r.EMS番号||'')+'|'+String(r.商品コード||'')+'|'+(Number(r.余り数)||0)).sort().join(';');
}

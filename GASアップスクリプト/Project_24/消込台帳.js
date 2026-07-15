// ===== 消込台帳: 受注明細から消えた注文(=発送済みでCSV対象外)を記憶して在庫を消し込む =====
// GoQのCSVは未発送の注文しか含まないため、他の人が先に発送した注文は次の取込でシートから消える。
// そのままだと「その注文が使った商品」がEMS在庫の余り(日本在庫)として二重計上される。
// → 取込のたびに取り寄せ行のスナップショットを台帳に残し、消えた注文を「出荷済み?」として検知。
//   引き当て実行時に出荷済み分をEMS在庫から差し引く。
// キャンセルで消えた(在庫が残っている)場合だけ、台帳の状態セルを「キャンセル」に手で直す。

const KESHIKOMI_CFG = {
  シート: '消込台帳',
  HDR: ['受注番号','商品コード','SKU','個数','入荷日','状態','初回記録','消滅日','メモ'],
  有効日数: 60 // これより古い出荷済みは在庫と無関係とみなして差し引かない(古い台帳が新しい入荷を食わないように)
};

function 台帳比較値_(value){
  if(value instanceof Date) return isNaN(value.getTime())?'':'D:'+ymd_(value);
  const text=String(value==null?'':value).trim();
  if(/^(?:20\d{2}|\d{2})[-\/.]\d{1,2}[-\/.]\d{1,2}/.test(text)){
    const date=ymd_(text);
    if(/^20\d{2}-\d{2}-\d{2}$/.test(date)) return 'D:'+date;
  }
  return text;
}

function 行列同一_(left,right){
  const a=left||[], b=right||[];
  if(a.length!==b.length) return false;
  for(let r=0;r<a.length;r++){
    if(a[r].length!==b[r].length) return false;
    for(let c=0;c<a[r].length;c++){
      if(台帳比較値_(a[r][c])!==台帳比較値_(b[r][c])) return false;
    }
  }
  return true;
}

// 台帳を最新化: 受注明細の取り寄せ行と突き合わせ、消えた注文を「出荷済み?」にする
function 消込台帳更新_(){
  const ss=SpreadsheetApp.getActive(), cfg=KESHIKOMI_CFG;
  const recv=ss.getSheetByName(HIKIATE_CFG.受注);
  if(!recv) return {新規:0, 新規出荷済:0, 復活:0};
  const M=列マップ_(recv);
  const R=recv.getDataRange().getValues();

  // 現在の受注明細にいる取り寄せ行: key=受注番号|コード|SKU
  const now={};
  for(let i=M.hr;i<R.length;i++){
    const row=R[i];
    const ban=String(row[M.番号]||'').trim(); if(!ban) continue;
    if(区分_(row[M.選択肢])!=='取り寄せ') continue;
    const qty=Number(row[M.個数])||0; if(qty<=0) continue;
    const code=String(row[M.コード]||'').trim();
    const sku=M.SKU>=0? String(row[M.SKU]||'').trim():'';
    const 入荷=M.入荷>=0? row[M.入荷]:'';
    now[ban+'|'+code+'|'+sku]={ban,code,sku,qty,入荷};
  }
  // 受注明細が空(消した直後・取込前)のときは何もしない。全注文を誤って「出荷済み?」にしないため
  if(!Object.keys(now).length) return {新規:0, 新規出荷済:0, 復活:0};
  // 受注明細に生きている受注番号の集合。発送されていない注文が「出荷済み?」で残っていたら受注中に戻すのに使う
  // (タグ付きコード等で商品コード/SKUの表記が変わりフルキーが一致しなくても、受注番号で幽霊を解消する)
  const liveBans=new Set(); Object.keys(now).forEach(k=> liveBans.add(now[k].ban));

  let sh=ss.getSheetByName(cfg.シート);
  if(!sh){
    sh=ss.insertSheet(cfg.シート);
    sh.getRange(1,1,1,cfg.HDR.length).setValues([cfg.HDR])
      .setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    [110,180,180,60,100,110,100,100,200].forEach((w,c)=> sh.setColumnWidth(c+1,w));
  }
  const last=sh.getLastRow();
  const vals= last>1? sh.getRange(2,1,last-1,cfg.HDR.length).getValues():[];
  const today=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd');

  const seen=new Set(); let 新規出荷済=0, 復活=0;
  const out=vals.map(r=>{
    const ban=String(r[0]||'').trim(), code=String(r[1]||'').trim(), sku=String(r[2]||'').trim();
    const key=ban+'|'+code+'|'+sku; seen.add(key);
    let qty=r[3], 入荷=r[4], 状態=String(r[5]||'').trim(), 初回=r[6], 消滅=r[7], メモ=r[8];
    const cur=now[key];
    if(cur){
      // まだ受注明細にいる → 最新値で更新。一度消えて戻ってきたら「受注中」へ復帰(キャンセル指定だけは人の判断を尊重)
      if(状態.indexOf('出荷済み')===0){ 復活++; }
      if(状態!=='キャンセル'){ 状態='受注中'; 消滅=''; }
      qty=cur.qty; 入荷=cur.入荷;
    } else if(liveBans.has(ban) && 状態!=='キャンセル'){
      // フルキーは一致しないが受注番号は受注明細に生きている(タグ等でコード表記が変わっただけ)=発送されていない。
      // 誤検知の「出荷済み?」を受注中に戻す(幽霊が今回の箱を食って本物が宙に出るのを防ぐ)
      if(状態.indexOf('出荷済み')===0){ 復活++; 状態='受注中'; 消滅=''; }
    } else if(状態==='受注中' || 状態===''){
      状態='出荷済み?'; 消滅=today; 新規出荷済++; // 消えた=発送済みの可能性大。キャンセルなら手で直す
    }
    return [ban,code,sku,qty,入荷,状態,初回,消滅,メモ];
  });
  let 新規=0;
  Object.keys(now).forEach(key=>{
    if(seen.has(key)) return;
    const c=now[key];
    out.push([c.ban,c.code,c.sku,c.qty,c.入荷,'受注中',today,'','']);
    新規++;
  });

  out.sort((a,b)=> String(b[0]).localeCompare(String(a[0])) || String(a[1]).localeCompare(String(b[1]))); // 新しい注文を上に
  if(!行列同一_(vals,out)){
    const clearRows=Math.max(vals.length,out.length);
    if(clearRows) sh.getRange(2,1,clearRows,cfg.HDR.length).clearContent().setBackground(null);
    if(out.length){
      sh.getRange(2,1,out.length,cfg.HDR.length).setValues(out);
      const bg=out.map(r=>{ const st=String(r[5]||'');
        const c= st.indexOf('出荷済み')===0? '#efefef' : st==='キャンセル'? '#f4cccc' : null;
        return new Array(cfg.HDR.length).fill(c); });
      sh.getRange(2,1,out.length,cfg.HDR.length).setBackgrounds(bg);
      sh.getRange(2,5,out.length,1).setNumberFormat('yyyy-mm-dd');
    }
  }
  return {新規, 新規出荷済, 復活};
}

// 引き当てで在庫から差し引くべき「出荷済み」行を返す → [{ban,code,sku,qty,入荷日,基準日}]
// 状態が「出荷済み」で始まる行(自動の「出荷済み?」/手動確定の「出荷済み」どちらも)。
// 入荷日(なければ消滅日)が有効日数より古いものは、もう今の在庫とは無関係なので除外。
// 入荷日=その注文が実際に受けた箱の到着日。呼び出し側は「入荷日=箱の到着日」が一致する箱だけに紐付けること
// (一致しない出荷済みは別の箱/即納在庫から出た分なので、新しく着いた箱を食わせない)。
// 基準日=発送日の目安(消滅日→なければ初回記録)。入荷日なし行を「発送日以前に到着した箱」だけに食わせる判定用。
function 消込台帳_出荷済み行_(){
  const ss=SpreadsheetApp.getActive(), cfg=KESHIKOMI_CFG;
  const sh=ss.getSheetByName(cfg.シート);
  if(!sh || sh.getLastRow()<2) return [];
  const vals=sh.getRange(2,1,sh.getLastRow()-1,cfg.HDR.length).getValues();
  const limit=new Date(); limit.setDate(limit.getDate()-cfg.有効日数);
  const out=[];
  vals.forEach(r=>{
    if(String(r[5]||'').trim().indexOf('出荷済み')!==0) return;
    const qty=Number(r[3])||0; if(qty<=0) return;
    const base=r[4]||r[7]; // 入荷日→なければ消滅日
    const d= base instanceof Date? base : new Date(String(base||''));
    if(!isNaN(d.getTime()) && d.getTime()<limit.getTime()) return;
    out.push({ban:String(r[0]||'').trim(), code:String(r[1]||'').trim(), sku:String(r[2]||'').trim(), qty, 入荷日:r[4], 基準日:r[7]||r[6], メモ:String(r[8]||'').trim()});
  });
  return out;
}

// 台帳のメモ列に「取り置き」と書かれた出荷済み行か。
// 開始前の取り置き在庫（EMS箱の外）から出荷した分は、今回の箱を消費させないための人為マーク。
function 取り置き出荷_(l){ return /取り?置き|取置/.test(String(l&&l.メモ||'')); }

// ===== 移行用: 発送済み注文CSVを消込台帳へ一括登録 =====
// GoQで「出荷日を期間指定」して出力した発送済みCSVを、いつものドライブフォルダに入れてから実行。
// CSV内の「取り寄せ」行を全部「出荷済み」として台帳へ追加する(受注明細は触らない)。
// 台帳が生まれる前に発送されて自動検知できなかった分(移行期の隙間)を一括で埋めるためのもの。
function 消込台帳_発送済みCSV取込(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi(), gcfg=GOQ_CFG, cfg=KESHIKOMI_CFG;

  // いつものGoQフォルダから最新CSVを選ぶ(取込前にファイル名を確認できる)
  let folder; try{ folder=DriveApp.getFolderById(gcfg.フォルダID); }
  catch(e){ ui.alert('フォルダが見つからへん\n'+gcfg.フォルダID); return; }
  const it=folder.getFiles(); let latest=null;
  while(it.hasNext()){ const f=it.next();
    const isCsv=/\.csv$/i.test(f.getName())||f.getMimeType()==='text/csv';
    if(!isCsv) continue;
    if(!latest||f.getLastUpdated()>latest.getLastUpdated()) latest=f; }
  if(!latest){ ui.alert('フォルダにCSVが無いで'); return; }
  const upd=Utilities.formatDate(latest.getLastUpdated(),'Asia/Tokyo','MM/dd HH:mm');
  const ans=ui.alert('発送済みCSV→消込台帳(移行用)',
    'ファイル：'+latest.getName()+'\n更新：'+upd+'\n\n'+
    'このCSVの「取り寄せ」行を、消込台帳へ「出荷済み」として追加します。\n'+
    '※受注明細は変更しません。\n'+
    '※GoQで出荷日を期間指定して出力した【発送済みだけのCSV】か確認してから押してな。ええ？',
    ui.ButtonSet.OK_CANCEL);
  if(ans!==ui.Button.OK) return;

  const rows=Utilities.parseCsv(latest.getBlob().getDataAsString(gcfg.文字コード));
  if(rows.length<2){ ui.alert('CSVが空やで'); return; }
  const head=rows[0].map(v=>String(v||'').trim());
  const f=(...names)=>{ for(const n of names){ const i=head.indexOf(n); if(i>=0) return i; } return -1; };
  const c番号=f('受注番号'), cコード=f('商品コード'), cSKU=f('商品SKU','SKU'),
        c個数=f('個数'), c選択肢=f('項目・選択肢','項目選択肢'), c入荷=f('入荷日'), c出荷=f('出荷日');
  if(c番号<0||cコード<0){ ui.alert('CSVに「受注番号」「商品コード」の列が見つからへん'); return; }

  // 台帳シート(無ければ作成)と既存キー
  let sh=ss.getSheetByName(cfg.シート);
  if(!sh){
    sh=ss.insertSheet(cfg.シート);
    sh.getRange(1,1,1,cfg.HDR.length).setValues([cfg.HDR])
      .setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    [110,180,180,60,100,110,100,100,200].forEach((w,c)=> sh.setColumnWidth(c+1,w));
  }
  const last=sh.getLastRow();
  const existing=new Set();
  if(last>1){ sh.getRange(2,1,last-1,3).getValues().forEach(r=>
    existing.add(String(r[0]).trim()+'|'+String(r[1]).trim()+'|'+String(r[2]).trim())); }

  const today=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd');
  const add=[]; let 既存=0, 対象外=0;
  for(let i=1;i<rows.length;i++){
    const r=rows[i];
    const ban=String(r[c番号]||'').replace(/^niyantarose-/i,'').trim(); if(!ban) continue;
    if(c選択肢>=0 && 区分_(r[c選択肢])!=='取り寄せ'){ 対象外++; continue; } // 即納などは在庫と無関係
    const qty=Number(r[c個数])||0; if(qty<=0) continue;
    const code=String(r[cコード]||'').trim();
    const sku=cSKU>=0? String(r[cSKU]||'').trim():'';
    const key=ban+'|'+code+'|'+sku;
    if(existing.has(key)){ 既存++; continue; } // もう台帳にいる(自動検知済み等)は触らない
    const 入荷v=(c入荷>=0 && String(r[c入荷]||'').trim())? r[c入荷] : (c出荷>=0? r[c出荷] : '');
    add.push([ban,code,sku,qty,入荷v,'出荷済み',today,today,'移行取込('+latest.getName()+')']);
    existing.add(key);
  }
  if(add.length){
    const start=sh.getLastRow()+1;
    sh.getRange(start,1,add.length,cfg.HDR.length).setValues(add)
      .setBackgrounds(add.map(()=>new Array(cfg.HDR.length).fill('#efefef')));
    sh.getRange(start,5,add.length,1).setNumberFormat('yyyy-mm-dd');
  }
  ss.setActiveSheet(sh);
  ui.alert('移行取込 完了',
    '出荷済みとして追加：'+add.length+'件\n既に台帳にいた：'+既存+'件\n取り寄せ以外(対象外)：'+対象外+'行\n\n'+
    'このあと「② 引き当て実行」で在庫から差し引かれます。',
    ui.ButtonSet.OK);
}

// ===== 注文をキャンセル扱いにする =====
// GoQでキャンセルになった注文の後始末を1発で:
//  ①消込台帳の該当行を「キャンセル」へ(出荷済み?の誤検知を訂正。在庫を食わなくなる)
//  ②EMSリストのP列からその番号の名指しを除去 ③引当履歴をキャンセル済みに
//  ④今回入荷EMSの在庫の該当行を連動更新
// 仕上げに「② 引き当て実行」を回すと、その分の在庫が次の注文/余りに解放される。
function 注文をキャンセル扱いにする(){ 直列_(注文をキャンセル扱い本体_); }
function 注文をキャンセル扱い本体_(){
  const ui=SpreadsheetApp.getUi();
  const resp=ui.prompt('🚫 注文をキャンセル扱いにする',
    'キャンセルになった受注番号を入力（複数はカンマ区切り）',ui.ButtonSet.OK_CANCEL);
  if(resp.getSelectedButton()!==ui.Button.OK) return;
  const bans=String(resp.getResponseText()||'').split(/[,、\s]+/).map(s=>s.trim()).filter(s=>/^\d{5,}$/.test(s));
  if(!bans.length){ ui.alert('受注番号が読めません。'); return; }
  const r=キャンセル処理_(bans);
  ui.alert('🚫 キャンセル扱い 完了',
    bans.join(', ')+'\n\n'+r.results.join('\n')+'\n\nこのあと「② 引き当て実行」を回すと、その分の在庫が解放されます。',
    ui.ButtonSet.OK);
}

// キャンセル処理の本体(受注番号の配列を受け取る。ダイアログなし)。取込のキャンセル自動仕分けからも使う。
//  ①消込台帳を「キャンセル」へ ②EMSリストP列から名指し除去 ③引当履歴キャンセル ④今回入荷EMSの在庫を連動更新
function キャンセル処理_(bans){
  const ss=SpreadsheetApp.getActive();
  const results=[]; let 台帳更新=0, P除去=0;

  // ① 消込台帳の状態を「キャンセル」へ
  const tsh=ss.getSheetByName(KESHIKOMI_CFG.シート);
  if(tsh && tsh.getLastRow()>1){
    const n=tsh.getLastRow()-1;
    const tv=tsh.getRange(2,1,n,KESHIKOMI_CFG.HDR.length).getValues();
    const states=tv.map(row=>[String(row[5]||'')]);
    for(let i=0;i<n;i++){
      const ban=String(tv[i][0]||'').trim();
      if(bans.indexOf(ban)<0) continue;
      if(states[i][0].trim()==='キャンセル') continue;
      states[i][0]='キャンセル';
      台帳更新++;
    }
    if(台帳更新) tsh.getRange(2,6,n,1).setValues(states);
  }
  results.push('消込台帳: '+台帳更新+'行を「キャンセル」に');

  // ② EMSリストのP列から名指しを除去(状態問わず) + ③履歴キャンセル + ④今回入荷EMSの在庫を連動更新
  try{
    const ems=個別対応_EMSリスト_();
    if(ems.error){ results.push('P列: EMSリストが読めずスキップ'); }
    else{
      const recv=ss.getSheetByName(HIKIATE_CFG.受注);
      const M=recv? 列マップ_(recv):null, R=recv? recv.getDataRange().getValues():[];
      ems.rows.forEach(rw=>{
        if(!rw.注文番号) return;
        const entries=個別対応_P注文展開_(rw.注文番号, rw.数量);
        const kept=entries.filter(e=> bans.indexOf(e.ban)<0);
        if(kept.length===entries.length) return;
        const cell=ems.sh.getRange(rw.row, ems.c.注文番号+1);
        const text=個別対応_P注文整形_(kept, rw.数量);
        cell.setValue(text);
        cell.setBackground(/[,:：、]/.test(text)?'#fff2cc':null);
        entries.filter(e=>bans.indexOf(e.ban)>=0).forEach(e=>{
          try{ 引当履歴_キャンセル_(個別対応_履歴Rec_(rw, e.ban, e.qty, '個別引当'), e.qty); }catch(err){}
          P除去+=e.qty;
        });
        if(M) try{ 受注個別_台帳更新_(ss, rw, kept, R, M); }catch(err){}
      });
      results.push('P列の名指し除去: '+P除去+'個（引当履歴もキャンセル済みに）');
    }
  }catch(e){ results.push('P列: エラーでスキップ('+e.message+')'); }

  const ledger=取り置き台帳_読む_();
  const now=new Date(); let 取り置き更新=0;
  const rows=ledger.map(r=>{
    if(r.状態!=='取り置き中' || bans.indexOf(String(r.受注番号))<0) return r;
    取り置き更新++;
    return Object.assign({},r,{状態:'キャンセル戻し',戻し処理結果:'未確認',更新日時:now});
  });
  if(取り置き更新) 取り置き台帳_保存_(rows);

  return {results, 台帳更新, P除去};
}

// 取込で仕分けた「処理済(発送済み)」行を消込台帳へ「出荷済み(確定)」として登録する。
// header=CSV見出し行(配列)、rows=処理済のデータ行(配列の配列)。取り寄せ行だけが在庫に効く。
// 既に台帳にいる(受注番号|コード|SKU)は触らない。入荷日はCSVにあれば事実として使い、無ければ空。
function 消込台帳_処理済を登録_(header, rows){
  if(!rows || !rows.length) return {追加:0, 既存:0, 対象外:0};
  const ss=SpreadsheetApp.getActive(), cfg=KESHIKOMI_CFG;
  const H=header.map(v=>String(v||'').trim());
  const f=(...names)=>{ for(const n of names){ const i=H.indexOf(n); if(i>=0) return i; } return -1; };
  const c番号=f('受注番号'), cコード=f('商品コード'), cSKU=f('商品SKU','SKU'),
        c個数=f('個数'), c選択肢=f('項目・選択肢','項目選択肢'), c入荷=f('入荷日'), c出荷=f('出荷日');
  if(c番号<0||cコード<0) return {追加:0, 既存:0, 対象外:0};

  let sh=ss.getSheetByName(cfg.シート);
  if(!sh){
    sh=ss.insertSheet(cfg.シート);
    sh.getRange(1,1,1,cfg.HDR.length).setValues([cfg.HDR])
      .setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    [110,180,180,60,100,110,100,100,200].forEach((w,c)=> sh.setColumnWidth(c+1,w));
  }
  // 現在値を1回だけ読み、CSV処理済み以外の恒久行と今回CSVの行から次の完成配列を作る。
  // 完成配列が同一ならシートへ一切書かない。
  const currentLast=sh.getLastRow();
  const current=currentLast>1?sh.getRange(2,1,currentLast-1,cfg.HDR.length).getValues():[];
  const permanent=current.filter(r=>String(r[8]||'').trim()!=='CSV処理済');
  const existing=new Set();
  permanent.forEach(r=>existing.add(
    String(r[0]).trim()+'|'+String(r[1]).trim()+'|'+String(r[2]).trim()
  ));

  const today=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd');
  const limit=new Date(); limit.setDate(limit.getDate()-cfg.有効日数); // これより前に発送された分は今の箱と無関係
  const add=[]; let 既存=0, 対象外=0, 古い=0;
  rows.forEach(r=>{
    const ban=String(r[c番号]||'').replace(/^niyantarose-/i,'').trim(); if(!ban) return;
    if(c選択肢>=0 && 区分_(r[c選択肢])!=='取り寄せ'){ 対象外++; return; } // 即納などはEMS在庫と無関係
    const qty=Number(r[c個数])||0; if(qty<=0) return;
    // 出荷日が有効日数(60日)より前=既に別の箱から出て完結済み。在庫差し引き対象外なので登録しない(古い発送が今の箱を食わない)
    const shipY=ymd_(c出荷>=0? r[c出荷] : '');
    if(shipY){ const d=new Date(shipY); if(!isNaN(d.getTime()) && d.getTime()<limit.getTime()){ 古い++; return; } }
    const code=String(r[cコード]||'').trim();
    const sku=cSKU>=0? String(r[cSKU]||'').trim():'';
    const key=ban+'|'+code+'|'+sku;
    if(existing.has(key)){ 既存++; return; }
    const 入荷v=(c入荷>=0 && String(r[c入荷]||'').trim())? r[c入荷] : '';
    const 発送日=shipY||today; // 消滅日に出荷日を入れる→有効日数の窓が正しく効く
    add.push([ban,code,sku,qty,入荷v,'出荷済み',発送日,発送日,'CSV処理済']);
    existing.add(key);
  });
  const next=permanent.concat(add);
  if(!行列同一_(current,next)){
    const clearRows=Math.max(current.length,next.length);
    if(clearRows) sh.getRange(2,1,clearRows,cfg.HDR.length).clearContent().setBackground(null);
    if(next.length){
      sh.getRange(2,1,next.length,cfg.HDR.length).setValues(next);
      const bg=next.map(r=>{
        const st=String(r[5]||'');
        const color=st.indexOf('出荷済み')===0?'#efefef':st==='キャンセル'?'#f4cccc':null;
        return new Array(cfg.HDR.length).fill(color);
      });
      sh.getRange(2,1,next.length,cfg.HDR.length).setBackgrounds(bg);
      sh.getRange(2,5,next.length,1).setNumberFormat('yyyy-mm-dd');
    }
  }
  return {追加:add.length, 既存, 対象外, 古い};
}

// 消込台帳から「CSV処理済」で登録した行だけを消す(消滅日=今日で登録された古い残骸の掃除用)。
// 消えた注文の出荷済み?・手動キャンセルは残す。このあと処理済含むCSVを取り込むと最近の発送だけ登録し直される。
function 消込台帳のCSV処理済をクリア(){ 直列_(消込台帳のCSV処理済をクリア本体_); }
function 消込台帳のCSV処理済をクリア本体_(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi(), cfg=KESHIKOMI_CFG;
  const sh=ss.getSheetByName(cfg.シート);
  if(!sh || sh.getLastRow()<2){ ui.alert('消込台帳がありません。'); return; }
  const n=sh.getLastRow()-1;
  const vals=sh.getRange(2,1,n,cfg.HDR.length).getValues();
  const keep=vals.filter(r=> String(r[8]||'').trim()!=='CSV処理済');
  const removed=n-keep.length;
  if(!removed){ ui.alert('「CSV処理済」で登録された行はありません。'); return; }
  const mr=sh.getMaxRows();
  sh.getRange(2,1,mr-1,cfg.HDR.length).clearContent().setBackground(null);
  if(keep.length){
    sh.getRange(2,1,keep.length,cfg.HDR.length).setValues(keep);
    const bg=keep.map(r=>{ const st=String(r[5]||''); const c= st.indexOf('出荷済み')===0?'#efefef':st==='キャンセル'?'#f4cccc':null; return new Array(cfg.HDR.length).fill(c); });
    sh.getRange(2,1,keep.length,cfg.HDR.length).setBackgrounds(bg);
    sh.getRange(2,5,keep.length,1).setNumberFormat('yyyy-mm-dd');
  }
  ui.alert('🧹 消込台帳のCSV処理済をクリア',
    '削除: '+removed+'行 / 残り: '+keep.length+'行\n\n'+
    'このあと「① 受注明細を更新」(処理済を含むCSV)を取り込むと、出荷日が直近'+cfg.有効日数+'日の発送だけ登録し直されます。',
    ui.ButtonSet.OK);
}

// ===== 出荷済みの重複排除(読み取り集計用) =====
// 受注/台帳側の基底コード: SKU優先→normCode_(受注番号タグ除去)→末尾のa/b在庫枝番を落とす(棚卸の需要側と同じ規則)
function 受注基底コード_(sku, code){
  let base=normCode_(String(sku||'').trim() || String(code||''));
  if(/[AB]$/.test(base) && base.length>2) base=base.slice(0,-1);
  return base;
}

// 同じ発送が表記ゆれ(受注番号タグ付きコード/a・b枝番SKU)で複数行になっている場合、
// 「受注番号+基底コード」で1件(数量が最大の行)にまとめる。受注番号が無い行は素通し。
// 適用先は読み取り専用の集計(全件検算レポート)のみ。②の消費ロジックには適用しない
// (実測で調整済みの挙動を変えないため)。
function 出荷済み重複排除_(rows){
  const best={}; const 素通し=[];
  (rows||[]).forEach(r=>{
    const ban=String(r.ban||'').trim();
    if(!ban){ 素通し.push(r); return; }
    const key=ban+'|'+受注基底コード_(r.sku, r.code);
    if(!(key in best) || (Number(r.qty)||0)>(Number(best[key].qty)||0)) best[key]=r;
  });
  return 素通し.concat(Object.keys(best).map(k=>best[k]));
}

// メニュー用: 台帳を更新してシートを開く
function 消込台帳を更新(){
  const r=消込台帳更新_();
  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName(KESHIKOMI_CFG.シート);
  if(sh) ss.setActiveSheet(sh);
  ss.toast('消込台帳: 新規'+r.新規+'件 / 出荷済み検知'+r.新規出荷済+'件 / 復活'+r.復活+'件','🧾消込台帳',6);
}

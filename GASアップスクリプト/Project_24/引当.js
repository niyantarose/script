// ===== GoQ受注CSV取込 設定 =====
const GOQ_CFG = {
  フォルダID: '1muJbMEtJzDIAJ4u4H31bjGhjKOep3AVo', // Googleドライブ「GoQsystem受注データ」
  受注シート: '受注明細',
  文字コード: 'windows-31j', // GoQのCSVはCP932。丸数字①等のWindows拡張文字も化けないようMS932で読む
  フォントサイズ: 13          // 取込後のデータ行のフォントサイズ(縦は中央にそろえる)
};

// ===== ダニエルEMS取込 設定 =====
const DANIEL_CFG = {
  フォルダID: '1XnWTXomCDe5fbmBPRz7fNWzgD9KZbYKz', // Googleドライブ「ダニエルEMSファイル」
  シート: 'ダニエルEMS',
  // xlsxの列(0始): C=発送日 / J=商品名(商品内容) / L=商品コード / M=数量 / Q=対象BOX / R=EMS番号
  X_発送日: 2, X_商品名: 9, X_商品コード: 11, X_数量: 12, X_BOX: 16, X_EMS: 17
};

// スプレッドを開いたときにメニューを出す(共通 / メイン引当 / ダニエル引当 の3メニュー)
function onOpen(){
  const ui=SpreadsheetApp.getUi();

  // ── 共通(受注明細・商品コード引当・データ管理。メインもダニエルも使う) ──
  ui.createMenu('📥 受注・共通')
    .addItem('📥 最新CSVを取込(受注明細)', '取込_最新CSV')
    .addItem('🔢 商品コードで引当(商品コード入力シートの在庫で)', '商品コード引当')
    .addItem('🧾 消込台帳を更新(発送済みの検知/確認)', '消込台帳を更新')
    .addItem('🚫 注文をキャンセル扱いにする(受注番号指定)', '注文をキャンセル扱いにする')
    .addItem('🧹 消込台帳のCSV処理済をクリア(残骸掃除)', '消込台帳のCSV処理済をクリア')
    .addItem('🧾 発送済みCSVを台帳へ一括取込(移行用)', '消込台帳_発送済みCSV取込')
    .addItem('📝 P列に注文番号を自動記入(発注共有へ)', 'P列に注文番号を自動記入')
    .addItem('♻️ P列を書き直す(到着済の残骸を一掃)', 'P列を書き直す')
    .addItem('🧾 引当履歴シートを作成', '引当履歴シートを作成')
    .addItem('🧾 引当履歴へ過去データ取込', '引当履歴_過去データを取込')
    .addItem('🧹 個別ボタン画像を削除(図形ボタン運用)', '受注明細個別ボタン画像を削除')
    .addSeparator()
    .addItem('🙈 色なし注文を除外/全表示(切替)', '色なし注文の除外切替')
    .addItem('🧱 注文ごとの罫線を引き直す', '注文罫線を引く')
    .addItem('🔎 色つき注文だけフィルタ(抽出=○)', '色つき注文でフィルタ')
    .addItem('🔻 非表示の確認/全解除(フィルタ＋🙈)', 'フィルタ確認解除')
    .addSubMenu(ui.createMenu('🗑 データを消す')
      .addItem('受注明細', '受注明細を消す')
      .addItem('引当待ち', '引当待ちを消す')
      .addItem('部分在庫', '部分在庫を消す')
      .addItem('出荷GO未入金', '出荷GO未入金を消す')
      .addItem('希望日待ち', '希望日待ちを消す')
      .addItem('出荷可能', '出荷可能を消す')
      .addItem('日本在庫', '日本在庫を消す')
      .addItem('ダニエル出荷可能', 'ダニエル出荷を消す')
      .addItem('商品コード引当', '商品コード引当を消す')
      .addItem('照合レポート', '照合レポートを消す')
      .addItem('入荷日チェック', '入荷日チェックを消す')
      .addSeparator()
      .addItem('全部まとめて', '全データを消す'))
    .addToUi();

  // ── メイン引当(在庫=EMS在庫) ──
  ui.createMenu('🏠 メイン引当(EMS在庫)')
    .addItem('🔄 EMS在庫を更新(色クリア＋最新化)', 'EMS在庫を更新')
    .addItem('🧱 EMS在庫/今回入荷EMSの在庫：EMS番号ごとに罫線', '在庫EMS番号ごとに罫線を引く')
    .addItem('🎨 日本在庫を罫線+色分け(EMS番号ごと)', '日本在庫の罫線と色分け')
    .addItem('☑ EMS番号罫線ボタンを設置', '在庫EMS番号罫線ボタンを設置')
    .addSeparator()
    .addItem('① 前段階チェック(即納に水色＋罫線)', '前段階チェック_即納')
    .addItem('② 引き当て実行(EMS在庫・入荷日で判定)', '引当実行')
    .addItem('📦 到着済を在庫反映済みへ(便の締め)', '到着済を在庫反映済みへ')
    .addItem('⚖️ この便の引当をやり直す(到着日指定)', '便の引当をやり直す')
    .addItem('🔎 引当診断(受注番号で調べる)', '引当診断')
    .addItem('🔎 商品診断(商品コードで調べる)', '商品診断')
    .addItem('🔎 入荷日の整合チェック(誤記入の検出)', '入荷日整合チェック')
    .addItem('🧹 チェック一覧の入荷日をクリア', '入荷日チェック_一覧をクリア')
    .addItem('🔍 在庫照合レポート', '在庫照合レポート')
    .addToUi();

  // ── ダニエル引当(サブ。在庫=ダニエルEMS) ──
  ui.createMenu('🚚 ダニエル引当(サブ)')
    .addItem('📦 ダニエルEMS取込', '取込_ダニエルEMS')
    .addItem('📦 ダニエルEMS引当', 'ダニエルEMS引当')
    .addItem('🔬 ダニエルファイル診断', 'ダニエルファイル診断')
    .addToUi();

  try { 在庫EMS番号罫線ボタンを設置_(true); } catch(e) {}
  // 個別引当/キャンセルのボタンは手動の図形(スクリプト割り当て)運用。
  // 挿入画像はフリーズ行内で消える/描画されない事があるため自動設置はしない。
}

function onEdit(e){
  if (在庫EMS番号罫線ボタン編集_(e)) return;
}

function onSelectionChange(e){
  if (個別対応セルボタン選択_(e)) return;
}

// 受注明細のヘッダー行(「受注番号」がある行)を探す。見つからなければ1行目
function 受注ヘッダー行_(sh){
  const n=Math.min(sh.getLastRow()||1, 50), w=sh.getLastColumn()||1;
  const vals=sh.getRange(1,1,n,w).getValues();
  for(let i=0;i<vals.length;i++){
    if(vals[i].some(c=>String(c).trim()==='受注番号')) return i+1; // 1始まりの行番号
  }
  return 1;
}

// 指定見出し(label)がある行番号(1始)を返す。無ければ1
function ヘッダー行検索_(sh, label){
  const n=Math.min(sh.getLastRow()||1, 50), w=sh.getLastColumn()||1;
  const vals=sh.getRange(1,1,n,w).getValues();
  for(let i=0;i<vals.length;i++){ if(vals[i].some(c=>String(c).trim()===label)) return i+1; }
  return 1;
}

// xlsx(=zip)を直接展開して1枚目シートの2次元配列を返す(Drive変換を使わない=サイズ制限なし)
function _colIdx(s){ let n=0; for(let i=0;i<s.length;i++) n=n*26+(s.charCodeAt(i)-64); return n-1; }
function _xmlDec(s){ return String(s).replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&#(\d+);/g,(_,d)=>String.fromCharCode(+d)).replace(/&amp;/g,'&'); }
function xlsx読む_(blob){
  blob.setContentType('application/zip');
  const files=Utilities.unzip(blob);
  let sharedXml='', sheetXml='';
  files.forEach(f=>{ const n=f.getName();
    if(n==='xl/sharedStrings.xml') sharedXml=f.getDataAsString('UTF-8');
    else if(n==='xl/worksheets/sheet1.xml') sheetXml=f.getDataAsString('UTF-8'); });
  if(!sheetXml){ files.forEach(f=>{ if(!sheetXml && /^xl\/worksheets\/sheet\d+\.xml$/.test(f.getName())) sheetXml=f.getDataAsString('UTF-8'); }); }
  if(!sheetXml) throw new Error('シートのデータが見つからへん(xlsx構造が想定外)');
  // 共有文字列
  const shared=[]; const siRe=/<si\b[^>]*>([\s\S]*?)<\/si>/g; let m;
  while((m=siRe.exec(sharedXml))){ const tRe=/<t\b[^>]*>([\s\S]*?)<\/t>/g; let t,s=''; while((t=tRe.exec(m[1]))) s+=t[1]; shared.push(_xmlDec(s)); }
  // セル
  const rows=[]; const rowRe=/<row\b[^>]*>([\s\S]*?)<\/row>/g; let r;
  while((r=rowRe.exec(sheetXml))){
    const cells=[]; const cRe=/<c\b([^>]*)>([\s\S]*?)<\/c>/g; let c;
    while((c=cRe.exec(r[1]))){
      const attrs=c[1], inner=c[2];
      const refM=/r="([A-Z]+)\d+"/.exec(attrs); if(!refM) continue;
      const col=_colIdx(refM[1]); const tM=/t="([^"]+)"/.exec(attrs); const t=tM?tM[1]:'';
      const vM=/<v\b[^>]*>([\s\S]*?)<\/v>/.exec(inner);
      let v='';
      if(t==='s'){ v=vM? (shared[Number(vM[1])]||'') : ''; }
      else if(t==='inlineStr'){ const isM=/<t\b[^>]*>([\s\S]*?)<\/t>/.exec(inner); v=isM?_xmlDec(isM[1]):''; }
      else { v=vM? _xmlDec(vM[1]) : ''; }
      cells[col]=v;
    }
    for(let i=0;i<cells.length;i++) if(cells[i]===undefined) cells[i]='';
    rows.push(cells);
  }
  return rows;
}

// xlsx直読みが失敗したとき用: Driveで既存ファイルをGoogleシートに変換して読む
// (サーバー側コピー変換なのでアップロード無し=413回避)。読み終えたら一時ファイルを削除
function ダニエル_変換読み_(fileId){
  let copyId=null;
  try{
    const copy=Drive.Files.copy({title:'_tmp_daniel_conv_'+Date.now()}, fileId, {convert:true}); // xlsx→Googleシート変換
    copyId=copy.id;
    return SpreadsheetApp.openById(copyId).getSheets()[0].getDataRange().getValues();
  } finally {
    if(copyId){ try{ Drive.Files.remove(copyId); }catch(e){} } // 一時ファイル掃除
  }
}

// 🔬 ダニエルファイル診断: フォルダ最新xlsxの中身(先頭バイト/zip構成/先頭テキスト)を見て本当の形式を特定する
function ダニエルファイル診断(){
  const ui=SpreadsheetApp.getUi(), cfg=DANIEL_CFG;
  let folder; try{ folder=DriveApp.getFolderById(cfg.フォルダID); }catch(e){ ui.alert('フォルダ無し'); return; }
  const XLSX='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const it=folder.getFiles(); let latest=null;
  while(it.hasNext()){ const f=it.next(); const nm=f.getName(), mm=f.getMimeType();
    if(!(/\.xlsx$/i.test(nm) || mm===XLSX || mm===MimeType.GOOGLE_SHEETS)) continue;
    if(!latest || f.getLastUpdated()>latest.getLastUpdated()) latest=f; }
  if(!latest){ ui.alert('xlsx/Googleシートが無いで'); return; }
  const bytes=latest.getBlob().getBytes();
  const head=bytes.slice(0,4).map(b=>((b&0xff).toString(16)).padStart(2,'0')).join(' ');
  let 形式='不明';
  if(head.indexOf('50 4b')===0) 形式='ZIP(本物のxlsx候補)';
  else if(head.indexOf('d0 cf 11 e0')===0) 形式='旧Excel .xls(バイナリ)';
  else if(head.indexOf('3c')===0) 形式='HTML/XMLテキスト(Excel偽装の可能性)';
  let zipList=''; try{ const b2=latest.getBlob(); b2.setContentType('application/zip'); zipList=Utilities.unzip(b2).map(f=>f.getName()).slice(0,25).join(', '); }catch(e){ zipList='unzip不可: '+e.message; }
  let text=''; try{ text=latest.getBlob().getDataAsString('UTF-8').slice(0,400); }catch(e){ text='(テキスト化不可=バイナリ)'; }
  try{ if(!text||text.indexOf('�')>=0){ text=latest.getBlob().getDataAsString('Shift_JIS').slice(0,400); } }catch(e){}
  ui.alert('ダニエルファイル診断',
    'ファイル：'+latest.getName()+'\nサイズ：'+Math.round(bytes.length/1024)+'KB\nMIME：'+latest.getMimeType()+'\n'+
    '先頭4バイト：'+head+' → 判定：'+形式+'\nzip内ファイル：'+zipList+'\n\n先頭テキスト：\n'+text.replace(/</g,'‹').slice(0,400),
    ui.ButtonSet.OK);
}

// 📦 ダニエルEMS取込: フォルダ内の最新xlsxを直接読み→DANIEL BOX行だけ整形してタブへ
function 取込_ダニエルEMS(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi(), cfg=DANIEL_CFG;

  let folder;
  try{ folder=DriveApp.getFolderById(cfg.フォルダID); }
  catch(e){ ui.alert('「ダニエルEMSファイル」フォルダが見つからへん\n'+cfg.フォルダID); return; }
  // xlsx(またはGoogleシート)だけを対象に、更新が最新の1件を選ぶ。CSV等(Yahoo商品登録など)は無視
  const XLSX='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const it=folder.getFiles(); let latest=null;
  while(it.hasNext()){ const f=it.next(); const nm=f.getName(), mm=f.getMimeType();
    if(!(/\.xlsx$/i.test(nm) || mm===XLSX || mm===MimeType.GOOGLE_SHEETS)) continue; // xlsx/Googleシートのみ
    if(!latest || f.getLastUpdated()>latest.getLastUpdated()) latest=f; }
  if(!latest){ ui.alert('「'+folder.getName()+'」にxlsx(EMS発送リスト)が見つからへん。\nCSV等は対象外や。ダニエルのxlsxを入れてな'); return; }

  const upd=Utilities.formatDate(latest.getLastUpdated(),'Asia/Tokyo','MM/dd HH:mm');
  const ans=ui.alert('ダニエルEMS取込','ファイル：'+latest.getName()+'\n更新：'+upd+'\n\nDANIEL BOXの行を取り込みます。ええ？',ui.ButtonSet.OK_CANCEL);
  if(ans!==ui.Button.OK) return;

  // 最新がGoogleシートならそのまま読む。xlsxは①直接展開→②ダメならDrive変換。少ししか読めない時はDrive変換にフォールバック
  let data=null, 方法='';
  try{
    if(latest.getMimeType()===MimeType.GOOGLE_SHEETS){
      data=SpreadsheetApp.openById(latest.getId()).getSheets()[0].getDataRange().getValues(); 方法='Googleシート直読み';
    } else {
      try{ data=xlsx読む_(latest.getBlob()); 方法='xlsx直読み'; }             // まずxlsxを直接展開して読む(速い)
      catch(e1){ data=ダニエル_変換読み_(latest.getId()); 方法='Drive変換(直読み失敗:'+e1.message+')'; } // ダメならDriveで自動変換
      // 直読みで極端に少ない行数しか取れない=構造が特殊 → Drive変換で読み直す
      if(方法==='xlsx直読み' && (!data || data.length<5)){
        try{ const d2=ダニエル_変換読み_(latest.getId()); if(d2 && d2.length>(data?data.length:0)){ data=d2; 方法='Drive変換(直読みは'+(data?data.length:0)+'行のみ)'; } }catch(e2){}
      }
    }
  }catch(e){
    ui.alert('読み込みに失敗：'+e.message+'\n(xlsx直読み・自動変換の両方が失敗)'); return;
  }
  if(!data || !data.length){ ui.alert('データが空やで'); return; }

  // ヘッダー行(C列に「발송일」がある行)を探す→その下からデータ
  let h=-1;
  for(let i=0;i<Math.min(data.length,30);i++){ if(String(data[i][cfg.X_発送日]||'').trim()==='발송일'){ h=i; break; } }
  const start = h>=0? h+1 : 4;

  // 全行を抽出 → 発送日の最新2つに絞る → 発送日ごとにEMS番号でBOXNo連番(対象BOX=Q列そのまま)
  const all=[];
  for(let i=start;i<data.length;i++){
    const r=data[i];
    const 発送日=String(r[cfg.X_発送日]||'').trim();
    if(!発送日) continue;
    all.push({
      発送日,
      name:String(r[cfg.X_商品名]||'').trim(),         // J列 商品名(商品内容)
      code:String(r[cfg.X_商品コード]||'').trim(),
      qty:Number(r[cfg.X_数量])||0,
      box:String(r[cfg.X_BOX]||'').trim(),            // Q列の値そのまま(DANIEL BOX or 空白)
      ems:String(r[cfg.X_EMS]||'').replace(/\s+/g,'') // R列 スペース除去
    });
  }
  const dates=[...new Set(all.map(a=>a.発送日))].sort().reverse().slice(0,2); // 最新2つの発送日(YYMMDDは文字列ソートで時系列)
  const dateSet=new Set(dates);
  const boxNo={}, counter={}; const rows=[];
  all.forEach(a=>{
    if(!dateSet.has(a.発送日)) return;
    const key=a.発送日+'|'+a.ems;
    if(!(key in boxNo)){ counter[a.発送日]=(counter[a.発送日]||0)+1; boxNo[key]=counter[a.発送日]; } // 発送日ごとに1,2,3...
    rows.push([boxNo[key], a.発送日, a.name, a.code, a.qty, a.box, a.ems, '']); // BOXNo/発送日/商品名/商品コード/数量/対象BOX/EMS番号/入荷後ステータス
  });

  // 0件のときは原因調査の情報を出す(データが読めているか・発送日が取れているか)
  if(!rows.length){
    const 種類=[...new Set(all.map(a=>a.発送日))];
    ui.alert('取込0件（原因調査）',
      'ファイル：'+latest.getName()+'\n'+
      'MIME：'+latest.getMimeType()+'\n'+
      '読込方法：'+方法+'\n'+
      '読めた総行数(data)：'+data.length+'\n'+
      'ヘッダー(発送일検出)：'+(h>=0?(h+1)+'行目':'見つからず→'+(start+1)+'行目から読む')+'\n'+
      '発送日ありの行数(all)：'+all.length+'\n'+
      '発送日の種類：'+(種類.slice(0,6).join(', ')||'(なし)')+'\n'+
      '最新2発送日：'+(dates.join(', ')||'(なし)')+'\n\n'+
      'この内容を教えてください。', ui.ButtonSet.OK);
    return;
  }

  // ダニエルEMSタブへ(ヘッダーは検出、その下のデータだけ入替・書式維持)
  let sh=ss.getSheetByName(cfg.シート); if(!sh) sh=ss.insertSheet(cfg.シート);
  const hdrRow=ヘッダー行検索_(sh,'BOXNo'), startRow=hdrRow+1;
  const mr=sh.getMaxRows(), mc=sh.getMaxColumns();
  if(mr>=startRow){
    const reset=sh.getRange(startRow,1,mr-startRow+1,mc);
    reset.clearContent();
    reset.setBackgrounds(Array.from({length:mr-startRow+1},()=>new Array(mc).fill(null)));
    reset.setBorder(false,false,false,false,false,false);
  }
  if(rows.length){
    const need=startRow+rows.length-1; // 行数が多いファイルでも書けるようシートを広げる
    if(sh.getMaxRows()<need) sh.insertRowsAfter(sh.getMaxRows(), need-sh.getMaxRows());
    sh.getRange(startRow,1,rows.length,8).setValues(rows)
      .setVerticalAlignment('middle').setFontSize(HIKIATE_CFG.字).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    sh.setRowHeights(startRow, rows.length, HIKIATE_CFG.行高);
  }
  ss.toast('ダニエルEMS取込完了：'+rows.length+'件 / 発送日'+dates.join(',')+'（'+latest.getName()+'）','📦ダニエルEMS',6);
}

// EMS在庫の生データを返す(1行ヘッダーがあれば除外)。offset=データ先頭のシート行番号(1始)
function EMS明細_(emv){
  const all=emv.getDataRange().getValues();
  let s=0, cols={コード:1, 数量:2, EMS番号:3, 到着:-1}; // 既定(ヘッダー無し時の固定位置)
  if(all.length){ const r=all[0].map(v=>String(v||'').trim());
    if(r.indexOf('商品コード')>=0 || r.indexOf('状態')>=0 || r.indexOf('EMS番号')>=0){ s=1; // ヘッダーあり→見出し名で列を特定(到着日を挿入して列がズレてもOK)
      const f=(...names)=>{ for(const n of names){ const i=r.indexOf(n); if(i>=0) return i; } return -1; };
      const ci=f('商品コード'), qi=f('数量','個数'), ei=f('EMS番号'), ai=f('EMS到着日','到着日','到着','入荷日');
      cols={ コード: ci>=0?ci:1, 数量: qi>=0?qi:2, EMS番号: ei>=0?ei:3, 到着: ai };
    }
  }
  return { rows: all.slice(s), offset: s+1, cols };
}

// 受注明細の各列の位置を「見出し名」から特定する(列の並び替え・追加に強くする)
// 見つからない列は -1。入荷日が無いCSVでも安全に動く。
function 列マップ_(sh){
  const hr=受注ヘッダー行_(sh);
  const head=sh.getRange(hr,1,1,sh.getLastColumn()||1).getValues()[0].map(v=>String(v||'').trim());
  const find=(...names)=>{ for(const n of names){ const i=head.indexOf(n); if(i>=0) return i; } return -1; };
  return {
    hr,
    番号: find('受注番号'),
    氏名: find('注文者氏名','氏名'),
    届: find('お届け日指定','お届け日'),
    商品名: find('商品名'),
    コード: find('商品コード'),
    SKU: find('商品SKU','SKU'),
    選択肢: find('項目・選択肢','項目選択肢'),
    個数: find('個数'),
    入金: find('入金日'),
    入荷: find('入荷日')
  };
}

// GoQの最新CSV(更新日時が一番新しいもの)を受注明細へ丸ごと入れ替え
function 取込_最新CSV(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi(), cfg=GOQ_CFG;

  // フォルダID内のCSVから更新日時が最新の1件を選ぶ
  let folder;
  try{ folder=DriveApp.getFolderById(cfg.フォルダID); }
  catch(e){ ui.alert('フォルダが見つからへん（ID間違い or 権限なし）\n'+cfg.フォルダID); return; }

  const it=folder.getFiles(); let latest=null;
  while(it.hasNext()){
    const f=it.next();
    const isCsv=/\.csv$/i.test(f.getName()) || f.getMimeType()==='text/csv';
    if(!isCsv) continue;
    if(!latest || f.getLastUpdated()>latest.getLastUpdated()) latest=f;
  }
  if(!latest){ ui.alert('「'+folder.getName()+'」にCSVが無いで'); return; }

  // 取り込む前に確認(古いファイルを誤って入れないため)
  const upd=Utilities.formatDate(latest.getLastUpdated(),'Asia/Tokyo','MM/dd HH:mm');
  const ans=ui.alert('受注明細を入れ替えます',
    '取込ファイル：'+latest.getName()+'\n更新：'+upd+'\n\n受注明細を全消去して、このCSVで丸ごと入れ替えます。ええ？',
    ui.ButtonSet.OK_CANCEL);
  if(ans!==ui.Button.OK) return;
  直列_(()=>取込_実行_(latest)); // 連続クリック対策: 他の処理と排他して実行
}

// 取込の本体(確認ダイアログの後)。直列_で他の処理と排他して実行する
function 取込_実行_(latest){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi(), cfg=GOQ_CFG;
  const upd=Utilities.formatDate(latest.getLastUpdated(),'Asia/Tokyo','MM/dd HH:mm');

  // Shift-JISで読み込み→CSVパース
  const text=latest.getBlob().getDataAsString(cfg.文字コード);
  const rows=Utilities.parseCsv(text);
  if(!rows.length){ ui.alert('CSVが空っぽやで'); return; }

  // 行ごとに列数がブレても矩形になるよう最大列数に揃える。
  // ＋セル内の改行を除去(スペース化)→セルが1行になり行高27pxが効くようにする
  const ncol=rows.reduce((m,r)=>Math.max(m,r.length),0);
  const data=rows.map(r=>{
    const a=r.slice(0,ncol).map(v=> String(v==null?'':v).replace(/[\r\n]+/g,' ').trim());
    while(a.length<ncol)a.push('');
    return a;
  });
  // 受注番号の先頭 "niyantarose-" を除去(例 niyantarose-10117052 → 10117052)
  const banCol = data.length? data[0].indexOf('受注番号') : -1;
  if(banCol>=0) for(let r=1;r<data.length;r++){ data[r][banCol]=String(data[r][banCol]||'').replace(/^niyantarose-/i,''); }

  // 【キャンセルの自動仕分け】CSVに「キャンセル」の受注ステータス行が含まれていたら、受注明細には入れず
  // (需要に化けるのを防ぐ)、受注番号だけ拾ってあとで一括キャンセル処理する。全件CSV運用で、
  // 誰がキャンセルしても取込だけで自動反映される(手で番号を聞かなくてよい)。
  const stCol = data.length? data[0].indexOf('受注ステータス') : -1;
  const キャンセル番号=[]; const キャンセル行=[]; // 別シート「キャンセル」へ残す証跡(取り除いた生データ)
  const 処理済行=[]; // 発送済み(処理済)。受注明細に入れず、消込台帳に確定登録して在庫から差し引く
  if(stCol>=0 && banCol>=0){
    for(let r=data.length-1; r>=1; r--){ // 後ろから消すのでインデックスがズレない
      const st=String(data[r][stCol]||'');
      if(/キャンセル/.test(st)){
        const ban=String(data[r][banCol]||'').trim();
        if(ban && キャンセル番号.indexOf(ban)<0) キャンセル番号.push(ban);
        キャンセル行.unshift(data[r].slice()); // 元の並び順で保存
        data.splice(r,1); // 受注明細に書かない
      } else if(/処理済|発送済|出荷済/.test(st)){ // 発送完了。出荷待/出荷GO等の未発送とは別
        処理済行.unshift(data[r].slice());
        data.splice(r,1); // 受注明細に書かない(需要に化けさせない)
      }
    }
  }

  // 書式・ヘッダーは残し、ヘッダー行の下のデータだけ入れ替える
  let sh=ss.getSheetByName(cfg.受注シート); if(!sh) sh=ss.insertSheet(cfg.受注シート);
  const hdrRow=受注ヘッダー行_(sh);       // 例:6行目(無ければ1)
  const startRow=hdrRow+1;               // データ開始行
  // CSVで列がズレないよう、取込前に既存の「EMS番号」列(前回②が個数の隣に作った分)を一旦削除する。
  // 取込の最後に個数の隣へ作り直す→常に「個数の隣」に固定される。
  { const h0=sh.getRange(hdrRow,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
    const e=h0.indexOf('EMS番号'); if(e>=0) sh.deleteColumn(e+1); }
  // 【入荷日を引き継ぐ】上書き前に、今の入荷日を 受注番号+商品コード+SKU をキーに退避(取込で消えないように)
  const 退避入荷={};
  { const M0=列マップ_(sh);
    if(M0.入荷>=0 && sh.getLastRow()>=startRow){
      const old=sh.getRange(startRow,1,sh.getLastRow()-startRow+1,sh.getLastColumn()).getValues();
      old.forEach(r=>{ const ban=String(r[M0.番号]||'').trim(), code=String(r[M0.コード]||'').trim();
        const sku=M0.SKU>=0?String(r[M0.SKU]||'').trim():''; const v=r[M0.入荷];
        if(ban && v!=='' && v!=null) 退避入荷[ban+'|'+code+'|'+sku]=v; });
    } }
  // ヘッダーより下を全リセット(値・背景色・罫線を消す)→ 前回の引当色や罫線が残らないように
  // ※ヘッダー行(例6行目)と上は触らないので書式は維持される
  // ヘッダーより下をリセット: 値・背景色・罫線だけ消す(フォント/配置は残す)
  const maxRow=sh.getMaxRows(), maxCol=sh.getMaxColumns();
  if(maxRow>=startRow){
    const nr=maxRow-startRow+1;
    const reset=sh.getRange(startRow,1,nr,maxCol);
    reset.clearContent();
    reset.setBackgrounds(Array.from({length:nr},()=>new Array(maxCol).fill(null))); // 背景色を確実に消す
    reset.setBorder(false,false,false,false,false,false);                           // 罫線を消す
  }
  // CSVの見出し行をヘッダー行に反映(列の追加・並び替えに自動追従。書式は保つ)
  sh.getRange(hdrRow,1,1,ncol).setValues([data[0]]);
  // データ行を書き込み → 縦中央・フォント統一(横は左のまま)
  const body=data.slice(1);
  if(body.length){
    sh.getRange(startRow,1,body.length,ncol).setValues(body)
      .setVerticalAlignment('middle').setFontSize(cfg.フォントサイズ)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); // 折返しを切って行高27を一定に保つ
    sh.setRowHeights(startRow, body.length, HIKIATE_CFG.行高);
    注文罫線_(sh, startRow, banCol); // 取込時点でも注文ごとに太線を引く(更新だけで罫線が出るように)
  }

  // 【入荷日を引き継ぐ】退避した入荷日を、同じ 受注番号+商品コード+SKU の行へ貼り直す(取込しても到着状態が消えない)
  let 引継=0;
  if(body.length && Object.keys(退避入荷).length){
    const M1=列マップ_(sh);
    let 入荷C=M1.入荷; // 0始まり。無ければ末尾に「入荷日」列を作る
    if(入荷C<0){ 入荷C=sh.getLastColumn(); sh.getRange(hdrRow,入荷C+1).setValue('入荷日'); }
    const cur=sh.getRange(startRow,1,body.length,sh.getLastColumn()).getValues();
    const out=cur.map(r=>{ const ban=String(r[M1.番号]||'').trim(), code=String(r[M1.コード]||'').trim();
      const sku=M1.SKU>=0?String(r[M1.SKU]||'').trim():''; const v=退避入荷[ban+'|'+code+'|'+sku];
      if(v!==undefined) 引継++; return [ v!==undefined ? v : '' ]; });
    sh.getRange(startRow, 入荷C+1, out.length, 1).setValues(out);
  }
  EMS番号列を用意_(sh); // 個数の隣に「EMS番号」列を作り直す(中身は②が書く)。取込のたびに個数の隣に固定される

  // 消込台帳を更新: 前回いて今回消えた注文=発送済み(他の人が発送した分)を検知
  const 台帳=消込台帳更新_();

  // 【処理済(発送済み)の確定登録】CSVに処理済で載っていた分を消込台帳へ「出荷済み(確定)」として登録。
  // 消えた=発送だろうの推測でなく、CSVの事実で在庫を差し引ける(土日にGoQで発送・入荷日未入力の分も確実に拾う)
  let 処理済結果=null;
  if(処理済行.length){
    try{ 処理済結果=消込台帳_処理済を登録_(data[0], 処理済行); }catch(e){}
    try{ 仕分け証跡シートへ_(ss, '発送済み', data[0], 処理済行, latest.getName()); }catch(e){}
  }

  // 【キャンセルの自動仕分け】CSVにあったキャンセル行を「キャンセル」シートに証跡として残しつつ一括処理。
  // 全件CSVには過去のキャンセルも毎回載り続けるので、処理済みの番号はドキュメントプロパティに覚えておき、
  // 証跡追記とお知らせダイアログは「新規のキャンセルだけ」にする(取込のたびに同じ件数が出続けない)。
  // キャンセル処理_自体は冪等なので毎回全番号に再適用する(P列に手で名前を戻した等の手戻りも自動で直る)
  let キャンセル結果=null, 新規キャンセル=[];
  if(キャンセル番号.length){
    const props=PropertiesService.getDocumentProperties();
    let 既知=[]; try{ 既知=JSON.parse(props.getProperty('処理済キャンセル番号')||'[]'); }catch(e){}
    新規キャンセル=キャンセル番号.filter(b=>既知.indexOf(b)<0);
    if(新規キャンセル.length){
      const 新規行=キャンセル行.filter(r=>新規キャンセル.indexOf(String(r[banCol]||'').trim())>=0);
      try{ 仕分け証跡シートへ_(ss, 'キャンセル', data[0], 新規行, latest.getName()); }catch(e){}
    }
    try{ キャンセル結果=キャンセル処理_(キャンセル番号); }catch(e){}
    // 今回CSVに載っていた番号だけ覚える(CSVの期間から外れた古い番号は自然に忘れる→肥大化しない)
    try{ props.setProperty('処理済キャンセル番号', JSON.stringify(キャンセル番号)); }catch(e){}
  }

  ss.toast('取込完了：'+latest.getName()+' / 受注'+body.length+'行（更新 '+upd+' / 入荷日引継 '+引継+'件'
    +(台帳.新規出荷済? ' / 🧾消えた注文の出荷済み検知'+台帳.新規出荷済+'件':'')
    +(処理済結果&&処理済結果.追加? ' / 📦処理済を確定登録'+処理済結果.追加+'件':'')
    +(キャンセル番号.length? ' / 🚫キャンセル'+キャンセル番号.length+'件(新規'+新規キャンセル.length+'件)':'')+'）','GoQ取込',6);
  if(新規キャンセル.length){
    SpreadsheetApp.getUi().alert('🚫 キャンセルを自動処理しました',
      '新しくキャンセルになった注文 '+新規キャンセル.length+'件を、受注明細に入れず後始末しました:\n'+
      新規キャンセル.slice(0,20).join(', ')+(新規キャンセル.length>20?' …他'+(新規キャンセル.length-20)+'件':'')+
      (キャンセル結果? '\n\n'+キャンセル結果.results.join('\n'):'')+
      '\n\nこのあと「② 引き当て実行」を回すと、その分の在庫が解放されます。',
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// 取込で仕分けた行を証跡シートへ追記する共通関数。取込のCSV見出し＋取込元＋日時を頭に付ける
function 仕分け証跡シートへ_(ss, NAME, header, rows, srcName){
  if(!rows || !rows.length) return;
  let sh=ss.getSheetByName(NAME);
  if(!sh) sh=ss.insertSheet(NAME);
  const ncol=header.length;
  const HDR=['取込日時','取込元'].concat(header.map(h=>String(h||'')));
  const cur=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),HDR.length)).getDisplayValues()[0].slice(0,HDR.length).join('\t');
  if(cur!==HDR.join('\t')){
    sh.getRange(1,1,1,HDR.length).setValues([HDR]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff').setFontSize(HIKIATE_CFG.字);
    sh.setFrozenRows(1);
  }
  const now=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm:ss');
  const out=rows.map(r=>{ const a=r.slice(0,ncol); while(a.length<ncol) a.push(''); return [now, srcName].concat(a); });
  const start=sh.getLastRow()+1;
  sh.getRange(start,1,out.length,HDR.length).setValues(out).setFontSize(HIKIATE_CFG.字);
}

// 受注明細に「EMS番号」列を個数の隣に用意し、その列番号(1始まり)を返す。既にあればその位置。
// 取込でCSV上書き前に一旦削除→取込後にこれで作り直す→常に「個数の隣」に固定される。
function EMS番号列を用意_(recv){
  const hr=受注ヘッダー行_(recv);
  const head=recv.getRange(hr,1,1,Math.max(recv.getLastColumn(),1)).getValues()[0].map(v=>String(v||'').trim());
  const idx=head.indexOf('EMS番号'); if(idx>=0) return idx+1;
  let g=head.indexOf('個数'); if(g<0) g=head.length-1; // 個数が無ければ末尾
  recv.insertColumnAfter(g+1);
  recv.getRange(hr, g+2).setValue('EMS番号').setFontWeight('bold');
  return g+2;
}

// 注文ごと(受注番号が変わる境目)に太い下線を引く。全体には薄い格子。
// シートから実データを直接読むので行ズレしない。太線は境目を getRangeList でまとめて1回で引く(大量注文でも全行に確実に入る)。
// banCol = 受注番号の列(0始まり)
function 注文罫線_(sh, startRow, banCol){
  const lastRow=sh.getLastRow(); if(lastRow<startRow) return {行:0, 注文:0};
  const n=lastRow-startRow+1, ncol=sh.getLastColumn();
  const whole=sh.getRange(startRow,1,n,ncol);
  whole.setBorder(true,true,true,true,true,true,'#cccccc',SpreadsheetApp.BorderStyle.SOLID); // 全体に薄い格子
  const bans=whole.getValues().map(r=>String(r[banCol]||'').trim());
  const a1=[];
  for(let i=0;i<n;i++){ if(i===n-1 || bans[i+1]!==bans[i]) a1.push(sh.getRange(startRow+i,1,1,ncol).getA1Notation()); }
  if(a1.length) sh.getRangeList(a1).setBorder(null,null,true,null,null,null,'#000000',SpreadsheetApp.BorderStyle.SOLID_THICK); // 境目に太い下線をまとめて(THICK=見やすく)
  return {行:n, 注文:a1.length}; // 引いた行数・注文の境目(太線)の数
}

// ボタン用(引数なし): 受注明細に注文ごとの罫線だけを引き直す
function 注文罫線を引く(){
  const ss=SpreadsheetApp.getActive(), cfg=HIKIATE_CFG;
  const recv=ss.getSheetByName(cfg.受注);
  if(!recv){ SpreadsheetApp.getUi().alert('「'+cfg.受注+'」タブが無いで'); return; }
  const M=列マップ_(recv);
  const r=注文罫線_(recv, M.hr+1, M.番号);
  // 引いた件数をポップアップで明示(=ちゃんと入ったか分かるように)
  SpreadsheetApp.getUi().alert('✅ 注文ごとの罫線を引きました\n\n・データ '+r.行+'行に薄い格子\n・注文の境目 '+r.注文+'件 に太線\n\n※フィルタ中は「各注文の最終行」に乗った太線が隠れて見えません。\nデータ→フィルタを解除すると全部見えます。'+(recv.getFilter()?'\n\n⚠️ いま受注明細にフィルタがかかっています。':''));
}

// 受注明細で行が隠れているか(フィルタ or 🙈色なし除外)を知らせる/まとめて全解除する
function フィルタ確認解除(){
  const ss=SpreadsheetApp.getActive(), cfg=HIKIATE_CFG, ui=SpreadsheetApp.getUi();
  const recv=ss.getSheetByName(cfg.受注);
  if(!recv){ ui.alert('「'+cfg.受注+'」タブが無いで'); return; }
  const M=列マップ_(recv), startRow=M.hr+1, lastRow=recv.getLastRow();
  const f=recv.getFilter();
  const props=PropertiesService.getDocumentProperties();
  const 色なし中 = props.getProperty('色なし除外中')==='1'; // 🙈で行を隠している状態
  if(!f && !色なし中){ ui.alert('✅ 全行表示中', '受注明細はフィルタも非表示行もありません（全行表示中）。', ui.ButtonSet.OK); return; }
  const msgs=[];
  if(f) msgs.push('・フィルタがかかっています');
  if(色なし中) msgs.push('・🙈「色なし注文を除外」で行を隠しています');
  const ans=ui.alert('⚠️ 行が隠れています', msgs.join('\n')+'\n\n全部解除して全行表示にしますか？\n（罫線や色を確認するなら解除推奨）', ui.ButtonSet.YES_NO);
  if(ans===ui.Button.YES){
    if(f) f.remove();
    if(lastRow>=startRow) recv.showRows(startRow, lastRow-startRow+1); // 🙈の非表示も戻す
    props.deleteProperty('色なし除外中');
    ss.toast('全部解除して全行表示にしました','表示',4);
  }
}

// 受注明細に「抽出=○(色つき=即納/引当など)」だけ表示するフィルタをかける(ボタン用)
function 色つき注文でフィルタ(){
  const ss=SpreadsheetApp.getActive(), cfg=HIKIATE_CFG, ui=SpreadsheetApp.getUi();
  const recv=ss.getSheetByName(cfg.受注);
  if(!recv){ ui.alert('「'+cfg.受注+'」タブが無いで'); return; }
  抽出フラグ更新_(); // 抽出列(○/×)を今の色で最新化
  const M=列マップ_(recv), nc=recv.getLastColumn(), lastRow=recv.getLastRow();
  if(lastRow<=M.hr){ ss.toast('データが無いで'); return; }
  const head=recv.getRange(M.hr,1,1,nc).getValues()[0].map(v=>String(v||'').trim());
  const flagCol=head.indexOf('抽出')+1; // 1始まり
  if(flagCol===0){ ui.alert('抽出列が無いで。先に②引き当て実行してな'); return; }
  const old=recv.getFilter(); if(old) old.remove(); // 既存フィルタは作り直し
  const filter=recv.getRange(M.hr,1,lastRow-M.hr+1,nc).createFilter(); // 見出し行から下にフィルタ
  filter.setColumnFilterCriteria(flagCol, SpreadsheetApp.newFilterCriteria().whenTextEqualTo('○').build()); // ○だけ表示
  ss.toast('色つき(抽出=○)の注文だけ表示にしました。元に戻すなら🔻フィルタ確認/解除','フィルタ',6);
}

// ===== ボタン連続クリック対策: 書き込み系の処理をドキュメントロックで直列化 =====
// ①受注明細を更新→②即納チェック→③EMS在庫を更新→④引当実行 と連続で押しても、
// 前の処理が終わってから次が走る(同時実行で「①のリセットが②の色を消す」等の競合を防ぐ)。
function 直列_(fn){
  const lock=LockService.getDocumentLock();
  try{ lock.waitLock(10*60*1000); } // 前の処理を最大10分待つ
  catch(e){ SpreadsheetApp.getActive().toast('前の処理が終わらないため中断しました。少し待ってからもう一度押してください','⏳順番待ち',8); return; }
  try{ return fn(); }
  finally{ try{ lock.releaseLock(); }catch(e){} }
}

// ①前段階チェック: 受注明細で「即納」行に水色＋注文ごとの太線(引当前の目視確認用)
function 前段階チェック_即納(){ 直列_(前段階チェック_即納_本体_); }
function 前段階チェック_即納_本体_(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi(), cfg=HIKIATE_CFG;
  const recv=ss.getSheetByName(cfg.受注);
  if(!recv){ ui.alert('「'+cfg.受注+'」タブが無いで'); return; }

  const M=列マップ_(recv), startRow=M.hr+1;
  const lastRow=recv.getLastRow(), lastCol=recv.getLastColumn();
  if(lastRow<startRow){ ss.toast('受注データが無いで'); return; }

  const nrows=lastRow-startRow+1;
  const R=recv.getRange(startRow,1,nrows,lastCol).getValues();

  // 即納=水色 / それ以外=白(リセット)
  let n即納=0;
  const bg=R.map(row=>{
    const kbn=区分_(row[M.選択肢]);
    const col=(kbn==='即納')?cfg.色_水:null;
    if(col) n即納++;
    return new Array(lastCol).fill(col);
  });
  recv.getRange(startRow,1,nrows,lastCol).setBackgrounds(bg);

  注文罫線_(recv, startRow, M.番号);
  ss.toast('前段階チェック完了：即納'+n即納+'行を水色＋注文ごとに罫線','①前段階',6);
}

// ===== 設定 =====
const HIKIATE_CFG = {
  受注: '受注明細', EMS在庫: 'EMS在庫', 待ち: '引当待ち', 部分: '部分在庫', 取置: '出荷GO未入金', 希望: '希望日待ち', 出荷: '出荷可能', 日本在庫: '今回入荷EMSの在庫', 純在庫: '日本在庫',
  ダニエル出荷: 'ダニエル出荷可能', // ダニエルEMS専用の引当結果(メインとは別枠)
  コード入力: '商品コード入力',     // 手動で商品コード・数量を貼り付ける在庫の入力シート(A=商品コード/B=数量)
  コード出荷: '商品コード引当',     // 商品コード(H列)一致での引当結果(在庫=商品コード入力シート・枝番まとめない)
  行高: 27, 字: 13, // 表の行の高さ・フォントサイズ

  // 受注明細の列は固定番号ではなく見出し名で特定する(列マップ_を参照)。並び替え・入荷日追加に強い。

  // EMS在庫タブの列(0始): QUERYで「到着済」だけ A状態 B商品コード C数量 D EMS番号
  E_コード: 1, E_数量: 2,

  色_緑:'#b7e1cd', 色_黄:'#fce8b2', 色_グレー:'#efefef', 色_橙:'#fcd5b4', 色_赤:'#f4cccc', 色_水:'#cfe2f3', 色_着:'#d9d2e9',
  色_ダ引当:'#ead1dc', 色_ダ着:'#d5a6bd', // ダニエル引当=ピンク系(メインの黄と区別)。ダ着=濃いめピンク(ダニエルの着済)
  色_今着:'#b4a7d6' // 今日到着(入荷日=②実行日=今回入荷)=濃いめ紫で目立たせる。過去の着済は薄ラベンダー(色_着)
};

const 在庫EMS番号罫線_CFG = {
  対象シート: ['EMS在庫', '今回入荷EMSの在庫'],
  ボタン行: 1,
  ボタン列: 10, // J列
  ラベル列: 11  // K列
};

function 在庫EMS番号罫線_EMSキー_(value){
  return String(value == null ? '' : value)
    .normalize('NFKC')
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, '');
}

function 在庫EMS番号罫線_グループ範囲_(emsValues){
  const groups = [];
  let start = null, key = '';
  for(let i=0; i<=emsValues.length; i++){
    const cur = i<emsValues.length ? 在庫EMS番号罫線_EMSキー_(emsValues[i]) : '';
    if(start === null){
      if(cur){ start = i; key = cur; }
      continue;
    }
    if(!cur || cur !== key){
      groups.push({ start, count: i - start, key });
      start = cur ? i : null;
      key = cur || '';
    }
  }
  return groups;
}

function 在庫EMS番号罫線_ヘッダー情報_(sh){
  const rows = Math.min(sh.getLastRow() || 1, 20);
  const cols = Math.max(sh.getLastColumn() || 1, 在庫EMS番号罫線_CFG.ラベル列);
  const values = sh.getRange(1, 1, rows, cols).getDisplayValues();
  for(let r=0; r<values.length; r++){
    for(let c=0; c<values[r].length; c++){
      if(String(values[r][c] || '').trim() === 'EMS番号'){
        return { headerRow: r + 1, emsCol: c + 1 };
      }
    }
  }
  return null;
}

function 在庫EMS番号罫線_データ範囲_(sh, header){
  const dataStart = header.headerRow + 1;
  const lastRow = sh.getLastRow();
  if(lastRow < dataStart) return null;

  const scanCols = Math.max(sh.getLastColumn() || 1, header.emsCol);
  const values = sh.getRange(dataStart, 1, lastRow - dataStart + 1, scanCols).getDisplayValues();
  let lastIdx = -1, lastCol = header.emsCol;
  for(let r=0; r<values.length; r++){
    let rowHasValue = false;
    for(let c=0; c<values[r].length; c++){
      if(String(values[r][c] || '').trim() !== ''){
        rowHasValue = true;
        if(c + 1 > lastCol) lastCol = c + 1;
      }
    }
    if(rowHasValue) lastIdx = r;
  }
  if(lastIdx < 0) return null;

  return {
    startRow: dataStart,
    numRows: lastIdx + 1,
    numCols: lastCol,
    emsValues: values.slice(0, lastIdx + 1).map(row => row[header.emsCol - 1])
  };
}

function 在庫EMS番号ごとに罫線を引く(){
  const result = 在庫EMS番号ごとに罫線を引く_(false);
  SpreadsheetApp.getActive().toast(
    'EMS番号ごとに罫線: ' + result.sheets + 'シート / ' + result.groups + 'グループ',
    'EMS在庫',
    6
  );
}

function 在庫EMS番号ごとに罫線を引く_(silent){
  const ss = SpreadsheetApp.getActive();
  let sheets = 0, groups = 0;
  在庫EMS番号罫線_CFG.対象シート.forEach(name => {
    const sh = ss.getSheetByName(name);
    if(!sh) return;
    const r = 在庫EMS番号罫線_シート_(sh);
    if(r.ok){ sheets++; groups += r.groups; }
  });
  if(!silent && !sheets) SpreadsheetApp.getUi().alert('対象シートにデータが見つかりません。');
  return { sheets, groups };
}

function 在庫EMS番号罫線_シート_(sh){
  const header = 在庫EMS番号罫線_ヘッダー情報_(sh);
  if(!header) return { ok:false, groups:0 };
  const data = 在庫EMS番号罫線_データ範囲_(sh, header);
  if(!data) return { ok:false, groups:0 };

  const range = sh.getRange(data.startRow, 1, data.numRows, data.numCols);
  range.setBorder(false, false, false, false, false, false);
  range.setBorder(true, true, true, true, true, true, '#d9d9d9', SpreadsheetApp.BorderStyle.SOLID);

  const groups = 在庫EMS番号罫線_グループ範囲_(data.emsValues);
  groups.forEach(g => {
    sh.getRange(data.startRow + g.start, 1, g.count, data.numCols)
      .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
  return { ok:true, groups:groups.length };
}

// ===== 日本在庫: EMS番号ごとに罫線+色分け(図形ボタン用) =====
// 余り(日本在庫)を箱ごとに見分けやすくする。②のたびにシートが書き直される(色も消える)ので、見たいときに押す。
const 日本在庫色分け_パレット=['#e8f0fe','#e6f4ea','#fef7e0','#fde9e9','#f3e8fd']; // 淡い5色を循環
function 日本在庫の罫線と色分け(){ 直列_(日本在庫の罫線と色分け本体_); }
function 日本在庫の罫線と色分け本体_(){
  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName(HIKIATE_CFG.純在庫);
  if(!sh){ SpreadsheetApp.getUi().alert('「'+HIKIATE_CFG.純在庫+'」タブがありません'); return; }
  const header=在庫EMS番号罫線_ヘッダー情報_(sh);
  if(!header){ SpreadsheetApp.getUi().alert('「EMS番号」の見出しが見つかりません'); return; }
  const data=在庫EMS番号罫線_データ範囲_(sh, header);
  if(!data){ ss.toast('日本在庫にデータがありません','🎨日本在庫',5); return; }
  在庫EMS番号罫線_シート_(sh); // 薄い格子+EMS番号グループの太枠
  // 色分け: EMS番号グループごとに淡い色を循環(EMS番号が空白の行は白のまま)
  const groups=在庫EMS番号罫線_グループ範囲_(data.emsValues);
  const bg=Array.from({length:data.numRows},()=>new Array(data.numCols).fill(null));
  groups.forEach((g,gi)=>{
    const col=日本在庫色分け_パレット[gi % 日本在庫色分け_パレット.length];
    for(let r=g.start; r<g.start+g.count; r++) bg[r].fill(col);
  });
  sh.getRange(data.startRow,1,data.numRows,data.numCols).setBackgrounds(bg);
  ss.toast('日本在庫: '+groups.length+'箱を罫線+色分けしました','🎨日本在庫',5);
}

function 在庫EMS番号罫線ボタンを設置(){
  const count = 在庫EMS番号罫線ボタンを設置_(false);
  SpreadsheetApp.getActive().toast('EMS番号罫線ボタンを設置: ' + count + 'シート', 'EMS在庫', 5);
}

function 在庫EMS番号罫線ボタンを設置_(silent){
  const ss = SpreadsheetApp.getActive();
  let count = 0;
  在庫EMS番号罫線_CFG.対象シート.forEach(name => {
    const sh = ss.getSheetByName(name);
    if(!sh) return;
    const check = sh.getRange(在庫EMS番号罫線_CFG.ボタン行, 在庫EMS番号罫線_CFG.ボタン列);
    check.insertCheckboxes().setValue(false)
      .setHorizontalAlignment('center')
      .setBackground('#d9ead3');
    sh.getRange(在庫EMS番号罫線_CFG.ボタン行, 在庫EMS番号罫線_CFG.ラベル列)
      .setValue('EMS番号ごとに罫線')
      .setFontWeight('bold')
      .setFontColor('#274e13')
      .setBackground('#d9ead3');
    if(sh.getColumnWidth(在庫EMS番号罫線_CFG.ラベル列) < 150){
      sh.setColumnWidth(在庫EMS番号罫線_CFG.ラベル列, 170);
    }
    count++;
  });
  if(!silent && !count) SpreadsheetApp.getUi().alert('対象シートが見つかりません。');
  return count;
}

function 在庫EMS番号罫線ボタン編集_(e){
  if(!e || !e.range || e.value !== 'TRUE') return false;
  const sh = e.range.getSheet();
  if(在庫EMS番号罫線_CFG.対象シート.indexOf(sh.getName()) < 0) return false;
  if(e.range.getRow() !== 在庫EMS番号罫線_CFG.ボタン行 || e.range.getColumn() !== 在庫EMS番号罫線_CFG.ボタン列) return false;

  e.range.setValue(false);
  在庫EMS番号ごとに罫線を引く();
  return true;
}

// 🔎 引当診断: 指定の受注番号について、入荷日の列・値・区分・EMS在庫との一致を調べる(白いままの原因特定用)
function 引当診断(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi(), cfg=HIKIATE_CFG;
  const recv=ss.getSheetByName(cfg.受注); if(!recv){ ui.alert('受注明細が無い'); return; }
  const resp=ui.prompt('引当診断','調べたい受注番号を入力（例 10116946）',ui.ButtonSet.OK_CANCEL);
  if(resp.getSelectedButton()!==ui.Button.OK) return;
  const ban=String(resp.getResponseText()||'').trim(); if(!ban) return;
  const M=列マップ_(recv), R=recv.getDataRange().getValues();
  const head=recv.getRange(M.hr,1,1,recv.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
  const 入荷列=[]; head.forEach((h,i)=>{ if(h==='入荷日') 入荷列.push('列'+(i+1)); }); // 入荷日ヘッダーの重複チェック
  const emv=ss.getSheetByName(cfg.EMS在庫); const emsD=emv?EMS明細_(emv):{rows:[],cols:{コード:1}};
  const stockKeys={}; emsD.rows.forEach(r=>{ const c=normCode_(r[emsD.cols.コード]); if(c) codeKeys_(c).forEach(k=>stockKeys[k]=true); });
  let msg='受注番号：'+ban+'\n入荷日ヘッダー：'+(入荷列.join(' , ')||'なし')+(入荷列.length>1?'  ⚠️重複あり！':'')+'\n\n■ 受注明細の行\n';
  let hit=0;
  for(let i=M.hr;i<R.length;i++){ if(String(R[i][M.番号]||'').trim()!==ban) continue; hit++;
    const r=R[i], code=String(r[M.コード]||'').trim(), sku=String(r[M.SKU]||'').trim();
    const 入荷raw=M.入荷>=0? r[M.入荷] : '';
    const 入荷v=M.入荷>=0? (String(入荷raw||'').trim()? (ymd_(入荷raw)||String(入荷raw)) : '') : '(列なし)';
    const 一致=受注候補コード_(sku,code).some(v=>codeKeys_(v).some(k=>stockKeys[k]));
    const 品表示=(code||sku)+(sku && sku!==code? '（SKU:'+sku+'）':''); // 同一コードの種類違いをSKUで見分けられるように
    msg+=hit+') 行'+(i+1)+' '+品表示+' x'+(Number(r[M.個数])||0)+' / '+区分_(r[M.選択肢])+' / 入荷日:「'+(入荷v||'空')+'」/ EMS在庫一致:'+(一致?'○':'×')+'\n';
  }
  if(!hit) msg+='(受注明細に見つからへん。🔻フィルタ確認/解除で非表示を解除して確認してや)\n';

  // EMSリスト(発注共有)のP列で、この受注が名指しされている箱と個数
  msg+='\n■ EMSリストのP列名指し(どの箱に何個)\n';
  try{
    const ems=個別対応_EMSリスト_();
    if(ems.error){ msg+='(EMSリストが読めません)\n'; }
    else{
      let total=0, cnt=0;
      ems.rows.forEach(r=>{
        if(!r.注文番号) return;
        const q=個別対応_P注文展開_(r.注文番号, r.数量).filter(e=>e.ban===ban).reduce((s,e)=>s+e.qty,0);
        if(q<=0) return;
        cnt++; total+=q;
        msg+='・['+(r.状態||'?')+'] '+(r.EMS到着日||'?')+'着 '+r.商品コード+' に '+q+'個\n';
      });
      msg+= cnt? ('　→ 名指し合計 '+total+'個\n') : '(P列にこの受注の名指しなし)\n';
    }
  }catch(e){ msg+='(P列確認でエラー: '+e.message+')\n'; }

  // 消込台帳
  msg+='\n■ 消込台帳\n';
  try{
    const tsh=ss.getSheetByName(KESHIKOMI_CFG.シート);
    if(tsh && tsh.getLastRow()>1){
      const tv=tsh.getRange(2,1,tsh.getLastRow()-1,KESHIKOMI_CFG.HDR.length).getDisplayValues();
      let n=0;
      tv.forEach(r=>{ if(String(r[0]||'').trim()!==ban) return; n++;
        msg+='・'+r[1]+' x'+r[3]+' / 状態:'+r[5]+' / 入荷日:'+(r[4]||'-')+'\n'; });
      if(!n) msg+='(記録なし)\n';
    } else msg+='(台帳なし)\n';
  }catch(e){ msg+='(台帳確認でエラー)\n'; }

  // 引当履歴(需要から差し引かれるものに印)
  msg+='\n■ 引当履歴\n';
  try{
    const hsh=ss.getSheetByName(HIKIATE_HISTORY_CFG.シート);
    if(hsh && hsh.getLastRow()>1){
      const hc=引当履歴_列_(hsh);
      const hv=hsh.getRange(2,1,hsh.getLastRow()-1,hsh.getLastColumn()).getDisplayValues();
      let n=0;
      hv.forEach(r=>{
        const b=hc['受注番号']?String(r[hc['受注番号']-1]||'').trim():'';
        if(b!==ban) return; n++;
        const kind=hc['取込区分']?String(r[hc['取込区分']-1]||''):'', st=hc['EMSリスト状態']?String(r[hc['EMSリスト状態']-1]||''):'',
              code=hc['商品コード']?String(r[hc['商品コード']-1]||''):'', q=hc['引当数']?String(r[hc['引当数']-1]||''):'',
              state=hc['状態']?String(r[hc['状態']-1]||''):'';
        const 差引き=(state!=='キャンセル済み') && (kind==='過去取込' || st==='在庫反映済み');
        msg+='・'+code+' x'+q+' / '+kind+' / EMS状態:'+st+' / '+(state||'有効')+(差引き?' ←需要から差引き':'')+'\n';
      });
      if(!n) msg+='(記録なし)\n';
    } else msg+='(履歴なし)\n';
  }catch(e){ msg+='(履歴確認でエラー)\n'; }

  ui.alert('引当診断', msg, ui.ButtonSet.OK);
}

// 商品コード末尾の受注番号タグ(中古品などの名指し買付)は照合前に取り除き、基底コードで一致させる。
// 対応書式: 「POEM65（10116569）」「POEM65(10116569)」「RECIPE42/10117126」(カッコは全角/半角、スラッシュは/／)
// タグ自体の名指しはP列自動記入が拾う(タグ受注番号_)。
function normCode_(v){ return String(v||'').trim().replace(/[\s　]*(?:[（(]\s*\d{7,}\s*[）)]|[/／]\s*\d{7,})\s*$/,'').toUpperCase().replace(/_/g,'-'); }
// コード末尾の受注番号タグ(（受注番号）または /受注番号)から受注番号を取り出す。
// 無ければ''(7桁以上の数字だけ=巻数カッコや枝番の誤検出防止)
function タグ受注番号_(v){ const m=String(v||'').trim().match(/(?:[（(]\s*(\d{7,})\s*[）)]|[/／]\s*(\d{7,}))\s*$/); return m? (m[1]||m[2]) : ''; }
function SKU枝番あり_(sku, code){
  const s=normCode_(sku), c=normCode_(code);
  return !!(s && c && s!==c && s.indexOf(c)===0 && /^[A-Z]+$/.test(s.slice(c.length)));
}
function 受注候補コード_(sku, code){
  // 候補の優先順: ①SKUそのまま(丸めない一致を最優先) ②商品コード列 ③末尾1字落とし(a/b枝番の保険)
  // かつて枝番あり時にSKUだけへ絞っていたが、それだと限定版(例 JPSJCM39-03S+SKU=…Sb)や
  // SKU=コード+枝番(例 MOFUN-IV-09a)がEMS在庫のコードと一致せず、
  // 「今回の便なのにラベンダー」「在庫が消費されず他へ二重割当」になる。
  // ③は受注番号タグを外した基底コードに対して行う(例 RECIPE42b/10117126 → RECIPE42B → RECIPE42。
  // 生の文字列のままだと末尾が数字なので枝番bが落ちない)。
  // 名指し(P列)の横取りは引当実行側の注文所有キー_ガードで防ぐので、ここで候補を削る必要はない。
  const s=String(sku||'').trim(), c=String(code||'').trim();
  const out=[];
  const push=v=>{ const t=String(v||'').trim(); if(t && !out.some(o=>normCode_(o)===normCode_(t))) out.push(t); };
  push(s);
  push(c);
  push(normCode_(s).replace(/[A-Za-z]$/,''));
  push(normCode_(c).replace(/[A-Za-z]$/,''));
  return out;
}
function codeKeys_(code){
  const c=normCode_(code); const keys=[c];
  const m=c.match(/^(.+)-(\d{2})(\d{2})-(\d{2})$/);
  if(m && (m[4]===m[2]||m[4]===m[3])) keys.push(m[1]+'-'+m[4]);
  // 月号付き雑誌コード: 末尾の -YYMM / -YYYYMM (例 EBS1504B-2607, CSG1504--202607) を落とした基底コードも候補に
  // GoQ側は月号なし(EBS1504B)で登録されているため。YY=2x・MM=01〜12の形だけ対象(巻数セット等の誤爆防止)
  const mg=c.match(/^(.+?)-(?:20)?(2\d)(0[1-9]|1[0-2])$/);
  if(mg){ const base=mg[1].replace(/-+$/,''); if(base && keys.indexOf(base)<0) keys.push(base); }
  return keys;
}
// ===== 定期購読の月号照合(例: EBS1504B_2607=7月号) =====
// 商品ページのルール「ご注文月の翌月号をお届け」に合わせて、
// 月号付き在庫(-YYMM/-YYYYMM)は「注文月+1＝その号」の注文にだけ引き当てる(号ズレ防止)。
// 月号なしの在庫・注文日時が読めない行は無制限(従来通り)。P列の名指しは人の判断なので月号チェックを通さない。
function 日付値_(v){ if(v instanceof Date) return isNaN(v.getTime())?null:v; const s=String(v||'').trim(); if(!s) return null; const d=new Date(s.replace(/\//g,'-')); return isNaN(d.getTime())?null:d; }
function 月号_(code){ const m=String(code||'').match(/-(?:20)?(2\d)(0[1-9]|1[0-2])$/); return m? m[1]+m[2] : null; }   // 'EBS1504B-2607'→'2607'
function 期待号_(d){ if(!d) return null; let y=d.getFullYear()%100, mo=d.getMonth()+2; if(mo>12){ mo-=12; y++; } return String(y)+('0'+mo).slice(-2); } // 注文月の翌月
function 月号OK_(l, stockCode){ const g=月号_(stockCode); if(!g) return true; const e=期待号_(l.日時); return !e || e===g; }

function 区分_(opt){
  const s=String(opt||'').trim();
  if(s==='') return '即納'; // 項目・選択肢が空 → 即納扱い
  if(s.indexOf('取り寄せ')>=0||s.indexOf('取寄')>=0) return '取り寄せ';
  if(s.indexOf('即納')>=0) return '即納';
  return '指定なし';
}
function 番号num_(v){ const m=String(v||'').match(/\d+/g); return m? Number(m[m.length-1]):0; }
function 入金済み_(v){ return !(v===''||v===null||v===undefined||String(v).trim()===''); }
// お届け日指定が今日より後(未来)か。空や日付でなければ false(=今出せる)
function 希望日未来_(v){
  const s=String(v||'').trim(); if(!s) return false;
  const d=new Date(s); if(isNaN(d.getTime())) return false;
  const t=new Date(); t.setHours(0,0,0,0);
  return d.getTime() > t.getTime();
}
// 入荷日が今日(=②引当を実行する日)か。今日の着済を「今回分」として扱うため。
// 手入力の「26/07/09(木)」のような2桁年・曜日付きは new Date が読めず今日扱いにならない
// (=当日入荷なのにラベンダー)ので、ymd_ で正規化してから比べる
function 入荷日今日_(v){
  const s=ymd_(v); if(!s) return false;
  return s===ymd_(new Date());
}
// 日付を yyyy-MM-dd 文字列に正規化(Date/文字列どちらでも)。入荷日と到着日の一致判定用。
// 手入力の「26/07/09(木)」のような2桁年・曜日付きも読む。パースできない文字列はそのまま返す
function ymd_(v){
  const p=n=>('0'+n).slice(-2);
  if(v instanceof Date){ if(isNaN(v.getTime())) return ''; return v.getFullYear()+'-'+p(v.getMonth()+1)+'-'+p(v.getDate()); }
  const s=String(v||'').trim(); if(!s) return '';
  const d=new Date(s.replace(/\//g,'-'));
  if(!isNaN(d.getTime())) return d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate());
  let m=s.match(/(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if(m) return m[1]+'-'+p(+m[2])+'-'+p(+m[3]);
  m=s.match(/(\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if(m) return '20'+m[1]+'-'+p(+m[2])+'-'+p(+m[3]);
  return s;
}

// 「今回入荷EMS」として扱う着済行か。実行日だけでなく、今回EMS在庫の到着日と入荷日が一致する行も対象。
function 今回到着扱い_(l, 入荷消費OK){
  return !!(l && (l.今回P || (l.入荷 && (typeof 入荷消費OK === 'function' ? 入荷消費OK(l) : 入荷日今日_(l.入荷日値)))));
}
function 今回行判定_(l, 入荷消費OK){
  return !!(l && (l.引当成立 || 今回到着扱い_(l, 入荷消費OK)));
}
function 引当行状態_(l, cfg, 入荷消費OK){
  if(l.キャンセル) return { st:'キャンセル', color:cfg.色_グレー };
  if(l.kbn==='即納') return { st:'即納', color:cfg.色_水 };
  if(l.kbn==='指定なし') return { st:'要確認', color:cfg.色_橙 };
  if(l.kbn==='取り寄せ'){
    if(今回到着扱い_(l, 入荷消費OK)) return { st:'着済(今回)', color:cfg.色_黄 };
    if(l.入荷) return { st:'着済', color:cfg.色_着 };
    if(l.引当成立) return { st:'引当(今回)', color:cfg.色_黄 };
    if(l.履歴成立) return { st:'着済(履歴)', color:cfg.色_着 };
  }
  return { st:'在庫待ち', color:null };
}
function 注文出荷準備OK_(arr){
  const active=(arr||[]).filter(l=>l && !l.キャンセル);
  return active.length>0 && active.every(l=> l.kbn==='即納' || (l.kbn==='取り寄せ' && (l.入荷 || l.引当成立 || l.履歴成立)) );
}
function 注文区分判定_(arr, paid, 入荷消費OK){
  const active=(arr||[]).filter(l=>l && !l.キャンセル);
  const 黄色あり=active.some(l=>今回行判定_(l, 入荷消費OK));
  if(!黄色あり) return 'wait';
  const 未入荷あり=active.some(l=> l.kbn==='取り寄せ' && !l.入荷 && !l.引当成立);
  if(未入荷あり) return 'part';
  if(!注文出荷準備OK_(active)) return 'part';
  if(!paid) return 'keep';
  if(active.some(l=>希望日未来_(l.届))) return 'hold';
  return 'ship';
}
function 発送可否判定_(arr, paid){
  if(!注文出荷準備OK_(arr)) return '';
  if((arr||[]).some(l=>l && !l.キャンセル && 希望日未来_(l.届))) return '発送可能希望日待ち';
  return paid ? '発送可能' : '発送可能入金待ち';
}

function 引当実行(){ 直列_(引当実行_本体_); }
function 引当実行_本体_(){
  const ss=SpreadsheetApp.getActive(), cfg=HIKIATE_CFG;
  const recv=ss.getSheetByName(cfg.受注);
  if(!recv){ SpreadsheetApp.getUi().alert('「'+cfg.受注+'」タブが無いで'); return; }
  消込台帳更新_(); // 受注明細から消えた注文(発送済み)を検知して台帳を最新化
  try{ 発注共有P列記入_(); }catch(e){} // 発注共有EMSリストのP列(注文番号)も自動記入(開けない時は黙ってスキップ)
  // 引当履歴への到着分記録は②では省略(発注共有EMSリストを丸読みして重い→タイムアウト要因)。
  // ⑤便の締め(到着済を在庫反映済みへ)で記録するので、反映前の履歴はそこで残る。
  const R=recv.getDataRange().getValues();

  // 受注明細→行オブジェクト(列は見出し名で特定。並び替え・列追加に強い)
  // 着済(着いたか)の判定: 即納 or (取り寄せ かつ 入荷日あり)。取り寄せで入荷日なし=未着。
  const M=列マップ_(recv), 受注hdr=M.hr;
  const 受注head=recv.getRange(M.hr,1,1,recv.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
  const c日時=受注head.indexOf('注文日時'); // 月号照合(定期購読)用
  const lines=[];
  for(let i=受注hdr;i<R.length;i++){
    const row=R[i];
    const ban=String(row[M.番号]||'').trim(), code=String(row[M.コード]||'').trim();
    if(!ban && !code) continue;
    const kbn=区分_(row[M.選択肢]);
    const 入荷日値=M.入荷>=0? row[M.入荷] : '', 入荷=String(入荷日値||'').trim()!=='';
    const 着=(kbn==='即納') || (kbn==='取り寄せ' && 入荷); // 着済かどうか
    const qty=Number(row[M.個数])||0;
    lines.push({ i, ban, sortKey:番号num_(ban),
      氏名:row[M.氏名], 届:row[M.届], 商品名:row[M.商品名],
      code, sku:String(row[M.SKU]||'').trim(), qty, kbn,
      日時: c日時>=0? 日付値_(row[c日時]) : null,
      paid:入金済み_(row[M.入金]), 入荷, 入荷日値, 着, alloc:0, 引当成立:false, キャンセル:qty<=0 });
  }

  // ===== EMS在庫引当: 入荷日あり(割当済)を在庫から差引き、残りを未着の取り寄せに古い注文順でFIFO引当 =====
  const emv=ss.getSheetByName(cfg.EMS在庫);
  const emsD= emv? EMS明細_(emv) : {rows:[], cols:{コード:1,数量:2,EMS番号:3,到着:-1}};
  const E= emsD.rows, EC=emsD.cols; // EC=EMS在庫の列位置(見出し名で特定。到着日を挿入して列がズレてもOK)
  const 到着_=row=>{ if(EC.到着<0) return ''; const v=row[EC.到着]; // EMS到着日(発注共有ファイルE列)。日付として扱い yyyy-mm-dd 表示
    if(v instanceof Date) return v;
    const s=String(v||'').trim(); if(!s) return ''; const d=new Date(s); return isNaN(d.getTime())? s : d; };
  const stock={};
  for(let i=0;i<E.length;i++){ const c=normCode_(E[i][EC.コード]); if(!c) continue; stock[c]=(stock[c]||0)+(Number(E[i][EC.数量])||0); }
  // 商品コード→EMS到着日(入荷日の自動記入用。同コードは最初の到着日を代表に)
  const code到着={};
  for(let i=0;i<E.length;i++){ const c=normCode_(E[i][EC.コード]); if(!c||(c in code到着)) continue; const a=到着_(E[i]); if(a!=='' && a!=null) code到着[c]=a; }
  const origStock=Object.assign({},stock); // 余り(今回入荷EMSの在庫=日本在庫)算出用
  const aliasMap={}; Object.keys(stock).forEach(c=> codeKeys_(c).forEach(k=>{ if(!(k in aliasMap)) aliasMap[k]=c; }));
  const candKeys=l=>{ const cands=受注候補コード_(l.sku,l.code);
    const keys=[]; cands.forEach(v=> codeKeys_(v).forEach(k=>{ if(keys.indexOf(k)<0) keys.push(k); })); return keys; };
  const keyInStock=l=>{ for(const k of candKeys(l)){ if(k in stock && 月号OK_(l,k)) return k; if(aliasMap[k] && aliasMap[k] in stock && 月号OK_(l,aliasMap[k])) return aliasMap[k]; } return null; };
  const findAvail=l=>{ for(const k of candKeys(l)){ if(stock[k]>0 && 月号OK_(l,k)) return k; if(aliasMap[k]&&stock[aliasMap[k]]>0&&月号OK_(l,aliasMap[k])) return aliasMap[k]; } return null; };
  const 履歴一致_= (l,e)=> candKeys(l).some(k=> k===e.key || aliasMap[k]===e.key || codeKeys_(e.key).indexOf(k)>=0);
  const 残必要_=l=> Math.max(0, l.qty-(l.alloc||0)-(l.履歴Alloc||0));
  // 各商品コードが「今回のEMS在庫」に到着日として存在する日付の集合(yyyy-MM-dd)
  const 到着日SetByCode={};
  for(let i=0;i<E.length;i++){ const c=normCode_(E[i][EC.コード]); if(!c) continue; const s=ymd_(到着_(E[i])); if(s){ (到着日SetByCode[c]=到着日SetByCode[c]||new Set()).add(s); } }
  // 入荷日付きの注文が「今回のEMS在庫」を消費してよいか。今日着(=今回の便)か、入荷日が現物の到着日と一致する分だけ。
  // 入荷日はマッチした箱の到着日から自動記入される(下部の【入荷日を自動記入】)ので「入荷日=箱の到着日」が保証される。
  // 前回の入荷で処理済み(入荷日が今回の到着日に無い)注文は、今回別便で来た同じ商品を消費しない=日本在庫として残す。
  const 到着日不明_ = EC.到着 < 0; // 到着日列が無いEMS在庫では日付判定できない→従来通り全消費(誤って全部を日本在庫にしない)
  const 入荷日一致_=l=>{ const k=keyInStock(l); if(k==null) return false; const set=到着日SetByCode[k]; return !!(set && set.has(ymd_(l.入荷日値))); };
  const 入荷消費OK_=l=> 到着日不明_ || 入荷日今日_(l.入荷日値) || 入荷日一致_(l);
  // 引当履歴の「在庫反映済み/過去取込」分を先に差し引く。
  //  ・入荷日付きの行: 過去の箱で確保済みの個数(過去箱分)を求め、今回の箱はその残りだけ消費する
  //    (2箱またがりの注文=入荷日欄が1つしかなく今回日付でも、過去箱の分まで今回の箱を食わない)
  //  ・未着の行: 従来通り必要数から差し引く(古い箱で割当済みの注文に今回の在庫を再度つかませない)
  let hist={}; try{ hist=引当履歴_反映済み割当マップ_(); }catch(e){}
  try{
    lines.filter(l=>l.kbn==='取り寄せ' && l.入荷 && l.qty>0).sort((a,b)=> a.sortKey-b.sortKey || a.i-b.i).forEach(l=>{
      const q=hist[l.ban]||[]; let used=0;
      for(const e of q){
        if(l.qty-(l.過去箱分||0)-used<=0) break;
        if(e.qty<=0 || !履歴一致_(l,e)) continue;
        const take=Math.min(l.qty-(l.過去箱分||0)-used, e.qty);
        e.qty-=take; used+=take;
      }
      if(used>0) l.過去箱分=(l.過去箱分||0)+used;
    });
    lines.filter(l=>l.kbn==='取り寄せ' && !l.入荷 && l.qty>0).sort((a,b)=> a.sortKey-b.sortKey || a.i-b.i).forEach(l=>{
      const q=hist[l.ban]||[]; let used=0;
      for(const e of q){
        if(残必要_(l)<=0) break;
        if(e.qty<=0 || !履歴一致_(l,e)) continue;
        const take=Math.min(残必要_(l), e.qty);
        e.qty-=take; used+=take;
        if(!l.履歴入荷日値 && e.date) l.履歴入荷日値=e.date;
      }
      if(used>0){ l.履歴Alloc=(l.履歴Alloc||0)+used; l.履歴成立=残必要_(l)<=0; }
    });
  }catch(e){}
  // 【入荷日ズレの自動修復(P列に無い注文のフォールバック)】先引当のあと箱の到着日が実日付に直る(例7/9→7/10)と、
  // ②が先に書いた入荷日がどの到着済の箱とも一致しなくなり、今回の便なのにラベンダー扱いになる。
  // P列(注文番号)で名指しされた注文は下の【確定引き当て】が箱の到着日へ正確に訂正するのでここでは触らない。
  // P列に載っていない注文だけ、「商品が今の到着済在庫にある」「履歴で過去箱の受取が判明していない」
  // 「その商品の箱の到着日が一意」「ズレが3日以内」を全て満たす場合にスタンプを箱の到着日へ貼り直す。
  // 日数が離れた行(=過去便で受け取り済みの可能性が高い)や箱が複数日にまたがる行は触らず、完了ダイアログの一覧で知らせる
  const 確定=P列確定マップ_(); // P列の名指し(ここのスキップ判定と、下の確定引き当ての両方で使う)
  let ズレ修復=0;
  lines.forEach(l=>{
    if(l.kbn!=='取り寄せ' || !l.入荷 || l.qty<=0) return;
    if(確定[l.ban]) return;              // P列で名指しあり=確定引き当て側が箱の到着日へ訂正する(そちらが正)
    if((l.過去箱分||0)>=l.qty) return;   // 在庫反映済み(過去箱の受取が履歴で判明)=正しいラベンダー
    if(入荷消費OK_(l)) return;           // 既に箱と一致している(黄色になる)行は触らない
    const k=keyInStock(l); if(k==null) return; // 商品自体が今の到着済に無い=過去便の正常な着済
    const set=到着日SetByCode[k]; if(!set || set.size!==1) return;
    const d0=Array.from(set)[0], s=ymd_(l.入荷日値);
    if(!/^20\d{2}-\d{2}-\d{2}$/.test(s) || !/^20\d{2}-\d{2}-\d{2}$/.test(d0)) return;
    const diff=Math.abs((new Date(d0+'T00:00:00').getTime()-new Date(s+'T00:00:00').getTime())/86400000);
    if(diff>3) return;
    l.入荷日値=new Date(d0+'T00:00:00'); l.修復入荷=true; ズレ修復++;
  });
  // 入荷日あり(もう割当済)を在庫から先に差し引く。ただし今回の便で着いた分だけ、かつ過去箱分を除いた残り
  lines.filter(l=>l.kbn==='取り寄せ' && l.入荷 && 入荷消費OK_(l)).forEach(l=>{
    const take=l.qty-(l.過去箱分||0); if(take<=0) return;
    const k=keyInStock(l); if(k!=null) stock[k]=Math.max(0,(stock[k]||0)-take);
  });
  // 出荷済み(受注明細から消えた注文=消込台帳)の分も差し引く(発送済み品が余り=日本在庫に二重計上されるのを防ぐ)。
  // ただし台帳の入荷日が今回の到着日と一致する分だけ(別の箱/即納在庫から出た出荷済みに、後から着いた箱を食わせない)
  // 引当履歴で「過去の箱で受け取り済み(反映済み)」の注文は、消込台帳の出荷済みでも今回の箱を食わない
  // (10116953のように過去箱で受け取り済みなのに出荷済みで今回の箱を二重に食い、本物の注文を宙に押し出すのを防ぐ)。
  // ↑のライブ行の差引きで消費されずに残った履歴分だけをカバー扱いにする(=発送済み注文の過去箱分)。
  const 履歴カバー_=l=>{ const q=hist[l.ban]||[]; if(!q.length) return false;
    const cands=[]; 受注候補コード_(l.sku,l.code).forEach(v=> codeKeys_(v).forEach(k=>{ if(cands.indexOf(k)<0) cands.push(k); }));
    return q.some(e=> e.qty>0 && (cands.indexOf(e.key)>=0 || codeKeys_(e.key).some(k=>cands.indexOf(k)>=0))); };
  const 出荷済行=消込台帳_出荷済み行_().filter(l=> !履歴カバー_(l));
  // 出荷済み(発送済みで受注明細から消えた注文)を今回の在庫から差し引いてよいか。
  //  ・入荷日あり → その日付が今回の到着日と一致する分だけ(別便で処理済みが今回の箱を食わない)
  //  ・入荷日なし → 土日等にGoQで発送・入荷日未入力の分。「発送日以前に到着していた箱」がある時だけ差し引く。
  //    発送より後に着いた箱はその発送の出どころではあり得ないので食わせない(かつてコード一致で無条件に
  //    差し引いていたら、60日分の発送済みが新しく着いた箱を食い尽くし、本物の注文が宙=要確認に
  //    押し出されて毎回⚠️が出続けた)。発送日が読めない時は従来通り差し引く(売り越し防止を優先)
  const 出荷済消費OK_=l=>{ if(到着日不明_) return true; const k=keyInStock(l); if(k==null) return false;
    const set=到着日SetByCode[k];
    const d=ymd_(l.入荷日);
    if(d) return !!(set && set.has(d));
    const ship=ymd_(l.基準日); if(!ship) return true; // 発送日不明=従来通りコード一致で差し引く
    return !!(set && Array.from(set).some(d0=> d0<=ship)); };
  // 出荷済み(発送済みで受注明細から消えた分)を在庫から差し引く(発送済み品が余り=日本在庫に二重計上されるのを防ぐ)。内訳は「発送済み」シートで見る
  出荷済行.forEach(l=>{ if(!出荷済消費OK_(l)) return; const k=keyInStock(l); if(k!=null) stock[k]=Math.max(0,(stock[k]||0)-l.qty); });
  // 各注文が実際に持つ商品キー(別名込み)。下の救済(コード不一致引き当て)で、同じ注文の
  // 別商品行に名指しされた分を、コードの合わない行が横取りしないためのガードに使う。
  const 注文所有キー_={};
  lines.forEach(l=>{ const s=注文所有キー_[l.ban]||(注文所有キー_[l.ban]=new Set());
    candKeys(l).forEach(k=>{ s.add(k); if(aliasMap[k]!=null) s.add(aliasMap[k]); }); });
  // 【確定引き当て】発注共有ファイルEMSリストのP列(注文番号)で名指しされた分を最優先で割り当てる
  let 確定行数=0; // 確定マップは上(入荷日ズレ修復)で読み込み済み
  // 入荷日が箱とズレていても、現在「到着済」のP列で名指しされた注文だけは今回便を正とする。
  // 同一コードというだけでは修復しないため、在庫買い分や後発注文を誤って黄色にしない。
  const 確定残必要_=l=>Math.max(0, 残必要_(l)-(l.入荷?(l.過去箱分||0):0));
  lines.filter(l=>l.kbn==='取り寄せ' && (!l.入荷 || !入荷消費OK_(l)) && l.qty>0 && 確定残必要_(l)>0 && 確定[l.ban])
    .sort((a,b)=> a.sortKey-b.sortKey || a.i-b.i).forEach(l=>{
    const cand=candKeys(l); let took=0;
    確定[l.ban].forEach(e=>{
      if(確定残必要_(l)<=0 || e.qty<=0) return;
      // P列の行のコード(e.key)がこの受注行の候補キー(別名込み)に含まれるか
      if(!cand.some(k=> k===e.key || aliasMap[k]===e.key)) return;
      const take=Math.min(確定残必要_(l), e.qty, stock[e.key]||0);
      if(take<=0) return;
      // 古い予定日が残った行は、P列で現在便への名指しが確認できた時だけ未着へ戻して再引当する。
      if(l.入荷 && !入荷消費OK_(l)){
        l.今回P旧入荷日値=l.入荷日値;
        if(l.過去箱分){ l.履歴Alloc=(l.履歴Alloc||0)+l.過去箱分; l.過去箱分=0; }
        l.入荷=false; l.着=false;
      }
      stock[e.key]-=take; e.qty-=take; l.alloc+=take; took+=take;
      l.今回P=true;
      if(!l.今回P入荷日値 && e.arrival) l.今回P入荷日値=e.arrival;
      if(!l.今回PEMS && e.ems) l.今回PEMS=e.ems;
      if(!l.matchedKey) l.matchedKey=e.key;
    });
    if(took>0) 確定行数++;
    l.引当成立 = l.alloc>0 && 残必要_(l)<=0 && l.qty>0;
  });
  // 【確定引き当て・コード不一致救済】REQ系(人名リクエスト)や月号付きコードは受注側とコードが一致しない。
  // その行の商品コードがEMS在庫のどこにも無い場合に限り、P列の名指し(受注番号)だけを信じて割り当てる。
  // ※「在庫がある別の行」が他の商品の名指し分を横取りしないよう keyInStock==null の行に限定
  lines.filter(l=>l.kbn==='取り寄せ' && !l.入荷 && l.qty>0 && 残必要_(l)>0 && 確定[l.ban] && keyInStock(l)==null)
    .sort((a,b)=> a.sortKey-b.sortKey || a.i-b.i).forEach(l=>{
    let took=0;
    確定[l.ban].forEach(e=>{
      if(残必要_(l)<=0 || e.qty<=0) return;
      // 名指しコード(e.key)が同じ注文の別商品行の分なら横取りしない(複数商品注文で、相方が入荷して
      // 名指しされた分を、まだ入荷していないこの行が奪って誤って入荷日が付くのを防ぐ)。
      // REQ系(どの行のコードにも一致しない名指し)だけが救済で引き当たる。
      if(注文所有キー_[l.ban] && 注文所有キー_[l.ban].has(e.key)) return;
      const take=Math.min(残必要_(l), e.qty, stock[e.key]||0);
      if(take<=0) return;
      stock[e.key]-=take; e.qty-=take; l.alloc+=take; took+=take;
      if(!l.matchedKey) l.matchedKey=e.key;
    });
    if(took>0) 確定行数++;
    l.引当成立 = l.alloc>0 && 残必要_(l)<=0 && l.qty>0;
  });
  // 残り(名指しなし・名指しで足りない分)を未着の取り寄せに引き当て(古い注文順のFIFO)
  lines.filter(l=>l.kbn==='取り寄せ' && !l.入荷 && 残必要_(l)>0).sort((a,b)=> a.sortKey-b.sortKey || a.i-b.i).forEach(l=>{
    const k=findAvail(l); if(k){ const got=Math.min(残必要_(l), stock[k]); stock[k]-=got; l.alloc+=got; if(!l.matchedKey) l.matchedKey=k; }
    l.引当成立 = l.alloc>0 && 残必要_(l)<=0 && l.qty>0;
  });

  // 【箱の割当を各行に付ける】どのEMS番号(箱)から出すかを、今回入荷EMSの在庫と同じFIFOで割り出して
  // 各注文行に l.箱EMS として持たせる(出荷可能などのシートにEMS番号列で出す→ピッキングで箱が分かる)
  {
    const consumersByCode={};
    const pushC=(l, qty, kind)=>{ const k=l.matchedKey||keyInStock(l); if(k==null||qty<=0) return; (consumersByCode[k]=consumersByCode[k]||[]).push({qty, ban:l.ban, k, kind}); };
    出荷済行.filter(出荷済消費OK_).forEach(l=> pushC(l, l.qty, '出荷済'));
    lines.filter(l=>l.kbn==='取り寄せ' && l.入荷 && l.qty>0 && 入荷消費OK_(l)).sort((a,b)=> a.sortKey-b.sortKey || a.i-b.i).forEach(l=> pushC(l, l.qty-(l.過去箱分||0), '今日着'));
    lines.filter(l=>l.kbn==='取り寄せ' && !l.入荷 && l.alloc>0).sort((a,b)=> a.sortKey-b.sortKey || a.i-b.i).forEach(l=> pushC(l, l.alloc, '引当'));
    const ptr={};
    const 箱Map={}; // ban|code -> Set(EMS番号)
    E.forEach(row=>{
      const c=normCode_(row[EC.コード]); const qty=Number(row[EC.数量])||0; if(!c||qty<=0) return;
      const ems=String(row[EC.EMS番号]||'').trim();
      const q=consumersByCode[c]||[]; const st=ptr[c]||{i:0,used:0}; let left=qty;
      while(left>0 && st.i<q.length){
        const e=q[st.i], avail=e.qty-st.used, take=Math.min(left, avail);
        if(take>0){ if(ems){ const key=e.ban+'|'+e.k; (箱Map[key]=箱Map[key]||new Set()).add(ems); } st.used+=take; left-=take; }
        if(st.used>=e.qty){ st.i++; st.used=0; }
      }
      ptr[c]=st;
    });
    lines.forEach(l=>{ const k=l.matchedKey||keyInStock(l); const s= k? 箱Map[l.ban+'|'+k]:null;
      l.箱EMS=l.今回PEMS || (s? Array.from(s).join(', '):''); });
  }

  // 注文ごと: 入金済み(1行でも入金日があれば入金済み。注文単位入金)
  const byOrder={};
  lines.forEach(l=> (byOrder[l.ban]=byOrder[l.ban]||[]).push(l));
  const paidOrder={};
  Object.keys(byOrder).forEach(ban=>{ paidOrder[ban]=byOrder[ban].some(l=>l.paid); });

  // 注文を分類:
  //  出荷可能=今回分(黄)あり&全行出せる&入金済 / 希望日待ち=全行出せるが希望日が未来
  //  出荷GO未入金=全行出せるが未入金 / 部分在庫=黄あり+未入荷(色なし)あり / 引当待ち=黄なし
  const 区分け=ban=> 注文区分判定_(byOrder[ban], paidOrder[ban], 入荷消費OK_);
  // 引当待ちの「発送可否」(K列): 揃っているのに今回分でない=実は発送できる注文を見分ける
  const 発送可否_=ban=> 発送可否判定_(byOrder[ban], paidOrder[ban]);

  // 出力(行の色=到着状況。未入金は受注番号セルだけ赤)
  const HDR=['受注番号','氏名','お届け日','商品コード','商品名','個数','区分','入荷日','入金','状態','EMS番号'];
  const HDR_待=HDR.concat(['発送可否']); // 引当待ちだけK列に発送可否を足す
  const seen=new Set(), seq=[]; lines.forEach(l=>{ if(!seen.has(l.ban)){ seen.add(l.ban); seq.push(l.ban);} });
  const waitRows=[], partRows=[], keepRows=[], holdRows=[], shipRows=[];
  seq.forEach(ban=>{
    const b=区分け(ban), paid=paidOrder[ban];
    const target = b==='ship'? shipRows : b==='keep'? keepRows : b==='hold'? holdRows : b==='part'? partRows : waitRows;
    const 可否 = (b==='wait')? 発送可否_(ban) : null; // 引当待ちのみK列に発送可否
    byOrder[ban].forEach(l=>{
      const 状態=引当行状態_(l, cfg, 入荷消費OK_);
      let st=状態.st, color=状態.color;
      if(!l.キャンセル && !paid) st+='・入金待ち';
      const vals=[ban,l.氏名,l.届,l.code,l.商品名,l.qty,l.kbn,l.入荷日値,paid?'済':'未',st,l.箱EMS||''];
      if(b==='wait') vals.push(可否);
      target.push({ vals, color, ban, paid, qty:l.qty });
    });
  });

  書き出し_(ss, cfg.待ち, HDR_待, waitRows, 受注hdr);
  書き出し_(ss, cfg.部分, HDR, partRows, 受注hdr);
  書き出し_(ss, cfg.取置, HDR, keepRows, 受注hdr);
  書き出し_(ss, cfg.希望, HDR, holdRows, 受注hdr);
  書き出し_(ss, cfg.出荷, HDR, shipRows, 受注hdr);

  // ===== 受注明細の色分け: 行=到着色(即納=水/到着=黄/未着=白/指定なし=橙)、未入金は受注番号セルを赤 =====
  const recvStart=受注hdr+1, recvLast=recv.getLastRow();
  if(recvLast>=recvStart){
    const nc=recv.getLastColumn(), rowColor={}, rowUnpaid={}, rowQty={};
    lines.forEach(l=>{
      let col=null;
      if(l.キャンセル) col=cfg.色_グレー;       // 個数0=キャンセル
      else if(l.kbn==='即納') col=cfg.色_水;
      else if(l.kbn==='指定なし') col=cfg.色_橙;
      else if(l.kbn==='取り寄せ'){
        col=引当行状態_(l, cfg, 入荷消費OK_).color; // 今回EMS到着日と一致する着済も黄。それ以外の過去着済は薄ラベンダー
        // else: 在庫待ち=白(未入金は受注番号だけ赤)
      }
      rowColor[l.i]=col;
      rowUnpaid[l.i]=!paidOrder[l.ban];
      rowQty[l.i]=l.qty;
    });
    const bg=[];
    for(let r=recvStart;r<=recvLast;r++){
      const idx=r-1, c=rowColor[idx];
      const arr=new Array(nc).fill(c!==undefined?c:null);
      if(rowUnpaid[idx]) arr[M.番号]=cfg.色_赤;                 // 未入金は受注番号セルを赤
      if(rowQty[idx]>=2 && M.個数>=0) arr[M.個数]=cfg.色_緑;    // 個数2以上は個数セル緑
      bg.push(arr);
    }
    const dataRng=recv.getRange(recvStart,1,bg.length,nc);
    dataRng.setBackgrounds(bg);
    dataRng.setFontSize(cfg.字).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    recv.setRowHeights(recvStart, bg.length, cfg.行高);
    注文罫線_(recv, recvStart, M.番号);
  }

  // EMS在庫タブも色分け(黄=未着へ全引当 / 緑=一部 / 白=入荷済 or 余り)
  EMS消込色付け_();

  const jpRows=[]; // 純日本在庫(行ごとの余り。EMS番号付き)
  const 突合宙=[]; let 突合式=''; // 突き合わせ検算: 到着済=出荷済+割当+余り / 箱に乗り切らなかった割当の一覧
  // ===== 今回入荷EMSの在庫(台帳・色付き): 各EMS行に 状態と引き当てた受注番号 を出力 =====
  {
    // コードごとの消費者キュー(出荷済み=もう発送 → 割当済=入荷日あり → 引当=今回 の順)
    const consumersByCode={};
    const pushCons=(l, qty, kind)=>{ const k=l.matchedKey||keyInStock(l); if(k==null||qty<=0) return; (consumersByCode[k]=consumersByCode[k]||[]).push({qty, ban:l.ban, kind}); }; // 確定引当(コード不一致救済含む)はmatchedKeyの行に記帳
    出荷済行.filter(出荷済消費OK_).forEach(l=> pushCons(l, l.qty, '出荷済'));                             // 消込台帳の出荷済み(入荷日=今回の到着日の分だけ。別便から出た出荷済みは今回の箱を食わない)
    lines.filter(l=>l.kbn==='取り寄せ' && l.入荷 && l.qty>0 && 入荷消費OK_(l))
      .sort((a,b)=> a.sortKey-b.sortKey || a.i-b.i) // 古い注文を先頭の箱行に乗せる(箱↔注文の対応を注文順に)
      .forEach(l=> pushCons(l, l.qty-(l.過去箱分||0), '今日着'));                                                         // ゲート通過=入荷日が今回の到着日と一致=今回の便で出す分。全て黄(今回出せる分)で見せる。別便で処理済み(入荷日≠今回到着日)はゲートで除外され日本在庫に残る。ラベンダー(割当済)は廃止=このシートは「黄=出す/色なし=在庫」に統一
    lines.filter(l=>l.kbn==='取り寄せ' && !l.入荷 && l.alloc>0).sort((a,b)=> a.sortKey-b.sortKey || a.i-b.i).forEach(l=> pushCons(l, l.alloc, '引当')); // 今回引き当て
    const ptr={};
    const consumeRow=(code, qty)=>{
      const q=consumersByCode[code]||[]; const st=ptr[code]||{i:0,used:0};
      const got={出荷済:[], 割当済:[], 今日着:[], 引当:[]}; let left=qty;
      while(left>0 && st.i<q.length){
        const e=q[st.i], avail=e.qty-st.used, take=Math.min(left, avail);
        if(take>0){
          const arr=got[e.kind]; const prev=arr.find(g=>g.ban===e.ban);
          if(prev) prev.qty+=take; else arr.push({ban:e.ban, qty:take}); // 同じ受注は個数を合算(2個以上は「番号:個数」で表示)
          st.used+=take; left-=take;
        }
        if(st.used>=e.qty){ st.i++; st.used=0; }
      }
      ptr[code]=st; return {got, surplus:left};
    };
    const EHDR=['状態','到着日','商品コード','数量','EMS番号','余り','引き当てた受注番号（左=古い → 右=新しい）'];
    const ledgerRows=[]; // {vals,bg} 受注番号を1セル1番号で右へ展開
    let 突_供給=0, 突_出荷=0, 突_割当=0, 突_余り=0; // 突き合わせ用の集計
    E.forEach(row=>{
      const c=normCode_(row[EC.コード]), qty=Number(row[EC.数量])||0;
      if(!c) return;
      let col=null, surplus=0; const cells=[]; // cells={ban,color}
      if(qty>0){
        const r=consumeRow(c, qty);
        // 出荷済みの消費は余りの計算にだけ効かせて、受注番号は表示しない(このシートは「これから出す分」の引き当て先を見る場所)
        r.got.割当済.forEach(g=> cells.push({ban:g.ban, qty:g.qty, color: paidOrder[g.ban]? cfg.色_着 : cfg.色_赤})); // 過去に着いた割当済=ラベンダー(未入金は赤)
        r.got.今日着.forEach(g=> cells.push({ban:g.ban, qty:g.qty, color: paidOrder[g.ban]? cfg.色_黄 : cfg.色_赤})); // 入荷日=今日(今回出せる分)=黄(未入金は赤)
        r.got.引当.forEach(g=> cells.push({ban:g.ban, qty:g.qty, color: paidOrder[g.ban]? cfg.色_黄 : cfg.色_赤}));   // 今回引当=黄(未入金は赤)
        cells.sort((a,b)=> 番号num_(a.ban)-番号num_(b.ban)); // 古い注文を左・新しいを右
        surplus=r.surplus;
        突_供給+=qty; 突_余り+=r.surplus;
        突_出荷+=r.got.出荷済.reduce((s,g)=>s+g.qty,0);
        突_割当+=r.got.割当済.concat(r.got.今日着, r.got.引当).reduce((s,g)=>s+g.qty,0);
        if(r.got.引当.length || r.got.今日着.length) col = r.surplus>0 ? cfg.色_緑 : cfg.色_黄; // 今回分あり: 余りあり=緑 / 全部今回分=黄
        else if(r.got.割当済.length) col=cfg.色_着;                   // 過去の割当済のみ=ラベンダー
        // 出荷済みが取った分は色なし(番号・余りも出ない=何もすることがない行。記録は消込台帳にある)
        if(r.surplus>0) jpRows.push([row[0], 到着_(row), c, r.surplus, row[EC.EMS番号]]); // 状態/到着日/商品コード/余り数/EMS番号
      }
      const vals=[row[0], 到着_(row), c, qty, row[EC.EMS番号], surplus>0?surplus:''].concat(cells.map(x=> x.ban+(x.qty>1?'*'+x.qty:'')));
      const bg=[col,col,col,col,col,col].concat(cells.map(x=>x.color));     // A〜F=行の状態色 / G以降=注文ごとの色
      ledgerRows.push({vals, bg});
    });
    // ===== 突き合わせ: 箱(到着済)に乗り切らなかった消費者を検出 =====
    // 入荷日付き/出荷済みの数量が箱の数量を超えている=二重引当・誤入荷日・台帳の重複の疑い。
    // (今回引当はstock残でキャップされるので、ここに出るのは入荷日付き/出荷済みだけのはず)
    Object.keys(consumersByCode).forEach(code=>{
      const q=consumersByCode[code], st=ptr[code]||{i:0,used:0};
      for(let i=st.i;i<q.length;i++){
        const left=q[i].qty-(i===st.i? st.used:0);
        if(left>0) 突合宙.push('・'+q[i].ban+' '+code+' x'+left+'（'+(q[i].kind==='今日着'?'着済':q[i].kind)+'）');
      }
    });
    突合式='到着済'+突_供給+'＝出荷済'+突_出荷+'＋割当'+突_割当+'＋余り'+突_余り;
    const maxCols=ledgerRows.reduce((m,r)=>Math.max(m,r.vals.length), EHDR.length);
    let esh=ss.getSheetByName(cfg.日本在庫), esh新規=!esh; if(!esh) esh=ss.insertSheet(cfg.日本在庫);
    シート値クリア_(esh); // 値・色だけ入れ替え(列幅・行高・中央揃えは保持)
    esh.getRange(1,1).setValue('最終引当: '+Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm:ss')+' / '+ledgerRows.length+'件'
      +' ｜ 突合せ '+(突合宙.length? '⚠️'+突合宙.length+'件不一致' : 'OK『'+突合式+'』'));
    const header=EHDR.slice(); while(header.length<maxCols) header.push('');
    esh.getRange(2,1,1,maxCols).setValues([header]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff').setFontSize(cfg.字);
    esh.setFrozenRows(2); if(esh新規) esh.setRowHeight(2, cfg.行高);
    if(ledgerRows.length){
      const vals=ledgerRows.map(r=>{ const v=r.vals.slice(); while(v.length<maxCols) v.push(''); return v; });
      const bgs=ledgerRows.map(r=>{ const b=r.bg.slice(); while(b.length<maxCols) b.push(null); return b; });
      const rng=esh.getRange(3,1,vals.length,maxCols);
      rng.setValues(vals).setFontSize(cfg.字);
      rng.setBackgrounds(bgs);
      if(esh新規) esh.setRowHeights(3, vals.length, cfg.行高);
      esh.getRange(3,2,vals.length,1).setNumberFormat('yyyy-mm-dd'); // 到着日(B列)を 2026-07-02 形式に
      if(esh新規){
        [90,90,210,70,150,60].forEach((w,c)=> esh.setColumnWidth(c+1,w)); // 状態/到着日/商品コード/数量/EMS番号/余り
        for(let c=6;c<maxCols;c++) esh.setColumnWidth(c+1,105); // 受注番号セルの幅
      }
    } else {
      esh.getRange(3,1).setValue('(EMS在庫なし)');
    }
  }

  // ===== 純粋な日本在庫(EMSで引き当たらず余った分)を行ごと(EMS番号付き)に出力 =====
  {
    const JHDR=['状態','到着日','商品コード','余り数(日本在庫)','EMS番号'];
    let jsh=ss.getSheetByName(cfg.純在庫), jsh新規=!jsh; if(!jsh) jsh=ss.insertSheet(cfg.純在庫);
    シート値クリア_(jsh); // 値・色だけ入れ替え(書式は保持)
    jsh.getRange(1,1).setValue('最終引当: '+Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm:ss')+' / '+jpRows.length+'件');
    jsh.getRange(2,1,1,JHDR.length).setValues([JHDR]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff').setFontSize(cfg.字);
    jsh.setFrozenRows(2); if(jsh新規) jsh.setRowHeight(2, cfg.行高);
    if(jpRows.length){
      jsh.getRange(3,1,jpRows.length,JHDR.length).setValues(jpRows).setFontSize(cfg.字);
      jsh.getRange(3,2,jpRows.length,1).setNumberFormat('yyyy-mm-dd'); // 到着日(B列)を 2026-07-02 形式に
      if(jsh新規){
        jsh.setRowHeights(3, jpRows.length, cfg.行高);
        [90,90,210,120,150].forEach((w,c)=> jsh.setColumnWidth(c+1,w)); // 状態/到着日/商品コード/余り数/EMS番号
      }
    } else { jsh.getRange(3,1).setValue('(日本在庫なし)'); }
  }

  const ord=seq.length;
  const ship=shipRows.length? new Set(shipRows.map(r=>r.ban)).size:0;
  const keep=keepRows.length? new Set(keepRows.map(r=>r.ban)).size:0;
  const hold=holdRows.length? new Set(holdRows.map(r=>r.ban)).size:0;
  const part=partRows.length? new Set(partRows.map(r=>r.ban)).size:0;
  const mp=Object.keys(paidOrder).filter(b=>!paidOrder[b]).length;
  // 【入荷日を自動記入】EMS在庫で引き当たった未着の注文へ、その商品のEMS到着日を入荷日として書く(空欄のみ・手入力は上書きしない)
  let 自動入荷=0, 入荷表示統一=0, P入荷日訂正=0;
  if(M.入荷>=0){
    const rs=受注hdr+1, rl=recv.getLastRow();
    if(rl>=rs){
      const 入荷col=recv.getRange(rs,M.入荷+1,rl-rs+1,1).getValues();
      lines.forEach(l=>{
        const idx=l.i-受注hdr; if(idx<0 || idx>=入荷col.length) return;
        if(l.今回P && l.引当成立){
          const v=l.今回P入荷日値 || (l.matchedKey?code到着[l.matchedKey]:'');
          const d=ymd_(v);
          if(/^20\d{2}-\d{2}-\d{2}$/.test(d)){
            入荷col[idx][0]=new Date(d+'T00:00:00'); l.入荷日値=入荷col[idx][0]; P入荷日訂正++;
          }
          return;
        }
        if(l.今回P旧入荷日値 && !l.引当成立){ 入荷col[idx][0]=''; P入荷日訂正++; return; }
        if(l.修復入荷){ 入荷col[idx][0]=l.入荷日値; return; } // ズレ修復(P列に無い注文): 箱の到着日への貼り直しをシートへ反映
        if(l.入荷) return;
        if(l.履歴成立 && l.履歴入荷日値){
          入荷col[idx][0]=l.履歴入荷日値; 自動入荷++;
        } else if(l.引当成立 && l.matchedKey && code到着[l.matchedKey]!==undefined){
          入荷col[idx][0]=code到着[l.matchedKey]; 自動入荷++;
        }
      });
      // 【表示ゆれの正規化】手入力の「26/07/09(木)」等の文字列日付を日付値に直し、列全体を yyyy-MM-dd 表示に統一する
      // (自動記入は日付値・手入力は文字列で形がバラバラになるため。日付らしい文字列だけ触り、読めないメモ等はそのまま)
      for(let idx=0; idx<入荷col.length; idx++){
        const v=入荷col[idx][0];
        if(v==='' || v==null || v instanceof Date) continue;
        const t=String(v).trim();
        if(!/^(?:20\d{2}|\d{2})[\/\-.]\d{1,2}[\/\-.]\d{1,2}/.test(t)) continue;
        const s=ymd_(t); if(!/^20\d{2}-\d{2}-\d{2}$/.test(s)) continue;
        const dd=new Date(s+'T00:00:00'); if(isNaN(dd.getTime())) continue;
        入荷col[idx][0]=dd; 入荷表示統一++;
      }
      if(自動入荷||入荷表示統一||P入荷日訂正||ズレ修復){ recv.getRange(rs,M.入荷+1,入荷col.length,1).setValues(入荷col).setNumberFormat('yyyy-mm-dd'); }
    }
  }

  // 【EMS番号を記入】各行がどの箱から出すかを、個数の隣の「EMS番号」列に書く(取込で個数の隣に作られている)
  {
    const rs=受注hdr+1, rl=recv.getLastRow();
    if(rl>=rs){
      const emsC=EMS番号列を用意_(recv); // 個数の隣。無ければ作る(1始まり)
      const col=recv.getRange(rs,emsC,rl-rs+1,1).getValues();
      for(let idx=0;idx<col.length;idx++) col[idx][0]=''; // 一旦クリア(前回の割当が残らないように)
      lines.forEach(l=>{ const idx=l.i-受注hdr; if(idx>=0 && idx<col.length && l.箱EMS) col[idx][0]=l.箱EMS; });
      recv.getRange(rs,emsC,col.length,1).setValues(col);
    }
  }
  抽出フラグ更新_(); // 受注明細の「抽出」列(○=色あり/×=色なし)を更新→普通のフィルタで絞れる
  const 隠れ = recv.getFilter()? 'フィルタON' : (PropertiesService.getDocumentProperties().getProperty('色なし除外中')==='1'? '🙈色なし除外中' : '');
  const filt = 隠れ? ' ⚠️'+隠れ+'(行が隠れてるかも→🔻フィルタ確認/解除で全解除)' : '';
  try{ 在庫EMS番号罫線ボタンを設置_(true); }catch(e){} // シート値クリアで消えた罫線チェックボックス＋ラベルを毎回復元
  // 誤入荷日チェックは②では走らせない(発注共有EMSリスト+受注明細を丸ごと再読み込みするため重く、
  // ②がタイムアウトしかける)。必要時にメニュー「🔎 入荷日の整合チェック」で手動実行する。
  // 完了サマリは大きなダイアログで表示(トーストは小さく末尾が切れるため)
  {
    const ui=SpreadsheetApp.getUi();
    // 【入荷日ズレの検知】入荷日はあるのに今の「到着済」の箱の到着日と一致せず(=ラベンダー扱い)、
    // かつ同じ商品が今回の到着済在庫に存在する行。EMS追跡の同期などで箱の到着日が後から実日付に
    // 直ると、②が先に書いたスタンプが古くなって一斉にこの状態になる(例: 7/9記入→箱が7/10に更新)。
    // 引当履歴で過去箱の受取が判明している行(過去箱分でカバー済み)は正常な着済なので除く
    const 入荷日ズレ=[];
    lines.forEach(l=>{
      if(l.kbn!=='取り寄せ' || !l.入荷 || l.qty<=0) return;
      if((l.過去箱分||0)>=l.qty) return;
      if(入荷消費OK_(l)) return;
      const k=keyInStock(l); if(k==null) return; // 商品自体が今の到着済に無い=過去便の正常な着済
      const set=到着日SetByCode[k]; if(!set || !set.size) return;
      入荷日ズレ.push('・'+l.ban+' '+(l.code||l.sku)+' 入荷日'+ymd_(l.入荷日値)+'（箱は'+Array.from(set).sort().join(',')+'）');
    });
    // 要確認(宙)だけを目立たせ、正常な出荷済み消込は件数だけ(詳細は発送済みシート)にして情報過多を避ける
    const 突合行= 突合宙.length
      ? '⚠️ 要確認：箱と合わない割当 '+突合宙.length+'件（二重引当/誤入荷日の疑い）\n'
        +突合宙.slice(0,15).join('\n')+(突合宙.length>15? '\n…他'+(突合宙.length-15)+'件':'')
        +'\n→ 気になる分だけ 🔎商品診断 で1個単位で確認（急ぎでなければ後回しでOK）'
      : '✅ 整合OK 『'+突合式+'』';
    const ok= 突合宙.length===0 && 入荷日ズレ.length===0;
    // 便の締め(到着済を在庫反映済みへ)のガード用に、②の整合状態を残す。
    // ⚠️が残ったまま締めると幽霊スタンプ(実物なしのラベンダー)が過去便として固定されるため。
    try{ PropertiesService.getDocumentProperties().setProperty('引当_整合状態',
      JSON.stringify({ts:Date.now(), 要確認:突合宙.length+入荷日ズレ.length})); }catch(e){}
    ui.alert(ok? '✅ 引当完了（整合OK）' : '⚠️ 引当完了（要確認 '+(突合宙.length+入荷日ズレ.length)+'件）',
      '■ 突合せ（引当＋在庫 と EMS到着済）\n'+突合行
      +(入荷日ズレ.length? '\n\n⚠️ 入荷日が箱の到着日とズレてラベンダーの行 '+入荷日ズレ.length+'件\n'
      +入荷日ズレ.slice(0,10).join('\n')+(入荷日ズレ.length>10?'\n…他'+(入荷日ズレ.length-10)+'件':'')
      +'\n→ 🔎入荷日の整合チェック → 🧹一括クリア → ②再実行 で正しい日付が入り直ります' : '')
      +'\n\n■ 注文の分類\n出荷可能 '+ship+' ／ 出荷GO未入金 '+keep+' ／ 希望日待ち '+hold+' ／ 部分在庫 '+part+' ／ 引当待ち '+(ord-ship-keep-hold-part)+'（うち入金待ち '+mp+'）'
      +'\n\n■ 処理内容\n入荷日自動記入 '+自動入荷+'件'+(入荷表示統一?'（表示をyyyy-MM-ddに統一 '+入荷表示統一+'件）':'')
      +(P入荷日訂正? ' ／ 🔧P列確定の入荷日訂正 '+P入荷日訂正+'件':'')
      +(ズレ修復? ' ／ 🔧入荷日ズレ修復 '+ズレ修復+'件（P列に無い注文を箱の到着日に貼り直し）':'')
      +' ／ P列確定 '+確定行数+'行 ／ 発送済み消込 '+出荷済行.length+'件（内訳は「発送済み」シート）'
      +(隠れ? '\n\n⚠️ '+隠れ+'：行が隠れています（🔻フィルタ確認/解除で全解除）':''),
      ui.ButtonSet.OK);
  }
}

// EMS在庫の引当色分け: 在庫から まず入荷日あり(割当済)を差引き、残りを入荷日なし(未着)の人へ引当
// 黄=その行まるごと未着の人へ引当 / 緑=一部 / 白=入荷済が取った分 or 余り
function EMS消込色付け_(){
  const ss=SpreadsheetApp.getActive(), cfg=HIKIATE_CFG;
  const recv=ss.getSheetByName(cfg.受注), emv=ss.getSheetByName(cfg.EMS在庫);
  if(!recv||!emv) return;
  const R=recv.getDataRange().getValues();
  const ems=EMS明細_(emv), E=ems.rows, EC=ems.cols; // 1行ヘッダーがあれば除外。EC=見出しで特定した列位置
  if(!E.length) return;
  const orig={};
  for(let i=0;i<E.length;i++){ const c=normCode_(E[i][EC.コード]); if(!c) continue; orig[c]=(orig[c]||0)+(Number(E[i][EC.数量])||0); }
  const aliasMap={}; Object.keys(orig).forEach(c=> codeKeys_(c).forEach(k=>{ if(!(k in aliasMap)) aliasMap[k]=c; }));
  const keyOf=l=>{ const cands=受注候補コード_(l.sku,l.code);
    const keys=[]; cands.forEach(v=> codeKeys_(v).forEach(k=>{ if(keys.indexOf(k)<0) keys.push(k); }));
    for(const k of keys){ if(k in orig) return k; if(aliasMap[k] && aliasMap[k] in orig) return aliasMap[k]; } return null; };
  // 取り寄せ需要を 入荷日あり(割当済) と 入荷日なし(未着) に分けて集計
  const M=列マップ_(recv), need済={}, need未={};
  for(let i=M.hr;i<R.length;i++){
    const row=R[i]; const ban=String(row[M.番号]||'').trim(), code=String(row[M.コード]||'').trim();
    if(!ban && !code) continue;
    if(区分_(row[M.選択肢])!=='取り寄せ') continue;
    const k=keyOf({code, sku:String(row[M.SKU]||'').trim()}); if(!k) continue;
    const qty=Number(row[M.個数])||0;
    const 入荷=M.入荷>=0 && String(row[M.入荷]||'').trim()!=='';
    if(入荷) need済[k]=(need済[k]||0)+qty; else need未[k]=(need未[k]||0)+qty;
  }
  // 出荷済み(消込台帳)の分も「処理済みの取り分」として色なし側に加える
  消込台帳_出荷済み行_().forEach(l=>{ const k=keyOf(l); if(k) need済[k]=(need済[k]||0)+l.qty; });
  // コードごと: 在庫から入荷済の取り分(色なし)を除き、残りを未着へ引当(色付け)
  const skipLeft={}, colorLeft={};
  Object.keys(orig).forEach(k=>{
    const s=orig[k]||0, d済=need済[k]||0, d未=need未[k]||0;
    skipLeft[k]=Math.min(d済, s);
    colorLeft[k]=Math.min(d未, Math.max(0, s-d済));
  });
  const ncolE=E[0].length;
  const emsBg=E.map(row=>{
    const c=normCode_(row[EC.コード]), qty=Number(row[EC.数量])||0;
    let col=null;
    if(c && qty>0){
      let q=qty;
      const sk=Math.min(skipLeft[c]||0, q); skipLeft[c]=(skipLeft[c]||0)-sk; q-=sk; // 入荷済の取り分はスキップ(色なし)
      if(q>0 && (colorLeft[c]||0)>0){
        const use=Math.min(q, colorLeft[c]); colorLeft[c]-=use;
        col = (sk===0 && use>=qty) ? cfg.色_黄 : cfg.色_緑; // 行まるごと未着引当=黄 / 一部=緑
      }
    }
    return new Array(ncolE).fill(col);
  });
  emv.getRange(ems.offset,1,E.length,ncolE).setBackgrounds(emsBg); // ヘッダーの下(データ行)から着色
}

// 🔍 在庫照合レポート: コードごとに 在庫/取り寄せ要求/引当済/残/不足 を一覧化
function 在庫照合レポート(){
  const ss=SpreadsheetApp.getActive(), cfg=HIKIATE_CFG;
  const recv=ss.getSheetByName(cfg.受注), emv=ss.getSheetByName(cfg.EMS在庫);
  if(!recv||!emv){ SpreadsheetApp.getUi().alert('受注明細/EMS在庫タブが無いで'); return; }
  const R=recv.getDataRange().getValues(); const emsD=EMS明細_(emv), E=emsD.rows, EC=emsD.cols;

  // 在庫
  const stock={};
  for(let i=0;i<E.length;i++){ const c=normCode_(E[i][EC.コード]); if(!c) continue; stock[c]=(stock[c]||0)+(Number(E[i][EC.数量])||0); }
  const orig=Object.assign({},stock);
  const aliasMap={};
  Object.keys(stock).forEach(c=> codeKeys_(c).forEach(k=>{ if(!(k in aliasMap)) aliasMap[k]=c; }));
  const candKeys=l=>{ const cands=受注候補コード_(l.sku,l.code);
    const keys=[]; cands.forEach(v=> codeKeys_(v).forEach(k=>{ if(keys.indexOf(k)<0) keys.push(k); })); return keys; };

  // 取り寄せ行を集める(列は見出し名で特定)
  const M=列マップ_(recv); const lines=[];
  for(let i=M.hr;i<R.length;i++){
    const row=R[i]; const ban=String(row[M.番号]||'').trim(), code=String(row[M.コード]||'').trim();
    if(!ban && !code) continue;
    if(区分_(row[M.選択肢])!=='取り寄せ') continue;
    lines.push({ ban, sortKey:番号num_(ban), code, sku:String(row[M.SKU]||'').trim(), qty:Number(row[M.個数])||0 });
  }

  // 要求集計(在庫に存在するキーへ寄せる。無ければ「在庫なし」コード扱い)
  const demand={};
  const demandKeyOf=l=>{ for(const k of candKeys(l)){ if(k in orig) return k; if(aliasMap[k] && aliasMap[k] in orig) return aliasMap[k]; }
    return '★在庫なし:'+(normCode_(l.sku)||normCode_(l.code)); };
  lines.forEach(l=>{ const k=demandKeyOf(l); demand[k]=(demand[k]||0)+l.qty; });

  // 本番と同じFIFOで消費(=引当済を算出)
  const findKey=l=>{ for(const k of candKeys(l)){ if(stock[k]>0) return k; if(aliasMap[k]&&stock[aliasMap[k]]>0) return aliasMap[k]; } return null; };
  lines.slice().sort((a,b)=> a.sortKey-b.sortKey).forEach(l=>{ const k=findKey(l); if(k){ const g=Math.min(l.qty,stock[k]); stock[k]-=g; } });

  // レポート行(在庫があるコード ∪ 要求キー)
  const allKeys={}; Object.keys(orig).forEach(k=>allKeys[k]=1); Object.keys(demand).forEach(k=>allKeys[k]=1);
  const rows=Object.keys(allKeys).map(k=>{
    const z=orig[k]||0, d=demand[k]||0, rem=(k in stock)?stock[k]:0, alloc=z-rem, sho=Math.max(0,d-alloc);
    return [k, z, d, alloc, rem, sho];
  }).sort((a,b)=> b[5]-a[5] || b[1]-a[1] || (a[0]<b[0]?-1:1)); // 不足多い順→在庫多い順→コード順

  const HDR=['在庫コード','在庫数','取り寄せ要求','引当済','残在庫','不足'];
  const hr=受注ヘッダー行_(recv), dr=hr+1; // 受注明細とヘッダー行をそろえる
  let sh=ss.getSheetByName('照合レポート'), 新規=!sh; if(!sh) sh=ss.insertSheet('照合レポート');
  シート値クリア_(sh); // 値・色だけ入れ替え(書式は保持)
  sh.getRange(hr,1,1,6).setValues([HDR]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff').setFontSize(cfg.字);
  sh.setFrozenRows(hr);
  if(新規) sh.setRowHeight(hr, cfg.行高);
  if(rows.length){
    sh.getRange(dr,1,rows.length,6).setValues(rows).setFontSize(cfg.字);
    if(新規) sh.setRowHeights(dr, rows.length, cfg.行高);
    const bg=rows.map(r=> new Array(6).fill( r[5]>0?cfg.色_赤 : (r[4]>0?cfg.色_黄:null) )); // 不足=赤 / 在庫余り=黄
    sh.getRange(dr,1,rows.length,6).setBackgrounds(bg);
    if(新規) [220,80,100,80,80,80].forEach((w,c)=> sh.setColumnWidth(c+1,w));
  }
  ss.toast('照合レポート作成：'+rows.length+'コード（不足=赤 / 在庫余り=黄）','🔍レポート',6);
}

// 受注明細で「色なしだけ(全行が白)の注文」を隠す/全表示する トグル
function 色なし注文の除外切替(){
  const ss=SpreadsheetApp.getActive(), cfg=HIKIATE_CFG, ui=SpreadsheetApp.getUi();
  const recv=ss.getSheetByName(cfg.受注);
  if(!recv){ ui.alert('「'+cfg.受注+'」タブが無いで'); return; }
  const M=列マップ_(recv), startRow=M.hr+1, lastRow=recv.getLastRow(), nc=recv.getLastColumn();
  if(lastRow<startRow){ ss.toast('データが無いで'); return; }
  const n=lastRow-startRow+1;
  const props=PropertiesService.getDocumentProperties();

  // 既に除外中なら全表示に戻す
  if(props.getProperty('色なし除外中')==='1'){
    recv.showRows(startRow, n);
    props.deleteProperty('色なし除外中');
    ss.toast('全注文を表示しました','表示',4);
    return;
  }

  // 色なしの注文を隠す。判定では受注番号セル(赤=未入金)と個数セル(緑)は無視=ステータス色が1つも無ければ色なし扱い
  recv.showRows(startRow, n); // いったん全表示
  const bg=recv.getRange(startRow,1,n,nc).getBackgrounds();
  const bans=recv.getRange(startRow,1,n,M.番号+1).getValues().map(r=>String(r[M.番号]||'').trim());
  const hasColor={};
  for(let i=0;i<n;i++){ const b=bans[i]; if(!b) continue;
    if(!(b in hasColor)) hasColor[b]=false;
    // 受注番号(未入金の赤)と個数(緑)以外の列に色が1つでもあれば「色あり」=表示
    if(bg[i].some((c,ci)=> ci!==M.番号 && ci!==M.個数 && c && c.toLowerCase()!=='#ffffff')) hasColor[b]=true;
  }
  const toHide=[];
  for(let i=0;i<n;i++){ const b=bans[i]; if(b && !hasColor[b]) toHide.push(startRow+i); }
  let k=0, cnt=0;
  while(k<toHide.length){
    let s=toHide[k], e=s;
    while(k+1<toHide.length && toHide[k+1]===e+1){ k++; e=toHide[k]; }
    recv.hideRows(s, e-s+1); cnt+=e-s+1; k++;
  }
  props.setProperty('色なし除外中','1');
  ss.toast('色なし(未入金・個数だけの行も含む)の注文を'+cnt+'行 隠しました(もう一度押すと全表示)','除外',5);
}

// 受注明細の末尾に「抽出」列を作り、注文ごとに ○(即納/引当など色あり) / ×(色なし・数量や未入金だけ) を入れる。
// Googleの普通のフィルタで「抽出=○」だけ表示できるようにする用。②引当・ダニエル引当の最後に呼ぶ(現在の色で判定)。
function 抽出フラグ更新_(){
  const ss=SpreadsheetApp.getActive(), cfg=HIKIATE_CFG;
  const recv=ss.getSheetByName(cfg.受注); if(!recv) return;
  const M=列マップ_(recv), startRow=M.hr+1, lastRow=recv.getLastRow();
  if(lastRow<startRow) return;
  let nc=recv.getLastColumn();
  // 「抽出」列を確保(見出し行から探す。無ければ末尾に作る)
  const head=recv.getRange(M.hr,1,1,nc).getValues()[0].map(v=>String(v||'').trim());
  let flagCol=head.indexOf('抽出')+1; // 1始まり / 0=無し
  if(flagCol===0){
    flagCol=nc+1;
    recv.getRange(M.hr,flagCol).setValue('抽出').setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff').setHorizontalAlignment('center');
    recv.setColumnWidth(flagCol,60); nc=flagCol;
  }
  const n=lastRow-startRow+1;
  const bg=recv.getRange(startRow,1,n,nc).getBackgrounds();
  const bans=recv.getRange(startRow,1,n,M.番号+1).getValues().map(r=>String(r[M.番号]||'').trim());
  const hasColor={};
  for(let i=0;i<n;i++){ const b=bans[i]; if(!b) continue; if(!(b in hasColor)) hasColor[b]=false;
    // 受注番号(未入金の赤)・個数(緑)・抽出列 以外に色が1つでもあれば「色あり」
    if(bg[i].some((c,ci)=> ci!==M.番号 && ci!==M.個数 && ci!==flagCol-1 && c && c.toLowerCase()!=='#ffffff')) hasColor[b]=true;
  }
  recv.getRange(startRow, flagCol, n, 1).setValues(bans.map(b=> [ b? (hasColor[b]?'○':'×') : '' ]));
}

// 共通のヘッダー行(受注明細の「受注番号」がある行)を返す
function 共通ヘッダー行_(ss){ const r=ss.getSheetByName(HIKIATE_CFG.受注); return r? 受注ヘッダー行_(r):1; }

// 指定シートのヘッダーより下のデータを消す(ヘッダー・書式は残す)
function シートデータ消去_(name, skipConfirm){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  const sh=ss.getSheetByName(name);
  if(!sh){ ss.toast('「'+name+'」タブが無いで'); return; }
  if(!skipConfirm){
    const ans=ui.alert('データ消去','「'+name+'」のヘッダーより下のデータを消します。ええ？',ui.ButtonSet.OK_CANCEL);
    if(ans!==ui.Button.OK) return;
  }
  const start=共通ヘッダー行_(ss)+1, mr=sh.getMaxRows(), mc=sh.getMaxColumns();
  if(mr>=start){
    const r=sh.getRange(start,1,mr-start+1,mc);
    r.clearContent();
    r.setBackgrounds(Array.from({length:mr-start+1},()=>new Array(mc).fill(null)));
    r.setBorder(false,false,false,false,false,false);
  }
  ss.toast('「'+name+'」のデータを消しました','🗑データ消去',5);
}
// シートに置いた図形ボタン用: 今開いているシートのデータを消す(EMS在庫なら色クリア＋最新化)
function アクティブシートを消す(){
  const ss=SpreadsheetApp.getActive(), name=ss.getActiveSheet().getName();
  if(name===HIKIATE_CFG.EMS在庫){ EMS在庫を更新(); return; }
  シートデータ消去_(name);
}
function 受注明細を消す(){ シートデータ消去_(HIKIATE_CFG.受注); }
function 引当待ちを消す(){ シートデータ消去_(HIKIATE_CFG.待ち); }
function 部分在庫を消す(){ シートデータ消去_(HIKIATE_CFG.部分); }
function 出荷GO未入金を消す(){ シートデータ消去_(HIKIATE_CFG.取置); }
function 希望日待ちを消す(){ シートデータ消去_(HIKIATE_CFG.希望); }
function 出荷可能を消す(){ シートデータ消去_(HIKIATE_CFG.出荷); }
function 日本在庫を消す(){ シートデータ消去_(HIKIATE_CFG.純在庫); }
function ダニエル出荷を消す(){ シートデータ消去_(HIKIATE_CFG.ダニエル出荷); }
function 商品コード引当を消す(){ シートデータ消去_(HIKIATE_CFG.コード出荷); }
function 照合レポートを消す(){ シートデータ消去_('照合レポート'); }
function 全データを消す(){
  const ui=SpreadsheetApp.getUi();
  const ans=ui.alert('全データ消去','受注明細・引当待ち・部分在庫・出荷GO未入金・希望日待ち・出荷可能・照合レポートのデータを全部消します。ええ？',ui.ButtonSet.OK_CANCEL);
  if(ans!==ui.Button.OK) return;
  [HIKIATE_CFG.受注, HIKIATE_CFG.待ち, HIKIATE_CFG.部分, HIKIATE_CFG.取置, HIKIATE_CFG.希望, HIKIATE_CFG.出荷, HIKIATE_CFG.純在庫, HIKIATE_CFG.ダニエル出荷, HIKIATE_CFG.コード出荷, '照合レポート'].forEach(n=> シートデータ消去_(n,true));
  SpreadsheetApp.getActive().toast('全シートのデータを消しました','🗑データ消去',5);
}

// 🔄 EMS在庫を更新: 残った色・罫線を全部消して、QUERY(IMPORTRANGE)を再計算させる
function EMS在庫を更新(){ 直列_(EMS在庫を更新_本体_); }
function EMS在庫を更新_本体_(){
  const ss=SpreadsheetApp.getActive(), cfg=HIKIATE_CFG;
  const emv=ss.getSheetByName(cfg.EMS在庫);
  if(!emv){ SpreadsheetApp.getUi().alert('「'+cfg.EMS在庫+'」タブが無いで'); return; }
  const ems=EMS明細_(emv); // offset=データ先頭行(1行ヘッダーがあれば2)
  const mr=emv.getMaxRows(), mc=emv.getMaxColumns();
  // ヘッダー行は触らず、データ行(offset)以降の色・罫線だけ消す
  if(mr>=ems.offset){
    const rng=emv.getRange(ems.offset,1,mr-ems.offset+1,mc);
    rng.setBackgrounds(Array.from({length:mr-ems.offset+1},()=>new Array(mc).fill(null)));
    rng.setBorder(false,false,false,false,false,false);
  }
  // QUERY数式(データ先頭セル=A{offset})を入れ直して最新化。ヘッダーがあればA2、無ければA1
  const fcell=emv.getRange(ems.offset,1);
  const f=fcell.getFormula();
  if(f){
    fcell.clearContent(); SpreadsheetApp.flush(); fcell.setFormula(f);
    // IMPORTRANGE/QUERYの再計算完了を待つ(③→④と連続クリックしても④が空/古い在庫を読まないように)
    SpreadsheetApp.flush();
    for(let i=0;i<30;i++){
      const v=String(fcell.getDisplayValue()||'');
      if(v && v.indexOf('Loading')<0) break; // '#N/A'(0件)も確定として扱う
      Utilities.sleep(1000); SpreadsheetApp.flush();
    }
  }
  ss.toast('EMS在庫を更新しました(色クリア＋最新化)','🔄EMS更新',6);
}

// ===== 📦 ダニエルEMS引当(別枠) =====
// メインのEMS在庫とは完全に別。ダニエルEMSタブの「選んだEMS番号」の在庫だけで取り寄せ未着に引き当て、
// 結果は「ダニエル出荷可能」シートと、ダニエルEMSタブの入荷後ステータス(H列)＋受注明細のピンクに書く。
// ① ダニエルEMS引当 = EMS番号を選ぶダイアログを出す → ② danielHikiateRun(選んだEMS番号) が実体
function ダニエルEMS引当(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  const dsh=ss.getSheetByName(DANIEL_CFG.シート);
  if(!dsh){ ui.alert('「'+DANIEL_CFG.シート+'」タブが無いで。先に📦ダニエルEMS取込してな'); return; }
  const dhr=ヘッダー行検索_(dsh,'BOXNo'), ds=dhr+1, dl=dsh.getLastRow();
  if(dl<ds){ ui.alert('ダニエルEMSにデータが無いで。先に📦ダニエルEMS取込してな'); return; }
  // EMS番号ごとに 発送日・件数・数量 を集計してチェックボックス一覧に
  const byEms={};
  dsh.getRange(ds,1,dl-ds+1,7).getValues().forEach(r=>{
    const ems=String(r[6]||'').trim(); if(!ems) return;
    const 発送日=String(r[1]||'').trim(), q=Number(r[4])||0;
    if(!byEms[ems]) byEms[ems]={発送日, 件数:0, 数量:0};
    byEms[ems].件数++; byEms[ems].数量+=q;
  });
  const list=Object.keys(byEms).sort((a,b)=> (byEms[a].発送日||'').localeCompare(byEms[b].発送日||'') || a.localeCompare(b));
  if(!list.length){ ui.alert('ダニエルEMSにEMS番号が無いで'); return; }
  let rowsHtml='';
  list.forEach(ems=>{ const o=byEms[ems];
    rowsHtml+='<label style="display:block;margin:6px 0;padding:5px 6px;border:1px solid #eee;border-radius:4px;cursor:pointer;">'
      +'<input type="checkbox" class="ems" value="'+ems+'" checked> '
      +'<b>'+o.発送日+'</b> &nbsp; '+ems+' <span style="color:#777">('+o.件数+'件 / 計'+o.数量+'個)</span></label>';
  });
  const html='<div style="font-family:sans-serif;font-size:13px;line-height:1.4;">'
    +'<p style="margin:0 0 8px;">引き当てる<b>EMS番号</b>にチェック。外したものは引当しません。</p>'
    +rowsHtml
    +'<div style="margin-top:6px;"><button type="button" onclick="all(true)">全部</button> '
    +'<button type="button" onclick="all(false)">全部外す</button></div>'
    +'<div style="margin-top:14px;text-align:right;"><button type="button" id="go" style="padding:7px 16px;font-weight:bold;background:#4472c4;color:#fff;border:none;border-radius:4px;cursor:pointer;" onclick="go(this)">この内容で引当</button></div>'
    +'<script>'
    +'function all(v){var cs=document.querySelectorAll(".ems");for(var i=0;i<cs.length;i++)cs[i].checked=v;}'
    +'function go(b){var a=[],cs=document.querySelectorAll(".ems:checked");for(var i=0;i<cs.length;i++)a.push(cs[i].value);'
    +'if(!a.length){alert("EMS番号を1つ以上選んでな");return;}'
    +'b.disabled=true;b.textContent="引当中...";'
    +'google.script.run.withSuccessHandler(function(){google.script.host.close();})'
    +'.withFailureHandler(function(e){b.disabled=false;b.textContent="この内容で引当";alert("エラー: "+e.message);}).danielHikiateRun(a);}'
    +'</script></div>';
  const out=HtmlService.createHtmlOutput(html).setWidth(450).setHeight(Math.min(130+list.length*42, 540));
  ui.showModalDialog(out, '📦 ダニエル引当：EMS番号を選ぶ');
}

// ダニエル引当の実体。選ばれたEMS番号(emsList)の行だけを在庫にして引き当てる
function danielHikiateRun(emsList){
  const ss=SpreadsheetApp.getActive(), cfg=HIKIATE_CFG, ui=SpreadsheetApp.getUi();
  const recv=ss.getSheetByName(cfg.受注), dsh=ss.getSheetByName(DANIEL_CFG.シート);
  if(!recv){ ui.alert('「'+cfg.受注+'」タブが無いで'); return; }
  if(!dsh){ ui.alert('「'+DANIEL_CFG.シート+'」タブが無いで'); return; }
  const sel=new Set((emsList||[]).map(s=>String(s).trim()).filter(Boolean));
  if(!sel.size){ ui.alert('EMS番号が選ばれてへん'); return; }

  // ダニエルEMSの各行(EMS番号付き)。在庫は「選んだEMS番号」の行だけ集計
  const dhr=ヘッダー行検索_(dsh,'BOXNo'), ds=dhr+1, dl=dsh.getLastRow();
  const danielRows=[], stock={};
  if(dl>=ds){
    dsh.getRange(ds,1,dl-ds+1,7).getValues().forEach((r,i)=>{
      const code=String(r[3]||'').trim(), qty=Number(r[4])||0, ems=String(r[6]||'').trim(); // D=商品コード/E=数量/G=EMS番号
      const on=sel.has(ems);
      const c=normCode_(code);
      if(code && qty>0 && on) stock[c]=(stock[c]||0)+qty; // 選んだEMS番号だけ在庫に入れる
      danielRows.push({sheetRow:ds+i, c, qty, ems, on});
    });
  }
  if(!Object.keys(stock).length){ ui.alert('選んだEMS番号に在庫(商品コード・数量)が無いで'); return; }

  // 受注明細→行(列は見出し名で特定)
  const M=列マップ_(recv), R=recv.getDataRange().getValues();
  const lines=[];
  for(let i=M.hr;i<R.length;i++){
    const row=R[i], ban=String(row[M.番号]||'').trim(), code=String(row[M.コード]||'').trim();
    if(!ban && !code) continue;
    const kbn=区分_(row[M.選択肢]);
    const 入荷=M.入荷>=0 && String(row[M.入荷]||'').trim()!=='';
    const qty=Number(row[M.個数])||0;
    lines.push({ i, ban, sortKey:番号num_(ban), 氏名:row[M.氏名], 届:row[M.届], 商品名:row[M.商品名],
      code, sku:String(row[M.SKU]||'').trim(), qty, kbn, 入荷, 入荷日値:M.入荷>=0?row[M.入荷]:'',
      paid:入金済み_(row[M.入金]), キャンセル:qty<=0, dAlloc:0, d成立:false });
  }

  // SKU突合(メインと同じ。枝番SKUは親番へ丸めない)＋ゆる照合(区切り文字を全部除去して表記ゆれを吸収)
  const loose_=v=>String(v||'').toUpperCase().replace(/[^A-Z0-9]/g,''); // -_スペース等を全部除去 例:MARI2607_B→MARI2607B
  const aliasMap={}; Object.keys(stock).forEach(c=>{
    codeKeys_(c).forEach(k=>{ if(!(k in aliasMap)) aliasMap[k]=c; });
    const lc=loose_(c); if(lc && !(lc in aliasMap)) aliasMap[lc]=c; // ダニエルの適当な区切りに対応(ゆる別名)
  });
  const cand=l=>{ const a=受注候補コード_(l.sku,l.code);
    const k=[]; a.forEach(v=>{ codeKeys_(v).forEach(x=>{ if(k.indexOf(x)<0)k.push(x); }); const lc=loose_(v); if(lc && k.indexOf(lc)<0) k.push(lc); }); return k; };
  const keyIn=l=>{ for(const k of cand(l)){ if(k in stock) return k; if(aliasMap[k] && aliasMap[k] in stock) return aliasMap[k]; } return null; };
  const findAv=l=>{ for(const k of cand(l)){ if(stock[k]>0) return k; if(aliasMap[k]&&stock[aliasMap[k]]>0) return aliasMap[k]; } return null; };

  // 着済(入荷日あり 取り寄せ)でダニエル商品のものは在庫から先に差引き(もう割当済)
  lines.filter(l=>l.kbn==='取り寄せ' && l.入荷).forEach(l=>{ const k=keyIn(l); if(k!=null) stock[k]=Math.max(0,(stock[k]||0)-l.qty); });
  // 未着の取り寄せに古い注文順でFIFO引当
  lines.filter(l=>l.kbn==='取り寄せ' && !l.入荷).sort((a,b)=> a.sortKey-b.sortKey || a.i-b.i).forEach(l=>{
    const k=findAv(l); if(k){ const got=Math.min(l.qty, stock[k]); stock[k]-=got; l.dAlloc=got; }
    l.d成立 = l.dAlloc>=l.qty && l.qty>0;
  });

  // 注文ごと(入金は注文単位)
  const byOrder={}; lines.forEach(l=> (byOrder[l.ban]=byOrder[l.ban]||[]).push(l));
  const paidOrder={}; Object.keys(byOrder).forEach(b=> paidOrder[b]=byOrder[b].some(l=>l.paid));
  const 完成_=arr=> arr.every(l=> l.キャンセル || l.kbn==='即納' || (l.kbn==='取り寄せ' && (l.入荷 || l.d成立)) );

  // ダニエルで1つでも引き当たった注文だけを「ダニエル出荷可能」へ。状態列に発送可否(発送可能/入金待ち/希望日待ち/部分)
  const HDR=['受注番号','氏名','お届け日','商品コード','商品名','個数','区分','入荷日','入金','状態'];
  const seen=new Set(), seq=[]; lines.forEach(l=>{ if(!seen.has(l.ban)){ seen.add(l.ban); seq.push(l.ban);} });
  const rows=[];
  seq.forEach(ban=>{
    const arr=byOrder[ban];
    if(!arr.some(l=>l.d成立)) return; // ダニエルで1つも引き当たってない注文は出さない
    const paid=paidOrder[ban], comp=完成_(arr);
    let 可否; if(!comp) 可否='部分'; else if(希望日未来_(arr[0].届)) 可否='希望日待ち'; else 可否= paid? '発送可能' : '入金待ち';
    arr.forEach(l=>{
      let st, color;
      if(l.キャンセル){ st='キャンセル'; color=cfg.色_グレー; }
      else if(l.kbn==='即納'){ st='即納'; color=cfg.色_水; }
      else if(l.kbn==='指定なし'){ st='要確認'; color=cfg.色_橙; }
      else if(l.入荷){ st='着済'; color=cfg.色_着; }
      else if(l.d成立){ st='ダニエル引当'; color=cfg.色_黄; } // ダニエルEMSが当たった=今回出せる
      else { st='在庫待ち'; color=null; }                      // ダニエルにも無い=白
      rows.push({ vals:[ban,l.氏名,l.届,l.code,l.商品名,l.qty,l.kbn,l.入荷日値,paid?'済':'未', st+'／'+可否], color, ban, paid, qty:l.qty });
    });
  });
  書き出し_(ss, cfg.ダニエル出荷, HDR, rows, M.hr);

  // ダニエルEMSタブの入荷後ステータス(H=8列)に、その箱の品が当たった受注番号を書く
  const cons={};
  lines.filter(l=>l.kbn==='取り寄せ' && l.入荷).forEach(l=>{ const k=keyIn(l); if(k!=null)(cons[k]=cons[k]||[]).push({qty:l.qty, ban:l.ban, kind:'割当済'}); });
  lines.filter(l=>l.kbn==='取り寄せ' && !l.入荷 && l.dAlloc>0).sort((a,b)=> a.sortKey-b.sortKey || a.i-b.i).forEach(l=>{ const k=keyIn(l); if(k!=null)(cons[k]=cons[k]||[]).push({qty:l.dAlloc, ban:l.ban, kind:'引当'}); });
  const ptr={};
  const consume=(code,need)=>{ const q=cons[code]||[], st=ptr[code]||{i:0,used:0}, got={割当済:[],引当:[]}; let left=need;
    while(left>0 && st.i<q.length){ const e=q[st.i], av=e.qty-st.used, t=Math.min(left,av);
      if(t>0){ if(got[e.kind].indexOf(e.ban)<0) got[e.kind].push(e.ban); st.used+=t; left-=t; }
      if(st.used>=e.qty){ st.i++; st.used=0; } }
    ptr[code]=st; return {got, surplus:left}; };
  const dcolor=[];
  const dstat=danielRows.map(d=>{
    if(!d.on){ dcolor.push(cfg.色_グレー); return ['対象外']; } // 選ばなかったEMS番号=対象外・グレー
    if(d.qty<=0){ dcolor.push(null); return ['']; }
    const rr=consume(d.c, d.qty), pp=[];
    const 使用 = rr.got.割当済.length || rr.got.引当.length;
    if(rr.got.割当済.length) pp.push('割当済:'+rr.got.割当済.join(','));
    if(rr.got.引当.length)   pp.push('引当:'+rr.got.引当.join(','));
    if(rr.surplus>0)         pp.push('余り'+rr.surplus);
    // 箱の色: 全部引き当て(余り0)=ピンク / 一部引当(余りあり)=緑 / 未使用(全部余り)=白
    dcolor.push(使用 ? (rr.surplus<=0 ? cfg.色_ダ引当 : cfg.色_緑) : null);
    return [pp.join(' / ') || '余り'];
  });
  if(danielRows.length){
    const sr=danielRows[0].sheetRow;
    dsh.getRange(sr, 8, dstat.length, 1).setValues(dstat);
    dsh.getRange(sr, 1, dcolor.length, 8).setBackgrounds(dcolor.map(c=> new Array(8).fill(c))); // ダニエルEMSタブも消込色付け
  }

  // 受注明細にダニエルの色を重ね塗り(メインの色は消さず、ダニエルが当てた取り寄せ行だけピンク。ダニエル色だけ毎回クリア)
  const recvStart=M.hr+1, recvLast=recv.getLastRow(), nc=recv.getLastColumn();
  if(recvLast>=recvStart){
    const rng=recv.getRange(recvStart,1,recvLast-recvStart+1,nc), bgs=rng.getBackgrounds();
    const dCols=[cfg.色_ダ引当.toLowerCase(), cfg.色_ダ着.toLowerCase()];
    const dHit={}; lines.forEach(l=>{ if(l.d成立) dHit[l.i]={paid:paidOrder[l.ban]}; });
    for(let r=0;r<bgs.length;r++){
      const idx=recvStart+r-1, hit=dHit[idx];
      if(hit){
        for(let c=0;c<nc;c++) bgs[r][c]=cfg.色_ダ引当;
        if(!hit.paid && M.番号>=0) bgs[r][M.番号]=cfg.色_ダ着; // ダニエルの未入金マーク(濃いめピンク)
      } else { // ダニエルが当ててない行=前回のダニエル色だけ白に戻す(メイン色には触れない)
        for(let c=0;c<nc;c++){ if(dCols.indexOf(String(bgs[r][c]||'').toLowerCase())>=0) bgs[r][c]=null; }
      }
    }
    rng.setBackgrounds(bgs);
  }

  const 注文数=new Set(rows.map(r=>r.ban)).size;
  抽出フラグ更新_(); // 受注明細の「抽出」列(ダニエルのピンクも色ありに反映)
  ss.toast('ダニエル引当 完了：EMS番号'+sel.size+'件で '+注文数+'注文。選んでないEMS番号は「対象外」表示', '📦 ダニエル', 6);
}

// ===== 🔢 商品コードで引当(別枠・手動入力の在庫を使う) =====
// 「商品コード入力」シートに手で貼り付けた 商品コード＋数量 を在庫にして、受注明細の商品コード(H列)に突き合わせる。
// SKUの枝番はまとめない(001A/001B/001Cは別物・各行個別)。結果は「商品コード引当」シートだけに出す。
function 商品コード引当(){
  const ss=SpreadsheetApp.getActive(), cfg=HIKIATE_CFG, ui=SpreadsheetApp.getUi();
  const recv=ss.getSheetByName(cfg.受注);
  if(!recv){ ui.alert('「'+cfg.受注+'」タブが無いで'); return; }

  // 在庫=「商品コード入力」シート(A=商品コード / B=数量)。無ければ作って案内
  let insh=ss.getSheetByName(cfg.コード入力);
  if(!insh){
    insh=ss.insertSheet(cfg.コード入力);
    insh.getRange(1,1,1,2).setValues([['商品コード','数量']])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground('#4472c4').setHorizontalAlignment('center');
    insh.setFrozenRows(1); insh.setColumnWidth(1,200); insh.setColumnWidth(2,80);
    insh.getRange(2,1).setValue('ここに商品コードと数量を貼り付け→もう一度ボタン');
    ui.alert('「'+cfg.コード入力+'」シートを作ったで。\nA列=商品コード / B列=数量 を貼り付けて、もう一度「🔢 商品コードで引当」を押してな');
    return;
  }
  const stock={};
  const han_=v=>String(v==null?'':v).replace(/[０-９．]/g,d=>String.fromCharCode(d.charCodeAt(0)-0xFEE0)); // 全角数字→半角
  const num_=v=>{ const n=Number(han_(v).replace(/[^0-9.\-]/g,'')); return isNaN(n)?0:n; };       // カンマ等が混ざっても数量を取る
  const iv=insh.getDataRange().getValues();
  for(let i=1;i<iv.length;i++){ const code=String(iv[i][0]||'').trim(), q=num_(iv[i][1]);
    if(code && q>0){ const c=normCode_(code); stock[c]=(stock[c]||0)+q; } }
  if(!Object.keys(stock).length){ ui.alert('「'+cfg.コード入力+'」に商品コードと数量を入れてな\n(A列=商品コード / B列=数量。数量は半角の数字で)'); return; }

  // 受注明細→行
  const M=列マップ_(recv), R=recv.getDataRange().getValues();
  const lines=[];
  for(let i=M.hr;i<R.length;i++){
    const row=R[i], ban=String(row[M.番号]||'').trim(), code=String(row[M.コード]||'').trim();
    if(!ban && !code) continue;
    const kbn=区分_(row[M.選択肢]);
    const 入荷=M.入荷>=0 && String(row[M.入荷]||'').trim()!=='';
    const qty=Number(row[M.個数])||0;
    lines.push({ i, ban, sortKey:番号num_(ban), 氏名:row[M.氏名], 届:row[M.届], 商品名:row[M.商品名],
      code, qty, kbn, 入荷, 入荷日値:M.入荷>=0?row[M.入荷]:'',
      paid:入金済み_(row[M.入金]), キャンセル:qty<=0, cAlloc:0, c成立:false });
  }

  // 商品コードで突合(枝番のサフィックスは除去しない。normCode_で_→-だけ正規化)
  const aliasMap={}; Object.keys(stock).forEach(c=> codeKeys_(c).forEach(k=>{ if(!(k in aliasMap)) aliasMap[k]=c; }));
  const cand=l=>{ const k=[]; if(l.code) codeKeys_(l.code).forEach(x=>{ if(k.indexOf(x)<0)k.push(x); }); return k; };
  const keyIn=l=>{ for(const k of cand(l)){ if(k in stock) return k; if(aliasMap[k]&&aliasMap[k] in stock) return aliasMap[k]; } return null; };
  const findAv=l=>{ for(const k of cand(l)){ if(stock[k]>0) return k; if(aliasMap[k]&&stock[aliasMap[k]]>0) return aliasMap[k]; } return null; };

  // 手動入力の在庫はそのまま全部、未着の取り寄せに古い注文順でFIFO引当(着済=入荷日ありは満たし済とみなし差引かない)
  const inputStock=Object.assign({},stock); // 在庫サマリ(投入数)用に控える
  // 未着の取り寄せに古い注文順でFIFO引当(各行の商品コードごと)
  lines.filter(l=>l.kbn==='取り寄せ' && !l.入荷).sort((a,b)=> a.sortKey-b.sortKey || a.i-b.i).forEach(l=>{
    const k=findAv(l); if(k){ const got=Math.min(l.qty, stock[k]); stock[k]-=got; l.cAlloc=got; }
    l.c成立 = l.cAlloc>=l.qty && l.qty>0;
  });

  // 注文ごと
  const byOrder={}; lines.forEach(l=> (byOrder[l.ban]=byOrder[l.ban]||[]).push(l));
  const paidOrder={}; Object.keys(byOrder).forEach(b=> paidOrder[b]=byOrder[b].some(l=>l.paid));
  const 完成_=arr=> arr.every(l=> l.キャンセル || l.kbn==='即納' || (l.kbn==='取り寄せ' && (l.入荷 || l.c成立)) );

  // 受注明細の色・書式をそのまま結果に映すため、全列の背景を読む(代表行のセル色をそのまま使う)
  const recvLast2=recv.getLastRow();
  const recvBg= recvLast2>M.hr ? recv.getRange(M.hr+1,1,recvLast2-M.hr,recv.getLastColumn()).getBackgrounds() : [];

  // 引き当たった注文を「商品コード引当」シートへ。★注文単位(1注文=1行)で表示
  const HDR=['受注番号','氏名','お届け日','商品コード','商品名','個数','区分','入荷日','入金','状態'];
  const seen=new Set(), seq=[]; lines.forEach(l=>{ if(!seen.has(l.ban)){ seen.add(l.ban); seq.push(l.ban);} });
  const rows=[];
  seq.forEach(ban=>{
    const arr=byOrder[ban];
    const torai=arr.filter(l=>l.kbn==='取り寄せ' && !l.キャンセル);
    if(!torai.some(l=>l.c成立)) return; // 入力した商品コードで1つも引き当たってない注文は出さない
    const paid=paidOrder[ban], comp=完成_(arr);
    const 引当数=torai.filter(l=>l.c成立).length, 総数=torai.length;
    let 状態, statusColor;
    if(!comp){ 状態='部分('+引当数+'/'+総数+')'; statusColor=cfg.色_グレー; }
    else if(希望日未来_(arr[0].届)){ 状態='希望日待ち'; statusColor=cfg.色_水; }
    else if(paid){ 状態='発送可能'; statusColor=cfg.色_緑; }
    else { 状態='入金待ち'; statusColor=cfg.色_黄; }
    const repIdx=(torai.find(l=>l.c成立)||arr[0]).i - M.hr; // 代表行(引き当てた取り寄せ行)の受注明細での位置=色を映す元
    // 注文の中身をまとめる(商品コードは結合、個数は合計、区分は混在なら「混在」)
    const codes=[...new Set(arr.filter(l=>!l.キャンセル).map(l=>l.code).filter(Boolean))].join(', ');
    const 商品名= arr.length>1 ? (String(arr[0].商品名||'')+' 他'+(arr.length-1)+'件') : (arr[0].商品名||'');
    const 個数= arr.reduce((s,l)=> s+(l.キャンセル?0:l.qty), 0);
    const kbns=[...new Set(arr.filter(l=>!l.キャンセル).map(l=>l.kbn))];
    const 区分= kbns.length>1 ? '混在' : (kbns[0]||'');
    rows.push({ vals:[ban, arr[0].氏名, arr[0].届, codes, 商品名, 個数, 区分, '', paid?'済':'未', 状態], color:statusColor, ban, paid, qty:個数, repIdx, statusColor });
  });
  書き出し_(ss, cfg.コード出荷, HDR, rows, M.hr);
  const outSh=ss.getSheetByName(cfg.コード出荷);

  // ★受注明細の色・書式をそのまま映す: 各結果行の背景を代表行(受注明細)のセル色で上書き(状態列だけはステータス色)
  if(rows.length && outSh){
    const colMap=[M.番号,M.氏名,M.届,M.コード,M.商品名,M.個数,M.選択肢,M.入荷,M.入金,-1]; // 結果10列→受注明細の対応列
    const bgGrid=rows.map(r=>{ const rb=recvBg[r.repIdx];
      return colMap.map((src,ci)=> ci===colMap.length-1 ? (r.statusColor||null) : (src>=0 && rb ? rb[src] : null)); });
    outSh.getRange(M.hr+1,1,bgGrid.length,colMap.length).setBackgrounds(bgGrid);
  }

  // 在庫サマリはA2へ(A1は時刻だけにして見やすく)。「在庫不足で出せない注文＝あと何個必要か」をはっきり書く
  let 不足数=0; const 不足注文=new Set();
  lines.filter(l=>l.kbn==='取り寄せ' && !l.入荷 && !l.c成立).forEach(l=>{ if(keyIn(l)!=null){ 不足注文.add(l.ban); 不足数+=Math.max(0,l.qty-l.cAlloc); } });
  const codeSum=Object.keys(inputStock).map(c=> c+'：投入'+inputStock[c]+'／使用'+(inputStock[c]-(stock[c]||0))+'／残'+(stock[c]||0)).join('　');
  const 不足番号=[...不足注文].sort((a,b)=> 番号num_(a)-番号num_(b)); // 古い順に受注番号を並べる
  const 不足文=不足番号.length? '在庫不足で出せない注文 '+不足番号.length+'件(あと'+不足数+'個あれば出せる)：'+不足番号.join('、') : '在庫不足なし(対象は全部出せた)';
  const summaryStr='引当'+rows.length+'注文　｜　在庫 '+codeSum+'　｜　'+不足文;
  if(outSh) outSh.getRange(2,1).setValue(summaryStr);

  ss.toast('商品コード引当 完了：'+summaryStr, '🔢 商品コード', 8);
}

// 出力シートの中身(値・背景色・罫線)だけを消す。列幅・行高・文字揃え・フォント等の書式は保つ
// → 一度手で整えたレイアウトが、引き当てのたびにリセットされないように
function シート値クリア_(sh){
  const mr=sh.getMaxRows(), mc=sh.getMaxColumns();
  const r=sh.getRange(1,1,mr,mc);
  r.clearContent();
  r.setBackgrounds(Array.from({length:mr},()=>new Array(mc).fill(null)));
  r.setBorder(false,false,false,false,false,false);
}

function 書き出し_(ss, name, hdr, rows, startRow){
  const hr=startRow||1, dr=hr+1; // hr=見出し行 / dr=データ開始行
  let sh=ss.getSheetByName(name), 新規=!sh; if(!sh) sh=ss.insertSheet(name);
  シート値クリア_(sh); // 値・色・罫線だけ入れ替え(列幅・行高・中央揃えは保持)
  // 押すたびに変わる印(最終引当の時刻＋行数)をA1に表示
  sh.getRange(1,1).setValue('最終引当: '+Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm:ss')+' / '+rows.length+'行');
  const ncol=hdr.length;
  sh.getRange(hr,1,1,ncol).setValues([hdr])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#4472c4').setHorizontalAlignment('center');
  sh.setFrozenRows(hr);
  sh.getRange(hr,1,1,ncol).setFontSize(HIKIATE_CFG.字);
  if(新規) sh.setRowHeight(hr, HIKIATE_CFG.行高);
  if(rows.length===0){ sh.getRange(dr,1).setValue('(対象なし)'); return; }

  const rng=sh.getRange(dr,1,rows.length,ncol);
  rng.setValues(rows.map(r=>r.vals));
  rng.setFontSize(HIKIATE_CFG.字);
  if(新規) sh.setRowHeights(dr, rows.length, HIKIATE_CFG.行高);
  rng.setBackgrounds(rows.map(r=>{ const a=new Array(ncol).fill(r.color);
    if(r.paid===false) a[0]=HIKIATE_CFG.色_赤;          // 未入金は受注番号セル赤
    if(r.qty>=2 && ncol>5) a[5]=HIKIATE_CFG.色_緑;      // 個数2以上は個数セル緑(見間違い防止)
    return a; }));
  rng.setBorder(true,true,true,true,true,true,'#cccccc',SpreadsheetApp.BorderStyle.SOLID);

  for(let i=0;i<rows.length;i++){
    if(i===rows.length-1 || rows[i+1].ban!==rows[i].ban)
      sh.getRange(dr+i,1,1,ncol).setBorder(null,null,true,null,null,null,'#000000',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  if(新規) [150,110,100,150,360,55,80,70,55,95,150,120].forEach((w,c)=> { if(c<ncol) sh.setColumnWidth(c+1,w); });
}

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
    .addSeparator()
    .addItem('📋 取り置き登録を更新', '取り置き初期登録を作成')
    .addItem('✅ 取り置き登録を反映(通常＋棚戻し)', '取り置き登録を反映')
    .addSubMenu(ui.createMenu('👀 取り置き登録の表示モード')
      .addItem('要作業(既定)', '取り置き表示_要作業')
      .addItem('部分在庫', '取り置き表示_部分在庫')
      .addItem('希望日待ち・現物あり', '取り置き表示_希望日現物')
      .addItem('先行引当', '取り置き表示_先行引当')
      .addItem('すべて', '取り置き表示_すべて'))
    .addItem('🧱 取り置き登録の罫線を引き直す', '取り置き登録の罫線を引く')
    .addSubMenu(ui.createMenu('🔧 取り置き内部管理(通常操作不要)')
      .addItem('キャンセル戻し確認を更新', 'キャンセル戻し確認を更新')
      .addItem('キャンセル戻し確認を確定', 'キャンセル戻し確認を確定')
      .addItem('Yahoo戻しを反映済みにする', 'キャンセル戻しをYahoo反映済みにする')
      .addItem('選択した取り置きを手動解除', '選択した取り置きを手動解除')
      .addItem('🧹 孤児取り置きをまとめて解除', '孤児取り置きをまとめて解除'))
    .addSeparator()
    .addItem('🔄 現物確認移行を作成(一度きりの仕分け)', '現物確認移行を作成')
    .addItem('✅ 現物確認移行を反映', '現物確認移行を反映')
    .addItem('🧹 バックアップシートを整理(最新2世代だけ残す)', 'バックアップシートを整理')
    .addSeparator()
    .addItem('🔍 引当切替差分を作成(移行プレビュー)', '引当切替差分を作成')
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
    .addItem('② 引き当て実行(実EMSリストのみ)', '引当実行')
    .addItem('🔁 更新してから引当(EMS更新→即納→引当を一括)', '更新してから引当')
    .addItem('📦 到着済を在庫反映済みへ(便の締め・📤出力もここから)', '到着済を在庫反映済みへ')
    .addItem('📤 Yahoo在庫変更を出力(出力だけやり直す時用)', 'Yahoo在庫変更を出力')
    .addItem('⚖️ この便の引当をやり直す(到着日指定・複数可)', '便の引当をやり直す')
    .addItem('🔎 引当診断(受注番号で調べる)', '引当診断')
    .addItem('🔎 商品診断(商品コードで調べる)', '商品診断')
    .addItem('🔎 入荷日の整合チェック(誤記入の検出)', '入荷日整合チェック')
    .addItem('🧹 チェック一覧の入荷日をクリア', '入荷日チェック_一覧をクリア')
    .addItem('🧹 旧棚卸割当だけを解除（再引当前）', '旧棚卸割当だけを解除して再引当')
    .addItem('🧮 全件検算レポート(EMS×台帳×受注×Yahooの突き合わせ)', '全件検算レポート')
    .addItem('🧪 全件再計算プレビュー(事実データから)', '全件再計算プレビュー')
    .addItem('✅ 全件再計算を反映(要プレビュー)', '全件再計算を反映')
    .addItem('🔓 全件再計算ガードを解除(復旧用)', '全件再計算ガードを解除')
    .addItem('🧹 全件再計算のブロックSKUをクリア', '全件再計算ブロックSKUをクリア')
    .addItem('🔍 在庫照合レポート', '在庫照合レポート')
    .addToUi();

  // ── ダニエル引当(サブ。在庫=ダニエルEMS) ──
  ui.createMenu('🚚 ダニエル引当(サブ)')
    .addItem('📦 ダニエルEMS取込', '取込_ダニエルEMS')
    .addItem('📦 ダニエルEMS引当', 'ダニエルEMS引当')
    .addItem('🧮 ダニエル余りを計算(便の余り推定)', 'ダニエル余りを計算')
    .addItem('🧾 ダニエル入荷記録を現タブから初期化(初回用)', 'ダニエル入荷記録_現タブから初期化')
    .addItem('📝 受注明細に取り置きメモ列を用意(初回用)', '受注明細_取り置きメモ列を用意')
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
  // 便の入荷を積み上げ台帳へ(EMS番号単位で洗い替え)→余り推定を更新(失敗しても取込は成立)
  try{ ダニエル入荷記録へ追記_(rows); }catch(e){}
  try{ ダニエル余りを計算(); }catch(e){}
  ss.toast('ダニエルEMS取込完了：'+rows.length+'件 / 発送日'+dates.join(',')+'（'+latest.getName()+'）','📦ダニエルEMS',6);
}

// 実在するEMS番号だけを供給として認める。
// 「棚卸...」は待ち需要から作った机上数量で、実物のEMS入荷ではないため必ず除外する。
// EMS番号空欄も出どころを追えないので、引当には使わない。
function 実EMS番号_(v){
  const s=String(v==null?'':v).trim();
  return !!s && !/^棚卸/i.test(s);
}

function 実EMS番号配列_(v){
  return String(v==null?'':v).split(/[\s,、\/／\.．]+/)
    .map(s=>s.trim()).filter(s=>実EMS番号_(s));
}

// 受注明細のEMS番号を書き戻す時、在庫反映済みなど過去の実EMS割当は保持する。
// 未割当行と旧棚卸番号は消し、今回の実EMSが決まった行だけ新しい番号へ置き換える。
function EMS番号書戻し値_(現在値, line){
  if(line && Array.isArray(line.入荷EMS候補)){
    const cand=new Set(line.入荷EMS候補.map(v=>String(v||'').trim()).filter(v=>実EMS番号_(v)));
    const box=実EMS番号配列_(line.箱EMS), current=実EMS番号配列_(現在値);
    if(box.length && box.every(v=>cand.has(v))) return box.join(', ');
    if(current.length && current.every(v=>cand.has(v))) return current.join(', ');
    if(line.入荷EMS候補.length===1) return String(line.入荷EMS候補[0]).trim();
    if(line.入荷EMS候補.length>1) return ''; // 複数便で特定不能なら誤った番号を残さない
  }
  if(line && 実EMS番号_(line.箱EMS)) return String(line.箱EMS).trim();
  if(line && line.入荷 && 実EMS番号_(現在値)) return String(現在値).trim();
  return '';
}

// EMS在庫の生データを返す(1行ヘッダーがあれば除外)。offset=データ先頭のシート行番号(1始)
// 行位置は色付け用に維持しつつ、実EMSでない行はコード空欄・数量0にして全引当処理から無効化する。
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
  let 除外=0;
  const rows=all.slice(s).map(row=>{
    if(実EMS番号_(row[cols.EMS番号])) return row;
    const copy=row.slice();
    if(cols.コード>=0) copy[cols.コード]='';
    if(cols.数量>=0) copy[cols.数量]=0;
    if(String(row[cols.コード]||'').trim() || (Number(row[cols.数量])||0)) 除外++;
    return copy;
  });
  return { rows, offset: s+1, cols, 除外 };
}

function EMS供給オブジェクト_(rows, cols, arrivalFn){
  return (rows||[]).map(row=>{
    const ems=String(row[cols.EMS番号]||'').trim(), raw=String(row[cols.コード]||'').trim(), qty=Number(row[cols.数量])||0;
    const arrival=ymd_(typeof arrivalFn==='function'?arrivalFn(row):'');
    if(!実EMS番号_(ems) || !raw || qty<=0 || !/^20\d{2}-\d{2}-\d{2}$/.test(arrival)) return null;
    return {ems,code:normCode_(raw),sourceCode:raw,qty,arrival,directBan:注文番号在庫コード_(raw)||タグ受注番号_(raw)};
  }).filter(Boolean);
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
    支払: find('お支払い方法','支払方法','支払い方法','決済方法'),
    入荷: find('入荷日'),
    出荷日: find('出荷日'),
    出荷日毎: find('出荷日(複数時には配送先毎)','出荷日（複数時には配送先毎）'),
    EMS: find('EMS番号'),
    入荷予定: find('入荷予定'),
    メモ: find('取り置きメモ'),
    旧取置数: find('取り置き数','既存取り置き数'),
    取置数: -1 // 旧列は表示用に残すが、移行後の引当計算では読まない
  };
}

// 分割出荷の行判定(2026-07-23): 出荷日(U)か出荷日(複数時には配送先毎)(W)に値がある行は、
// その分割がもう出荷済み=需要から外す。GoQのステータスは注文単位で分割ごとに変えられないため、
// 「持ち戻り/問合せ中」等でも行単位の出荷日を事実として使う(10117477の分割出荷で導入)。
function 引当_行出荷済み_(row, M){
  const u=(M&&M.出荷日>=0)?String(row[M.出荷日]==null?'':row[M.出荷日]).trim():'';
  const w=(M&&M.出荷日毎>=0)?String(row[M.出荷日毎]==null?'':row[M.出荷日毎]).trim():'';
  return u!=='' || w!=='';
}

// 台湾/中国(別ルート): 台帳(この画面の棚登録)と受注明細入荷日の二重計上を防ぐ。
// 台帳で確保した分だけ入荷日ベースの別ルート済数量を目減りさせる(実効確保=両者の大きい方)
function 別ルート二重控除_(lines, activeByKey){
  (lines||[]).forEach(l=>{
    if(!l || !l.別ルート) return;
    const held=Number(activeByKey&&activeByKey[取り置き_行キー_(l)])||0;
    if(held>0) l.別ルート済数量=Math.max(0,(Number(l.別ルート済数量)||0)-held);
  });
  return lines;
}

function 残必要計算_(l){
  return Math.max(0,(Number(l&&l.qty)||0)-(Number(l&&l.取り置き中数量)||0)-(Number(l&&l.alloc)||0)-(Number(l&&l.別ルート済数量)||0));
}

// 台湾・中国ルート(韓国EMSに照合先が無い)。手入力の入荷日が正で、取り置き台帳の対象外。
function 引当_別ルート判定_(選択肢, 商品名){
  return /台湾|中国/.test(String(選択肢||'')) || /台湾|中国/.test(String(商品名||''));
}

// 受注明細の末尾に「取り置きメモ」列を用意する(あれば何もしない)。
// GoQ取込は取り込みのたびにこの列を退避→貼り直しするので、書いたメモは消えない。
// エディタからも実行できるようUIは使わない。
function 受注明細_取り置きメモ列を用意(){
  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName(HIKIATE_CFG.受注);
  if(!sh){ Logger.log('受注明細タブがありません'); return; }
  const hr=受注ヘッダー行_(sh);
  const made=[];
  ['取り置き数','取り置きメモ'].forEach(name=>{
    const head=sh.getRange(hr,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0].map(v=>String(v||'').trim());
    if(head.indexOf(name)>=0) return;
    const col=sh.getLastColumn()+1;
    sh.getRange(hr,col).setValue(name).setFontWeight('bold');
    made.push(name+'('+col+'列目)');
  });
  const msg = made.length? '作成: '+made.join(' / ') : '取り置き数・取り置きメモ列は既にあります';
  Logger.log(msg);
  ss.toast(msg);
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
  const 開始ms=Date.now();
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

  // 受注明細へ書く前に、全ステータスCSVと現在台帳から遷移を検証する。
  // 行の消滅だけでは発送済みにせず、取り置き要確認へ残す。
  let 取り置き遷移;
  try{
    取り置き遷移=取り置き_CSV遷移計画_(CSV行を受注行オブジェクトへ_(data[0],data.slice(1)),取り置き台帳_読む_(),new Date());
  }catch(e){
    ui.alert('取り置き台帳の更新を中止しました',e.message,ui.ButtonSet.OK);
    return;
  }
  if(取り置き遷移.errors.length){
    ui.alert('取り置き台帳の更新を中止しました',取り置き遷移.errors.join('\n'),ui.ButtonSet.OK);
    return;
  }

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
  // CSVで列がズレないよう、取込前に既存の「EMS番号」「確保済み」「不足」「確保内訳」列
  // (前回②/取り置き登録が個数の隣に作った分)を一旦削除する。取込の最後に個数の隣へ作り直す。
  { ['EMS番号','確保済み','不足','確保内訳'].forEach(name=>{
      const h0=sh.getRange(hdrRow,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
      const e=h0.indexOf(name); if(e>=0) sh.deleteColumn(e+1); }); }
  // 【入荷日を引き継ぐ】上書き前に、今の入荷日を 受注番号+商品コード+SKU をキーに退避(取込で消えないように)
  // 【取り置きメモも引き継ぐ】同じキーで退避→取込後に貼り直す(手書きメモが取込で消えないように)
  const 退避入荷={}; const 退避メモ={}; let メモ列あり=false;
  { const M0=列マップ_(sh);
    メモ列あり = M0.メモ>=0;
    if((M0.入荷>=0 || M0.メモ>=0) && sh.getLastRow()>=startRow){
      const old=sh.getRange(startRow,1,sh.getLastRow()-startRow+1,sh.getLastColumn()).getValues();
      old.forEach(r=>{ const ban=String(r[M0.番号]||'').trim(), code=String(r[M0.コード]||'').trim();
        const sku=M0.SKU>=0?String(r[M0.SKU]||'').trim():''; const key=ban+'|'+code+'|'+sku;
        if(M0.入荷>=0){ const v=r[M0.入荷]; if(ban && v!=='' && v!=null) 退避入荷[key]=v; }
        if(M0.メモ>=0){ const m=String(r[M0.メモ]==null?'':r[M0.メモ]).trim(); if(ban && m!=='') 退避メモ[key]=m; } });
    } }
  // ヘッダーより下を全リセット(値・背景色・罫線を消す)→ 前回の引当色や罫線が残らないように
  // ※ヘッダー行(例6行目)と上は触らないので書式は維持される
  // ヘッダーより下をリセット: 値・背景色・罫線だけ消す(フォント/配置は残す)
  const oldLastRow=sh.getLastRow(), oldLastCol=sh.getLastColumn();
  if(oldLastRow>=startRow){
    const nr=oldLastRow-startRow+1;
    const reset=sh.getRange(startRow,1,nr,Math.max(1,oldLastCol));
    reset.clearContent();
    reset.setBackground(null);                                  // 前回使用範囲の背景色だけ消す
    reset.setBorder(false,false,false,false,false,false);         // 前回使用範囲の罫線だけ消す
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
  取り置き_CSV遷移を反映_(取り置き遷移);

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
  受注明細_確保列を用意_(sh); // 個数の隣に確保済み/不足/確保内訳を作り直す(中身は取り置き登録の更新が書く)

  // 【取り置きメモを引き継ぐ】列を末尾に確保し直して、同じキーの行へ貼り直す
  let メモ引継=0;
  if(body.length && (メモ列あり || Object.keys(退避メモ).length)){
    const M2=列マップ_(sh);
    let メモC=M2.メモ;
    if(メモC<0){ メモC=sh.getLastColumn(); sh.getRange(hdrRow,メモC+1).setValue('取り置きメモ').setFontWeight('bold'); }
    if(Object.keys(退避メモ).length){
      const cur2=sh.getRange(startRow,1,body.length,sh.getLastColumn()).getValues();
      const out2=cur2.map(r=>{ const ban=String(r[M2.番号]||'').trim(), code=String(r[M2.コード]||'').trim();
        const sku=M2.SKU>=0?String(r[M2.SKU]||'').trim():''; const m=退避メモ[ban+'|'+code+'|'+sku];
        if(m!==undefined) メモ引継++; return [ m!==undefined ? m : '' ]; });
      sh.getRange(startRow, メモC+1, out2.length, 1).setValues(out2);
    }
  }

  // 消込台帳を更新: 前回いて今回消えた注文=発送済み(他の人が発送した分)を検知
  const 台帳=消込台帳更新_();

  // ダニエル便の余り推定を更新(出荷が台帳に載った分を自動で差し引く。記録が空なら何もしない)
  try{ ダニエル余りを計算(true); }catch(e){}

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
    try{ キャンセル結果=キャンセル処理_(キャンセル番号,{取り置き台帳を更新:false}); }catch(e){}
    // 今回CSVに載っていた番号だけ覚える(CSVの期間から外れた古い番号は自然に忘れる→肥大化しない)
    try{ props.setProperty('処理済キャンセル番号', JSON.stringify(キャンセル番号)); }catch(e){}
  }

  // 日常画面をCSVの最新状態へ追従させる。数量0/注文キャンセルで確保があった行は、
  // ここで「取り置き登録」の赤い棚戻し待ちへ自動表示される。
  let 取り置き登録更新=null, 取り置き登録更新エラー='';
  try{ 取り置き登録更新=取り置き初期登録を作成本体_({silent:true}); }
  catch(e){ 取り置き登録更新エラー=String(e&&e.message||e); console.error('取り置き登録の自動更新失敗: '+取り置き登録更新エラー); }

  const 処理ms=Date.now()-開始ms, 処理秒=Math.round(処理ms/1000);
  console.log('①受注明細更新 処理時間ms='+処理ms);
  ss.toast('取込完了：'+latest.getName()+' / 受注'+body.length+'行（更新 '+upd+' / 入荷日引継 '+引継+'件'
    +(メモ引継? ' / メモ引継 '+メモ引継+'件':'')
    +(台帳.新規出荷済? ' / 🧾消えた注文の出荷済み検知'+台帳.新規出荷済+'件':'')
    +(処理済結果&&処理済結果.追加? ' / 📦処理済を確定登録'+処理済結果.追加+'件':'')
    +(キャンセル番号.length? ' / 🚫キャンセル'+キャンセル番号.length+'件(新規'+新規キャンセル.length+'件)':'')
    +(取り置き登録更新? ' / 🔴棚戻し待ち'+取り置き登録更新.棚戻し待ち+'件':'')
    +(取り置き登録更新エラー? ' / ⚠取り置き登録更新失敗':'')
    +' / 処理 '+処理秒+'秒）','GoQ取込',6);
  if(新規キャンセル.length){
    SpreadsheetApp.getUi().alert('🚫 キャンセルを自動処理しました',
      '新しくキャンセルになった注文 '+新規キャンセル.length+'件を、受注明細に入れず後始末しました:\n'+
      新規キャンセル.slice(0,20).join(', ')+(新規キャンセル.length>20?' …他'+(新規キャンセル.length-20)+'件':'')+
      (キャンセル結果? '\n\n'+キャンセル結果.results.join('\n'):'')+
      '\n\n台帳確保がある商品は「取り置き登録」に赤い棚戻し待ちで出ます。棚を確認して右端の処理を選び、「取り置き登録を反映」を押してください。',
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
  if(取り置き登録更新エラー){
    SpreadsheetApp.getUi().alert('⚠️ 取り置き登録の自動更新だけ失敗しました',
      'CSV取込と台帳遷移は完了しています。メニューの「取り置き登録を更新」を押してください。\n\n'+取り置き登録更新エラー,
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function 仕分け証跡キー_(header,row){
  const H=(header||[]).map(v=>String(v||'').trim());
  const find=(...names)=>{
    for(const name of names){
      const index=H.indexOf(name);
      if(index>=0) return index;
    }
    return -1;
  };
  const indexes=[
    find('受注番号'),
    find('商品コード'),
    find('商品SKU','SKU'),
    find('受注ステータス')
  ];
  if(indexes.some(index=>index<0)) return '';
  return indexes.map((index,pos)=>{
    const value=String((row||[])[index]||'').trim();
    return pos===0?value.replace(/^niyantarose-/i,''):value;
  }).join('|');
}

function 仕分け証跡重複除外_(header,rows,existingKeys){
  const seen=new Set(existingKeys||[]), out=[];
  (rows||[]).forEach(row=>{
    const key=仕分け証跡キー_(header,row);
    if(!key){ out.push(row); return; }
    if(seen.has(key)) return;
    seen.add(key); out.push(row);
  });
  return out;
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
  const H=header.map(v=>String(v||'').trim());
  const find=(...names)=>{
    for(const name of names){ const index=H.indexOf(name); if(index>=0) return index; }
    return -1;
  };
  const keyIndexes=[find('受注番号'),find('商品コード'),find('商品SKU','SKU'),find('受注ステータス')];
  const existingKeys=new Set();
  if(keyIndexes.every(index=>index>=0) && sh.getLastRow()>1){
    const count=sh.getLastRow()-1;
    const keyColumns=keyIndexes.map(index=>index>=0
      ? sh.getRange(2,3+index,count,1).getDisplayValues().map(row=>row[0])
      : new Array(count).fill(''));
    for(let i=0;i<count;i++){
      const values=keyColumns.map((column,pos)=>{
        const value=String(column[i]||'').trim();
        return pos===0?value.replace(/^niyantarose-/i,''):value;
      });
      existingKeys.add(values.join('|'));
    }
  }
  const fresh=仕分け証跡重複除外_(header,rows,existingKeys);
  if(!fresh.length) return;
  const now=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm:ss');
  const out=fresh.map(r=>{ const a=r.slice(0,ncol); while(a.length<ncol) a.push(''); return [now, srcName].concat(a); });
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

// 受注明細に「確保済み」「不足」「確保内訳」列を個数の隣へ用意する(2026-07-23 ユーザー要望
// 「受注明細でもこの3列を確認したい。注文数量の近くに」)。取込で削除→ここで作り直し=常に個数の隣。
// 逆順に挿すので並びは 個数|確保済み|不足|確保内訳|EMS番号 (取り置き登録と同じ順)。
function 受注明細_確保列を用意_(recv){
  const hr=受注ヘッダー行_(recv);
  const read=()=>recv.getRange(hr,1,1,Math.max(recv.getLastColumn(),1)).getValues()[0].map(v=>String(v||'').trim());
  const names=['確保済み','不足','確保内訳'];
  if(names.some(n=>read().indexOf(n)<0)){
    names.forEach(n=>{ const i=read().indexOf(n); if(i>=0) recv.deleteColumn(i+1); }); // 中途半端な残骸は作り直す
    let g=read().indexOf('個数'); if(g<0) g=read().length-1;
    ['確保内訳','不足','確保済み'].forEach(n=>{
      recv.insertColumnAfter(g+1);
      recv.getRange(hr,g+2).setValue(n).setFontWeight('bold');
    });
  }
  const head=read();
  return {確保済み:head.indexOf('確保済み')+1,不足:head.indexOf('不足')+1,確保内訳:head.indexOf('確保内訳')+1};
}

// 取り置き登録の候補(確保済み/不足/確保内訳を計算済み)を受注明細の同キー行へ書き戻す。
// 取り置き登録の更新のたびに呼ばれ、受注明細でも同じ数字が見える。候補に無い行(出荷GO等)は空欄。
// 同一キーの分割行には同じ集計値が入る(取り置き登録側で受注番号|SKU単位に束ねているため)。
function 受注明細_確保列を書く_(candidates){
  const ss=SpreadsheetApp.getActive(), recv=ss.getSheetByName(HIKIATE_CFG.受注);
  if(!recv) return;
  const M=列マップ_(recv), last=recv.getLastRow();
  if(last<=M.hr) return;
  const cols=受注明細_確保列を用意_(recv);
  const byKey={};
  (candidates||[]).forEach(c=>{
    if(String(c.判定||'')==='棚戻し待ち') return; // キャンセル戻し行は受注明細に無い
    const k=取り置き_行キー_(c);
    if(!(k in byKey)) byKey[k]=c;
  });
  const R=recv.getRange(M.hr+1,1,last-M.hr,recv.getLastColumn()).getValues();
  const 確保out=[],不足out=[],内訳out=[];
  R.forEach(row=>{
    const ban=String(row[M.番号]||'').trim();
    const c=ban? byKey[取り置き_行キー_({ban,sku:M.SKU>=0?row[M.SKU]:'',code:row[M.コード]})] : null;
    確保out.push([c&&c.確保済み!=null&&c.確保済み!==''?c.確保済み:'']);
    不足out.push([c&&c.不足!=null&&c.不足!==''?c.不足:'']);
    内訳out.push([c?String(c.確保内訳||''):'']);
  });
  recv.getRange(M.hr+1,cols.確保済み,確保out.length,1).setValues(確保out);
  recv.getRange(M.hr+1,cols.不足,不足out.length,1).setValues(不足out);
  recv.getRange(M.hr+1,cols.確保内訳,内訳out.length,1).setValues(内訳out);
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
  const a1=注文境界A1_(bans.map(ban=>({ban})),startRow,ncol);
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
  色_代引:'#ffd966', // 代引き注文の受注番号マーク(入金前でも発送対象。未入金の赤と区別する濃いめ黄)
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
        msg+='・'+code+' x'+q+' / '+kind+' / EMS状態:'+st+' / '+(state||'有効')+' (記録のみ・需要には影響しない)\n';
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
// 商品コード欄そのものが注文番号の在庫(中古品・名指し買付)を検出する。
// ただの短い数字コードを誤認しないよう、GoQ受注番号と同じ7桁以上の数字だけを候補にする。
function 注文番号在庫コード_(v){ const s=String(v||'').trim(); return /^\d{7,}$/.test(s)?s:''; }
// 注文番号在庫の引当先を、現在の受注明細に残る取り寄せ行から一意に決める。
// 取り寄せ行が複数なら、商品コード/SKU自体が注文番号の行だけを採用する。
function 注文番号指定引当先_(code, lines){
  const ban=注文番号在庫コード_(code);
  if(!ban) return {ban:'', line:null, reason:''};
  const target=(lines||[]).filter(l=>String(l.ban||'').trim()===ban && !l.キャンセル && l.kbn==='取り寄せ' && (Number(l.qty)||0)>0);
  if(!target.length) return {ban, line:null, reason:'対象注文または取り寄せ行なし'};
  if(target.length===1) return {ban, line:target[0], reason:''};
  const exact=target.filter(l=> normCode_(l.code)===ban || normCode_(l.sku)===ban);
  if(exact.length===1) return {ban, line:exact[0], reason:''};
  return {ban, line:null, reason:'取り寄せ行を一意に特定できない'};
}
// 注文番号在庫は商品数量ではなくEMS行数量の全数を、その注文の専用品として割り当てる。
function 注文番号指定割当一覧_(items, lines){
  const assignments=[], warnings=[];
  (items||[]).forEach((item,index)=>{
    const ban=注文番号在庫コード_(item.code); if(!ban) return;
    const hit=注文番号指定引当先_(ban,lines), qty=Math.max(0,Number(item.qty)||0);
    if(!hit.line){ warnings.push({index:item.index==null?index:item.index, ban, reason:hit.reason}); return; }
    if(qty<=0) return;
    assignments.push({index:item.index==null?index:item.index, ban, key:ban, qty,
      arrival:item.arrival||'', ems:item.ems||'', line:hit.line});
  });
  return {assignments,warnings};
}
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
// ===== 商品コード別名(旧コード→現行コード) =====
// 韓国側の発注リスト/EMSリストが旧コードのまま書いてくる商品を、現行コードとして全照合系
// (②EMS引当・P列計画・入荷日チェック・全件再計算・商品診断)で同一商品と認識させる。
// 機械的な末尾落としは別バリエーション(例 MRBLUE44-11/-12)を誤引当するため、宣言した組だけを対象にする。
// 追加は発注共有「コード別名」シート(A列=旧コード / B列=現行コード、2行目以降)か、下の既定マップへ。
let _コード別名キャッシュ=null;
function 引当_コード別名マップ_(){
  if(_コード別名キャッシュ) return _コード別名キャッシュ;
  const map={'AISTALT01S':'AISTALT01S-0'}; // エイリアンステージ アートブック(特別版): 韓国側は旧コードで書いてくる(2026-07-23)
  try{
    if(typeof 発注共有を開く_==='function'){
      const sh=発注共有を開く_().getSheetByName('コード別名');
      if(sh && sh.getLastRow()>=2){
        sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(r=>{
          const from=normCode_(r[0]), to=normCode_(r[1]);
          if(from && to && from!==to) map[from]=to;
        });
      }
    }
  }catch(e){} // 発注共有を開けない時は既定マップだけで動く(照合が狭くなるだけで安全側)
  _コード別名キャッシュ=map;
  return map;
}
// 照合キー配列へ別名を双方向(旧→現行・現行→旧)で追加する。どちら側が旧コードでも出会える
function 引当_コード別名展開_(keys){
  const map=引当_コード別名マップ_();
  keys.slice().forEach(k=>{ const to=map[k]; if(to && keys.indexOf(to)<0) keys.push(to); });
  keys.slice().forEach(k=>{ Object.keys(map).forEach(from=>{ if(map[from]===k && keys.indexOf(from)<0) keys.push(from); }); });
  return keys;
}
// 新しい取り置き割当/P列計画では、推測による複合・月号コード展開を行わない。
// 販売SKUの末尾a/bだけを在庫基底へ寄せ、商品コード自体は完全一致(＋宣言済み別名)だけを許可する。
function 引当用照合キー一覧_(sku, code){
  const out=[];
  const add=v=>{ const key=normCode_(v); if(key && out.indexOf(key)<0) out.push(key); };
  const salesSku=normCode_(sku);
  add(salesSku);
  if(/[AB]$/.test(salesSku)) add(salesSku.slice(0,-1));
  add(code);
  return 引当_コード別名展開_(out);
}
function codeKeys_(code){
  const c=normCode_(code); const keys=[c];
  const m=c.match(/^(.+)-(\d{2})(\d{2})-(\d{2})$/);
  if(m && (m[4]===m[2]||m[4]===m[3])) keys.push(m[1]+'-'+m[4]);
  // 月号付き雑誌コード: 末尾の -YYMM / -YYYYMM (例 EBS1504B-2607, CSG1504--202607) を落とした基底コードも候補に
  // GoQ側は月号なし(EBS1504B)で登録されているため。YY=2x・MM=01〜12の形だけ対象(巻数セット等の誤爆防止)
  const mg=c.match(/^(.+?)-(?:20)?(2\d)(0[1-9]|1[0-2])$/);
  if(mg){ const base=mg[1].replace(/-+$/,''); if(base && keys.indexOf(base)<0) keys.push(base); }
  return 引当_コード別名展開_(keys);
}
// ===== 定期購読の月号照合(例: EBS1504B_2607=7月号) =====
// 商品ページのルール「ご注文月の翌月号をお届け」に合わせて、
// 月号付き在庫(-YYMM/-YYYYMM)は「注文月+1＝その号」の注文にだけ引き当てる(号ズレ防止)。
// 月号なしの在庫・注文日時が読めない行は無制限(従来通り)。P列の名指しは人の判断なので月号チェックを通さない。
function 日付値_(v){
  if(v instanceof Date) return isNaN(v.getTime())?null:v;
  const s=ymd_(v);
  if(/^20\d{2}-\d{2}-\d{2}$/.test(s)) return new Date(s+'T00:00:00');
  return null;
}
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
// Google/Excelシリアル日(46212等)を yyyy-MM-dd に。QUERY/IMPORTRANGEが数値で返す到着日用。
function ymdFromSheetSerial_(serial){
  if(typeof serial!=='number' || !isFinite(serial) || serial<20000 || serial>500000) return '';
  const d=new Date(Math.round((serial-25569)*86400000));
  if(isNaN(d.getTime())) return '';
  const p=n=>('0'+n).slice(-2);
  return d.getUTCFullYear()+'-'+p(d.getUTCMonth()+1)+'-'+p(d.getUTCDate());
}
function 実在YMD_(year,month,day){
  const y=Number(year),m=Number(month),d=Number(day);
  if(!Number.isInteger(y)||!Number.isInteger(m)||!Number.isInteger(d)||m<1||m>12||d<1||d>31) return '';
  const date=new Date(Date.UTC(y,m-1,d));
  if(date.getUTCFullYear()!==y || date.getUTCMonth()!==m-1 || date.getUTCDate()!==d) return '';
  const p=n=>('0'+n).slice(-2);
  return y+'-'+p(m)+'-'+p(d);
}
// 旧処理がシリアル値 46213 を new Date("46213") と解釈して作った
// 「西暦46213年1月1日」のDateを、元のシリアル日(2026-07-10)へ戻す。
function 旧シリアル年Date補正日_(v){
  if(!(v instanceof Date) || isNaN(v.getTime()) || v.getMonth()!==0 || v.getDate()!==1) return '';
  const s=ymdFromSheetSerial_(v.getFullYear());
  return /^20\d{2}-\d{2}-\d{2}$/.test(s)? s : '';
}
function 入荷日シート値補正_(v){
  const s=旧シリアル年Date補正日_(v);
  return s? new Date(s+'T00:00:00') : null;
}
// 日付を yyyy-MM-dd 文字列に正規化(Date/文字列どちらでも)。入荷日と到着日の一致判定用。
// 手入力の「26/07/09(木)」のような2桁年・曜日付きも読む。パースできない文字列はそのまま返す
function ymd_(v){
  const p=n=>('0'+n).slice(-2);
  if(v instanceof Date){
    if(isNaN(v.getTime())) return '';
    const recovered=旧シリアル年Date補正日_(v); if(recovered) return recovered;
    return v.getFullYear()+'-'+p(v.getMonth()+1)+'-'+p(v.getDate());
  }
  if(typeof v==='number' && isFinite(v)){
    const fromSerial=ymdFromSheetSerial_(v);
    if(fromSerial) return fromSerial;
  }
  const s=String(v||'').trim(); if(!s) return '';
  // 46212 のようなシリアル数値文字列を new Date("46212")→46212-01-01 と誤読しない
  if(/^\d{1,6}(\.\d+)?$/.test(s)){
    const fromSerial=ymdFromSheetSerial_(parseFloat(s));
    if(fromSerial) return fromSerial;
  }
  // 旧バグがシートへ書いた "46213-01-01" を元のシリアル日へ戻す。
  // 通常の日付として末尾4桁だけを拾う前に判定する。
  let m=s.match(/^(\d{5,6})-0?1-0?1$/);
  if(m){
    const recovered=ymdFromSheetSerial_(parseFloat(m[1]));
    if(recovered) return recovered;
  }
  m=s.match(/(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if(m) return 実在YMD_(m[1],m[2],m[3]);
  m=s.match(/(\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if(m) return 実在YMD_('20'+m[1],m[2],m[3]);
  if(/[-\/.]/.test(s)){
    const d=new Date(s.replace(/\//g,'-'));
    if(!isNaN(d.getTime())) return d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate());
  }
  return s;
}

// 「今回入荷EMS」として扱う着済行か。実行日だけでなく、今回EMS在庫の到着日と入荷日が一致する行も対象。
function 今回到着扱い_(l, 入荷消費OK){
  return !!(l && (l.今回P || (l.入荷 && (typeof 入荷消費OK === 'function' ? 入荷消費OK(l) : 入荷日今日_(l.入荷日値)))));
}
function 今回行判定_(l, 入荷消費OK){
  if(!l) return false;
  if((Number(l.alloc)||0)>0) return true;
  // 台帳で確保済み(取り置き中)の行は、再実行でも「出せる状態」を維持する
  // (この判定が新規割当だけだと、④を回し直すたびに出荷可能が引当待ちへ落ちてしまう)
  if((Number(l.取り置き中数量)||0)>0) return true;
  // 台湾・中国ルートは当日手入力の入荷日だけ「今回」扱い(当日中に④を回せば出荷可能に上がる旧運用)
  return !!(l.別ルート && l.入荷 && 入荷日今日_(l.入荷日値));
}
// 注文一覧にだけ使う表示コード。今回引当で確定した実コードを優先し、
// それ以外は受注明細の元コードを維持する(受注明細自体は変更しない)。
function 注文一覧表示コード_(line, isCurrent, stockKey){
  return String((line&&line.matchedKey) || (isCurrent&&stockKey) || (line&&line.code) || '').trim();
}
function 引当行状態_(l,cfg){
  if(l.キャンセル) return {st:'キャンセル',color:cfg.色_グレー};
  if(l.kbn==='即納') return {st:'即納',color:cfg.色_水};
  if(l.kbn==='指定なし') return {st:'要確認',color:cfg.色_橙};
  // 三段階(2026-07-22): 段階付与済みの行は段階から表示を決める。現物=緑 / 先行含み=薄青 / 到着済=ラベンダー / 不足=色なし
  if(l.段階付与){
    const qty=Math.max(0,Number(l.qty)||0);
    const phys=Math.max(0,Number(l.現物確認済み数量)||0);
    const arrived=Math.max(0,Number(l.到着済引当数量)||0);
    const planned=Math.max(0,Number(l.先行引当数量)||0);
    const total=phys+arrived+planned+Math.max(0,Number(l.別ルート済数量)||0);
    if(total<qty) return {st:'不足'+(qty-total)+'個',color:null};
    if(phys>=qty) return {st:'現物確認済み',color:cfg.色_緑};
    if(planned>0) return {st:arrived+phys>0?'一部先行':'先行',color:cfg.色_水};
    return {st:'到着済',color:cfg.色_着};
  }
  if((Number(l.alloc)||0)>0) return {st:'引当(今回)',color:cfg.色_黄};
  if((Number(l.取り置き中数量)||0)>0) return {st:'取り置き中',color:cfg.色_着};
  if(l.別ルート && l.入荷) return {st:'着済',color:cfg.色_着}; // 台湾・中国の手入力入荷
  return {st:'在庫待ち',color:null};
}
// 「なぜこの状態なのか」を1列で読める文言にする(分類シートの状態の理由列)。
function 行状態理由_(l){
  if(!l || l.キャンセル) return 'キャンセル';
  if(l.kbn==='即納') return '即納(店頭現物)';
  if(!l.段階付与) return '';
  const qty=Math.max(0,Number(l.qty)||0), parts=[];
  const phys=Math.max(0,Number(l.現物確認済み数量)||0);
  const arrived=Math.max(0,Number(l.到着済引当数量)||0);
  const planned=Math.max(0,Number(l.先行引当数量)||0);
  const beturoute=Math.max(0,Number(l.別ルート済数量)||0);
  if(phys>0) parts.push('現物'+phys+'個確認済み');
  if(arrived>0) parts.push('到着済'+arrived+'個');
  if(planned>0) parts.push('先行'+planned+'個'+(l.先行到着予定?'('+l.先行到着予定+'到着予定)':''));
  if(beturoute>0) parts.push('別ルート'+beturoute+'個');
  if(!parts.length) return '未引当';
  const lack=Math.max(0,qty-(phys+arrived+planned+beturoute));
  if(lack>0) parts.push(lack+'個不足');
  return '注文'+qty+'個/'+parts.join('・');
}

function 注文出荷準備OK_(arr){
  const active=(arr||[]).filter(l=>l && !l.キャンセル);
  return active.length>0 && active.every(l=>l.kbn==='即納'||(l.kbn==='取り寄せ'&&残必要計算_(l)===0));
}
// 代引き(代金引換)注文の判定。お支払い方法列の値で見る。列が無いシートでは呼び出し側でfalse=従来動作
function 代引き支払_(value){
  return /代引|代金引換|コレクト/.test(String(value==null?'':value));
}

// 注文内全行の注文数と三段階数量の集計。段階付与済みの行は段階列を、
// 未付与の旧経路行は台帳確保(取り置き中数量)+今回割当(alloc)を到着済相当として数える(移行期互換)。
function 注文充足集計_(arr){
  const active=(arr||[]).filter(l=>l && !l.キャンセル);
  const out={注文数量:0,現物確認済み:0,到着済引当:0,先行引当:0,別ルート済:0,確保総数:0,不足:0,行数:active.length};
  active.forEach(l=>{
    const qty=Math.max(0,Number(l.qty)||0);
    const phys=Math.max(0,Number(l.現物確認済み数量)||0);
    const planned=Math.max(0,Number(l.先行引当数量)||0);
    const beturoute=Math.max(0,Number(l.別ルート済数量)||0);
    const arrived=l.段階付与?Math.max(0,Number(l.到着済引当数量)||0)
      :Math.max(0,Number(l.取り置き中数量)||0)+Math.max(0,Number(l.alloc)||0);
    const 即納分=l.kbn==='即納'?qty:0; // 即納は現物が店にある前提(従来どおり確保扱い)
    const secured=Math.min(qty,phys+arrived+planned+beturoute+即納分);
    out.注文数量+=qty; out.現物確認済み+=Math.min(qty,phys); out.到着済引当+=arrived;
    out.先行引当+=planned; out.別ルート済+=beturoute;
    out.確保総数+=secured; out.不足+=Math.max(0,qty-secured);
  });
  return out;
}

// 分類は「今回新規割当の有無」ではなく先行を含む確保総数で決める(三段階 2026-07-22)。
// 優先順: 確保0→一部→希望日→代引→未入金→入金済。必ず1分類だけを返す。
function 注文区分判定_(arr, paid, cod){
  const s=注文充足集計_(arr);
  if(s.行数===0 || s.確保総数<=0) return 'wait';
  if(s.不足>0) return 'part';
  const active=(arr||[]).filter(l=>l && !l.キャンセル);
  if(active.some(l=>希望日未来_(l.届))) return 'hold';
  if(cod) return 'ship'; // 代引きは未入金でも出荷対象(受注番号セル黄で表示)
  if(!paid) return 'keep';
  return 'ship';
}

// 物理に出荷できる注文=不足0かつ先行0。分類上の「出荷可能」とは独立に持つ
// (先行だけで揃った注文は分類は出荷可能でも、ピック・納品書には出さない)。
function 注文物理出荷可_(arr){
  const s=注文充足集計_(arr);
  return s.行数>0 && s.不足===0 && s.先行引当===0;
}
function 発送可否判定_(arr, paid){
  if(!注文出荷準備OK_(arr)) return '';
  if((arr||[]).some(l=>l && !l.キャンセル && 希望日未来_(l.届))) return '発送可能希望日待ち';
  return paid ? '発送可能' : '発送可能入金待ち';
}

// 入荷日が現在便と数日だけズレた時の救済は、旧日付に実在便が無い場合だけ許可する。
// 旧日付の実EMS便が存在するなら、その注文は旧便で受取済みの可能性があるため日付を保持する。
function 入荷日ズレ自動修復可_(旧日, 候補日, 実在到着日Set, 過去箱分){
  if((Number(過去箱分)||0)>0) return false;
  const old=ymd_(旧日), next=ymd_(候補日);
  if(!/^20\d{2}-\d{2}-\d{2}$/.test(old) || !/^20\d{2}-\d{2}-\d{2}$/.test(next) || old===next) return false;
  if(!実在到着日Set || typeof 実在到着日Set.has!=='function') return false;
  if(実在到着日Set.has(old)) return false;
  const diff=Math.abs((new Date(next+'T00:00:00').getTime()-new Date(old+'T00:00:00').getTime())/86400000);
  return diff<=3;
}

function P列計画_新規確定割当_(plan,ledgerRows){
  const fixed={};
  (ledgerRows||[]).forEach(r=>{
    if(String(r.状態||'')!=='取り置き中' || !実EMS番号_(r.元EMS番号)) return;
    const key=取り置き_供給キー_(r.元EMS番号,r.元EMS商品コード||r.商品コード)+'|'+String(r.受注番号||'').trim();
    fixed[key]=(fixed[key]||0)+(Number(r.取り置き数量)||0);
  });
  const out=[];
  (plan&&plan.rows||[]).filter(r=>!String(r.directBan||'').trim()).forEach(r=>(r.entries||[]).forEach(e=>{
    const sourceCode=String(r.sourceCode||r.code||'').trim();
    const key=取り置き_供給キー_(r.ems,sourceCode)+'|'+String(e.ban||'').trim();
    const qty=Number(e.qty)||0,take=Math.min(qty,fixed[key]||0),left=qty-take;
    fixed[key]=Math.max(0,(fixed[key]||0)-take);
    if(left>0) out.push({ems:r.ems,code:r.code,sourceCode,ban:String(e.ban||''),qty:left});
  }));
  return out;
}

function 引当出力計画_(supplies,allocationPlan,projectedLedger){
  const byKey={},bySource={};
  (supplies||[]).forEach(s=>{
    const sourceCode=String(s.sourceCode||s.code||'').trim(),directBan=String(s.directBan||'').trim();
    const sourceKey=取り置き_供給キー_(s.ems,sourceCode),key=sourceKey+'|'+directBan;
    if(!byKey[key]){
      byKey[key]={ems:String(s.ems||''),code:sourceCode,matchCode:normCode_(s.code),sourceCode,directBan,qty:0,arrival:s.arrival,consumers:[],surplus:0};
      (bySource[sourceKey]=bySource[sourceKey]||[]).push(byKey[key]);
    }
    byKey[key].qty+=Number(s.qty)||0;
  });
  const newIds=new Set((allocationPlan&&allocationPlan.newRows||[]).map(r=>String(r.取置ID||'')));
  const 当日=ymd_(new Date()); // 黄色=「この実行の新規」だけだと、反映や前の②が作った当日確保が次の②で薄紫に落ちて見失う
  (projectedLedger||[]).forEach(r=>{
    if(String(r.状態||'')!=='取り置き中' || !実EMS番号_(r.元EMS番号)) return;
    const sourceKey=取り置き_供給キー_(r.元EMS番号,r.元EMS商品コード||r.商品コード),owner=String(r.受注番号||'').trim();
    const target=byKey[sourceKey+'|'+owner] || byKey[sourceKey+'|'] || ((bySource[sourceKey]||[]).length===1?bySource[sourceKey][0]:null);
    if(target) target.consumers.push({ban:String(r.受注番号||''),qty:Number(r.取り置き数量)||0,
      current:newIds.has(String(r.取置ID||'')) || (!!当日 && ymd_(r.登録日時)===当日),row:r});
  });
  (allocationPlan&&allocationPlan.surplus||[]).forEach(s=>{
    const key=取り置き_供給キー_(s.ems,s.sourceCode||s.code)+'|'+String(s.directBan||'').trim(),target=byKey[key];
    if(target) target.surplus=Number(s.qty)||0;
  });
  return {supplies:Object.keys(byKey).map(key=>byKey[key]),surplus:(allocationPlan&&allocationPlan.surplus||[]).map(r=>Object.assign({},r))};
}

function 引当計画_行へ反映_(lines,ledgerSummary,projectedSummary,newRows){
  const newRowsByKey={};
  (newRows||[]).forEach(r=>{ const key=取り置き_行キー_(r); (newRowsByKey[key]=newRowsByKey[key]||[]).push(r); });
  (lines||[]).forEach(l=>{
    const key=取り置き_行キー_(l),activeRows=projectedSummary.activeRowsByKey[key]||[],added=newRowsByKey[key]||[];
    l.取り置き中数量=ledgerSummary.activeByKey[key]||0;
    l.計画後取り置き中数量=projectedSummary.activeByKey[key]||0;
    l.alloc=added.reduce((sum,r)=>sum+(Number(r.取り置き数量)||0),0);
    l.引当成立=l.alloc>0&&残必要計算_(l)===0;
    // 三段階(2026-07-22): 計画後(このランの新規行込み)の段階別数量を行へ付与する。
    // 旧開始前在庫(要移行)は移行完了まで物理確保として到着済相当に数える(残必要と同じ根拠)。
    const stage=(projectedSummary.stageByKey||{})[key];
    if(stage){
      l.段階付与=true;
      l.現物確認済み数量=Number(stage.現物確認済み数量)||0;
      l.到着済引当数量=(Number(stage.到着済引当数量)||0)+(Number(stage.要移行数量)||0);
      l.先行引当数量=Number(stage.先行引当数量)||0;
      const qty=Math.max(0,Number(l.qty)||0);
      const total=l.現物確認済み数量+l.到着済引当数量+l.先行引当数量
        +Math.max(0,Number(l.別ルート済数量)||0)+(l.kbn==='即納'?qty:0);
      l.未引当数量=Math.max(0,qty-total);
      l.主段階=qty>0&&l.現物確認済み数量>=qty?'現物確認済み':(l.先行引当数量>0?'先行':(l.到着済引当数量>0?'到着済':''));
      // 先行行の到着予定日(最小)を表示用に写す
      const plannedRow=activeRows.filter(r=>String(r.引当段階||'')==='先行'&&String(r.EMS到着予定日||'').trim())
        .sort((a,b)=>String(a.EMS到着予定日).localeCompare(String(b.EMS到着予定日)))[0];
      if(plannedRow) l.先行到着予定=String(plannedRow.EMS到着予定日);
    }
    const matched=added[0]||activeRows[0];
    if(matched){
      const rawSource=String(matched.元EMS商品コード||'').trim();
      const source=normCode_(rawSource);
      l.matchedKey=(注文番号在庫コード_(rawSource)||タグ受注番号_(rawSource))?normCode_(matched.商品コード):source||normCode_(matched.商品コード);
      l.台帳入荷日=typeof 全件再計算_台帳到着日_==='function'?全件再計算_台帳到着日_(matched):'';
    }
    l.箱EMS=Array.from(new Set(activeRows.map(r=>String(r.元EMS番号||'').trim()).filter(v=>実EMS番号_(v)))).join(', ');
  });
  return lines;
}

function 引当切替差分_純計算_(plannedRows,currentRows){
  const keyOf=r=>[String(r&&r.ban||''),normCode_(r&&(r.matchCode||r.code)),normCode_(r&&r.sku)].join('|');
  const valueOf=r=>JSON.stringify({code:normCode_(r&&r.code),qty:Number(r&&r.qty)||0,state:String(r&&r.state||''),ems:String(r&&r.ems||'')});
  const planned={},current={},keys=[];
  const add=(groups,row)=>{
    const key=keyOf(row);
    if(!(key in groups)){ groups[key]=[]; if(keys.indexOf(key)<0) keys.push(key); }
    groups[key].push(row);
  };
  (plannedRows||[]).forEach(r=>add(planned,r));
  (currentRows||[]).forEach(r=>add(current,r));
  const out=[];
  keys.forEach(key=>{
    const p=(planned[key]||[]).slice(), c=(current[key]||[]).slice();
    // 同一行を先に相殺し、同じ受注・商品が複数行あっても上書きで消さない。
    for(let i=p.length-1;i>=0;i--){
      const found=c.findIndex(row=>valueOf(row)===valueOf(p[i]));
      if(found>=0){ p.splice(i,1); c.splice(found,1); }
    }
    while(p.length&&c.length){
      const after=p.shift(),before=c.shift();
      out.push({key,change:'更新',before,after});
    }
    p.forEach((after,index)=>out.push({key:key+'#追加'+index,change:'追加',before:null,after}));
    c.forEach((before,index)=>out.push({key:key+'#削除'+index,change:'削除',before,after:null}));
  });
  return out.sort((a,b)=>{ const rank={'更新':0,'追加':1,'削除':2}; return rank[a.change]-rank[b.change]||a.key.localeCompare(b.key); });
}

function 引当切替_計画行_(lines){
  const paidByBan={}, codByBan={};
  (lines||[]).forEach(l=>{ if(!l||l.キャンセル) return;
    if(l.paid) paidByBan[String(l.ban||'')]=true;
    if(l.代引き) codByBan[String(l.ban||'')]=true; });
  return (lines||[]).filter(l=>l&&!l.キャンセル).map(l=>({
    ban:String(l.ban||''),氏名:String(l.氏名||''),code:注文一覧表示コード_(l,false,null),sku:String(l.sku||''),
    商品名:String(l.商品名||''),qty:Number(l.qty)||0,
    state:引当行状態_(l,HIKIATE_CFG).st+(paidByBan[String(l.ban||'')]?'':(codByBan[String(l.ban||'')]?'・代引き':'・入金待ち')),ems:String(l.箱EMS||'')
  }));
}

function 引当切替_現行出力行_(ss,plannedRows){
  const out=[],byBanCode={},byBan={},used=new Set();
  (plannedRows||[]).forEach((r,index)=>{
    const item={row:r,index},ban=String(r.ban),key=ban+'|'+normCode_(r.code);
    (byBanCode[key]=byBanCode[key]||[]).push(item);
    (byBan[ban]=byBan[ban]||[]).push(item);
  });
  [HIKIATE_CFG.待ち,HIKIATE_CFG.部分,HIKIATE_CFG.取置,HIKIATE_CFG.希望,HIKIATE_CFG.出荷].forEach(name=>{
    const sh=ss.getSheetByName(name); if(!sh) return;
    const values=sh.getDataRange().getValues(),hr=values.findIndex(row=>row.map(v=>String(v||'').trim()).indexOf('受注番号')>=0);
    if(hr<0) return;
    const head=values[hr].map(v=>String(v||'').trim()),col=n=>head.indexOf(n);
    const cBan=col('受注番号'),cName=col('氏名'),cCode=col('商品コード'),cItem=col('商品名'),cSku=col('SKU'),cQty=col('個数'),cState=col('状態'),cEms=col('EMS番号');
    values.slice(hr+1).forEach(row=>{
      const ban=String(row[cBan]||'').trim(),code=String(row[cCode]||'').trim(); if(!ban||!code) return;
      const qty=cQty>=0?Number(row[cQty])||0:0,state=cState>=0?String(row[cState]||''):'',ems=cEms>=0?String(row[cEms]||''):'';
      const 氏名=cName>=0?String(row[cName]||'').trim():'',商品名=cItem>=0?String(row[cItem]||'').trim():'';
      const planKey=ban+'|'+normCode_(code);
      let cand=(byBanCode[planKey]||[]).filter(item=>!used.has(item.index));
      // 注文番号在庫などで表示コード自体が切り替わる場合は、同じ受注番号の未使用行へフォールバックする。
      if(!cand.length) cand=(byBan[ban]||[]).filter(item=>!used.has(item.index));
      let sku=cSku>=0?String(row[cSku]||'').trim():'';
      let found=null;
      if(sku) found=cand.find(item=>normCode_(item.row.sku)===normCode_(sku))||null;
      else {
        found=cand.find(item=>(Number(item.row.qty)||0)===qty&&String(item.row.state||'')===state&&String(item.row.ems||'')===ems)||cand[0]||null;
        if(found) sku=String(found.row.sku||'');
      }
      if(found) used.add(found.index);
      out.push({ban,氏名,code,sku,商品名,qty,state,ems,matchCode:found?String(found.row.code||''):''});
    });
  });
  return out;
}

function 引当切替差分を作成(){
  const preview=引当実行_本体_({preview:true}); if(!preview) return;
  const ss=SpreadsheetApp.getActive(),planned=引当切替_計画行_(preview.lines);
  const current=引当切替_現行出力行_(ss,planned),diff=引当切替差分_純計算_(planned,current);
  let sh=ss.getSheetByName('引当切替差分'); if(!sh) sh=ss.insertSheet('引当切替差分');
  const clear=sh.getRange(1,1,Math.max(1,sh.getMaxRows()),Math.max(1,sh.getMaxColumns())); if(clear&&typeof clear.clearContent==='function') clear.clearContent();
  // 初期登録と同じ調子で読めるよう、キーは受注番号/氏名/商品コード/SKU/商品名へ分解して出す
  const fieldOf=(r,name)=>String((r.after&&r.after[name]!=null&&r.after[name]!==''?r.after[name]:(r.before?r.before[name]:''))||'');
  const values=[['変更','受注番号','氏名','商品コード','SKU','商品名','現行状態','計画状態','現行数量','計画数量','現行EMS','計画EMS']]
    .concat(diff.map(r=>[r.change,fieldOf(r,'ban'),fieldOf(r,'氏名'),fieldOf(r,'code'),fieldOf(r,'sku'),fieldOf(r,'商品名'),
      r.before?r.before.state:'',r.after?r.after.state:'',r.before?r.before.qty:'',r.after?r.after.qty:'',r.before?r.before.ems:'',r.after?r.after.ems:'']));
  sh.getRange(1,1,values.length,values[0].length).setValues(values);
  sh.getRange(1,1,1,values[0].length).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return {preview,diff};
}

function 引当実行(){ 直列_(()=>引当実行_本体_({preview:false})); }
function 引当実行_本体_(options){
  options=options||{};
  if(!options.ignoreRebuildGuard && typeof 全件再計算_通常処理ガード_==='function'){
    const guard=全件再計算_通常処理ガード_();
    if(guard){
      SpreadsheetApp.getUi().alert('通常引当を停止しています',
        '全件再計算が完了していないため、台帳を保護しています。\n停止段階: '+String(guard.stage||'不明')+'\n'+String(guard.error||''),SpreadsheetApp.getUi().ButtonSet.OK);
      return {success:false,error:'全件再計算ガード中'};
    }
  }
  const 開始ms=Date.now();
  const ss=SpreadsheetApp.getActive(), cfg=HIKIATE_CFG;
  const recv=ss.getSheetByName(cfg.受注);
  if(!recv){ SpreadsheetApp.getUi().alert('「'+cfg.受注+'」タブが無いで'); return; }
  const R=recv.getDataRange().getValues();

  // 受注明細→行オブジェクト(列は見出し名で特定。並び替え・列追加に強い)
  // 着済(着いたか)の判定: 即納 or (取り寄せ かつ 入荷日あり)。取り寄せで入荷日なし=未着。
  const M=列マップ_(recv), 受注hdr=M.hr;
  const 受注head=recv.getRange(M.hr,1,1,recv.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
  const c日時=受注head.indexOf('注文日時'); // 月号照合(定期購読)用
  const cステータス=受注head.indexOf('受注ステータス'); // 引当状況一覧のGoQ差分用
  const lines=[];
  for(let i=受注hdr;i<R.length;i++){
    const row=R[i];
    const ban=String(row[M.番号]||'').trim(), code=String(row[M.コード]||'').trim();
    if(!ban && !code) continue;
    const kbn=区分_(row[M.選択肢]);
    const 入荷日値=M.入荷>=0? row[M.入荷] : '', 入荷=String(入荷日値||'').trim()!=='';
    const 着=(kbn==='即納') || (kbn==='取り寄せ' && 入荷); // 着済かどうか
    const qty=Number(row[M.個数])||0;
    // 台湾・中国ルートは韓国EMS供給に乗らないため、手入力の入荷日を確保数量として扱う(台帳の対象外)
    const 別ルート=kbn==='取り寄せ' && 引当_別ルート判定_(row[M.選択肢], row[M.商品名]);
    lines.push({ i, ban, sortKey:番号num_(ban),
      氏名:row[M.氏名], 届:row[M.届], 商品名:row[M.商品名],
      code, sku:String(row[M.SKU]||'').trim(), qty, kbn,
      日時: c日時>=0? 日付値_(row[c日時]) : null,
      メモ: M.メモ>=0? String(row[M.メモ]==null?'':row[M.メモ]).trim() : '',
      paid:入金済み_(row[M.入金]), 代引き:M.支払>=0 && 代引き支払_(row[M.支払]),
      入荷, 入荷日値, 着, alloc:0, 引当成立:false, キャンセル:qty<=0,
      受注ステータス: cステータス>=0? String(row[cステータス]||'') : '',
      別ルート, 別ルート済数量: 別ルート && 入荷? qty : 0,
      出荷済分割: 引当_行出荷済み_(row,M) });
  }

  // ===== EMS在庫引当: 到着実績のある実EMSを、台帳の有効取り置きと新規注文へFIFO引当 =====
  const emv=ss.getSheetByName(cfg.EMS在庫);
  const emsD= emv? EMS明細_(emv) : {rows:[], cols:{コード:1,数量:2,EMS番号:3,到着:-1}};
  const E= emsD.rows, EC=emsD.cols; // EC=EMS在庫の列位置(見出し名で特定。到着日を挿入して列がズレてもOK)
  const 到着_=row=>{ if(EC.到着<0) return ''; const v=row[EC.到着]; // EMS到着日。QUERYがシリアル数値で返す場合あり
    if(v instanceof Date) return isNaN(v.getTime())? '' : v;
    const s=ymd_(v);
    if(/^20\d{2}-\d{2}-\d{2}$/.test(s)) return new Date(s+'T00:00:00');
    return String(v||'').trim(); };
  const candKeys=l=>引当用照合キー一覧_(l.sku,l.code);

  // A. 現在の取り置き台帳を受注明細へ重ねる。旧取り置き数・入荷日・履歴は必要数へ使わない。
  const ledgerRows=取り置き台帳_読む_();
  const movementRows=EMS在庫移動台帳_読む_();
  const ledgerSummary=取り置き_集計_(ledgerRows,movementRows);
  lines.forEach(l=>{ l.取り置き中数量=ledgerSummary.activeByKey[取り置き_行キー_(l)]||0; });
  別ルート二重控除_(lines,ledgerSummary.activeByKey); // 台湾/中国の台帳登録と入荷日方式の二重計上防止

  // A2. 到着実績のある実EMSだけを供給へ変換する。
  const allSupplies=EMS供給オブジェクト_(E,EC,到着_);
  const rebuildBlocked=typeof 全件再計算_ブロックSKU集合_==='function'?全件再計算_ブロックSKU集合_():new Set();
  const supplies=typeof 全件再計算_ブロック供給_==='function'?全件再計算_ブロック供給_(allSupplies,rebuildBlocked):allSupplies;

  // A3. 台帳に載らない出荷(週末GoQ直接発送など)を検知し、発送済み行として自動登録する。
  //     箱残を実態へ合わせてから割当・P列計画を作る(旧9a8f8d8の売り越し防止の台帳版)。
  let 出荷自動={newRows:[],review:[]};
  try{ 出荷自動=取り置き_未台帳出荷計画_(消込台帳_出荷済み行_(),ledgerRows,movementRows,supplies); }
  catch(e){ 出荷自動={newRows:[],review:[{受注番号:'-',商品コード:'-',理由:'消込台帳を読めずスキップ: '+e.message}]}; }
  const ledgerRowsForPlan=出荷自動.newRows.length? ledgerRows.concat(出荷自動.newRows) : ledgerRows;

  // B. P列はここでは書かず、表示計画だけを作る(自動登録分も使用済みとして渡す)。
  const pOptions=出荷自動.newRows.length? {追加台帳行:出荷自動.newRows} : {};
  if(options.clearCurrentP) pOptions.clearCurrentP=true;
  const pPlan=発注共有P列計画_(pOptions);
  if(pPlan.error){ SpreadsheetApp.getUi().alert(pPlan.error); return; }

  // B2. P列の手動名指し(説明文コード等のコード不一致救済)を供給のdirect名指しへ引き継ぐ。
  {
    const 救済=P列救済供給マップ_(pPlan.rows);
    supplies.forEach(s=>{
      if(s.directBan) return;
      const ban=救済[取り置き_供給キー_(s.ems,s.sourceCode||s.code)];
      if(ban) s.directBan=ban;
    });
  }

  // C. 純粋割当を計算・検証する。
  const supplyKeys=new Set(supplies.map(s=>取り置き_供給キー_(s.ems,s.sourceCode||s.code)));
  const explicit=P列計画_新規確定割当_(pPlan,ledgerRowsForPlan)
    .filter(e=>supplyKeys.has(取り置き_供給キー_(e.ems,e.sourceCode||e.code)));
  const allocationPlan=取り置き_割当計算_({
    // 台湾・中国の入荷済み分(別ルート済数量)は確保済みとして注文数量から除き、韓国EMSを消費させない
    // 分割出荷済みの行(出荷日あり)はもう送った分。需要に数えると箱を二重に食う(持ち戻り注文等)
    orders:lines.filter(l=>l.kbn==='取り寄せ'&&!l.キャンセル&&!l.出荷済分割).map(l=>({
      ban:l.ban,code:l.code,sku:l.sku,qty:Math.max(0,l.qty-(Number(l.別ルート済数量)||0)),
      sortKey:l.sortKey,i:l.i,keys:candKeys(l),paid:l.paid
    })),
    ledger:ledgerRowsForPlan,movements:movementRows,supplies,explicit
  });
  if(allocationPlan.errors.length){
    const ui=SpreadsheetApp.getUi();
    ui.alert('引当を中止しました',allocationPlan.errors.join('\n'),ui.ButtonSet.OK);
    return;
  }
  // タグ異常(名指し不在・多め買付)と台帳外出荷の不一致。完走して完了ダイアログで知らせる
  const 割当警告=(allocationPlan.warnings||[])
    .concat(出荷自動.review.map(r=>'台帳外出荷の要確認: '+r.受注番号+' '+r.商品コード+' '+r.理由))
    .concat(E.length && !supplies.length? ['⚠️ EMS在庫に行はあるのに供給0件です(到着日列の欠落・日付形式を確認。全行が引当対象外になっています)'] : []);

  const now=new Date();
  const projectedLedger=取り置き台帳_割当計画後行_(allocationPlan,ledgerRowsForPlan,now);
  const projectedSummary=取り置き_集計_(projectedLedger,movementRows);
  引当計画_行へ反映_(lines,ledgerSummary,projectedSummary,allocationPlan.newRows);
  const outputPlan=引当出力計画_(supplies,allocationPlan,projectedLedger);

  // preview は純粋計画を返すだけ。ここより前に運用シートへの書込みは無い。
  if(options.preview) return {allocationPlan,pPlan,lines,ledgerRows:ledgerRowsForPlan,movementRows,projectedLedger,outputPlan};

  // D. 検証成功後だけ、台帳を一括1回保存してからP列へ反映する(台帳外出荷の自動登録分も同じ保存に含める)。
  取り置き台帳_割当計画を反映_(allocationPlan,ledgerRowsForPlan,now);
  発注共有P列計画を反映_(pPlan);
  消込台帳更新_(); // 監査用にだけ更新し、割当数量や今回EMS消費の推測には使わない。

  const stock={}; outputPlan.surplus.forEach(r=>{ const k=normCode_(r.code); stock[k]=(stock[k]||0)+(Number(r.qty)||0); });
  const aliasMap={}; supplies.forEach(s=>引当用照合キー一覧_('',s.code).forEach(k=>{ if(!(k in aliasMap)) aliasMap[k]=normCode_(s.code); }));
  const keyInStock=l=>{ if(l&&l.matchedKey) return normCode_(l.matchedKey); for(const k of candKeys(l)){ if(k in stock) return k; if(aliasMap[k]) return aliasMap[k]; } return null; };
  const code到着={}; supplies.forEach(s=>{ const k=normCode_(s.code); if(!(k in code到着)) code到着[k]=s.arrival; });
  const 入荷消費OK_=()=>false;
  const 確定行数=(allocationPlan.newRows||[]).length, ズレ修復=0;


  // 注文ごと: 入金済み(1行でも入金日があれば入金済み。注文単位入金)
  const byOrder={};
  lines.forEach(l=> (byOrder[l.ban]=byOrder[l.ban]||[]).push(l));
  const paidOrder={};
  Object.keys(byOrder).forEach(ban=>{ paidOrder[ban]=byOrder[ban].some(l=>l.paid); });
  // 代引きは入金前でも発送対象(支払いは受け取り時)。分類だけ入金済み扱いにし、表示は黄色+「代引」で区別する
  const codOrder={};
  Object.keys(byOrder).forEach(ban=>{ codOrder[ban]=byOrder[ban].some(l=>l.代引き); });

  // 注文を分類:
  //  出荷可能=新規割当あり&全行出せる&入金済 / 希望日待ち=全行出せるが希望日が未来
  //  出荷GO未入金=全行出せるが未入金 / 部分在庫=新規割当あり+不足あり / 引当待ち=新規割当なし
  const 区分け=ban=> 注文区分判定_(byOrder[ban], paidOrder[ban], codOrder[ban]);
  // 引当待ちの「発送可否」(K列): 揃っているのに今回分でない=実は発送できる注文を見分ける
  const 発送可否_=ban=> 発送可否判定_(byOrder[ban], paidOrder[ban]||codOrder[ban]);

  // 出力(行の色=到着状況。未入金は受注番号セルだけ赤)
  // 取り置きメモ(受注明細の手書き列)は各出力の末尾に転記する(毎回作り直しても消えない見せ方)
  const HDR=['受注番号','氏名','お届け日','商品コード','商品名','個数','区分','入荷日','入金','状態','EMS番号'];
  // 三段階(2026-07-22): 段階別数量と「状態の理由」を全分類シートの共通末尾列へ出す
  const HDR_段階=['引当段階','現物確認済み','到着済引当','先行引当','不足','状態の理由'];
  const HDR_共=HDR.concat(['取り置き中数量','取り置きメモ'],HDR_段階);
  const HDR_待=HDR.concat(['発送可否','取り置き中数量','取り置きメモ'],HDR_段階);
  const seen=new Set(), seq=[]; lines.forEach(l=>{ if(!seen.has(l.ban)){ seen.add(l.ban); seq.push(l.ban);} });
  const waitRows=[], partRows=[], keepRows=[], holdRows=[], shipRows=[];
  seq.forEach(ban=>{
    const b=区分け(ban), paid=paidOrder[ban], 代引=codOrder[ban];
    const target = b==='ship'? shipRows : b==='keep'? keepRows : b==='hold'? holdRows : b==='part'? partRows : waitRows;
    const 可否 = (b==='wait')? 発送可否_(ban) : null; // 引当待ちのみK列に発送可否
    // 先行を含んで揃った注文は分類上出荷可能でも現物が未着=納品書・ピック対象外(状態列で明示)
    const 物理可=注文物理出荷可_(byOrder[ban]);
    byOrder[ban].forEach(l=>{
      const 状態=引当行状態_(l, cfg, 入荷消費OK_);
      let st=状態.st, color=状態.color;
      if(b==='ship' && !物理可) st+='・先行待ち(現物未着)';
      if(!l.キャンセル && !paid) st+=(代引?'・代引き':'・入金待ち');
      const stockKey=keyInStock(l);
      const 表示コード=注文一覧表示コード_(l, l.kbn==='取り寄せ' && l.入荷 && 入荷消費OK_(l), stockKey);
      const vals=[ban,l.氏名,l.届,表示コード,l.商品名,l.qty,l.kbn,l.入荷日値,paid?'済':(代引?'代引':'未'),st,l.箱EMS||''];
      if(b==='wait') vals.push(可否);
      vals.push(l.計画後取り置き中数量||'');
      vals.push(l.メモ||'');
      vals.push(l.主段階||'', l.現物確認済み数量||'', l.到着済引当数量||'', l.先行引当数量||'', l.未引当数量||'', 行状態理由_(l));
      target.push({ vals, color, ban, paid, qty:l.qty });
    });
  });

  書き出し_(ss, cfg.待ち, HDR_待, waitRows, 受注hdr);
  書き出し_(ss, cfg.部分, HDR_共, partRows, 受注hdr);
  書き出し_(ss, cfg.取置, HDR_共, keepRows, 受注hdr);
  書き出し_(ss, cfg.希望, HDR_共, holdRows, 受注hdr);
  書き出し_(ss, cfg.出荷, HDR_共, shipRows, 受注hdr);

  // 引当状況一覧(読み取り専用・全注文): 5分類と同じ計画データから生成し別計算をしない
  try{
    const 一覧linesByBan={}, 一覧status={};
    seq.forEach(ban=>{ 一覧linesByBan[ban]=byOrder[ban];
      一覧status[ban]=(byOrder[ban].find(l=>String(l.受注ステータス||'').trim())||{}).受注ステータス||''; });
    const 一覧rows=引当状況_一覧行_(一覧linesByBan,paidOrder,codOrder,一覧status);
    let 一覧sh=ss.getSheetByName(引当状況_CFG.シート); if(!一覧sh) 一覧sh=ss.insertSheet(引当状況_CFG.シート);
    一覧sh.clearContents();
    const 一覧HDR=引当状況_CFG.HDR;
    const 一覧data=[一覧HDR].concat(一覧rows.map(r=>一覧HDR.map(h=>r[h]!=null?r[h]:'')));
    一覧sh.getRange(1,1,一覧data.length,一覧HDR.length).setValues(一覧data);
    一覧sh.getRange(1,1,1,一覧HDR.length).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
    一覧sh.setFrozenRows(1);
    // 差異ありの行だけ背景を薄橙で目立たせる(範囲一括で適用)
    if(一覧rows.length){
      const bg=一覧rows.map(r=>new Array(一覧HDR.length).fill(r.GoQ差分==='差異あり'?cfg.色_橙:null));
      一覧sh.getRange(2,1,一覧rows.length,一覧HDR.length).setBackgrounds(bg);
    }
  }catch(e){ ss.toast('引当状況一覧の生成に失敗: '+e.message,'⚠️',8); }

  // ===== 受注明細の色分け: 即納=水/新規割当=黄/既存取り置き=薄紫/未割当=白/指定なし=橙 =====
  const recvStart=受注hdr+1, recvLast=recv.getLastRow();
  if(recvLast>=recvStart){
    const nc=recv.getLastColumn(), rowColor={}, rowUnpaid={}, rowCod={}, rowQty={};
    lines.forEach(l=>{
      let col=null;
      if(l.キャンセル) col=cfg.色_グレー;       // 個数0=キャンセル
      else if(l.kbn==='即納') col=cfg.色_水;
      else if(l.kbn==='指定なし') col=cfg.色_橙;
      else if(l.kbn==='取り寄せ'){
        col=引当行状態_(l, cfg, 入荷消費OK_).color;
        // else: 在庫待ち=白(未入金は受注番号だけ赤)
      }
      rowColor[l.i]=col;
      rowUnpaid[l.i]=!paidOrder[l.ban] && !codOrder[l.ban];
      rowCod[l.i]=!paidOrder[l.ban] && codOrder[l.ban];
      rowQty[l.i]=l.qty;
    });
    const bg=[];
    for(let r=recvStart;r<=recvLast;r++){
      const idx=r-1, c=rowColor[idx];
      const arr=new Array(nc).fill(c!==undefined?c:null);
      if(rowUnpaid[idx]) arr[M.番号]=cfg.色_赤;                 // 未入金は受注番号セルを赤
      else if(rowCod[idx]) arr[M.番号]=cfg.色_代引;             // 代引きは黄(入金前でも発送対象)
      if(rowQty[idx]>=2 && M.個数>=0) arr[M.個数]=cfg.色_緑;    // 個数2以上は個数セル緑
      bg.push(arr);
    }
    const dataRng=recv.getRange(recvStart,1,bg.length,nc);
    dataRng.setBackgrounds(bg);
    dataRng.setFontSize(cfg.字).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    recv.setRowHeights(recvStart, bg.length, cfg.行高);
    注文罫線_(recv, recvStart, M.番号);
  }

  // 突合せ検算: 現供給キーごとに 供給＝取り置き中＋発送済み＋戻し未処理＋在庫なし確定＋Yahoo移動済み＋余り
  // (全件検算・⑤便締めと同じ台帳式。検証済み計画なら超過0のはずで、超過>0は要確認として⑤をブロックする)
  const 突合={供給:0,取り置き中:0,発送済み:0,戻し未処理:0,在庫なし確定:0,Yahoo移動済み:0,余り:0}, 突合超過=[];
  {
    const qtyByKey={};
    supplies.forEach(s=>{ const key=取り置き_供給キー_(s.ems,s.sourceCode||s.code); qtyByKey[key]=(qtyByKey[key]||0)+(Number(s.qty)||0); });
    Object.keys(qtyByKey).forEach(key=>{
      const u=projectedSummary.usageBySupply[key]||{取り置き中:0,発送済み:0,戻し未処理:0,在庫なし確定:0,Yahoo移動済み:0};
      const supplied=qtyByKey[key], used=取り置き_使用合計_(u);
      突合.供給+=supplied; 突合.取り置き中+=u.取り置き中||0; 突合.発送済み+=u.発送済み||0;
      突合.戻し未処理+=u.戻し未処理||0; 突合.在庫なし確定+=u.在庫なし確定||0; 突合.Yahoo移動済み+=u.Yahoo移動済み||0;
      突合.余り+=supplied-used;
      if(used>supplied) 突合超過.push(key.replace('|',' ')+' 使用'+used+'／供給'+supplied);
    });
  }
  const 突合式='供給'+突合.供給+'＝取り置き中'+突合.取り置き中+'＋発送済み'+突合.発送済み+'＋戻し未処理'+突合.戻し未処理
    +'＋在庫なし確定'+突合.在庫なし確定+'＋Yahoo移動済み'+突合.Yahoo移動済み+'＋余り'+突合.余り;
  const jpRows=outputPlan.surplus.map(r=>['到着済',r.arrival,r.code,r.qty,r.ems]);

  // 今回EMS消費者は、同じallocationPlanから作った台帳投影の元EMSで表示する。
  {
    const EHDR=['状態','到着日','商品コード','数量','EMS番号','余り','引き当てた受注番号（左=古い → 右=新しい）'];
    const rows=outputPlan.supplies.map(s=>{
      const consumers=s.consumers.slice().sort((a,b)=>番号num_(a.ban)-番号num_(b.ban));
      const current=consumers.some(c=>c.current), color=current?(s.surplus>0?cfg.色_緑:cfg.色_黄):(consumers.length?cfg.色_着:null);
      const vals=['到着済',s.arrival,s.code,s.qty,s.ems,s.surplus>0?s.surplus:'']
        .concat(consumers.map(c=>c.ban+(c.qty>1?'*'+c.qty:'')));
      const bg=[color,color,color,color,color,color]
        .concat(consumers.map(c=>(paidOrder[c.ban]||codOrder[c.ban])?(c.current?cfg.色_黄:cfg.色_着):cfg.色_赤));
      return {vals,bg};
    });
    const maxCols=rows.reduce((m,r)=>Math.max(m,r.vals.length),EHDR.length);
    let sh=ss.getSheetByName(cfg.日本在庫),fresh=!sh; if(!sh) sh=ss.insertSheet(cfg.日本在庫);
    シート値クリア_(sh);
    sh.getRange(1,1).setValue('最終引当: '+Utilities.formatDate(now,'Asia/Tokyo','yyyy/MM/dd HH:mm:ss')+' / '+rows.length+'件 ｜ '
      +(突合超過.length? '⚠️超過消費'+突合超過.length+'件' : '整合OK『'+突合式+'』'));
    const header=EHDR.slice(); while(header.length<maxCols) header.push('');
    sh.getRange(2,1,1,maxCols).setValues([header]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff').setFontSize(cfg.字);
    sh.setFrozenRows(2); if(fresh) sh.setRowHeight(2,cfg.行高);
    if(rows.length){
      const values=rows.map(r=>{ const v=r.vals.slice(); while(v.length<maxCols) v.push(''); return v; });
      const colors=rows.map(r=>{ const v=r.bg.slice(); while(v.length<maxCols) v.push(null); return v; });
      const range=sh.getRange(3,1,rows.length,maxCols); range.setValues(values).setFontSize(cfg.字); range.setBackgrounds(colors);
      sh.getRange(3,2,rows.length,1).setNumberFormat('yyyy-mm-dd');
      if(fresh){ sh.setRowHeights(3,rows.length,cfg.行高); [90,90,210,70,150,60].forEach((w,c)=>sh.setColumnWidth(c+1,w)); }
    }else sh.getRange(3,1).setValue('(EMS在庫なし)');
  }

  // 日本在庫はallocationPlan.surplusをそのまま出力する。
  {
    const JHDR=['状態','到着日','商品コード','余り数(日本在庫)','EMS番号'];
    let sh=ss.getSheetByName(cfg.純在庫),fresh=!sh; if(!sh) sh=ss.insertSheet(cfg.純在庫);
    シート値クリア_(sh);
    sh.getRange(1,1).setValue('最終引当: '+Utilities.formatDate(now,'Asia/Tokyo','yyyy/MM/dd HH:mm:ss')+' / '+jpRows.length+'件');
    sh.getRange(2,1,1,JHDR.length).setValues([JHDR]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff').setFontSize(cfg.字);
    sh.setFrozenRows(2); if(fresh) sh.setRowHeight(2,cfg.行高);
    if(jpRows.length){ sh.getRange(3,1,jpRows.length,JHDR.length).setValues(jpRows).setFontSize(cfg.字); sh.getRange(3,2,jpRows.length,1).setNumberFormat('yyyy-mm-dd');
      if(fresh){ sh.setRowHeights(3,jpRows.length,cfg.行高); [90,90,210,120,150].forEach((w,c)=>sh.setColumnWidth(c+1,w)); }
    }else sh.getRange(3,1).setValue('(日本在庫なし)');
  }

  const ord=seq.length;
  const ship=shipRows.length? new Set(shipRows.map(r=>r.ban)).size:0;
  const keep=keepRows.length? new Set(keepRows.map(r=>r.ban)).size:0;
  const hold=holdRows.length? new Set(holdRows.map(r=>r.ban)).size:0;
  const part=partRows.length? new Set(partRows.map(r=>r.ban)).size:0;
  const mp=Object.keys(paidOrder).filter(b=>!paidOrder[b] && !codOrder[b]).length; // 代引きは入金待ちに数えない
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
        if(l.台帳入荷日 && 残必要計算_(l)===0){
          const d=ymd_(l.台帳入荷日);
          if(/^20\d{2}-\d{2}-\d{2}$/.test(d)){ 入荷col[idx][0]=new Date(d+'T00:00:00'); 自動入荷++; }
          return;
        }
        // 旧履歴ベースのスタンプ(履歴成立)は台帳一本化で廃止(2026-07-21 遺物掃除。代入箇所も既に無い)
        if(l.引当成立 && l.matchedKey && code到着[l.matchedKey]!==undefined){
          入荷col[idx][0]=code到着[l.matchedKey]; 自動入荷++;
        }
      });
      // 【表示ゆれの正規化】手入力の「26/07/09(木)」等の文字列日付を日付値に直し、列全体を yyyy-MM-dd 表示に統一する
      // (自動記入は日付値・手入力は文字列で形がバラバラになるため。日付らしい文字列だけ触り、読めないメモ等はそのまま)
      for(let idx=0; idx<入荷col.length; idx++){
        const v=入荷col[idx][0];
        if(v==='' || v==null) continue;
        const corrected=入荷日シート値補正_(v);
        if(corrected){ 入荷col[idx][0]=corrected; 入荷表示統一++; continue; }
        if(v instanceof Date) continue;
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
      const byIdx={}; lines.forEach(l=>byIdx[l.i-受注hdr]=l);
      for(let idx=0;idx<col.length;idx++) col[idx][0]=EMS番号書戻し値_(col[idx][0], byIdx[idx]);
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
    const ui=SpreadsheetApp.getUi(),処理ms=Date.now()-開始ms,処理秒=Math.round(処理ms/1000);
    console.log('④引当実行 処理時間ms='+処理ms);
    // 台帳・P列・受注明細・全出力が成功した最後にだけ整合状態を確定する。
    // 突合せ超過と台帳外出荷の不一致(要確認>0)が残っている間は⑤便締めがブロックされる。
    PropertiesService.getDocumentProperties().setProperty('引当_整合状態',JSON.stringify({ts:Date.now(),要確認:突合超過.length+出荷自動.review.length,台帳版:'v1'}));
    // 取り置き登録も新しい台帳から作り直す(②のたびに手動で📋更新を押さなくて済むように 2026-07-23)。
    // 失敗しても②本体の完了は妨げない。直列_の再入不可のため本体_を直接呼ぶ
    try{ if(typeof 取り置き初期登録を作成本体_==='function') 取り置き初期登録を作成本体_({silent:true}); }
    catch(e){ ss.toast('取り置き登録の自動更新に失敗: '+e.message,'⚠️',6); }
    if(!options.silentSummary) ui.alert(突合超過.length? '⚠️ 引当完了（要確認 '+突合超過.length+'件）'
        : 割当警告.length? '✅ 引当完了（タグ注意 '+割当警告.length+'件）' : '✅ 引当完了（整合OK）',
      '■ 処理時間\n'+処理秒+'秒\n\n■ 突合せ\n'
      +(突合超過.length? '⚠️ 箱の供給を超えた消費 '+突合超過.length+'件（⑤便締めはブロックされます）\n'+突合超過.slice(0,10).join('\n')+(突合超過.length>10?'\n…他'+(突合超過.length-10)+'件':'')
        : '✅ 整合OK『'+突合式+'』')
      +(割当警告.length? '\n\n■ ⚠️ 注文番号タグの注意 '+割当警告.length+'件\n'+割当警告.slice(0,10).join('\n')+(割当警告.length>10?'\n…他'+(割当警告.length-10)+'件':''):'')
      +'\n\n■ 注文の分類\n出荷可能 '+ship+' ／ 出荷GO未入金 '+keep+' ／ 希望日待 '+hold+' ／ 部分在庫 '+part+' ／ 引当待ち '+(ord-ship-keep-hold-part)+'（うち入金待ち '+mp+'）'
      +'\n\n■ 処理内容\n取り置き台帳新規 '+確定行数+'行 ／ 入荷日自動記入 '+自動入荷+'件'
      +(出荷自動.newRows.length? ' ／ 📦台帳外出荷の自動登録 '+出荷自動.newRows.length+'行':'')
      +(emsD.除外? ' ／ 🚫棚卸・EMS番号なしを供給除外 '+emsD.除外+'行':'')
      +(隠れ?'\n\n⚠️ '+隠れ+'：行が隠れています（🔻フィルタ確認/解除で全解除）':''),ui.ButtonSet.OK);
  }
  return {success:true,allocationPlan,pPlan,outputPlan,blockedSkus:Array.from(rebuildBlocked)};
}

// EMS在庫の引当色分け: 在庫から まず入荷日あり(割当済)を差引き、残りを入荷日なし(未着)の人へ引当
// 黄=その行まるごと未着の人へ引当 / 緑=一部 / 白=入荷済が取った分 or 余り
function EMS消込色付け_(注文番号指定行){
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
  const direct=注文番号指定行 instanceof Set?注文番号指定行:new Set();
  const emsBg=E.map((row,index)=>{
    const c=normCode_(row[EC.コード]), qty=Number(row[EC.数量])||0;
    let col=null;
    if(direct.has(index) && qty>0) return new Array(ncolE).fill(cfg.色_黄);
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
// QUERY/IMPORTRANGE の再計算がまだ終わっていないか。表示値のどこかに "Loading" が残っていれば未完了(2026-07-20)。
// 先頭セルだけ見る旧判定では、先頭が読めてもJMEE167等の後続行/列がまだLoadingで②が不完全な供給を読む事故が起きた。
function EMS在庫_読込中_(displayValues){
  return (displayValues||[]).some(row=>(row||[]).some(cell=>String(cell==null?'':cell).indexOf('Loading')>=0));
}

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
    // IMPORTRANGE/QUERYの再計算完了を待つ(③→④と連続クリックしても④が空/古い在庫を読まないように)。
    // 先頭セルだけでなくデータ範囲全体に "Loading" が残っていないかを見る(2026-07-20 ラグ由来の中止対策)。
    SpreadsheetApp.flush();
    for(let i=0;i<45;i++){
      const head=String(fcell.getDisplayValue()||''); // '#N/A'(0件)もLoadingでないので確定扱い
      if(head && head.indexOf('Loading')<0){
        const last=emv.getLastRow();
        const vals=last>=ems.offset ? emv.getRange(ems.offset,1,last-ems.offset+1,mc).getDisplayValues() : [[head]];
        if(!EMS在庫_読込中_(vals)) break; // 範囲全体が読み込み終わった
      }
      Utilities.sleep(1000); SpreadsheetApp.flush();
    }
  }
  // 到着日列はQUERYがシリアル数値(46226等)で返すことがあり生数字表示になる。
  // 計算はymd_が数値も解釈するので実害はないが、表示を日付書式に整える(2026-07-21)
  if(ems.cols && ems.cols.到着>=0){
    const last=emv.getLastRow();
    if(last>=ems.offset) emv.getRange(ems.offset,ems.cols.到着+1,last-ems.offset+1,1).setNumberFormat('yyyy-mm-dd');
  }
  ss.toast('EMS在庫を更新しました(色クリア＋最新化)','🔄EMS更新',6);
}

// 🔁 更新してから引当(2026-07-20): EMS在庫の堅牢更新→即納チェック→引当実行を1ボタンで順に実行する。
// 「トーストを待って②を押したのに供給が古かった」ラグ事故を根本回避する(EMS更新の待ちが範囲全体で完了する)。
// 同一の直列_ロック内で _本体_ を順に呼ぶ(EMS更新が確実に読み込み終わってから引当が供給を読む)。
function 更新してから引当(){ 直列_(更新してから引当_本体_); }
function 更新してから引当_本体_(){
  EMS在庫を更新_本体_();      // ② EMS在庫を確実に最新化(範囲全体が読み込み終わるまで待つ)
  前段階チェック_即納_本体_(); // ③ 即納チェック(水色＋罫線)
  引当実行_本体_();           // ④ 引き当て実行(最新の供給で)
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
      paid:入金済み_(row[M.入金]), 代引き:M.支払>=0 && 代引き支払_(row[M.支払]), キャンセル:qty<=0, dAlloc:0, d成立:false });
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
  const codOrder={}; Object.keys(byOrder).forEach(b=> codOrder[b]=byOrder[b].some(l=>l.代引き)); // 代引き=入金前でも発送対象
  const 完成_=arr=> arr.every(l=> l.キャンセル || l.kbn==='即納' || (l.kbn==='取り寄せ' && (l.入荷 || l.d成立)) );

  // ダニエルで1つでも引き当たった注文だけを「ダニエル出荷可能」へ。状態列に発送可否(発送可能/入金待ち/希望日待ち/部分)
  const HDR=['受注番号','氏名','お届け日','商品コード','商品名','個数','区分','入荷日','入金','状態'];
  const seen=new Set(), seq=[]; lines.forEach(l=>{ if(!seen.has(l.ban)){ seen.add(l.ban); seq.push(l.ban);} });
  const rows=[];
  seq.forEach(ban=>{
    const arr=byOrder[ban];
    if(!arr.some(l=>l.d成立)) return; // ダニエルで1つも引き当たってない注文は出さない
    const paid=paidOrder[ban], 代引=codOrder[ban], comp=完成_(arr);
    let 可否; if(!comp) 可否='部分'; else if(希望日未来_(arr[0].届)) 可否='希望日待ち'; else 可否= (paid||代引)? '発送可能' : '入金待ち';
    arr.forEach(l=>{
      let st, color;
      if(l.キャンセル){ st='キャンセル'; color=cfg.色_グレー; }
      else if(l.kbn==='即納'){ st='即納'; color=cfg.色_水; }
      else if(l.kbn==='指定なし'){ st='要確認'; color=cfg.色_橙; }
      else if(l.入荷){ st='着済'; color=cfg.色_着; }
      else if(l.d成立){ st='ダニエル引当'; color=cfg.色_黄; } // ダニエルEMSが当たった=今回出せる
      else { st='在庫待ち'; color=null; }                      // ダニエルにも無い=白
      rows.push({ vals:[ban,l.氏名,l.届,l.code,l.商品名,l.qty,l.kbn,l.入荷日値,paid?'済':(代引?'代引':'未'), st+'／'+可否], color, ban, paid, qty:l.qty });
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
    const dHit={}; lines.forEach(l=>{ if(l.d成立) dHit[l.i]={paid:paidOrder[l.ban]||codOrder[l.ban]}; }); // 代引きは未入金マークにしない
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
  try{ ダニエル余りを計算(); }catch(e){} // 引当が変わったので余り推定を更新(失敗しても引当は成立)
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
function シート消去範囲_(lastRow,lastCol){
  return {
    rows:Math.max(1,Number(lastRow)||0),
    cols:Math.max(1,Number(lastCol)||0)
  };
}

function 列記号_(n){
  let s='';
  for(let x=Number(n)||0;x>0;x=Math.floor((x-1)/26)){
    s=String.fromCharCode(65+(x-1)%26)+s;
  }
  return s;
}

function 注文境界A1_(rows,startRow,ncol){
  const end=列記号_(ncol), out=[];
  for(let i=0;i<(rows||[]).length;i++){
    if(i===rows.length-1 || rows[i+1].ban!==rows[i].ban){
      const row=Number(startRow)+i;
      out.push('A'+row+':'+end+row);
    }
  }
  return out;
}

function シート値クリア_(sh){
  const size=シート消去範囲_(sh.getLastRow(),sh.getLastColumn());
  const range=sh.getRange(1,1,size.rows,size.cols);
  range.clearContent();
  range.setBackground(null);
  range.setBorder(false,false,false,false,false,false);
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

  const borders=注文境界A1_(rows,dr,ncol);
  if(borders.length) sh.getRangeList(borders).setBorder(
    null,null,true,null,null,null,'#000000',SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  if(新規) [150,110,100,150,360,55,80,70,55,95,150,120].forEach((w,c)=> { if(c<ncol) sh.setColumnWidth(c+1,w); });
}

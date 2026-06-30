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
  // xlsxの列(0始): C=発送日 / L=商品コード / M=数量 / Q=対象BOX / R=EMS番号
  X_発送日: 2, X_商品コード: 11, X_数量: 12, X_BOX: 16, X_EMS: 17
};

// スプレッドを開いたときにメニューを出す
function onOpen(){
  const ui=SpreadsheetApp.getUi();
  ui.createMenu('📦 受注引当')
    // ── EMS在庫(役割:在庫の準備) ──
    .addItem('🔄 EMS在庫を更新(色クリア＋最新化)', 'EMS在庫を更新')
    .addSeparator()
    // ── 受注明細から引き当て(役割:引当) ──
    .addItem('📥 最新CSVを取込', '取込_最新CSV')
    .addItem('📦 ダニエルEMS取込', '取込_ダニエルEMS')
    .addItem('① 前段階チェック(即納に水色＋罫線)', '前段階チェック_即納')
    .addItem('② 引き当て実行(入荷日で判定)', '引当実行')
    .addItem('🔍 在庫照合レポート', '在庫照合レポート')
    .addItem('🙈 色なし注文を除外/全表示(切替)', '色なし注文の除外切替')
    .addSeparator()
    // ── データを消す(役割:リセット) ──
    .addSubMenu(ui.createMenu('🗑 データを消す')
      .addItem('受注明細', '受注明細を消す')
      .addItem('引当待ち', '引当待ちを消す')
      .addItem('部分在庫', '部分在庫を消す')
      .addItem('出荷GO未入金', '出荷GO未入金を消す')
      .addItem('希望日待ち', '希望日待ちを消す')
      .addItem('出荷可能', '出荷可能を消す')
      .addItem('日本在庫', '日本在庫を消す')
      .addItem('照合レポート', '照合レポートを消す')
      .addSeparator()
      .addItem('全部まとめて', '全データを消す'))
    .addToUi();
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

// 📦 ダニエルEMS取込: フォルダ内の最新xlsxを直接読み→DANIEL BOX行だけ整形してタブへ
function 取込_ダニエルEMS(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi(), cfg=DANIEL_CFG;

  let folder;
  try{ folder=DriveApp.getFolderById(cfg.フォルダID); }
  catch(e){ ui.alert('「ダニエルEMSファイル」フォルダが見つからへん\n'+cfg.フォルダID); return; }
  const it=folder.getFiles(); let latest=null;
  while(it.hasNext()){ const f=it.next(); if(!latest || f.getLastUpdated()>latest.getLastUpdated()) latest=f; }
  if(!latest){ ui.alert('フォルダにファイルが無いで'); return; }

  const upd=Utilities.formatDate(latest.getLastUpdated(),'Asia/Tokyo','MM/dd HH:mm');
  const ans=ui.alert('ダニエルEMS取込','ファイル：'+latest.getName()+'\n更新：'+upd+'\n\nDANIEL BOXの行を取り込みます。ええ？',ui.ButtonSet.OK_CANCEL);
  if(ans!==ui.Button.OK) return;

  // 最新がGoogleシートならそのまま読む。xlsxはZIPとして直接展開して読む(Drive変換を使わない=サイズ制限なし)
  let data=null;
  try{
    if(latest.getMimeType()===MimeType.GOOGLE_SHEETS){
      data=SpreadsheetApp.openById(latest.getId()).getSheets()[0].getDataRange().getValues();
    } else {
      data=xlsx読む_(latest.getBlob());
    }
  }catch(e){
    ui.alert('読み込みに失敗：'+e.message); return;
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
    rows.push([boxNo[key], a.発送日, a.code, a.qty, a.box, a.ems, '']); // BOXNo/発送日/商品コード/数量/対象BOX/EMS番号/入荷後ステータス
  });

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
    sh.getRange(startRow,1,rows.length,7).setValues(rows)
      .setVerticalAlignment('middle').setFontSize(HIKIATE_CFG.字).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    sh.setRowHeights(startRow, rows.length, HIKIATE_CFG.行高);
  }
  ss.toast('ダニエルEMS取込完了：'+rows.length+'件 / 発送日'+dates.join(',')+'（'+latest.getName()+'）','📦ダニエルEMS',6);
}

// EMS在庫の生データを返す(1行ヘッダーがあれば除外)。offset=データ先頭のシート行番号(1始)
function EMS明細_(emv){
  const all=emv.getDataRange().getValues();
  let s=0;
  if(all.length){ const r=all[0].map(v=>String(v||'').trim());
    if(r.indexOf('商品コード')>=0 || r.indexOf('状態')>=0 || r.indexOf('EMS番号')>=0) s=1; } // ヘッダー行を検出
  return { rows: all.slice(s), offset: s+1 };
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

  // 書式・ヘッダーは残し、ヘッダー行の下のデータだけ入れ替える
  let sh=ss.getSheetByName(cfg.受注シート); if(!sh) sh=ss.insertSheet(cfg.受注シート);
  const hdrRow=受注ヘッダー行_(sh);       // 例:6行目(無ければ1)
  const startRow=hdrRow+1;               // データ開始行
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
  }

  ss.toast('取込完了：'+latest.getName()+' / 受注'+body.length+'行（更新 '+upd+'）','GoQ取込',6);
}

// 注文ごと(受注番号が変わる境目)に太い下線を引く。全体には薄い罫線
function 注文罫線_(sh, startRow, rows, ncol, banCol){
  if(!rows.length) return;
  sh.getRange(startRow,1,rows.length,ncol)
    .setBorder(true,true,true,true,true,true,'#cccccc',SpreadsheetApp.BorderStyle.SOLID);
  for(let i=0;i<rows.length;i++){
    const cur=String(rows[i][banCol]||'').trim();
    const nxt=(i+1<rows.length)?String(rows[i+1][banCol]||'').trim():null;
    if(i===rows.length-1 || nxt!==cur)
      sh.getRange(startRow+i,1,1,ncol)
        .setBorder(null,null,true,null,null,null,'#000000',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
}

// ①前段階チェック: 受注明細で「即納」行に水色＋注文ごとの太線(引当前の目視確認用)
function 前段階チェック_即納(){
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

  注文罫線_(recv, startRow, R, lastCol, M.番号);
  ss.toast('前段階チェック完了：即納'+n即納+'行を水色＋注文ごとに罫線','①前段階',6);
}

// ===== 設定 =====
const HIKIATE_CFG = {
  受注: '受注明細', EMS在庫: 'EMS在庫', 待ち: '引当待ち', 部分: '部分在庫', 取置: '出荷GO未入金', 希望: '希望日待ち', 出荷: '出荷可能', 日本在庫: '今回入荷EMSの在庫', 純在庫: '日本在庫',
  行高: 27, 字: 13, // 表の行の高さ・フォントサイズ

  // 受注明細の列は固定番号ではなく見出し名で特定する(列マップ_を参照)。並び替え・入荷日追加に強い。

  // EMS在庫タブの列(0始): QUERYで「到着済」だけ A状態 B商品コード C数量 D EMS番号
  E_コード: 1, E_数量: 2,

  色_緑:'#b7e1cd', 色_黄:'#fce8b2', 色_グレー:'#efefef', 色_橙:'#fcd5b4', 色_赤:'#f4cccc', 色_水:'#cfe2f3', 色_着:'#d9d2e9'
};

function normCode_(v){ return String(v||'').trim().toUpperCase().replace(/_/g,'-'); }
function codeKeys_(code){
  const c=normCode_(code); const keys=[c];
  const m=c.match(/^(.+)-(\d{2})(\d{2})-(\d{2})$/);
  if(m && (m[4]===m[2]||m[4]===m[3])) keys.push(m[1]+'-'+m[4]);
  return keys;
}
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
// 入荷日が今日(=②引当を実行する日)か。今日の着済を「今回分」として扱うため
function 入荷日今日_(v){
  const s=String(v||'').trim(); if(!s) return false;
  const d=new Date(s); if(isNaN(d.getTime())) return false;
  const t=new Date();
  return d.getFullYear()===t.getFullYear() && d.getMonth()===t.getMonth() && d.getDate()===t.getDate();
}

function 引当実行(){
  const ss=SpreadsheetApp.getActive(), cfg=HIKIATE_CFG;
  const recv=ss.getSheetByName(cfg.受注);
  if(!recv){ SpreadsheetApp.getUi().alert('「'+cfg.受注+'」タブが無いで'); return; }
  const R=recv.getDataRange().getValues();

  // 受注明細→行オブジェクト(列は見出し名で特定。並び替え・列追加に強い)
  // 着済(着いたか)の判定: 即納 or (取り寄せ かつ 入荷日あり)。取り寄せで入荷日なし=未着。
  const M=列マップ_(recv), 受注hdr=M.hr;
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
      paid:入金済み_(row[M.入金]), 入荷, 入荷日値, 着, alloc:0, 引当成立:false, キャンセル:qty<=0 });
  }

  // ===== EMS在庫引当: 入荷日あり(割当済)を在庫から差引き、残りを未着の取り寄せに古い注文順でFIFO引当 =====
  const emv=ss.getSheetByName(cfg.EMS在庫);
  const E= emv? EMS明細_(emv).rows : []; // 1行ヘッダーがあれば除外
  const stock={};
  for(let i=0;i<E.length;i++){ const c=normCode_(E[i][cfg.E_コード]); if(!c) continue; stock[c]=(stock[c]||0)+(Number(E[i][cfg.E_数量])||0); }
  const origStock=Object.assign({},stock); // 余り(今回入荷EMSの在庫=日本在庫)算出用
  const aliasMap={}; Object.keys(stock).forEach(c=> codeKeys_(c).forEach(k=>{ if(!(k in aliasMap)) aliasMap[k]=c; }));
  const skuStrip_=s=>String(s||'').replace(/[A-Za-z]$/,'');
  const candKeys=l=>{ const cands=[]; if(l.sku){cands.push(skuStrip_(l.sku));cands.push(l.sku);} if(l.code)cands.push(l.code);
    const keys=[]; cands.forEach(v=> codeKeys_(v).forEach(k=>{ if(keys.indexOf(k)<0) keys.push(k); })); return keys; };
  const keyInStock=l=>{ for(const k of candKeys(l)){ if(k in stock) return k; if(aliasMap[k] && aliasMap[k] in stock) return aliasMap[k]; } return null; };
  const findAvail=l=>{ for(const k of candKeys(l)){ if(stock[k]>0) return k; if(aliasMap[k]&&stock[aliasMap[k]]>0) return aliasMap[k]; } return null; };
  // 入荷日あり(もう割当済)を在庫から先に差し引く
  lines.filter(l=>l.kbn==='取り寄せ' && l.入荷).forEach(l=>{ const k=keyInStock(l); if(k!=null) stock[k]=Math.max(0,(stock[k]||0)-l.qty); });
  // 残りを未着の取り寄せに引き当て(古い注文順)
  lines.filter(l=>l.kbn==='取り寄せ' && !l.入荷).sort((a,b)=> a.sortKey-b.sortKey || a.i-b.i).forEach(l=>{
    const k=findAvail(l); if(k){ const got=Math.min(l.qty, stock[k]); stock[k]-=got; l.alloc=got; }
    l.引当成立 = l.alloc>=l.qty && l.qty>0;
  });

  // 注文ごと: 入金済み(1行でも入金日があれば入金済み。注文単位入金)
  const byOrder={};
  lines.forEach(l=> (byOrder[l.ban]=byOrder[l.ban]||[]).push(l));
  const paidOrder={};
  Object.keys(byOrder).forEach(ban=>{ paidOrder[ban]=byOrder[ban].some(l=>l.paid); });

  // 注文を4分類(今回EMSで引き当たった=黄 がある注文だけを上3シートへ動かす):
  //  出荷可能=黄あり&完成&入金済 / 出荷GO未入金=黄あり&完成&未入金 / 部分在庫=黄ありだが未完成 / 引当待ち=黄なし
  // 今回分=今回EMSで引き当たった(黄) or 入荷日が今日(=②実行日)の着済。先に入荷日を入れた当日分も出荷可能に出すための例外。
  const 今回行_=l=> l.引当成立 || (l.入荷 && 入荷日今日_(l.入荷日値));
  const 区分け=ban=>{
    const arr=byOrder[ban];
    if(!arr.some(今回行_)) return 'wait';                      // 今回分(黄 or 入荷日=今日)が無い→引当待ち
    const 完成=arr.every(l=> l.キャンセル || l.kbn==='即納' || (l.kbn==='取り寄せ' && (l.入荷 || l.引当成立)) );
    if(!完成) return 'part';                                   // 一部だけ揃う→部分在庫
    if(!paidOrder[ban]) return 'keep';                         // 完成・未入金→出荷GO未入金
    if(希望日未来_(arr[0].届)) return 'hold';                  // 完成・入金済だがお届け日指定が未来→希望日待ち
    return 'ship';                                            // 完成・入金済・今出せる→出荷可能
  };
  // 引当待ちの「発送可否」(K列): 揃っているのに今回分でない=実は発送できる注文を見分ける
  const 発送可否_=ban=>{
    const arr=byOrder[ban];
    const 完成=arr.every(l=> l.キャンセル || l.kbn==='即納' || (l.kbn==='取り寄せ' && (l.入荷 || l.引当成立)) );
    if(!完成) return '';                                       // まだ揃ってない=発送不可(空欄)
    if(希望日未来_(arr[0].届)) return '発送可能希望日待ち';
    return paidOrder[ban]? '発送可能' : '発送可能入金待ち';
  };

  // 出力(行の色=到着状況。未入金は受注番号セルだけ赤)
  const HDR=['受注番号','氏名','お届け日','商品コード','商品名','個数','区分','入荷日','入金','状態'];
  const HDR_待=HDR.concat(['発送可否']); // 引当待ちだけK列に発送可否を足す
  const seen=new Set(), seq=[]; lines.forEach(l=>{ if(!seen.has(l.ban)){ seen.add(l.ban); seq.push(l.ban);} });
  const waitRows=[], partRows=[], keepRows=[], holdRows=[], shipRows=[];
  seq.forEach(ban=>{
    const b=区分け(ban), paid=paidOrder[ban];
    const target = b==='ship'? shipRows : b==='keep'? keepRows : b==='hold'? holdRows : b==='part'? partRows : waitRows;
    const 可否 = (b==='wait')? 発送可否_(ban) : null; // 引当待ちのみK列に発送可否
    byOrder[ban].forEach(l=>{
      let st, color;
      if(l.キャンセル){ st='キャンセル'; color=cfg.色_グレー; } // 個数0=キャンセル(完成判定はブロックしない)
      else if(l.kbn==='即納'){ st='即納'; color=cfg.色_水; }
      else if(l.kbn==='指定なし'){ st='要確認'; color=cfg.色_橙; }
      else if(l.入荷){ st='着済'; color=cfg.色_着; }        // 取り寄せ+入荷日=着済(処理済)=ラベンダー
      else if(l.引当成立){ st='引当(今回)'; color=cfg.色_黄; }  // 未着だがEMS在庫が当たった=今回出せる
      else { st='在庫待ち'; color=null; }                       // 在庫なし=白
      if(!l.キャンセル && !paid) st+='・入金待ち';
      const vals=[ban,l.氏名,l.届,l.code,l.商品名,l.qty,l.kbn,l.入荷日値,paid?'済':'未',st];
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
        if(l.入荷) col=cfg.色_着;       // 着済(入荷日あり・処理済)=ラベンダー
        else if(l.引当成立) col=cfg.色_黄;  // 今回EMS在庫が当たった=今回出せる
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
    注文罫線_(recv, recvStart, R.slice(recvStart-1), nc, M.番号);
  }

  // EMS在庫タブも色分け(黄=未着へ全引当 / 緑=一部 / 白=入荷済 or 余り)
  EMS消込色付け_();

  const jpRows=[]; // 純日本在庫(行ごとの余り。EMS番号付き)
  // ===== 今回入荷EMSの在庫(台帳・色付き): 各EMS行に 状態と引き当てた受注番号 を出力 =====
  {
    // コードごとの消費者キュー(割当済=入荷日あり → 引当=今回 の順)
    const consumersByCode={};
    const pushCons=(l, qty, kind)=>{ const k=keyInStock(l); if(k==null||qty<=0) return; (consumersByCode[k]=consumersByCode[k]||[]).push({qty, ban:l.ban, kind}); };
    lines.filter(l=>l.kbn==='取り寄せ' && l.入荷 && l.qty>0).forEach(l=> pushCons(l, l.qty, '割当済'));   // 入荷日あり(もう割当済)
    lines.filter(l=>l.kbn==='取り寄せ' && !l.入荷 && l.alloc>0).sort((a,b)=> a.sortKey-b.sortKey || a.i-b.i).forEach(l=> pushCons(l, l.alloc, '引当')); // 今回引き当て
    const ptr={};
    const consumeRow=(code, qty)=>{
      const q=consumersByCode[code]||[]; const st=ptr[code]||{i:0,used:0};
      const got={割当済:[], 引当:[]}; let left=qty;
      while(left>0 && st.i<q.length){
        const e=q[st.i], avail=e.qty-st.used, take=Math.min(left, avail);
        if(take>0){ if(got[e.kind].indexOf(e.ban)<0) got[e.kind].push(e.ban); st.used+=take; left-=take; }
        if(st.used>=e.qty){ st.i++; st.used=0; }
      }
      ptr[code]=st; return {got, surplus:left};
    };
    const EHDR=['状態','商品コード','数量','EMS番号','ステータス変化'];
    const ledger=[], ledgerColor=[];
    E.forEach(row=>{
      const c=normCode_(row[cfg.E_コード]), qty=Number(row[cfg.E_数量])||0;
      if(!c) return;
      let stt='', col=null;
      if(qty>0){
        const r=consumeRow(c, qty), parts=[];
        if(r.got.割当済.length) parts.push('割当済:'+r.got.割当済.join(','));
        if(r.got.引当.length)   parts.push('引当:'+r.got.引当.join(','));
        if(r.surplus>0)         parts.push('余り'+r.surplus+'(日本在庫)');
        stt=parts.join(' / ') || '余り(日本在庫)';
        if(r.got.引当.length) col = r.surplus>0 ? cfg.色_緑 : cfg.色_黄; // 引当+余りあり=緑 / 余りなし全引当=黄
        else if(r.got.割当済.length) col=cfg.色_着;                   // もう割当済=ラベンダー
        // else: 余りのみ → 白(日本在庫)
        if(r.surplus>0) jpRows.push([row[0], c, r.surplus, row[3]]); // 状態/商品コード/余り数/EMS番号
      }
      ledger.push([row[0], c, qty, row[3], stt]); ledgerColor.push(col); // 状態/商品コード/数量/EMS番号/ステータス変化
    });
    let esh=ss.getSheetByName(cfg.日本在庫)||ss.insertSheet(cfg.日本在庫);
    esh.clear();
    esh.getRange(1,1).setValue('最終引当: '+Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm:ss')+' / '+ledger.length+'件');
    esh.getRange(2,1,1,EHDR.length).setValues([EHDR]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff').setFontSize(cfg.字);
    esh.setFrozenRows(2); esh.setRowHeight(2, cfg.行高);
    if(ledger.length){
      const rng=esh.getRange(3,1,ledger.length,EHDR.length);
      rng.setValues(ledger).setFontSize(cfg.字);
      rng.setBackgrounds(ledgerColor.map(c=> new Array(EHDR.length).fill(c))); // 引当=黄/割当済=グレー/余り=白
      esh.setRowHeights(3, ledger.length, cfg.行高);
      [90,210,70,150,420].forEach((w,c)=> esh.setColumnWidth(c+1,w));
    } else {
      esh.getRange(3,1).setValue('(EMS在庫なし)');
    }
  }

  // ===== 純粋な日本在庫(EMSで引き当たらず余った分)を行ごと(EMS番号付き)に出力 =====
  {
    const JHDR=['状態','商品コード','余り数(日本在庫)','EMS番号'];
    let jsh=ss.getSheetByName(cfg.純在庫)||ss.insertSheet(cfg.純在庫);
    jsh.clear();
    jsh.getRange(1,1).setValue('最終引当: '+Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm:ss')+' / '+jpRows.length+'件');
    jsh.getRange(2,1,1,JHDR.length).setValues([JHDR]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff').setFontSize(cfg.字);
    jsh.setFrozenRows(2); jsh.setRowHeight(2, cfg.行高);
    if(jpRows.length){
      jsh.getRange(3,1,jpRows.length,JHDR.length).setValues(jpRows).setFontSize(cfg.字);
      jsh.setRowHeights(3, jpRows.length, cfg.行高);
      [90,210,120,150].forEach((w,c)=> jsh.setColumnWidth(c+1,w));
    } else { jsh.getRange(3,1).setValue('(日本在庫なし)'); }
  }

  const ord=seq.length;
  const ship=shipRows.length? new Set(shipRows.map(r=>r.ban)).size:0;
  const keep=keepRows.length? new Set(keepRows.map(r=>r.ban)).size:0;
  const hold=holdRows.length? new Set(holdRows.map(r=>r.ban)).size:0;
  const part=partRows.length? new Set(partRows.map(r=>r.ban)).size:0;
  const mp=Object.keys(paidOrder).filter(b=>!paidOrder[b]).length;
  SpreadsheetApp.getActive().toast(`引当完了：出荷可能${ship} / 出荷GO未入金${keep} / 希望日待ち${hold} / 部分在庫${part} / 引当待ち${ord-ship-keep-hold-part}（うち入金待ち${mp}）`);
}

// EMS在庫の引当色分け: 在庫から まず入荷日あり(割当済)を差引き、残りを入荷日なし(未着)の人へ引当
// 黄=その行まるごと未着の人へ引当 / 緑=一部 / 白=入荷済が取った分 or 余り
function EMS消込色付け_(){
  const ss=SpreadsheetApp.getActive(), cfg=HIKIATE_CFG;
  const recv=ss.getSheetByName(cfg.受注), emv=ss.getSheetByName(cfg.EMS在庫);
  if(!recv||!emv) return;
  const R=recv.getDataRange().getValues();
  const ems=EMS明細_(emv), E=ems.rows; // 1行ヘッダーがあれば除外
  if(!E.length) return;
  const orig={};
  for(let i=0;i<E.length;i++){ const c=normCode_(E[i][cfg.E_コード]); if(!c) continue; orig[c]=(orig[c]||0)+(Number(E[i][cfg.E_数量])||0); }
  const aliasMap={}; Object.keys(orig).forEach(c=> codeKeys_(c).forEach(k=>{ if(!(k in aliasMap)) aliasMap[k]=c; }));
  const skuStrip_=s=>String(s||'').replace(/[A-Za-z]$/,'');
  const keyOf=l=>{ const cands=[]; if(l.sku){cands.push(skuStrip_(l.sku));cands.push(l.sku);} if(l.code)cands.push(l.code);
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
  // コードごと: 在庫から入荷済の取り分(色なし)を除き、残りを未着へ引当(色付け)
  const skipLeft={}, colorLeft={};
  Object.keys(orig).forEach(k=>{
    const s=orig[k]||0, d済=need済[k]||0, d未=need未[k]||0;
    skipLeft[k]=Math.min(d済, s);
    colorLeft[k]=Math.min(d未, Math.max(0, s-d済));
  });
  const ncolE=E[0].length;
  const emsBg=E.map(row=>{
    const c=normCode_(row[cfg.E_コード]), qty=Number(row[cfg.E_数量])||0;
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
  const R=recv.getDataRange().getValues(), E=emv.getDataRange().getValues();

  // 在庫
  const stock={};
  for(let i=0;i<E.length;i++){ const c=normCode_(E[i][cfg.E_コード]); if(!c) continue; stock[c]=(stock[c]||0)+(Number(E[i][cfg.E_数量])||0); }
  const orig=Object.assign({},stock);
  const aliasMap={};
  Object.keys(stock).forEach(c=> codeKeys_(c).forEach(k=>{ if(!(k in aliasMap)) aliasMap[k]=c; }));
  const skuStrip_=s=>String(s||'').replace(/[A-Za-z]$/,'');
  const candKeys=l=>{ const cands=[]; if(l.sku){cands.push(skuStrip_(l.sku));cands.push(l.sku);} if(l.code)cands.push(l.code);
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
    return '★在庫なし:'+(normCode_(skuStrip_(l.sku))||normCode_(l.code)); };
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
  let sh=ss.getSheetByName('照合レポート')||ss.insertSheet('照合レポート');
  sh.clear();
  sh.getRange(hr,1,1,6).setValues([HDR]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff').setFontSize(cfg.字);
  sh.setFrozenRows(hr);
  sh.setRowHeight(hr, cfg.行高);
  if(rows.length){
    sh.getRange(dr,1,rows.length,6).setValues(rows).setFontSize(cfg.字);
    sh.setRowHeights(dr, rows.length, cfg.行高);
    const bg=rows.map(r=> new Array(6).fill( r[5]>0?cfg.色_赤 : (r[4]>0?cfg.色_黄:null) )); // 不足=赤 / 在庫余り=黄
    sh.getRange(dr,1,rows.length,6).setBackgrounds(bg);
    [220,80,100,80,80,80].forEach((w,c)=> sh.setColumnWidth(c+1,w));
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

  // 色なしだけの注文を隠す(受注番号が赤=未入金の注文は残す)
  recv.showRows(startRow, n); // いったん全表示
  const bg=recv.getRange(startRow,1,n,nc).getBackgrounds();
  const bans=recv.getRange(startRow,1,n,M.番号+1).getValues().map(r=>String(r[M.番号]||'').trim());
  const hasColor={};
  for(let i=0;i<n;i++){ const b=bans[i]; if(!b) continue;
    if(!(b in hasColor)) hasColor[b]=false;
    if(bg[i].some(c=> c && c.toLowerCase()!=='#ffffff')) hasColor[b]=true;
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
  ss.toast('色なしの注文を'+cnt+'行 隠しました(もう一度押すと全表示)','除外',5);
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
function 照合レポートを消す(){ シートデータ消去_('照合レポート'); }
function 全データを消す(){
  const ui=SpreadsheetApp.getUi();
  const ans=ui.alert('全データ消去','受注明細・引当待ち・部分在庫・出荷GO未入金・希望日待ち・出荷可能・照合レポートのデータを全部消します。ええ？',ui.ButtonSet.OK_CANCEL);
  if(ans!==ui.Button.OK) return;
  [HIKIATE_CFG.受注, HIKIATE_CFG.待ち, HIKIATE_CFG.部分, HIKIATE_CFG.取置, HIKIATE_CFG.希望, HIKIATE_CFG.出荷, HIKIATE_CFG.純在庫, '照合レポート'].forEach(n=> シートデータ消去_(n,true));
  SpreadsheetApp.getActive().toast('全シートのデータを消しました','🗑データ消去',5);
}

// 🔄 EMS在庫を更新: 残った色・罫線を全部消して、QUERY(IMPORTRANGE)を再計算させる
function EMS在庫を更新(){
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
  if(f){ fcell.clearContent(); SpreadsheetApp.flush(); fcell.setFormula(f); }
  ss.toast('EMS在庫を更新しました(色クリア＋最新化)','🔄EMS更新',6);
}

function 書き出し_(ss, name, hdr, rows, startRow){
  const hr=startRow||1, dr=hr+1; // hr=見出し行 / dr=データ開始行
  let sh=ss.getSheetByName(name); if(!sh) sh=ss.insertSheet(name);
  sh.clear();
  // 押すたびに変わる印(最終引当の時刻＋行数)をA1に表示
  sh.getRange(1,1).setValue('最終引当: '+Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm:ss')+' / '+rows.length+'行');
  const ncol=hdr.length;
  sh.getRange(hr,1,1,ncol).setValues([hdr])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#4472c4').setHorizontalAlignment('center');
  sh.setFrozenRows(hr);
  sh.getRange(hr,1,1,ncol).setFontSize(HIKIATE_CFG.字);
  sh.setRowHeight(hr, HIKIATE_CFG.行高);
  if(rows.length===0){ sh.getRange(dr,1).setValue('(対象なし)'); return; }

  const rng=sh.getRange(dr,1,rows.length,ncol);
  rng.setValues(rows.map(r=>r.vals));
  rng.setFontSize(HIKIATE_CFG.字);
  sh.setRowHeights(dr, rows.length, HIKIATE_CFG.行高);
  rng.setBackgrounds(rows.map(r=>{ const a=new Array(ncol).fill(r.color);
    if(r.paid===false) a[0]=HIKIATE_CFG.色_赤;          // 未入金は受注番号セル赤
    if(r.qty>=2 && ncol>5) a[5]=HIKIATE_CFG.色_緑;      // 個数2以上は個数セル緑(見間違い防止)
    return a; }));
  rng.setBorder(true,true,true,true,true,true,'#cccccc',SpreadsheetApp.BorderStyle.SOLID);

  for(let i=0;i<rows.length;i++){
    if(i===rows.length-1 || rows[i+1].ban!==rows[i].ban)
      sh.getRange(dr+i,1,1,ncol).setBorder(null,null,true,null,null,null,'#000000',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  [150,110,100,150,360,55,80,70,55,95,150].forEach((w,c)=> { if(c<ncol) sh.setColumnWidth(c+1,w); });
}
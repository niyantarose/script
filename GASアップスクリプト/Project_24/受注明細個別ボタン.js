// ===== 受注明細から直接 個別引当/引当キャンセル =====
// 受注明細で対象の行を選択して、シート上部のボタン(個別引当=緑/引当キャンセル=赤)を押すだけ。
// 照合キーの転記・個別対応シートへの入力は不要(個別対応シート運用は廃止。関数は残置)。
//
// 個別引当: 選択行の商品を発注共有EMSリストの「到着済」の箱から探してP列に名指しを追記。
//           引当履歴に記録し、入荷日が空欄ならEMS到着日を自動記入する。
// キャンセル: EMSリストP列からこの受注番号×この商品の名指しを削除。履歴を更新し、
//           他に有効な割当が無ければ受注明細の入荷日もクリアする。

const UKE_KOBETSU_CFG = {
  ボタン行: 5,      // 5行目(凡例1〜4行目と見出し6行目の間の空き行。既存ボタンと重ならない)
  引当ボタン列: 3,  // C列(ボタン画像のアンカー)
  取消ボタン列: 5,  // E列
  ボタン幅: 150,
  ボタン高さ: 36
};

// ボタン画像(PNG/base64埋め込み)。SheetsのinsertImageはSVG非対応(PNG/JPEG/GIFのみ)なので
// 事前生成したPNG(300x72=2倍解像度)を埋め込み、挿入後に150x36へ縮小表示する。
const UKE_KOBETSU_PNG = {
  個別引当: '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',
  引当キャンセル: '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'
};

function 受注明細個別ボタンを設置(){ 受注明細個別ボタンを設置_(false); }

// 図形ボタン運用に切り替えたときの掃除用: スクリプトが挿入したボタン画像を全部消す
function 受注明細個別ボタン画像を削除(){
  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName(HIKIATE_CFG.受注);
  if(!sh){ SpreadsheetApp.getUi().alert('「'+HIKIATE_CFG.受注+'」タブがありません'); return; }
  let n=0;
  try{
    sh.getImages().forEach(img=>{
      const t=String(img.getAltTextTitle&&img.getAltTextTitle()||'');
      if(t.indexOf('受注個別ボタン_')===0){ img.remove(); n++; }
    });
  }catch(e){}
  ss.toast('ボタン画像を'+n+'個削除しました。図形ボタン(スクリプト割り当て)をお使いください','🧹個別',6);
}

function 受注明細個別ボタンを設置_(silent){
  const ss=SpreadsheetApp.getActive(), cfg=UKE_KOBETSU_CFG;
  const sh=ss.getSheetByName(HIKIATE_CFG.受注);
  if(!sh){ if(!silent) SpreadsheetApp.getUi().alert('「'+HIKIATE_CFG.受注+'」タブがありません'); return 0; }
  try{
    sh.getImages().forEach(img=>{
      const t=String(img.getAltTextTitle&&img.getAltTextTitle()||'');
      if(t.indexOf('受注個別ボタン_')===0) img.remove();
    });
  }catch(e){}
  const add=(col,label,fn)=>{
    const blob=Utilities.newBlob(Utilities.base64Decode(UKE_KOBETSU_PNG[label]),'image/png','受注個別ボタン_'+label+'.png');
    const img=sh.insertImage(blob, col, cfg.ボタン行);
    img.setAltTextTitle('受注個別ボタン_'+label);
    img.setAltTextDescription(label+'を実行(先に対象の行を選択)');
    img.assignScript(fn);
    img.setWidth(cfg.ボタン幅);
    img.setHeight(cfg.ボタン高さ);
  };
  try{
    add(cfg.引当ボタン列,'個別引当','選択行を個別引当');
    add(cfg.取消ボタン列,'引当キャンセル','選択行の引当キャンセル');
  }catch(e){
    if(!silent) SpreadsheetApp.getUi().alert(
      'ボタン画像を設置できませんでした:\n'+e.message+'\n\n'+
      '代わりに手動で設置できます:\n挿入→図形描画でボタンを作り、図形の右上「⋮」→「スクリプトを割り当て」で\n'+
      '「選択行を個別引当」「選択行の引当キャンセル」を割り当ててください。');
    return 0;
  }
  sh.setRowHeight(cfg.ボタン行, Math.max(sh.getRowHeight(cfg.ボタン行), cfg.ボタン高さ+8));
  if(!silent) ss.toast('受注明細に個別ボタンを設置しました','🎯個別',5);
  return 1;
}

// 選択範囲からデータ行(ヘッダーより下)の行番号一覧を取る。画像クリックは選択を動かさないので
// 「行を選択→ボタン」がそのまま成立する。
function 受注個別_選択行_(sh, M){
  const rows=new Set();
  const list=sh.getActiveRangeList? sh.getActiveRangeList() : null;
  const ranges=list? list.getRanges() : (sh.getActiveRange()? [sh.getActiveRange()] : []);
  ranges.forEach(r=>{ for(let i=r.getRow(); i<=r.getLastRow(); i++){ if(i>M.hr) rows.add(i); } });
  return Array.from(rows).sort((a,b)=>a-b);
}

// 純ロジック: EMSリスト行から、この商品に一致する「到着済・残あり」の候補を返す
// (照合は②と同じ: 受注候補コード_→codeKeys_。残数=行数量-P列既存割当)
function 受注個別_候補_(emsRows, sku, code){
  const targetKeys=[];
  受注候補コード_(sku,code).forEach(v=> codeKeys_(v).forEach(k=>{ if(targetKeys.indexOf(k)<0) targetKeys.push(k); }));
  const out=[];
  emsRows.forEach(r=>{
    if(String(r.状態||'').trim()!=='到着済') return;
    if(!codeKeys_(r.商品コード).some(k=>targetKeys.indexOf(k)>=0)) return;
    const used=個別対応_P注文展開_(r.注文番号, r.数量).reduce((s,e)=>s+e.qty,0);
    const 残=Math.max(0,(Number(r.数量)||0)-used);
    if(残>0) out.push({item:r, 残});
  });
  return out;
}

// P列にこの受注番号×この商品の割当がある行を返す(状態問わず。キャンセル用)
function 受注個別_割当行_(emsRows, sku, code, ban){
  const targetKeys=[];
  受注候補コード_(sku,code).forEach(v=> codeKeys_(v).forEach(k=>{ if(targetKeys.indexOf(k)<0) targetKeys.push(k); }));
  const out=[];
  emsRows.forEach(r=>{
    if(!r.注文番号) return;
    if(!codeKeys_(r.商品コード).some(k=>targetKeys.indexOf(k)>=0)) return;
    const cur=個別対応_P注文展開_(r.注文番号, r.数量).filter(e=>e.ban===ban).reduce((s,e)=>s+e.qty,0);
    if(cur>0) out.push({item:r, cur});
  });
  return out;
}

// 受注明細の1行を塗る/クリアする(color=nullでクリア)。未入金の受注番号セル(赤)と
// 個数>=2の個数セル(緑)のマークは維持する(②の色分けと同じ約束事)
function 受注個別_行色_(sh, M, rowNo, color){
  const nc=sh.getLastColumn();
  const rng=sh.getRange(rowNo,1,1,nc);
  const cur=rng.getBackgrounds()[0];
  const 赤=HIKIATE_CFG.色_赤.toLowerCase(), 緑=HIKIATE_CFG.色_緑.toLowerCase();
  const out=cur.map((c,i)=>{
    const lc=String(c||'').toLowerCase();
    if(M.番号>=0 && i===M.番号 && lc===赤) return c;
    if(M.個数>=0 && i===M.個数 && lc===緑) return c;
    return color;
  });
  rng.setBackgrounds([out]);
}

function 受注個別_行情報_(row, M){
  return {
    ban: String(row[M.番号]||'').trim(),
    code: String(row[M.コード]||'').trim(),
    sku: M.SKU>=0? String(row[M.SKU]||'').trim() : '',
    qty: Number(row[M.個数])||0,
    name: M.商品名>=0? String(row[M.商品名]||'').trim() : '',
    kbn: 区分_(row[M.選択肢])
  };
}

function 選択行を個別引当(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  const sh=ss.getActiveSheet();
  if(sh.getName()!==HIKIATE_CFG.受注){ ui.alert('受注明細で対象の行を選択してからボタンを押してください。'); return; }
  const M=列マップ_(sh);
  const rows=受注個別_選択行_(sh,M);
  if(!rows.length){ ui.alert('対象のデータ行を選択してからボタンを押してください。'); return; }
  const ems=個別対応_EMSリスト_();
  if(ems.error){ ui.alert(ems.error); return; }
  const R=sh.getDataRange().getValues();
  const results=[];
  rows.forEach(rowNo=>{
    const l=受注個別_行情報_(R[rowNo-1], M);
    const label=l.ban+' '+(l.code||l.sku);
    if(!l.ban || (!l.code && !l.sku)){ results.push('行'+rowNo+': 受注番号/商品コードが読めません'); return; }
    if(l.kbn!=='取り寄せ'){ results.push(label+': 取り寄せ行ではないのでスキップ'); return; }
    if(l.qty<=0){ results.push(label+': 個数0のためスキップ'); return; }
    const cands=受注個別_候補_(ems.rows, l.sku, l.code);
    if(!cands.length){
      // 残あり箱が無くても、既にこの受注が到着済の箱に名指しされていれば「割当済み=出せる」なので黄に塗る
      const 既=受注個別_割当行_(ems.rows, l.sku, l.code, l.ban).filter(h=>String(h.item.状態||'').trim()==='到着済');
      if(既.length){
        受注個別_行色_(sh, M, rowNo, HIKIATE_CFG.色_黄);
        results.push(label+': 既に割当済み('+既.reduce((s,h)=>s+h.cur,0)+'個)。行を黄にしました');
      } else {
        results.push(label+': 到着済の箱にこの商品(残あり)が見つかりません');
      }
      return;
    }
    let pick=cands[0];
    if(cands.length>1){
      const menu=cands.map((c,i)=>(i+1)+') '+(c.item.EMS到着日||'?')+'着 '+c.item.EMS番号+(c.item.BoxNo?' Box'+c.item.BoxNo:'')+' 残'+c.残).join('\n');
      const resp=ui.prompt('どの箱から引き当てる？', l.ban+' '+l.name+'\n\n'+menu+'\n\n番号を入力', ui.ButtonSet.OK_CANCEL);
      if(resp.getSelectedButton()!==ui.Button.OK){ results.push(label+': 中止'); return; }
      const n=Number(String(resp.getResponseText()||'').trim());
      if(!(n>=1 && n<=cands.length)){ results.push(label+': 番号が不正のため中止'); return; }
      pick=cands[n-1];
    }
    // 書き込み直前にP列を読み直す(連続操作でスナップショットが古くても上書きしない)
    const cell=ems.sh.getRange(pick.item.row, ems.c.注文番号+1);
    const entries=個別対応_P注文展開_(cell.getDisplayValue(), pick.item.数量);
    const used=entries.reduce((s,e)=>s+e.qty,0);
    const 残=Math.max(0,(Number(pick.item.数量)||0)-used);
    const 既存=entries.filter(e=>e.ban===l.ban).reduce((s,e)=>s+e.qty,0);
    if(既存>=l.qty){ results.push(label+': 既にこの箱へ'+既存+'個割当済み'); return; }
    const take=Math.min(l.qty-既存, 残);
    if(take<=0){ results.push(label+': 箱の残数がありません'); return; }
    const 入荷済=M.入荷>=0 && String(R[rowNo-1][M.入荷]||'').trim()!=='';
    const confirm=ui.alert('個別引当の確認',
      l.ban+' '+l.name+'\n商品: '+(l.code||l.sku)+' × '+take+'個'+(take<l.qty-既存?'（注文'+l.qty+'個に対して部分）':'')
      +'\n箱: '+(pick.item.EMS到着日||'?')+'着 '+pick.item.EMS番号+(pick.item.BoxNo?' Box'+pick.item.BoxNo:'')+'（残'+残+'）'
      +(入荷済?'\n※入荷日は既に入っているので変更しません':'\n※入荷日にEMS到着日を記入します'),
      ui.ButtonSet.OK_CANCEL);
    if(confirm!==ui.Button.OK){ results.push(label+': 中止'); return; }
    entries.push({ban:l.ban, qty:take});
    const text=個別対応_P注文整形_(entries, pick.item.数量);
    cell.setValue(text);
    cell.setBackground(/[,:：、]/.test(text)?'#fff2cc':null);
    引当履歴_個別記録_(個別対応_履歴Rec_(pick.item, l.ban, take, '個別引当'));
    let msg=label+': OK '+take+'個 → '+pick.item.EMS番号;
    if(!入荷済 && M.入荷>=0 && pick.item.EMS到着日){
      sh.getRange(rowNo, M.入荷+1).setValue(pick.item.EMS到着日).setNumberFormat('yyyy-mm-dd');
      msg+=' / 入荷日 '+pick.item.EMS到着日;
    }
    受注個別_行色_(sh, M, rowNo, HIKIATE_CFG.色_黄); // 引き当たった=今回出せる分として即座に黄
    results.push(msg);
  });
  ui.alert('個別引当の結果', results.join('\n'), ui.ButtonSet.OK);
}

function 選択行の引当キャンセル(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  const sh=ss.getActiveSheet();
  if(sh.getName()!==HIKIATE_CFG.受注){ ui.alert('受注明細で対象の行を選択してからボタンを押してください。'); return; }
  const M=列マップ_(sh);
  const rows=受注個別_選択行_(sh,M);
  if(!rows.length){ ui.alert('対象のデータ行を選択してからボタンを押してください。'); return; }
  const ems=個別対応_EMSリスト_();
  if(ems.error){ ui.alert(ems.error); return; }
  const R=sh.getDataRange().getValues();
  const results=[];
  rows.forEach(rowNo=>{
    const l=受注個別_行情報_(R[rowNo-1], M);
    const label=l.ban+' '+(l.code||l.sku);
    if(!l.ban || (!l.code && !l.sku)){ results.push('行'+rowNo+': 受注番号/商品コードが読めません'); return; }
    const hits=受注個別_割当行_(ems.rows, l.sku, l.code, l.ban);
    if(!hits.length){ results.push(label+': P列にこの受注の割当が見つかりません'); return; }
    const total=hits.reduce((s,h)=>s+h.cur,0);
    const confirm=ui.alert('引当キャンセルの確認',
      l.ban+' '+(l.code||l.sku)+'\n合計 '+total+'個の割当を取り消します。\n'
      +hits.map(h=>'・'+(h.item.EMS到着日||'?')+'着 '+h.item.EMS番号+' x'+h.cur).join('\n'),
      ui.ButtonSet.OK_CANCEL);
    if(confirm!==ui.Button.OK){ results.push(label+': 中止'); return; }
    let removed=0;
    hits.forEach(h=>{
      const cell=ems.sh.getRange(h.item.row, ems.c.注文番号+1);
      const entries=個別対応_P注文展開_(cell.getDisplayValue(), h.item.数量);
      const kept=[]; let cut=0;
      entries.forEach(e=>{ if(e.ban===l.ban){ cut+=e.qty; } else kept.push(e); });
      if(cut<=0) return;
      const text=個別対応_P注文整形_(kept, h.item.数量);
      cell.setValue(text);
      cell.setBackground(/[,:：、]/.test(text)?'#fff2cc':null);
      引当履歴_キャンセル_(個別対応_履歴Rec_(h.item, l.ban, cut, '個別引当'), cut);
      removed+=cut;
    });
    SpreadsheetApp.flush(); // 入荷日クリアの「他に有効割当があるか」判定が書き込み後のP列を見るように
    const cleared=個別対応_入荷日クリア_(l.ban, l.code||l.sku, '');
    受注個別_行色_(sh, M, rowNo, null); // キャンセルした行は色をクリア(未入金の赤・個数の緑マークは残す)
    results.push(label+': キャンセル '+removed+'個'+(cleared?' / 入荷日クリア '+cleared+'行':''));
  });
  ui.alert('引当キャンセルの結果', results.join('\n'), ui.ButtonSet.OK);
}

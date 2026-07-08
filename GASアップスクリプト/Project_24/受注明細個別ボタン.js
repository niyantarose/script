// ===== 受注明細から直接 個別引当/引当キャンセル =====
// 受注明細で対象の行を選択して、シート上部のボタン(個別引当=緑/引当キャンセル=赤)を押すだけ。
// 照合キーの転記・個別対応シートへの入力は不要(個別対応シート運用は廃止。関数は残置)。
//
// 個別引当: 選択行の商品を発注共有EMSリストの「到着済」の箱から探してP列に名指しを追記。
//           引当履歴に記録し、入荷日が空欄ならEMS到着日を自動記入する。
// キャンセル: EMSリストP列からこの受注番号×この商品の名指しを削除。履歴を更新し、
//           他に有効な割当が無ければ受注明細の入荷日もクリアする。

const UKE_KOBETSU_CFG = {
  ボタン行: 1,
  引当ボタン列: 3,  // C列(ボタン画像のアンカー。凡例より上のフリーズ域)
  取消ボタン列: 5,  // E列
  ボタン幅: 150,
  ボタン高さ: 36
};

function 受注明細個別ボタンを設置(){ 受注明細個別ボタンを設置_(false); }

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
  const add=(col,label,fill,stroke,fn)=>{
    const img=sh.insertImage(個別対応_ボタン画像Blob_(label,fill,stroke,'#ffffff'), col, cfg.ボタン行);
    img.setAltTextTitle('受注個別ボタン_'+label);
    img.setAltTextDescription(label+'を実行(先に対象の行を選択)');
    img.assignScript(fn);
    img.setWidth(cfg.ボタン幅);
    img.setHeight(cfg.ボタン高さ);
  };
  try{
    add(cfg.引当ボタン列,'個別引当','#6aa84f','#38761d','選択行を個別引当');
    add(cfg.取消ボタン列,'引当キャンセル','#cc0000','#990000','選択行の引当キャンセル');
  }catch(e){
    if(!silent) SpreadsheetApp.getUi().alert('ボタン画像を設置できませんでした:\n'+e.message);
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
    if(!cands.length){ results.push(label+': 到着済の箱にこの商品(残あり)が見つかりません'); return; }
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
    results.push(label+': キャンセル '+removed+'個'+(cleared?' / 入荷日クリア '+cleared+'行':''));
  });
  ui.alert('引当キャンセルの結果', results.join('\n'), ui.ButtonSet.OK);
}

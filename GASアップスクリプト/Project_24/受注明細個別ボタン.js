// ===== 受注明細から直接 個別引当/引当キャンセル =====
// 受注明細で対象の行を選択して、シート上部のボタン(個別引当=緑/引当キャンセル=赤)を押すだけ。
// 照合キーの転記・個別対応シートへの入力は不要(個別対応シート運用は廃止。関数は残置)。
//
// 取り置き台帳が唯一の正になったため、ボタンは台帳へ直接書く(P列は④が計画から書き直す派生表示):
// 個別引当: 選択行の商品を「到着済」の箱(台帳使用を差し引いた残あり)から選び、
//           取り置き中の行を台帳へ直接作成。履歴に記録し、入荷日が空欄ならEMS到着日を自動記入。
//           コード不一致でも受注番号タグ(REQ系・名指し買付)の箱はこの受注の候補になる(救済)。
// キャンセル: この受注番号×この商品の取り置き中を台帳で手動解除。履歴を更新し、
//           取り置きが残らなければ受注明細の入荷日もクリアする。
// どちらもP列・各一覧シートは次の④実行で台帳から自動整合する。

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
  引当キャンセル: 'iVBORw0KGgoAAAANSUhEUgAAASwAAABICAYAAABMb8iNAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAmCSURBVHhe7Zy/yxRHGMefMoVFythZSUrBxtK/QFJaWFgGbFKKlb2FTfC9e3mJWhiNEH+QyBuIYGFCCIaEQIhEQkIgYCCFRQrTXfi+3Fye+97M7Mze7p377veBD8rN7ezu7MznnpkdNauMXbMTO2and8zOT80uCyFELROz9+GRa2an2DFrxXWztyCnidntidnrqdlMCCE6Zh8Su272NjuoOKZm703Nfo9ULoQQnTMxezU1u4hEiX2UDFhuYnafKxNCiE0wMXu+Y3aM3bQS+BK+zBUEbpjNPjabfWI2+0wIIVpwd+6RmxHHOGm9wjoXO2oR88xqRVa7ZrMHZrOnZrPvhBCiQ76eS2wvIa2p2bvsqoOITQNvmc2+iZxECCG6BJ75NC6t5ytrWlOzs/xFHPwsUrEQQvTF5xFpTc2uLmQFe03NXnJmJVkJIbbBvdUs6/ViER77H3wh1qw0DRRCbAskS5EF+Z0wHdz3BQ8jFQghxCZ5vCqsl+HN4NIO9q8iBwshxCZBlsVvDg3/lsd/gDSMDxRCiG2A/Z5Lwpr/85vFB3ciBwkhxDbA/qwlYU3MPvAfYHWeDxJCiG3wiIU1/+8eFh/AaHyQEEJsg5UMS8ISQrypSFhCiMEgYQkhBoOEJYQYDBKWEGIwSFhCiMEgYXVMiH9fvJj9deXKSnlbXt2710u9QgyJ0Qvrp+PHD2RQw69nzqzUE/BRK5Yfjx6dfX/kyMrnIAgrVy/uBfDnNfx56dLiXPg7l48V3/5/XLiwUi42w+iF9fPJk4uOWBq5DosMqOR7HkgKEkL88/TpSjkoGTChDnwX98XlJUBSIVAPl4+VkvZvC34Aw48h/0jgXHoe/yNhdSysNh2bryF2XFO9yM58cMcv5bdz55bq4fI3hb/39g7aIZWRdk1T+69DTkq5sjEyemGVUtphS7/H+MwGAQHV1IsBHAJZXtuBzPJsW0+f+EGM+OX06ZXv5GiTgTa1/zrkpJQrGyMSViGlHbb0ewzE4KeT3Dlz9bJkcmtsTeA6fCDjQv34E+fFtBPXgqkrgo+PgeuBUHF/tXJhsEYXC1wTSz6Gv7/SY0Cu/T2hrdBOaKOSTDcnpVzZGJGwCintsKXfi4GB7cNnAal6WTClnRrHoX6A+gCk4s9TElxvDJ891rZJjPCiJBZNQmS5ryMsllMqms7hpcSCk7CWkbDmoDPk8JEbdLGOXYPPsvzbwFS9fG0lbwl50K4TJefzIu5y0PF6W4jUW1TgBYDg8hQ5ITVF0z2nni2QsJaRsDLTjFRwp/LkOl8JGNyQFk/rUvXi1zusX2EAc30xeIG+KXBuSADnDVPEmjUgbl8ux/WETA/3AjmAkvUz1M0yyQ1s3444F5eHOtH+TZlTaeB58jlS18R9RsJaRsJqkXFwp/LkOt86NNVbkul4EBiMXkYAbRG2RyCaBlspPnAev98rFaktHgzE5l86pETH0+eY4HlaXhK4D9wTjsVz4B8E/vFhcs9WwlpGwqKpxbqdItf5YtTKsib4XKXwNXF5CSFLQRs0iSkVtc8Cos2tYbGMYpJn2cQC14U2Sq1N1Qo/12ckrGUkLOoUufWPEnKdLwbLocvgc5XCg7Zm+gdSa0uxQBaFNg9TTb9AH8uA1oG3fnB5AGV4jriWcF2lz5Wnv03ZFcjVLWEtI2FRh1l3kOQ6X4w3UVjARxh0fq0pTOkwuPk+m+4Jx6Qk6F86pDKYNrCEa59z6XP1a145KZbWXSosfrt4WJGwaHCGdZ0mUoMp1/n6hkXB5SWE7Q41i82x+0SE3eior+Ta/Hdyg7MNfpqGSK1zpSh5rrz5NyVlJld3ibDCVHcM0hq9sDiFL42hCwv3jfUeXGPYDNo2UgOJ8RFbP/LXUDKVKoXbpc20v+m58vpYjTx83Zz5lQjLZ6W5NbzDwOiFVbPe4oPrCTR17D7hgcnlAc42agL3FzInDI7SLALkhIS6QpROpUpAJsXZYurHJkfuuXK7l77dDPjg9mwSFvffNvc2JEYvLL8Qm/tV5E7J5YFcx86B49aFByafI+AHQSzCPigftVOoGH7K5NuatxuwzNaB5VzzTDxo31Qdvqx0/5jHR42wuN34ug4joxeWXzzmzuLxGQCCywO5jp2jj+BzBHCf/k0YrhOf+WkaT5VjU7haUjveu9xW4mFZ1WY+nqbnGjb81rZT0w9hTlj81rNWlENk9MIqJddxPE0dO0UfweeoxUcXayMsQV+GgQuh1A74FCwrxDrTpbbPtQkv65hQU/2Op4JdPJ8hIGEV4qczfQirC5p+rWvp4158dCUnD6TE09kuztVHWwB/rbEliZiweIG/zUuEoSJhFVLaYUu/1wddC8tnKTlJ1+Dbp8u1KkyH/OD2sa6sQB/PlZ9XLANkYXFm1WbNbMhIWIX4yA20Pjp2KTwAuLwWP1i6enOXWnhvCwZ5SlRdTjG7fq6QjN+OkPpBSN0bYmyyAhJWAbzgnhsEXXfsGroWVtf1Ad+WqUHaBNofmUZs6hcCMuxyMPvnCtHgPtapn9fYYtkVSAlrjLICEhaB7AkD1eMHRlOmUSusVIfsIvhctfA/Z8m9RS2F6+TyEnxmwoH2z/2gtIWnYjXB7cayyvWTWP/APY5RVkDCIvifV3DkOhc4TMICPnJT4Rp8tJELZ7wISIDF0DW5jC4XfF2+jzUtmHP/4J3wY0PCIvgNjI+SNPywCQv3g/vGn129Ou9i4R11YB8Sjm96Jl2C84X/ZLAkUhkf2rJkDQ/3FvYKpqaNY0LCItC5eMoR/keCkoFRK6wu6WPNqQ/QLhjwGIicfQiRQ8ISQgwGCUsIMRgkLCHEYJCwhBCDQcISQgwGCUsIMRgkLCHEYJCwhBCDQcISQgyGmLAu+g8eRA4SQoht8CgirLP+g7uRg4QQYhusZFg7Zqf9B7ciBwkhxDa4x8L60Owd/wH4NnKgEEJsmhssLMTE7Af/IeaNfKAQQmySpySridnrA2Hxm8KPzGbPIhUIIcSmuLMqrPsHwtoxOwZ7+UJtbxBCbIsvSVZz3jsQ1jzLuspfwEFckRBC9AmmgrurstpfyAoxX3x/ydJSpiWE2BSPzWZ75CDM/nbNTiwJC3HN7BRPDQG2OjyJVC6EEF2ArIrXrBxn2VWLQGFMWgCvGLEvAlkX3iTiTyGEqOWL+Z8PzWY3I65xXGZHrQQyrdj0UAghNgGSph2z8+ymZGBNa2J2mysSQoieeYKkiZ1UFPN1rdupaaIQQnTE/tLWhUj8B8R19Yj1pXRvAAAAAElFTkSuQmCC'
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

// (旧P列ベースの候補/割当行ヘルパーは台帳直書き化で廃止。候補は取り置き計算.jsの個別_台帳候補_)

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

// 「今回入荷EMSの在庫」の該当行を、EMSリストP列の今の割当(entries)に合わせて更新する。
// 余り(F)・受注番号セル(G以降)・行の色(A〜F)を②と同じ約束事で書き直す:
//   割当あり: 余りあり=緑/余りなし=黄、受注番号セルは入金済=黄・未入金=赤
//   割当なし: 色クリア・余り=数量
// 行が見つからない/シートが無い場合は黙ってスキップ(②が正なので次回②で必ず揃う)
function 受注個別_台帳更新_(ss, item, entries, R, M){
  try{
    const sh=ss.getSheetByName(HIKIATE_CFG.日本在庫);
    if(!sh || sh.getLastRow()<3) return;
    const nc=sh.getLastColumn();
    const vals=sh.getRange(3,1,sh.getLastRow()-2,nc).getDisplayValues();
    const key=normCode_(item.商品コード), 着=ymd_(item.EMS到着日), ems=String(item.EMS番号||'').trim();
    let rowNo=-1;
    for(let i=0;i<vals.length;i++){
      const r=vals[i];
      if(String(r[4]||'').trim()!==ems) continue;
      if(normCode_(r[2])!==key) continue;
      if(着 && ymd_(r[1])!==着) continue;
      rowNo=3+i; break;
    }
    if(rowNo<0) return;
    const qty=Number(item.数量)||0;
    const used=entries.reduce((s,e)=>s+e.qty,0);
    const surplus=Math.max(0, qty-used);
    const paid=ban=>{ for(let i=M.hr;i<R.length;i++){ const row=R[i];
      if(String(row[M.番号]||'').trim()===ban && 入金済み_(row[M.入金])) return true; } return false; };
    const agg={}; entries.forEach(e=>{ agg[e.ban]=(agg[e.ban]||0)+e.qty; });
    const bans=Object.keys(agg).sort((a,b)=>番号num_(a)-番号num_(b));
    // F=余り
    sh.getRange(rowNo,6).setValue(surplus>0?surplus:'');
    // G以降=受注番号(2個以上は「番号:個数」表示。余った古いセルはクリア)
    const nBanCells=Math.max(0,nc-6);
    if(nBanCells>0){
      const v=[], bg=[];
      for(let i=0;i<nBanCells;i++){
        v.push(i<bans.length? (bans[i]+(agg[bans[i]]>1?'*'+agg[bans[i]]:'')) : '');
        bg.push(i<bans.length? (paid(bans[i])? HIKIATE_CFG.色_黄 : HIKIATE_CFG.色_赤) : null);
      }
      sh.getRange(rowNo,7,1,nBanCells).setValues([v]).setBackgrounds([bg]);
    }
    // A〜F=行の状態色
    const col= bans.length? (surplus>0? HIKIATE_CFG.色_緑 : HIKIATE_CFG.色_黄) : null;
    sh.getRange(rowNo,1,1,6).setBackgrounds([new Array(6).fill(col)]);
  }catch(e){}
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

// この供給キー(箱×元コード)の台帳ACTIVE行を {ban,qty} に集計(今回入荷EMSの在庫の表示更新用)
function 受注個別_台帳消費者_(ledgerRows, item){
  const key=取り置き_供給キー_(item.EMS番号, item.商品コード);
  const agg={};
  (ledgerRows||[]).forEach(r=>{
    if(r.状態!=='取り置き中') return;
    if(取り置き_供給キー_(r.元EMS番号, r.元EMS商品コード||r.商品コード)!==key) return;
    const ban=String(r.受注番号||'').trim(); if(!ban) return;
    agg[ban]=(agg[ban]||0)+(Number(r.取り置き数量)||0);
  });
  return Object.keys(agg).map(ban=>({ban, qty:agg[ban]}));
}

function 選択行を個別引当(){ 直列_(選択行を個別引当_本体_); }
function 選択行を個別引当_本体_(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  const sh=ss.getActiveSheet();
  if(sh.getName()!==HIKIATE_CFG.受注){ ui.alert('受注明細で対象の行を選択してからボタンを押してください。'); return; }
  const M=列マップ_(sh);
  const rows=受注個別_選択行_(sh,M);
  if(!rows.length){ ui.alert('対象のデータ行を選択してからボタンを押してください。'); return; }
  const ems=個別対応_EMSリスト_();
  if(ems.error){ ui.alert(ems.error); return; }
  let ledgerRows, movementRows;
  try{ ledgerRows=取り置き台帳_読む_(); movementRows=EMS在庫移動台帳_読む_(); }
  catch(e){ ui.alert('取り置き台帳を読み込めません:\n'+e.message); return; }
  const R=sh.getDataRange().getValues();
  const results=[];
  rows.forEach(rowNo=>{
    const l=受注個別_行情報_(R[rowNo-1], M);
    const label=l.ban+' '+(l.code||l.sku);
    if(!l.ban || (!l.code && !l.sku)){ results.push('行'+rowNo+': 受注番号/商品コードが読めません'); return; }
    if(l.kbn!=='取り寄せ'){ results.push(label+': 取り寄せ行ではないのでスキップ'); return; }
    if(l.qty<=0){ results.push(label+': 個数0のためスキップ'); return; }
    const summary=取り置き_集計_(ledgerRows, movementRows);
    const key=取り置き_行キー_({ban:l.ban,code:l.code,sku:l.sku});
    const 確保済=Number(summary.activeByKey[key])||0;
    if(確保済>=l.qty){
      受注個別_行色_(sh, M, rowNo, HIKIATE_CFG.色_黄);
      results.push(label+': 既に取り置き'+確保済+'個で注文数を満たしています。行を黄にしました');
      return;
    }
    const targetKeys=[];
    受注候補コード_(l.sku,l.code).forEach(v=> codeKeys_(v).forEach(k=>{ if(targetKeys.indexOf(k)<0) targetKeys.push(k); }));
    const cands=個別_台帳候補_(ems.rows, targetKeys, l.ban, summary);
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
    const take=Math.min(l.qty-確保済, pick.残);
    if(take<=0){ results.push(label+': 箱の残数がありません'); return; }
    const 入荷済=M.入荷>=0 && String(R[rowNo-1][M.入荷]||'').trim()!=='';
    const confirm=ui.alert('個別引当の確認',
      l.ban+' '+l.name+'\n商品: '+(l.code||l.sku)+' × '+take+'個'+(take<l.qty-確保済?'（注文'+l.qty+'個に対して部分）':'')
      +'\n箱: '+(pick.item.EMS到着日||'?')+'着 '+pick.item.EMS番号+(pick.item.BoxNo?' Box'+pick.item.BoxNo:'')+'（残'+pick.残+'）'
      +'\n※取り置き台帳へ直接登録します(P列・一覧は次の④で整合)'
      +(入荷済?'\n※入荷日は既に入っているので変更しません':'\n※入荷日にEMS到着日を記入します'),
      ui.ButtonSet.OK_CANCEL);
    if(confirm!==ui.Button.OK){ results.push(label+': 中止'); return; }
    const now=new Date();
    const applied=個別_台帳引当行_(ledgerRows, {ban:l.ban,code:l.code,sku:l.sku}, take, pick.item.EMS番号, String(pick.item.商品コード||'').trim(), now);
    if(applied.error){ results.push(label+': '+applied.error); return; }
    try{ 取り置き台帳_保存_(applied.rows); }
    catch(e){ results.push(label+': 台帳保存に失敗 '+e.message); return; }
    ledgerRows=applied.rows; // 連続操作は保存済みの台帳を引き継ぐ
    try{ 引当履歴_個別記録_(個別対応_履歴Rec_(pick.item, l.ban, take, '個別引当')); }catch(e){}
    let msg=label+': OK '+take+'個 → '+pick.item.EMS番号+'（台帳登録）';
    if(!入荷済 && M.入荷>=0 && pick.item.EMS到着日){
      sh.getRange(rowNo, M.入荷+1).setValue(pick.item.EMS到着日).setNumberFormat('yyyy-mm-dd');
      msg+=' / 入荷日 '+pick.item.EMS到着日;
    }
    受注個別_行色_(sh, M, rowNo, HIKIATE_CFG.色_黄); // 引き当たった=今回出せる分として即座に黄
    受注個別_台帳更新_(ss, pick.item, 受注個別_台帳消費者_(ledgerRows, pick.item), R, M); // 今回入荷EMSの在庫の該当行も連動更新
    results.push(msg);
  });
  ui.alert('個別引当の結果', results.join('\n'), ui.ButtonSet.OK);
}

function 選択行の引当キャンセル(){ 直列_(選択行の引当キャンセル_本体_); }
function 選択行の引当キャンセル_本体_(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  const sh=ss.getActiveSheet();
  if(sh.getName()!==HIKIATE_CFG.受注){ ui.alert('受注明細で対象の行を選択してからボタンを押してください。'); return; }
  const M=列マップ_(sh);
  const rows=受注個別_選択行_(sh,M);
  if(!rows.length){ ui.alert('対象のデータ行を選択してからボタンを押してください。'); return; }
  let ems=null;
  try{ ems=個別対応_EMSリスト_(); }catch(e){ ems={error:e.message}; }
  let ledgerRows;
  try{ ledgerRows=取り置き台帳_読む_(); }
  catch(e){ ui.alert('取り置き台帳を読み込めません:\n'+e.message); return; }
  const R=sh.getDataRange().getValues();
  const results=[];
  rows.forEach(rowNo=>{
    const l=受注個別_行情報_(R[rowNo-1], M);
    const label=l.ban+' '+(l.code||l.sku);
    if(!l.ban || (!l.code && !l.sku)){ results.push('行'+rowNo+': 受注番号/商品コードが読めません'); return; }
    const key=取り置き_行キー_({ban:l.ban,code:l.code,sku:l.sku});
    const targets=ledgerRows.filter(r=>r.状態==='取り置き中' && 取り置き_行キー_(r)===key);
    if(!targets.length){ results.push(label+': 取り置き台帳にこの受注の取り置き中がありません'); return; }
    const total=targets.reduce((s,r)=>s+(Number(r.取り置き数量)||0),0);
    const confirm=ui.alert('引当キャンセルの確認',
      l.ban+' '+(l.code||l.sku)+'\n合計 '+total+'個の取り置きを手動解除します。\n'
      +targets.map(r=>'・'+(r.元EMS番号||'(EMSなし)')+' x'+r.取り置き数量).join('\n')
      +'\n※P列・一覧は次の④で整合します',
      ui.ButtonSet.OK_CANCEL);
    if(confirm!==ui.Button.OK){ results.push(label+': 中止'); return; }
    const now=new Date();
    const plan=個別_台帳解除計画_(ledgerRows, key, '個別ボタンからキャンセル', now);
    try{ 取り置き台帳_保存_(plan.rows); }
    catch(e){ results.push(label+': 台帳保存に失敗 '+e.message); return; }
    ledgerRows=plan.rows; // 連続操作は保存済みの台帳を引き継ぐ
    // 履歴・今回入荷EMSの在庫の表示も追随(読めない環境では黙ってスキップ。④が正)
    targets.forEach(r=>{
      const ems番号=String(r.元EMS番号||'').trim(); if(!ems番号 || !ems || ems.error) return;
      const item=(ems.rows||[]).find(x=>String(x.EMS番号||'').trim()===ems番号 && normCode_(x.商品コード)===normCode_(r.元EMS商品コード||r.商品コード));
      if(!item) return;
      try{ 引当履歴_キャンセル_(個別対応_履歴Rec_(item, l.ban, Number(r.取り置き数量)||0, '個別引当'), Number(r.取り置き数量)||0); }catch(e){}
      受注個別_台帳更新_(ss, item, 受注個別_台帳消費者_(ledgerRows, item), R, M);
    });
    // 取り置きが残らなければ入荷日をクリア(台湾・中国の手入力入荷日は別ルートなのでここへ来ない)
    let cleared=0;
    if(M.入荷>=0 && String(R[rowNo-1][M.入荷]||'').trim()!==''){
      sh.getRange(rowNo, M.入荷+1).setValue('');
      cleared=1;
    }
    受注個別_行色_(sh, M, rowNo, null); // キャンセルした行は色をクリア(未入金の赤・個数の緑マークは残す)
    results.push(label+': 手動解除 '+plan.qty+'個'+(cleared?' / 入荷日クリア':''));
  });
  ui.alert('引当キャンセルの結果', results.join('\n'), ui.ButtonSet.OK);
}

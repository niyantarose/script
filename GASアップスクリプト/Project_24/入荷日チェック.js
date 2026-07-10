// ===== 入荷日の整合チェック(誤記入の検出) =====
// 受注明細の「取り寄せ＋入荷日あり」行について、その入荷日と同じ到着日で
// その商品が発注共有ファイルのEMSリストに存在するかを照合する。
// 見つからない行＝引当バグ(救済横取り)や手入力ミスで付いた可能性のある入荷日として一覧に出す。
// 読み取り専用: 受注明細は一切書き換えない。消すかどうかは一覧を見て手で判断する。
//
// 誤検出になりうるもの(正しい場合もある):
//   ・手入力した入荷日(EMS便以外の入荷など)
//   ・EMSリスト側の到着日を後から書き換えた場合

// ===== 一覧を消す(図形ボタン割り当て用) =====
// タイトル行(1)と見出し行(2)は残して、3行目以降の一覧だけ消す。
// 一覧は🔎整合チェックでいつでも作り直せる診断結果なので確認ダイアログなしの1クリック。
function 入荷日チェックを消す(){
  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName('入荷日チェック');
  if(!sh){ ss.toast('「入荷日チェック」タブが無いで','🗑',5); return; }
  const mr=sh.getMaxRows(), mc=sh.getMaxColumns();
  if(mr>=3){
    const r=sh.getRange(3,1,mr-2,mc);
    r.clearContent();
    r.setBackgrounds(Array.from({length:mr-2},()=>new Array(mc).fill(null)));
  }
  sh.getRange(1,1).setValue('入荷日の整合チェック: (未実行) 🔎整合チェックを実行すると一覧が出ます');
  ss.toast('入荷日チェックの一覧を消しました','🗑入荷日',5);
}

// ===== 一覧の入荷日を一括クリア =====
// 入荷日整合チェックの一覧に出た行の入荷日を、確認のうえ受注明細から一括で消す。
// 行ズレ対策: 受注番号+商品コード/SKU+入荷日が一覧と一致する行だけをクリアし、結果をI列に書く。
// 残したい行(手入力の正しい入荷日など)は、実行前に一覧からその行を削除しておく。
function 入荷日チェック_一覧をクリア(){ 直列_(入荷日チェック_一覧をクリア本体_); }
function 入荷日チェック_一覧をクリア本体_(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  const rep=ss.getSheetByName('入荷日チェック');
  if(!rep || rep.getLastRow()<3){ ui.alert('「入荷日チェック」の一覧がありません。先に 🔎 入荷日の整合チェック を実行してください。'); return; }
  const recv=ss.getSheetByName(HIKIATE_CFG.受注);
  if(!recv){ ui.alert('「'+HIKIATE_CFG.受注+'」タブがありません'); return; }
  const M=列マップ_(recv);
  if(M.入荷<0){ ui.alert('受注明細に「入荷日」列がありません'); return; }
  const list=rep.getRange(3,1,rep.getLastRow()-2,8).getDisplayValues().filter(r=>String(r[1]||'').trim());
  if(!list.length){ ui.alert('一覧が空です(問題なし)。'); return; }
  const ans=ui.alert('入荷日の一括クリア',
    '一覧の '+list.length+' 行の入荷日を受注明細から消します。\n\n'+
    '※「これは手入力の正しい入荷日」という行が一覧にあれば、\n'+
    '　キャンセルして一覧からその行を削除してから実行してください。ええ？',
    ui.ButtonSet.OK_CANCEL);
  if(ans!==ui.Button.OK) return;
  const R=recv.getDataRange().getValues();
  let ok=0, ng=0, 保留=0;
  const 結果=[];
  list.forEach(r=>{
    const rowNo=Number(r[0])||0, ban=String(r[1]||'').trim(), code=String(r[3]||'').trim(), sku=String(r[4]||'').trim(), listed=ymd_(r[6]);
    let res='';
    if(String(r[7]||'').indexOf('台湾/中国ルート')===0){
      結果.push(['台湾/中国ルートのためスキップ(消す場合は手動で)']); 保留++;
      return;
    }
    if(String(r[7]||'').indexOf('未到着扱い:')===0){
      結果.push(['入荷日は正しい可能性が高いためスキップ(EMSリスト側のステータスを「到着済」に直してください)']); 保留++;
      return;
    }
    if(rowNo>M.hr && rowNo<=R.length){
      const row=R[rowNo-1];
      const banOk=String(row[M.番号]||'').trim()===ban;
      const rowCode=String(row[M.コード]||'').trim(), rowSku=M.SKU>=0?String(row[M.SKU]||'').trim():'';
      const codeOk=(code!=='' && rowCode===code)||(sku!=='' && rowSku===sku);
      const dateOk=listed!=='' && ymd_(row[M.入荷])===listed;
      if(banOk && codeOk && dateOk){
        recv.getRange(rowNo, M.入荷+1).clearContent();
        try{ 受注個別_行色_(recv, M, rowNo, null); }catch(e){} // 着済色も一緒にクリア
        res='クリア済み'; ok++;
      } else {
        res='不一致のためスキップ(🔎整合チェックを実行し直してください)'; ng++;
      }
    } else { res='行番号が範囲外(🔎整合チェックを実行し直してください)'; ng++; }
    結果.push([res]);
  });
  rep.getRange(2,9).setValue('処理結果').setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
  rep.getRange(3,9,結果.length,1).setValues(結果);
  ss.toast('入荷日クリア: '+ok+'行 / 対象外スキップ '+保留+'行 / 不一致スキップ '+ng+'行。仕上げに②引き当て実行を回してください','🧹入荷日',8);
}

// 照合スキャン(共通): 受注明細の入荷日×EMSリスト到着日を突き合わせ、疑わしい行の一覧を返す
// 戻り: {error:'…'} または {list:[[受注明細の行,受注番号,氏名,商品コード,SKU,個数,入荷日,理由],…]}
function 入荷日整合_スキャン_(){
  const ss=SpreadsheetApp.getActive();
  const recv=ss.getSheetByName(HIKIATE_CFG.受注);
  if(!recv) return {error:'「'+HIKIATE_CFG.受注+'」タブが無いで'};

  // --- 発注共有EMSリスト: 到着日(yyyy-MM-dd) → その日に到着した商品キー集合 ---
  let sh;
  try{ sh=発注共有を開く_().getSheetByName(P_KAKUTEI_CFG.シート); }
  catch(e){ return {error:'発注共有ファイルが開けません:\n'+e.message}; }
  if(!sh) return {error:'発注共有ファイルに「'+P_KAKUTEI_CFG.シート+'」がありません'};
  const hr=P_KAKUTEI_CFG.ヘッダー行, last=sh.getLastRow();
  if(last<=hr) return {error:'EMSリストにデータがありません'};
  const head=sh.getRange(hr,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
  const f=(...names)=>{ for(const n of names){ const i=head.indexOf(n); if(i>=0) return i; } return -1; };
  const cC=f('商品コード');
  let cA=f('EMS到着日','到着日','到着'); if(cA<0) cA=4; // 既定E列
  const cSt=f('ステータス列');
  if(cC<0) return {error:'EMSリストの'+hr+'行目に「商品コード」見出しがありません'};
  const vals=sh.getRange(hr+1,1,last-hr,sh.getLastColumn()).getValues();
  // ②の引当は「到着済」の箱しか見ない(EMS在庫タブのQUERYと同じ)ので、照合もステータスで分ける。
  // 到着済でない箱に日付だけ合っていても②ではラベンダーのまま=「一覧に出ないのにラベンダー」の見逃しになる
  const byDate={};   // ymd -> Set(正規化キー。別名込み)。ステータス=到着済 の箱だけ(②と同じ見え方)
  const byDate他={}; // ymd -> {正規化キー: ステータス}。到着済以外の箱(ステータス違いの検出用)
  vals.forEach(r=>{
    const code=normCode_(r[cC]); if(!code) return;
    const d=ymd_(r[cA]); if(!d) return;
    const st=cSt>=0? String(r[cSt]||'').trim() : '到着済'; // ステータス列が無い古い形式は従来通り全行を到着済扱い
    if(st==='到着済'){
      const set=byDate[d]||(byDate[d]=new Set());
      codeKeys_(code).forEach(k=>set.add(k));
    } else {
      const m=byDate他[d]||(byDate他[d]={});
      codeKeys_(code).forEach(k=>{ if(!(k in m)) m[k]=st; });
    }
  });

  // --- 受注明細の入荷日付き行を照合 ---
  const M=列マップ_(recv);
  if(M.入荷<0) return {error:'受注明細に「入荷日」列がありません'};
  const R=recv.getDataRange().getValues();
  const out=[];
  // 「在庫反映済み」の箱を指す行は便の締め後の正常状態(正しいラベンダー)。
  // 行ごとに並べてもやることが無い(🧹対象外・対処するなら箱単位)ので、
  // 日付別サマリに集約して一覧には出さない。締めたばかりの箱が紛れていないかは
  // サマリの日付を見れば分かる(今日/直近の日付が出ていたら要確認)。
  const 反映済みサマリ={};
  let 反映済み行数=0;
  for(let i=M.hr;i<R.length;i++){
    const row=R[i];
    const ban=String(row[M.番号]||'').trim(); if(!ban) continue;
    if(区分_(row[M.選択肢])!=='取り寄せ') continue;
    const 入荷日値=row[M.入荷]; if(String(入荷日値||'').trim()==='') continue;
    const code=String(row[M.コード]||'').trim(), sku=M.SKU>=0?String(row[M.SKU]||'').trim():'';
    const d=ymd_(入荷日値);
    const set=byDate[d];
    const keys=[]; 受注候補コード_(sku,code).forEach(v=> codeKeys_(v).forEach(k=>{ if(keys.indexOf(k)<0) keys.push(k); }));
    const hit = !!(set && keys.some(k=>set.has(k)));
    if(hit) continue;
    // EMSリストのこの日にはあるが、箱のステータス列が「到着済」でない → ②は在庫として見ないのでラベンダーのまま。
    // 入荷日は正しい可能性が高い(箱側の直し忘れ)ので🧹の対象外にし、箱のステータスを直すよう案内する
    const m他=byDate他[d];
    const kSt= m他? keys.find(k=> k in m他) : undefined;
    if(kSt!==undefined){
      const 箱St=m他[kSt]||'空';
      if(箱St==='在庫反映済み'){
        反映済みサマリ[d]=(反映済みサマリ[d]||0)+1;
        反映済み行数++;
        continue;
      }
      out.push([i+1, ban, String(row[M.氏名]||''), code, sku, Number(row[M.個数])||0, d||String(入荷日値||''),
        '未到着扱い: EMSリストのこの日にあるがステータスが「'+箱St+'」。箱が着いていれば「到着済」に直すと②で黄色になります(🧹では消しません)']);
      continue;
    }
    // 台湾・中国ルートは韓国EMSに照合先が無い=手入力の入荷日が正。理由を分けて🧹の対象外にする
    const 別ルート=/台湾|中国/.test(String(row[M.選択肢]||'')) || (M.商品名>=0 && /台湾|中国/.test(String(row[M.商品名]||'')));
    out.push([i+1, ban, String(row[M.氏名]||''), code, sku, Number(row[M.個数])||0, d||String(入荷日値||''),
      別ルート? '台湾/中国ルート: 手入力の入荷日なら正しい(🧹では消しません)'
        : (set? 'この日の到着EMSに、この商品が無い' : 'この日に到着したEMSが無い')]);
  }
  return {list:out, 反映済みサマリ:反映済みサマリ, 反映済み行数:反映済み行数};
}

// ②の完了時に呼ぶ件数版(台湾/中国ルートは正扱いなので除外)。読めない時は-1
function 入荷日整合_件数_(){
  const r=入荷日整合_スキャン_();
  if(r.error) return -1;
  return r.list.filter(x=>String(x[7]||'').indexOf('台湾/中国ルート')!==0).length;
}

function 入荷日整合チェック(){
  const ss=SpreadsheetApp.getActive();
  const r=入荷日整合_スキャン_();
  if(r.error){ SpreadsheetApp.getUi().alert(r.error); return; }
  const out=r.list;

  // --- 結果を「入荷日チェック」シートへ(受注明細は触らない) ---
  const NAME='入荷日チェック';
  let rep=ss.getSheetByName(NAME); if(!rep) rep=ss.insertSheet(NAME);
  rep.clearContents();
  rep.getRange(1,1).setValue('入荷日の整合チェック: '+Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm')
    +' / 疑わしい行 '+out.length+'件（消す前に必ず目視確認。手入力の入荷日は正しい場合あり）'
    +((r.反映済み行数||0)>0? '／在庫反映済み(締め済み過去便)の '+r.反映済み行数+'件は下部サマリに集約':''));
  const HDR=['受注明細の行','受注番号','氏名','商品コード','SKU','個数','入荷日','理由'];
  rep.getRange(2,1,1,HDR.length).setValues([HDR]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff').setFontSize(HIKIATE_CFG.字);
  rep.setFrozenRows(2);
  if(out.length){
    rep.getRange(3,1,out.length,HDR.length).setValues(out).setFontSize(HIKIATE_CFG.字);
  } else {
    rep.getRange(3,1).setValue('(問題なし: 一覧に出すべき入荷日ズレはありません)');
  }
  // 在庫反映済み(締め済み過去便)の日付別サマリ。受注番号列(B)が空なので🧹の対象にはならない
  const サマリ日付=Object.keys(r.反映済みサマリ||{}).sort();
  if(サマリ日付.length){
    let base=3+Math.max(out.length,1)+1;
    rep.getRange(base,1).setValue('■ 在庫反映済み(締め済み過去便)の箱を指す行: '+r.反映済み行数+'件 — 正常のため一覧から除外。'
      +'箱をまだ使う場合だけ発注共有のEMSリストで「到着済」に戻してください（直近の日付が出ていたら締めが早すぎた可能性）')
      .setFontWeight('bold').setFontSize(HIKIATE_CFG.字);
    const rows=サマリ日付.map(dd=>[dd,'この日の箱(在庫反映済み)を指す行 '+r.反映済みサマリ[dd]+'件']);
    rep.getRange(base+1,7,rows.length,2).setValues(rows).setFontSize(HIKIATE_CFG.字);
  }
  ss.setActiveSheet(rep);
  ss.toast('入荷日チェック: 疑わしい行 '+out.length+'件'
    +((r.反映済み行数||0)>0? '（在庫反映済み'+r.反映済み行数+'件はサマリ集約）':''),'🔎入荷日',8);
}

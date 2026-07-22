// ===== 一度きりの現物確認移行 =====
// 旧「開始前在庫」61行の仕分け(現物確認済みへ変換/解除/保留)と、
// 全件再計算で消えた確保(基準バックアップとの正差分)の復元を、一つの移行シートで行う。
// 仕様: docs/superpowers/specs/2026-07-22-preallocation-physical-lock-design.md §13
// 比較基準は「取り置き台帳_全件再計算前_20260721_123350」(最後の旧④状態)だけを使う。

const 現物移行_CFG = Object.freeze({
  シート:'現物確認移行', 差分:'現物確認移行_差分',
  基準シート:'取り置き台帳_全件再計算前_20260721_123350',
  // 作業者目線の3点(注文数・確保済み・不足)で表示する(2026-07-22ユーザー要望)。旧記録=旧④帳簿の主張(参考)
  HDR:['入力キー','種別','受注番号','商品コード','SKU','注文数','確保済み','不足','移行数量','選択','メモ','旧記録'],
  旧HDR:['入力キー','種別','受注番号','商品コード','SKU','以前','現在','差','移行数量','選択','メモ'],
  選択肢:['現物確認済みにする','解除','保留']
});

// 現役注文(受注番号+SKU)ごとの注文数量。numeric ban(10117428.0等)も文字列へ寄せる
function 現物移行_注文数量_(currentOrders){
  const out={};
  (currentOrders||[]).forEach(o=>{
    const key=取り置き_行キー_(o);
    out[key]=(out[key]||0)+Math.max(0,Number(o.qty||o.個数)||0);
  });
  return out;
}

// 候補計算: 旧開始前在庫(現状のまま)+基準比の消えた確保(現役注文だけ)。数量は勝手に変えない。
function 現物確認移行_候補計算_(currentLedger, baselineLedger, currentOrders){
  const orderQty=現物移行_注文数量_(currentOrders);
  const out=[], seen=new Set();
  const activeQty=rows=>{
    const sum={};
    (rows||[]).forEach(r=>{
      if(!r || String(r.状態||'')!=='取り置き中') return;
      const q=取り置き_整数_(r.取り置き数量); if(!q) return;
      const key=取り置き_行キー_(r);
      sum[key]=(sum[key]||0)+q;
    });
    return sum;
  };
  const currentActive=activeQty(currentLedger);
  // 1) 旧開始前在庫: 1行=1候補(取置ID単位。変換/解除の対象を特定できるように)
  (currentLedger||[]).forEach(r=>{
    if(!r || String(r.状態||'')!=='取り置き中' || String(r.取置元種別||'')!=='開始前在庫') return;
    if(String(r.引当段階||'')==='現物確認済み') return; // 変換済み(移行完了)の行は候補に出し続けない
    const key=取り置き_行キー_(r);
    const 注文=orderQty[key]||0, 確保=currentActive[key]||0;
    out.push({入力キー:key,種別:'旧開始前在庫',取置ID:String(r.取置ID||''),受注番号:String(r.受注番号||''),
      商品コード:String(r.商品コード||''),SKU:String(r.SKU||''),
      注文数:注文,確保済み:確保,不足:Math.max(0,注文-確保),旧記録:'',
      移行数量:取り置き_整数_(r.取り置き数量),選択:'',メモ:''});
    seen.add(key);
  });
  // 2) 消えた確保: 基準(最後の旧④)の確保 - 現在の確保 の正差分。現役注文だけ。
  const baseActive=activeQty(baselineLedger);
  Object.keys(baseActive).forEach(key=>{
    if(!(key in orderQty)) return; // 現役注文に無い=過去の幽霊確保。自動採用しない
    if(seen.has(key)) return;      // 開始前在庫として既に候補化済み
    const before=baseActive[key], now=currentActive[key]||0, diff=before-now;
    if(diff<=0) return;
    const sample=(baselineLedger||[]).find(r=>取り置き_行キー_(r)===key)||{};
    const 注文=orderQty[key]||0, 不足=Math.max(0,注文-now);
    out.push({入力キー:key,種別:'消えた確保',取置ID:'',受注番号:String(sample.受注番号||''),
      商品コード:String(sample.商品コード||''),SKU:String(sample.SKU||''),
      注文数:注文,確保済み:now,不足:不足,旧記録:before,
      移行数量:Math.min(diff,不足),選択:'',メモ:''});
  });
  return out;
}

// 反映計画: 選択を入力キー単位で検証し、適用可能な変更だけで新台帳を作る。
// 1キーのエラーは他キーを破棄しない(errorsへ理由を残す)。保留は台帳を変えない。
function 現物確認移行_反映計画_(candidates, choices, ledger, orders, supplies, now){
  const rows=(ledger||[]).map(r=>Object.assign({},r));
  const orderQty=現物移行_注文数量_(orders);
  const errors=[], applied=[], 保留=[], diff=[];
  const byKey={};
  (candidates||[]).forEach(c=>{ (byKey[c.入力キー]=byKey[c.入力キー]||[]).push(c); });
  Object.keys(choices||{}).forEach(key=>{
    const choice=choices[key]||{}, 選択=String(choice.選択||'').trim();
    const cands=byKey[key]||[];
    if(!選択 || 選択==='保留'){ if(選択==='保留') cands.forEach(c=>保留.push(c)); return; }
    if(現物移行_CFG.選択肢.indexOf(選択)<0){ errors.push(key+': 選択が不正です('+選択+')'); return; }
    if(!cands.length){ errors.push(key+': 候補にありません'); return; }
    if(選択==='解除'){
      // 旧開始前在庫の行だけを手動解除にする。EMS由来の確保はこのフローでは触らない
      const targets=rows.filter(r=>String(r.状態||'')==='取り置き中'
        && String(r.取置元種別||'')==='開始前在庫' && 取り置き_行キー_(r)===key);
      if(!targets.length){ errors.push(key+': 解除できる旧開始前在庫がありません'); return; }
      targets.forEach(r=>{ r.状態='手動解除'; r.更新日時=now;
        r['終了理由・メモ']=String(r['終了理由・メモ']||'')+(r['終了理由・メモ']?' / ':'')+'現物確認移行で解除'; });
      applied.push({入力キー:key,選択,行数:targets.length});
      diff.push({入力キー:key,内容:'解除 '+targets.length+'行'});
      return;
    }
    // 現物確認済みにする
    const holdRows=rows.filter(r=>String(r.状態||'')==='取り置き中'
      && String(r.取置元種別||'')==='開始前在庫' && 取り置き_行キー_(r)===key);
    if(holdRows.length){
      // 旧開始前在庫: 数量を加算せず段階だけ現物確認済みへ変換。棚の現物の箱は締め済みのため供給解放
      holdRows.forEach(r=>{ r.引当段階='現物確認済み'; r.供給処理='供給解放';
        r.現物確認日時=now; r.更新日時=now; });
      applied.push({入力キー:key,選択,行数:holdRows.length});
      diff.push({入力キー:key,内容:'開始前在庫'+holdRows.length+'行を現物確認済みへ変換'});
      return;
    }
    // 消えた確保の復元: 注文数量-現在確保 を上限に新規現物行を1行作る
    const cand=cands.find(c=>c.種別==='消えた確保');
    if(!cand){ errors.push(key+': 復元候補がありません'); return; }
    const want=取り置き_整数_(choice.数量!=null?choice.数量:cand.移行数量);
    if(!want){ errors.push(key+': 移行数量は正の整数で入力してください'); return; }
    const currentHeld=rows.filter(r=>String(r.状態||'')==='取り置き中'&&取り置き_行キー_(r)===key)
      .reduce((n,r)=>n+取り置き_整数_(r.取り置き数量),0);
    const cap=Math.max(0,(orderQty[key]||0)-currentHeld);
    if(want>cap){ errors.push(key+': 復元'+want+'個は注文数量を超えます(現在確保'+currentHeld+'/注文'+(orderQty[key]||0)+'。上限'+cap+'個)'); return; }
    rows.push({取置ID:'MIG|'+key+'|'+want,状態:'取り置き中',受注番号:cand.受注番号,商品コード:cand.商品コード,SKU:cand.SKU,
      取り置き数量:want,取置元種別:'EMS',引当段階:'現物確認済み',元EMS番号:'',元EMS商品コード:'',元取置ID:'',
      供給処理:'供給解放',現物確認日時:now,登録日時:now,更新日時:now,戻し処理結果:'',
      '終了理由・メモ':'現物確認移行で復元(基準'+現物移行_CFG.基準シート.slice(-13)+')'});
    applied.push({入力キー:key,選択,行数:1,数量:want});
    diff.push({入力キー:key,内容:'現物'+want+'個を復元'});
  });
  return {rows,errors,applied,保留,diff};
}

// ===== GAS入口 =====

function 現物移行_受注読む_(){
  const recv=SpreadsheetApp.getActive().getSheetByName(HIKIATE_CFG.受注);
  if(!recv) throw new Error('受注明細がありません');
  const M=列マップ_(recv), values=recv.getDataRange().getValues(), out=[];
  for(let i=M.hr;i<values.length;i++){
    const row=values[i], ban=String(row[M.番号]||'').trim(), qty=Number(row[M.個数])||0;
    if(!ban || qty<=0) continue;
    out.push({ban,code:String(row[M.コード]||''),sku:M.SKU>=0?String(row[M.SKU]||''):'',qty});
  }
  return out;
}

function 現物確認移行を作成(){ 直列_(現物確認移行を作成本体_); }
function 現物確認移行を作成本体_(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  const base=ss.getSheetByName(現物移行_CFG.基準シート);
  if(!base){ ui.alert('基準シート「'+現物移行_CFG.基準シート+'」が見つかりません。比較基準(最後の旧④状態)が無いと移行候補を作れません。'); return; }
  const baseline=取り置き_表を読む_(現物移行_CFG.基準シート,TORIOKI_CFG.台帳HDR);
  const ledger=取り置き台帳_読む_();
  const orders=現物移行_受注読む_();
  const candidates=現物確認移行_候補計算_(ledger,baseline,orders);
  // 前回入力(選択・移行数量・メモ)を入力キーで引き継ぐ(洗い替えで消さない)。旧レイアウトからも読める
  let prev=[];
  try{ prev=取り置き_表を読む_(現物移行_CFG.シート,現物移行_CFG.HDR); }
  catch(e){ try{ prev=取り置き_表を読む_(現物移行_CFG.シート,現物移行_CFG.旧HDR); }catch(e2){} }
  const prevByKey={}; prev.forEach(r=>{ const k=String(r.入力キー||''); if(k) prevByKey[k]=r; });
  candidates.forEach(c=>{
    const p=prevByKey[c.入力キー]; if(!p) return;
    if(String(p.選択||'').trim()) c.選択=String(p.選択).trim();
    if(String(p.移行数量||'').trim()!=='') c.移行数量=p.移行数量;
    if(String(p.メモ||'').trim()) c.メモ=String(p.メモ).trim();
  });
  取り置き_表を保存_(現物移行_CFG.シート,現物移行_CFG.HDR,candidates);
  const sh=ss.getSheetByName(現物移行_CFG.シート);
  if(candidates.length){
    const col=現物移行_CFG.HDR.indexOf('選択')+1;
    sh.getRange(2,col,candidates.length,1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(現物移行_CFG.選択肢.slice(),true).setAllowInvalid(false).build());
  }
  // 反映時の署名再検証用に、候補作成時点の台帳署名を保存する
  PropertiesService.getDocumentProperties().setProperty('現物確認移行_署名',全件再計算_署名_(ledger));
  const 旧=candidates.filter(c=>c.種別==='旧開始前在庫').length, 消=candidates.filter(c=>c.種別==='消えた確保').length;
  ui.alert('現物確認移行を作成しました',
    '旧開始前在庫 '+旧+'件 ／ 消えた確保 '+消+'件\n\n'
    +'「選択」列で 現物確認済みにする／解除／保留 を選び、復元は「移行数量」を確認してから\n'
    +'「✅ 現物確認移行を反映」を実行してください。\n'
    +'※旧開始前在庫が残っている間は全件再計算の反映は停止します。',ui.ButtonSet.OK);
}

function 現物確認移行を反映(){ 直列_(現物確認移行を反映本体_); }
function 現物確認移行を反映本体_(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  let inputs=[]; try{ inputs=取り置き_表を読む_(現物移行_CFG.シート,現物移行_CFG.HDR); }
  catch(e){ ui.alert('現物確認移行シートがありません。先に「🔄 現物確認移行を作成」を実行してください。'); return; }
  const choices={};
  inputs.forEach(r=>{ const k=String(r.入力キー||''),sel=String(r.選択||'').trim();
    if(k&&sel) choices[k]={選択:sel,数量:r.移行数量}; });
  if(!Object.keys(choices).length){ ui.alert('選択が1件もありません。「選択」列を入力してください。'); return; }
  const ledger=取り置き台帳_読む_();
  // 署名再検証: 候補作成後に台帳が変わっていたら中止(古い候補で上書きしない)
  const saved=PropertiesService.getDocumentProperties().getProperty('現物確認移行_署名');
  if(!saved || saved!==全件再計算_署名_(ledger)){
    ui.alert('反映を中止しました','候補作成後に取り置き台帳が変わっています。「🔄 現物確認移行を作成」からやり直してください。',ui.ButtonSet.OK); return;
  }
  const baseline=取り置き_表を読む_(現物移行_CFG.基準シート,TORIOKI_CFG.台帳HDR);
  const orders=現物移行_受注読む_();
  const candidates=現物確認移行_候補計算_(ledger,baseline,orders);
  const now=new Date();
  const plan=現物確認移行_反映計画_(candidates,choices,ledger,orders,[],now);
  if(!plan.applied.length){
    ui.alert('適用できる変更がありません',plan.errors.length?plan.errors.join('\n'):'選択内容を確認してください。',ui.ButtonSet.OK); return;
  }
  const 確認=ui.alert('現物確認移行を反映します',
    '適用 '+plan.applied.length+'キー ／ 保留 '+plan.保留.length+'件 ／ エラー '+plan.errors.length+'件\n'
    +(plan.errors.length?'\nエラー(この键は適用されません):\n'+plan.errors.slice(0,8).join('\n')+'\n':'')
    +'\nバックアップを作成してから台帳を保存します。続けますか？',ui.ButtonSet.OK_CANCEL);
  if(確認!==ui.Button.OK) return;
  // バックアップ失敗時は保存処理を開始しない
  const timestamp=Utilities.formatDate(now,'Asia/Tokyo','yyyyMMdd_HHmmss');
  try{
    全件再計算_ファイルをバックアップ_(ss.getId(),ss.getName()+'_現物移行前_'+timestamp);
    const copy=ss.getSheetByName(TORIOKI_CFG.台帳).copyTo(ss);
    copy.setName(('取り置き台帳_移行前_'+timestamp).slice(0,99)); copy.hideSheet();
  }catch(e){ ui.alert('バックアップに失敗したため中止しました:\n'+e.message); return; }
  取り置き台帳_保存_(plan.rows);
  // 差分を監査シートへ追記(明示消去以外で消さない)
  let dsh=ss.getSheetByName(現物移行_CFG.差分);
  if(!dsh){ dsh=ss.insertSheet(現物移行_CFG.差分); dsh.getRange(1,1,1,3).setValues([['日時','入力キー','内容']]).setFontWeight('bold'); dsh.hideSheet(); }
  if(plan.diff.length){
    const stamp=Utilities.formatDate(now,'Asia/Tokyo','yyyy-MM-dd HH:mm:ss');
    dsh.getRange(dsh.getLastRow()+1,1,plan.diff.length,3).setValues(plan.diff.map(d=>[stamp,d.入力キー,d.内容]));
  }
  PropertiesService.getDocumentProperties().deleteProperty('現物確認移行_署名');
  現物確認移行を作成本体_(); // 残候補(保留・エラー・未選択)で作り直す
  ss.toast('適用'+plan.applied.length+'キー / 保留'+plan.保留.length+' / エラー'+plan.errors.length,'現物確認移行',8);
}

// ===== 一度きりの現物確認移行 =====
// 旧「開始前在庫」61行の仕分け(現物確認済みへ変換/解除/保留)と、
// 全件再計算で消えた確保(基準バックアップとの正差分)の復元を、一つの移行シートで行う。
// 仕様: docs/superpowers/specs/2026-07-22-preallocation-physical-lock-design.md §13
// 比較基準は「取り置き台帳_全件再計算前_20260721_123350」(最後の旧④状態)だけを使う。

const 現物移行_CFG = Object.freeze({
  シート:'現物確認移行', 差分:'現物確認移行_差分',
  基準シート:'取り置き台帳_全件再計算前_20260721_123350',
  HDR:['入力キー','種別','受注番号','商品コード','SKU','以前','現在','差','移行数量','選択','メモ'],
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
    const key=取り置き_行キー_(r);
    out.push({入力キー:key,種別:'旧開始前在庫',取置ID:String(r.取置ID||''),受注番号:String(r.受注番号||''),
      商品コード:String(r.商品コード||''),SKU:String(r.SKU||''),以前:'',現在:取り置き_整数_(r.取り置き数量),差:'',
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
    out.push({入力キー:key,種別:'消えた確保',取置ID:'',受注番号:String(sample.受注番号||''),
      商品コード:String(sample.商品コード||''),SKU:String(sample.SKU||''),以前:before,現在:now,差:diff,
      移行数量:diff,選択:'',メモ:''});
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

const TORIOKI_CFG = Object.freeze({
  台帳:'取り置き台帳', 初期:'取り置き登録', 要確認:'取り置き要確認', 戻し:'キャンセル戻し確認', Yahoo候補:'Yahoo戻し候補', 移動:'EMS在庫移動台帳',
  台帳HDR:['取置ID','状態','受注番号','商品コード','SKU','取り置き数量','取置元種別','元EMS番号','元EMS商品コード','元取置ID','登録日時','更新日時','戻し処理結果','終了理由・メモ','引当段階','EMS到着予定日','現物確認日時','現物確認メモ','供給控除EMS','引当系譜ID','引当系譜数量','供給処理'],
  初期HDR:['取置ID','受注番号','氏名','商品コード','SKU','注文数量','現在の状態','受注ステータス','旧入荷日','旧EMS','台帳確保','棚確認','現物取り置き数量','メモ','判定','要対応','処理'],
  要確認HDR:['取置ID','受注番号','商品コード','理由'],
  移動HDR:['処理ID','EMS番号','商品コード','数量','移動先','処理日時']
});
const 戻しHDR=['取置ID','受注番号','商品コード','数量','元EMS番号','現物確認','メモ'];
// 手入力の永続保存(非表示シート)。GoQ取込・更新・引当・全件再計算で入力を消さないための内部保存。仕様§8
const TORIOKI_INPUT_CFG=Object.freeze({シート:'取り置き入力保存',履歴:'取り置き入力履歴',
  HDR:['入力キー','受注番号','SKU','商品コード','棚確認','取り置きメモ','確認メモ','注文作業メモ','未反映現物確認数量','入力エラー','最終表示日時','更新日時']});

// 入力キー=受注番号|正規化SKU(行キーと同じ規則)。取置IDと違い受注明細の再取込で揺れない
function 取り置き_入力キー_(row){ return 取り置き_行キー_(row); }

// 洗い替え前のシート入力を保存rowsへupsertし、生成行へ復元する(純粋関数)。
// 生成行に無いキーの保存も消さない(注文が一時的に一覧から消えても入力を失わない)。
function 取り置き_入力保存マージ_(generatedRows, savedRows, sheetRows, now){
  const store={};
  (savedRows||[]).forEach(r=>{ const k=String(r&&r.入力キー||''); if(k) store[k]=Object.assign({},r); });
  (sheetRows||[]).forEach(r=>{
    const k=取り置き_入力キー_(r); if(!k||k==='|') return;
    const cur=store[k]||(store[k]={入力キー:k});
    cur.受注番号=String(r.受注番号||cur.受注番号||''); cur.SKU=String(r.SKU||cur.SKU||''); cur.商品コード=String(r.商品コード||cur.商品コード||'');
    const 棚=String(r.棚確認==null?'':r.棚確認).trim(); if(棚) cur.棚確認=棚;
    const memo=String(r.メモ==null?'':r.メモ).trim(); if(memo) cur.取り置きメモ=memo;
    const qty=r.現物取り置き数量; if(qty!=null&&String(qty).trim()!=='') cur.未反映現物確認数量=qty;
    cur.更新日時=now||'';
  });
  const rows=(generatedRows||[]).map(g=>{
    const k=取り置き_入力キー_(g), s=store[k];
    if(!s) return g;
    s.最終表示日時=now||'';
    return Object.assign({},g,{
      棚確認:String(g.棚確認||'').trim()||String(s.棚確認||''),
      メモ:String(g.メモ||'').trim()||String(s.取り置きメモ||''),
      現物取り置き数量:(g.現物取り置き数量!=null&&String(g.現物取り置き数量).trim()!=='')?g.現物取り置き数量
        :(s.未反映現物確認数量!=null&&String(s.未反映現物確認数量).trim()!==''?s.未反映現物確認数量:g.現物取り置き数量)
    });
  });
  return {rows,store};
}

// 注文(同一受注番号の候補行の配列)が表示モードの対象か(純粋関数)。既定「要作業」は
// 部分在庫・要棚確認/棚戻し待ち/要対応あり・現物ありの希望日待ちだけを出す。仕様§10
function 取り置き_作業対象判定_(rows, viewMode){
  const arr=rows||[], mode=String(viewMode||'要作業');
  if(mode==='すべて') return true;
  const has=f=>arr.some(r=>r&&f(r));
  const 現物あり=has(r=>(Number(r.台帳確保数)||0)>0||String(r.棚確認||'')==='部分在庫'
    ||(r.現物取り置き数量!=null&&String(r.現物取り置き数量).trim()!==''&&Number(r.現物取り置き数量)>0));
  const 部分=has(r=>String(r.現在の状態||'')==='部分在庫');
  const 希望=has(r=>String(r.現在の状態||'')==='希望日待ち');
  if(mode==='部分在庫') return 部分;
  if(mode==='希望日待ち・現物あり') return 希望&&現物あり;
  if(mode==='先行引当') return has(r=>/先行/.test(String(r.台帳確保||'')+String(r.状態の理由||'')));
  const 要作業=has(r=>String(r.判定||'')==='要棚確認'||String(r.判定||'')==='棚戻し待ち'||String(r.要対応||'').trim()!=='');
  return 部分||要作業||(希望&&現物あり);
}
// 取り置き登録の「棚確認」プルダウン。出荷済み/未着/予約は数量なし(登録しない)の目印
const TORIOKI_棚確認=Object.freeze(['発送待ち','部分在庫','出荷済み','未着','予約']);
const TORIOKI_戻し処理=Object.freeze(['棚へ戻した','現物なし']);

function 取り置き_棚確認書式定義_(棚確認列, 開始行){
  let n=棚確認列, 列記号='';
  while(n>0){ n--; 列記号=String.fromCharCode(65+n%26)+列記号; n=Math.floor(n/26); }
  return [
    {値:'発送待ち',背景:'#cfe2f3'},
    {値:'部分在庫',背景:'#d9ead3'},
    {値:'出荷済み',背景:'#f4cccc',文字色:'#990000',太字:true},
    {値:'未着',背景:'#d9d9d9'},
    {値:'予約',背景:'#d9d2e9'}
  ].map(def=>Object.assign({条件:'=$'+列記号+開始行+'="'+def.値+'"'},def));
}

function 取り置き_棚確認書式を設定_(sh, rowCount){
  if(rowCount<=0){ sh.setConditionalFormatRules([]); return; }
  const 棚確認列=TORIOKI_CFG.初期HDR.indexOf('棚確認')+1, 幅=TORIOKI_CFG.初期HDR.length;
  const target=sh.getRange(2,1,rowCount,幅);
  const rules=取り置き_棚確認書式定義_(棚確認列,2).map(def=>{
    let b=SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(def.条件)
      .setBackground(def.背景)
      .setRanges([target]);
    if(def.文字色) b=b.setFontColor(def.文字色);
    if(def.太字) b=b.setBold(true);
    return b.build();
  });
  sh.setConditionalFormatRules(rules);
}

function 取り置き_登録行書式を更新_(sh, candidates){
  const rows=candidates||[], 幅=TORIOKI_CFG.初期HDR.length;
  const 管理行数=Math.max(0,sh.getMaxRows()-1);
  if(管理行数>0){
    const 管理範囲=sh.getRange(2,1,管理行数,幅);
    管理範囲.setBackground(null);
    管理範囲.setBorder(false,false,false,false,false,false);
  }
  注文罫線_(sh,2,1);
  if(rows.length){
    sh.getRange(2,1,rows.length,幅).setBackgrounds(
      rows.map(c=>new Array(幅).fill(c.判定==='棚戻し待ち'? '#f4cccc' : c.判定==='要棚確認'? '#fff2cc' : c.判定==='即納'? '#cfe2f3' : c.判定==='別ルート'? '#fce5cd' : null)));
  }
  取り置き_棚確認書式を設定_(sh,rows.length);
}

// 商品名・選択肢に「予約」と未来の発売予定がある場合だけ自動予約にする。
// 過去日・日付不明は入荷済みの可能性があるため自動では除外しない。
function 取り置き_予約判定_(選択肢, 商品名, today){
  const text=String(選択肢||'')+' '+String(商品名||'');
  if(text.indexOf('予約')<0) return false;
  const base=today instanceof Date && !isNaN(today.getTime())? today : new Date();
  const full=text.match(/(20\d{2})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
  if(full){
    const y=Number(full[1]), month=Number(full[2]), day=Number(full[3]);
    const date=new Date(y,month-1,day);
    if(date.getFullYear()!==y || date.getMonth()!==month-1 || date.getDate()!==day) return false;
    return date.getTime()>base.getTime();
  }
  const m=text.match(/(?:^|[^\d年])(\d{1,2})月(\s*末)?/);
  if(!m) return false;
  const month=Number(m[1]), currentMonth=base.getMonth()+1;
  if(month<1 || month>12) return false;
  if(month>currentMonth) return true;
  if(month!==currentMonth || !m[2]) return false;
  return new Date(base.getFullYear(),month,0).getTime()>base.getTime();
}

// 予約表記より強い「現物が既に来ている」証拠。早着した予約品を非表示のままにしない。
function 取り置き_入荷証拠あり_(row){
  const r=row||{};
  const qtyText=String(r.現物取り置き数量==null?'':r.現物取り置き数量).trim();
  return String(r.旧入荷日==null?'':r.旧入荷日).trim()!=='' ||
    String(r.旧EMS==null?'':r.旧EMS).trim()!=='' ||
    Number(r.台帳確保数)>0 ||
    (qtyText!=='' && Number(qtyText)>0);
}
const Yahoo候補HDR=['取置ID','商品コード','数量','元EMS番号','処理ID','確認'];

function CSV行を受注行オブジェクトへ_(header, rows){
  const head=(header||[]).map(v=>String(v||'').trim()), index=name=>head.indexOf(name);
  const cBan=index('受注番号'), cStatus=index('受注ステータス'), cCode=index('商品コード'), cQty=index('個数');
  const cSku=index('商品SKU')>=0?index('商品SKU'):index('SKU');
  const missing=[];
  if(cBan<0) missing.push('受注番号'); if(cStatus<0) missing.push('受注ステータス');
  if(cCode<0) missing.push('商品コード'); if(cQty<0) missing.push('個数'); if(cSku<0) missing.push('商品SKU/SKU');
  if(missing.length) throw new Error('全ステータスCSVの見出し不足: '+missing.join(','));
  return (rows||[]).map(row=>({受注番号:String(row[cBan]||'').replace(/^niyantarose-/i,''),受注ステータス:String(row[cStatus]||''),
    商品コード:String(row[cCode]||''),SKU:String(row[cSku]||''),個数:Number(row[cQty])||0}));
}

// 状態遷移の理由は既存の「終了理由・メモ」を消さずに追記する(棚の場所などの手書きメモを守る)。
// 同じ理由が既に入っていれば足さない=取込のたびにメモが伸びない。
function 取り置き_メモ追記_(existing, addition){
  const cur=String(existing==null?'':existing).trim(), add=String(addition==null?'':addition).trim();
  if(!add) return cur;
  if(!cur) return add;
  if(cur.indexOf(add)>=0) return cur;
  return cur+' / '+add;
}

function 取り置き_CSV遷移計画_(csvRows, ledgerRows, now){
  const groups={}, errors=[], review=[], counts={棚戻し待ち:0,棚戻し数量:0,発送済み:0};
  (csvRows||[]).forEach(r=>{
    const ban=String(r.受注番号||'').replace(/^niyantarose-/i,'').trim();
    const key=取り置き_行キー_({ban,code:r.商品コード,sku:r.SKU});
    if(!ban || !取り置き_商品コード_(r.SKU,r.商品コード)) return;
    if(!groups[key]) groups[key]={発送済み:false,キャンセル:false,生存数量:false,数量既知:false,有効数量:0,処理済数量:0,キャンセル数量:0};
    const group=groups[key], status=String(r.受注ステータス||'');
    const shipped=/処理済|発送済|出荷済/.test(status), cancelled=/キャンセル/.test(status);
    const rawQty=r.個数, qtyKnown=rawQty!==undefined && rawQty!==null && String(rawQty).trim()!=='' && isFinite(Number(rawQty));
    const qty=qtyKnown?Number(rawQty):0;
    group.発送済み=group.発送済み||shipped;
    group.キャンセル=group.キャンセル||cancelled;
    group.数量既知=group.数量既知||qtyKnown;
    // キャンセル行の正の数量を注文数へ混ぜない(有効・処理済み・キャンセルを別々に集計)
    if(qtyKnown){
      if(cancelled) group.キャンセル数量+=qty;
      else if(shipped) group.処理済数量+=qty;
      else group.有効数量+=qty;
    }
    // 同じ受注商品が分割され、数量0のキャンセル行と数量1以上の生きた行が共存する場合、
    // 商品全体の取り置きは解除しない。全分割行が0になったときだけ棚戻しへ進める。
    if(qtyKnown && qty>0 && !shipped && !cancelled) group.生存数量=true;
  });
  const statuses={};
  Object.keys(groups).forEach(key=>{
    const group=groups[key];
    if(group.発送済み && (group.キャンセル||group.生存数量)){
      errors.push('同じ受注行にステータス競合: '+key); return;
    }
    if(group.発送済み){ statuses[key]={状態:'発送済み',理由:'CSV処理済み'}; return; }
    if(group.生存数量){ statuses[key]={状態:'継続',理由:''}; return; }
    if(group.キャンセル){ statuses[key]={状態:'キャンセル',理由:'CSV注文キャンセル'}; return; }
    if(group.数量既知 && group.有効数量+group.処理済数量<=0){ statuses[key]={状態:'キャンセル',理由:'CSV数量0'}; return; }
    statuses[key]={状態:'継続',理由:''};
  });
  if(errors.length) return {rows:[],review,errors,counts};
  const rows=(ledgerRows||[]).map(r=>Object.assign({},r));
  // 【先に超過分を分離する】処理済みと数量減が同じCSVで来ても、出荷していない超過分を
  // 発送済みへ巻き込まず棚戻し待ちへ残すため、状態遷移より前に行う。
  //   継続:     注文数=有効数量           → 取り置き中(注文数)+棚戻し(超過)
  //   発送済み: 注文数=処理済数量+有効数量 → 発送済み(注文数)+棚戻し(超過)
  //   キャンセル: 遷移側で全数棚戻し
  // 新しい登録から先に外して古い確定を守り、行の途中に掛かる場合は分割する。
  const 分離キー={};
  rows.forEach((r,index)=>{
    if(r.状態!==TORIOKI_STATUS.ACTIVE) return;
    const key=取り置き_行キー_(r), st=statuses[key];
    if(!st || (st.状態!=='継続' && st.状態!=='発送済み')) return;
    if(!groups[key] || !groups[key].数量既知) return; // 数量を読めないCSVでは解除しない
    (分離キー[key]=分離キー[key]||[]).push(index);
  });
  const 既存ID=new Set(rows.map(r=>String(r.取置ID||'')));
  const 時刻=v=>{ if(v instanceof Date) return v.getTime(); const t=Date.parse(String(v||'')); return isFinite(t)?t:0; };
  Object.keys(分離キー).forEach(key=>{
    const indexes=分離キー[key], group=groups[key];
    const 注文数=Math.max(0, group.有効数量+(statuses[key].状態==='発送済み'? group.処理済数量:0));
    const 確保=indexes.reduce((sum,i)=>sum+取り置き_整数_(rows[i].取り置き数量),0);
    let excess=確保-注文数;
    if(excess<=0) return;
    const 理由='CSV数量減(注文'+注文数+'/確保'+確保+')（棚戻し待ち）';
    indexes.sort((a,b)=>時刻(rows[b].登録日時)-時刻(rows[a].登録日時)
      || String(rows[b].取置ID||'').localeCompare(String(rows[a].取置ID||'')));
    for(const i of indexes){
      if(excess<=0) break;
      const qty=取り置き_整数_(rows[i].取り置き数量);
      if(qty<=excess){
        rows[i]=Object.assign({},rows[i],{状態:TORIOKI_STATUS.RETURN,戻し処理結果:TORIOKI_RETURN.UNCHECKED,
          '終了理由・メモ':取り置き_メモ追記_(rows[i]['終了理由・メモ'],理由),更新日時:now});
        excess-=qty; counts.棚戻し待ち++; counts.棚戻し数量+=qty;
      } else {
        const 分割元=rows[i];
        rows[i]=Object.assign({},分割元,{取り置き数量:qty-excess,更新日時:now});
        let id=String(分割元.取置ID||'')+'|RTN', n=1;
        while(既存ID.has(id)){ n++; id=String(分割元.取置ID||'')+'|RTN'+n; }
        既存ID.add(id);
        rows.push(Object.assign({},分割元,{取置ID:id,取り置き数量:excess,状態:TORIOKI_STATUS.RETURN,
          戻し処理結果:TORIOKI_RETURN.UNCHECKED,'終了理由・メモ':取り置き_メモ追記_(分割元['終了理由・メモ'],理由),
          登録日時:now,更新日時:now}));
        counts.棚戻し待ち++; counts.棚戻し数量+=excess;
        excess=0;
      }
    }
  });
  // 【状態遷移】分離後に残った取り置き中だけを対象にする
  rows.forEach(copy=>{
    if(copy.状態!==TORIOKI_STATUS.ACTIVE) return;
    const key=取り置き_行キー_(copy), next=statuses[key];
    if(next&&next.状態==='発送済み'){
      copy.状態=TORIOKI_STATUS.SHIPPED; copy.更新日時=now;
      copy['終了理由・メモ']=取り置き_メモ追記_(copy['終了理由・メモ'],next.理由); counts.発送済み++;
    }
    else if(next&&next.状態==='キャンセル'){
      copy.状態=TORIOKI_STATUS.RETURN; copy.戻し処理結果=TORIOKI_RETURN.UNCHECKED; copy.更新日時=now;
      copy['終了理由・メモ']=取り置き_メモ追記_(copy['終了理由・メモ'],next.理由+'（棚戻し待ち）');
      counts.棚戻し待ち++; counts.棚戻し数量+=取り置き_整数_(copy.取り置き数量);
    }
    else if(!next) review.push({取置ID:copy.取置ID,受注番号:copy.受注番号,商品コード:copy.商品コード,理由:'最新CSVに受注行なし'});
  });
  return {rows,review,errors:[],counts};
}

function 取り置き_CSV遷移を反映_(plan){
  if(plan.errors && plan.errors.length) throw new Error(plan.errors.join('\n'));
  取り置き台帳_保存_(plan.rows||[]);
  取り置き_表を保存_(TORIOKI_CFG.要確認,TORIOKI_CFG.要確認HDR,plan.review||[]);
}

// sources: [{状態:'部分在庫', bans:Set}, ...] 一覧シート由来の受注番号集合(表示用の状態ラベル付き)。
// 未入金滞留中に同一商品の新箱が来ても二重引当にならないよう、棚に現物がある注文は全て候補にする。
// 同じ受注×商品の分割行(例: 同一商品を2行に分けた注文)は注文数量を合算して1候補にする。
// 並びはsourcesの順(出荷可能→出荷GO未入金→部分在庫→希望日待ち)＝棚で数えやすいグループ順。
function 取り置き_初期候補_(orders, sources){
  const list=(sources||[]).filter(s=>s && s.bans && typeof s.bans.has==='function');
  const stateOf=ban=>{ for(const s of list){ if(s.bans.has(String(ban))) return String(s.状態||''); } return ''; };
  const byKey={}, keys=[];
  (orders||[]).forEach(o=>{
    const state=stateOf(o.ban); if(!state) return;
    const key=取り置き_ID部_(o); // 取置ID(INIT|…)の互換のため従来3部形式で束ねる
    if(!byKey[key]){ byKey[key]={o,qty:0,state,入荷日:'',EMS:'',予約:false,ステータス一覧:[]}; keys.push(key); }
    byKey[key].qty+=Number(o.qty)||0;
    // 旧帳簿の着済情報(入荷日スタンプ/EMS番号)。棚を確認すべき行の目印として表示する
    if(!byKey[key].入荷日){ const d=ymd_(o.入荷日); if(d) byKey[key].入荷日=d; }
    if(!byKey[key].EMS && String(o.EMS||'').trim()) byKey[key].EMS=String(o.EMS).trim();
    if(o.予約) byKey[key].予約=true;
    const status=String(o.ステータス||'').trim();
    if(status && byKey[key].ステータス一覧.indexOf(status)<0) byKey[key].ステータス一覧.push(status);
  });
  const rank={}; list.forEach((s,i)=>{ rank[String(s.状態||'')]=i; });
  return keys
    .sort((a,b)=>(rank[byKey[a].state]-rank[byKey[b].state]) || (a<b?-1:a>b?1:0))
    .map(key=>{
      const c=byKey[key], o=c.o, statuses=c.ステータス一覧.join(' / ');
      const 自動予約=c.予約 && !c.入荷日 && !c.EMS && !/部分包装/.test(statuses);
      return {取置ID:'INIT|'+key,受注番号:String(o.ban),氏名:String(o.氏名||''),商品コード:取り置き_商品コード_(o.sku,o.code),SKU:String(o.sku||''),
        注文数量:c.qty,現在の状態:c.state,受注ステータス:statuses,旧入荷日:c.入荷日,旧EMS:c.EMS,
        棚確認:自動予約?'予約':'',自動予約,現物取り置き数量:'',メモ:'',判定:''};
    });
}

function 取り置き_注文単位で並べる_(rows){
  const groups=[], byBan=new Map();
  (rows||[]).forEach(row=>{
    const ban=String(row&&row.受注番号!=null?row.受注番号:'');
    let group=byBan.get(ban);
    if(!group){ group={rows:[],優先:2,i:groups.length}; byBan.set(ban,group); groups.push(group); }
    group.rows.push(row);
    if(row && row.判定==='棚戻し待ち') group.優先=0;
    else if(row && row.判定==='要棚確認') group.優先=Math.min(group.優先,1);
  });
  groups.sort((a,b)=>a.優先-b.優先 || a.i-b.i);
  return groups.reduce((out,group)=>out.concat(group.rows),[]);
}

function 取り置き_注文境界A1_(rows){
  return 注文境界A1_((rows||[]).map(row=>({ban:row.受注番号})),2,TORIOKI_CFG.初期HDR.length);
}

// ボタン・メニュー用。取り置き登録の値や色は触らず、受注番号のまとまりごとに罫線だけ引き直す。
function 取り置き登録の罫線を引く(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  const sh=ss.getSheetByName(TORIOKI_CFG.初期);
  if(!sh){ ui.alert('「'+TORIOKI_CFG.初期+'」タブがありません'); return; }
  const startRow=2, lastRow=sh.getLastRow(), width=TORIOKI_CFG.初期HDR.length;
  if(lastRow<startRow){ ui.alert('取り置き登録にデータがありません'); return; }
  const rowCount=lastRow-startRow+1, banCol=TORIOKI_CFG.初期HDR.indexOf('受注番号')+1;
  const target=sh.getRange(startRow,1,rowCount,width);
  target.setBorder(true,true,true,true,true,true,'#cccccc',SpreadsheetApp.BorderStyle.SOLID);
  const rows=sh.getRange(startRow,banCol,rowCount,1).getDisplayValues().map(row=>({受注番号:String(row[0]||'').trim()}));
  const borders=取り置き_注文境界A1_(rows);
  if(borders.length){
    sh.getRangeList(borders).setBorder(null,null,true,null,null,null,'#000000',SpreadsheetApp.BorderStyle.SOLID_THICK);
  }
  ui.alert('✅ 取り置き登録の罫線を引きました\n\n・データ '+rowCount+'行に薄い格子\n・受注番号の境目 '+borders.length+'件に太線');
}

// 再作成しても確定済みの数量が消えないよう、台帳の開始前在庫(取り置き中)を候補へ差し込む
function 取り置き_初期候補へ既存数量_(candidates, ledgerRows){
  const qtyById={};
  (ledgerRows||[]).forEach(r=>{
    if(r.取置元種別!=='開始前在庫' || r.状態!==TORIOKI_STATUS.ACTIVE) return;
    const id=String(r.取置ID||'');
    qtyById[id]=(qtyById[id]||0)+取り置き_整数_(r.取り置き数量);
  });
  return (candidates||[]).map(c=>{
    const id=String(c.取置ID||'');
    return Object.assign({},c,{現物取り置き数量: qtyById[id]!=null? qtyById[id] : c.現物取り置き数量});
  });
}

// 全行表示化v2(2026-07-20): 受注明細1行を取り置き登録のどの枠へ入れるかを1か所で決める純粋関数。
// 台湾/中国(別ルート)は区分に関わらず最優先、指定なしはほぼ即納なので即納扱い。対象外はどの枠にも入れない。
function 取り置き_明細区分_(kbn, 別ルート){
  if(別ルート) return '別ルート';
  if(kbn==='即納' || kbn==='指定なし') return '即納';
  if(kbn==='取り寄せ') return '取り寄せ';
  return '対象外';
}

// 全行表示化(2026-07-20): 取り寄せ候補がある注文の即納行を表示専用で差し込む。
// 受注明細と同じ水色(判定=即納)で「注文にどの現物が揃うか」を1画面で見せる(例: 10117188の見落とし事故対策)。
// 数量・棚確認は入力対象外(反映時にガード)。台帳へは載せない=全件再計算の開始前在庫を汚さない。
function 取り置き_即納行を付与_(candidates, sokunoOrders, sheetRows){
  const list=candidates||[];
  const bans=new Set(list.map(c=>String(c.受注番号)));
  const bySheet={};
  (sheetRows||[]).forEach(r=>{ const id=String(r&&r.取置ID||''); if(id) bySheet[id]=r; });
  const byKey={}, keys=[];
  (sokunoOrders||[]).forEach(o=>{
    if(!bans.has(String(o.ban))) return;
    const key=取り置き_ID部_(o); // 取置ID(即納|…)の互換のため従来3部形式で束ねる
    if(!byKey[key]){ byKey[key]={o,qty:0,ステータス一覧:[]}; keys.push(key); }
    byKey[key].qty+=Number(o.qty)||0;
    const status=String(o.ステータス||'').trim();
    if(status && byKey[key].ステータス一覧.indexOf(status)<0) byKey[key].ステータス一覧.push(status);
  });
  return list.concat(keys.map(key=>{
    const c=byKey[key], o=c.o, id='即納|'+key, prev=bySheet[id];
    return {取置ID:id,受注番号:String(o.ban),氏名:String(o.氏名||''),商品コード:取り置き_商品コード_(o.sku,o.code),SKU:String(o.sku||''),
      注文数量:c.qty,現在の状態:'即納',受注ステータス:c.ステータス一覧.join(' / '),旧入荷日:'',旧EMS:'',
      台帳確保:'',台帳確保数:0,棚確認:'',現物取り置き数量:'',メモ:prev?String(prev.メモ==null?'':prev.メモ).trim():'',判定:'即納',要対応:'',処理:''};
  }));
}

// 全行表示化v2(2026-07-20): 台湾/中国(別ルート)行を表示専用で差し込む。GoQで部分在庫=別ルートで
// 現物が到着した合図なのに候補から消えて気づけない、を防ぐ。確保は受注明細の入荷日(別ルート済数量)で
// 行うため、この画面では数量入力させない(台帳の開始前在庫と二重引当になる)。旧入荷日欄で確保状態が分かる。
function 取り置き_別ルート行を付与_(candidates, betsuOrders, sheetRows){
  const list=candidates||[];
  const bans=new Set(list.map(c=>String(c.受注番号)));
  const bySheet={};
  (sheetRows||[]).forEach(r=>{ const id=String(r&&r.取置ID||''); if(id) bySheet[id]=r; });
  const byKey={}, keys=[];
  (betsuOrders||[]).forEach(o=>{
    if(!bans.has(String(o.ban))) return;
    const key=取り置き_ID部_(o); // 取置ID(別ルート|…)の互換のため従来3部形式で束ねる
    if(!byKey[key]){ byKey[key]={o,qty:0,入荷日:'',ステータス一覧:[]}; keys.push(key); }
    byKey[key].qty+=Number(o.qty)||0;
    if(!byKey[key].入荷日){ const d=ymd_(o.入荷日); if(d) byKey[key].入荷日=d; }
    const status=String(o.ステータス||'').trim();
    if(status && byKey[key].ステータス一覧.indexOf(status)<0) byKey[key].ステータス一覧.push(status);
  });
  return list.concat(keys.map(key=>{
    const c=byKey[key], o=c.o, id='別ルート|'+key, prev=bySheet[id];
    return {取置ID:id,受注番号:String(o.ban),氏名:String(o.氏名||''),商品コード:取り置き_商品コード_(o.sku,o.code),SKU:String(o.sku||''),
      注文数量:c.qty,現在の状態:'別ルート',受注ステータス:c.ステータス一覧.join(' / '),旧入荷日:c.入荷日,旧EMS:'',
      台帳確保:'',台帳確保数:0,棚確認:'',現物取り置き数量:'',メモ:prev?String(prev.メモ==null?'':prev.メモ).trim():'',
      判定:'別ルート',要対応:c.入荷日?'':'入荷日を受注明細に入れて確保',処理:''};
  }));
}

// 棚まで見に行くべき行の自動判定。
// 旧帳簿が「着いているはず」(旧入荷日/旧EMSあり)なのに、数量も棚確認も未入力の行だけ「要棚確認」。
// 数量を入れる or 棚確認(出荷済み/未着など)を選べば解決扱い。
// 台帳(EMS引当等)が注文の全数を確保済みの行は、棚の現物が④の管理下にあるため対象外。
// 一部だけ確保済みの行は残りの現物を確かめるべきなので、旧情報が無くても要棚確認に残す。
function 取り置き_棚確認判定_(c){
  const qty=String(c&&c.現物取り置き数量==null?'':c.現物取り置き数量).trim();
  const check=String(c&&c.棚確認==null?'':c.棚確認).trim();
  if(qty!=='' || check!=='') return '';
  const ordered=Number(c&&c.注文数量)||0, secured=Number(c&&c.台帳確保数)||0;
  if(ordered>0 && secured>=ordered) return '';
  const 旧あり=String(c&&c.旧入荷日==null?'':c.旧入荷日).trim()!=='' || String(c&&c.旧EMS==null?'':c.旧EMS).trim()!=='';
  return (旧あり || secured>0)? '要棚確認' : '';
}

// 台帳確保(取り置き中のEMS引当等)を候補へ重ねる。
//   全数確保 → 入力欄を空へ戻す(開始前在庫として確定させない=EMS確保分の二重登録防止)
//   一部確保 → 「残り◯個要確認」を表示し、入力欄はそのまま(残り分だけ登録できる)
// 台帳の数量を現物取り置き数量へコピーはしない(コピーすると確定でEMS分が開始前在庫へ化ける)。
function 取り置き_台帳確保を適用_(candidates, securedByKey){
  return (candidates||[]).map(c=>{
    const secured=Number(securedByKey&&securedByKey[取り置き_行キー_(c)])||0;
    if(!secured) return Object.assign({},c,{台帳確保数:0,台帳確保:''});
    const ordered=Number(c.注文数量)||0, full=ordered>0 && secured>=ordered;
    return Object.assign({},c,{
      台帳確保数:secured,
      台帳確保: full? '台帳確保済み'+secured+'個'
                    : '台帳確保済み'+secured+'個／残り'+Math.max(0,ordered-secured)+'個要確認',
      現物取り置き数量: full? '' : c.現物取り置き数量
    });
  });
}

// 全行表示化v2(2026-07-20): 自動除外は「出荷GO(確定・発送作業中)」だけにする。
// 予約中は証拠が無くても隠さない=早着して棚にあるのに行が出ない(入れられない)を防ぐ。
// 候補は既にアクティブ注文の行だけなので、この緩和で純粋予約注文のノイズは増えない。
function 取り置き_登録絞り込み_(rows){
  return (rows||[]).filter(c=>{
    const st=String(c.受注ステータス||'');
    if(st.indexOf('出荷GO')>=0) return false;
    return true; // 予約中も表示。判断済み(出荷済み/未着/予約)は棚確認セルに残して条件付き書式で目立たせる
  });
}

// 全行表示化(2026-07-20): 判断済み(出荷済み/未着/予約)も隠さず、判断は棚確認列に残して条件付き書式で目立たせる。
// 記憶はシートの棚確認列の引き継ぎ(取り置き_登録シート引き継ぎ_)に一本化し、DocumentPropertiesの
// 非表示スイッチは廃止(セルを空にすれば次回から消える)。旧版が残した記憶は既定値として一度だけ復元する。
// 到着証拠が付いた「予約」だけは既定値を外し、要棚確認として棚へ出す(早着予約の隠れ事故防止: WTF-SMJ-28-2b)。
function 取り置き_棚確認記憶を適用_(candidates, store){
  const memo=store||{};
  const rows=(candidates||[]).map(c=>{
    let check=String(c.棚確認||'').trim();
    const stored=String(memo[String(c.取置ID||'')]||'');
    if(!check && stored) check=stored; // 旧・非表示スイッチからの一度きりの引き継ぎ
    if(check==='予約' && 取り置き_入荷証拠あり_(c)) check='';
    return Object.assign({},c,{棚確認:check});
  });
  return {rows, store:{}};
}

// 洗い替えの引き継ぎ: 台帳の確定数量を土台に、シートへ手入力済みの数量・メモを上書きで残す。
// 消えるのは「候補から外れた行(出荷済み・キャンセルで注文自体が消えた)」だけ。
function 取り置き_登録シート引き継ぎ_(candidates, sheetRows, ledgerRows){
  const withLedger=取り置き_初期候補へ既存数量_(candidates, ledgerRows);
  const bySheet={};
  (sheetRows||[]).forEach(r=>{ const id=String(r.取置ID||''); if(id) bySheet[id]=r; });
  return withLedger.map(c=>{
    const prev=bySheet[String(c.取置ID||'')];
    if(!prev) return c;
    const qty=prev.現物取り置き数量;
    let check=String(prev.棚確認==null?'':prev.棚確認).trim();
    let memo=String(prev.メモ==null?'':prev.メモ).trim();
    if(!check && TORIOKI_棚確認.indexOf(memo)>=0){ check=memo; memo=''; } // 旧メモの分類語はプルダウンへ移す
    return Object.assign({},c,{
      現物取り置き数量: (qty!=null && String(qty).trim()!=='')? qty : c.現物取り置き数量,
      棚確認: check!==''? check : c.棚確認,
      自動予約: check!==''? false : c.自動予約,
      メモ: memo!==''? memo : c.メモ
    });
  });
}

// 台帳の未確認キャンセルだけを、日常画面「取り置き登録」の赤い要対応行へ変換する。
// 取置IDで前回入力を引き継ぐため、CSV取込や手動更新で洗い替えても選択中の処理は消えない。
function 取り置き_棚戻し候補_(ledgerRows, sheetRows){
  const previous={};
  (sheetRows||[]).forEach(row=>{ const id=String(row.取置ID||''); if(id) previous[id]=row; });
  return (ledgerRows||[])
    .filter(row=>row.状態===TORIOKI_STATUS.RETURN && row.戻し処理結果===TORIOKI_RETURN.UNCHECKED)
    .map(row=>{
      const id=String(row.取置ID||''), prev=previous[id]||{}, qty=取り置き_整数_(row.取り置き数量);
      const reason=String(row['終了理由・メモ']||'キャンセル').trim();
      return {
        取置ID:id,受注番号:String(row.受注番号||''),氏名:String(prev.氏名||''),商品コード:String(row.商品コード||''),SKU:String(row.SKU||''),
        注文数量:qty,現在の状態:'キャンセル戻し',受注ステータス:reason,旧入荷日:'',旧EMS:String(row.元EMS番号||''),
        台帳確保:'確保済み'+qty+'個',台帳確保数:qty,棚確認:'',現物取り置き数量:'',
        メモ:String(prev.メモ||'').trim()||reason,判定:'棚戻し待ち',要対応:'棚へ戻す '+qty+'個',処理:String(prev.処理||'').trim()
      };
    });
}

function 取り置き_初期確定計画_(inputRows, existingRows, now){
  const errors=[], targets={}, inputIds=new Set();
  // 発送済み・解除済みになった開始前在庫行は履歴。再確定で復活も消滅もさせない(誤操作ガード)
  const lockedIds=new Set((existingRows||[]).filter(r=>r.取置元種別==='開始前在庫' && r.状態!==TORIOKI_STATUS.ACTIVE).map(r=>String(r.取置ID||'')));
  // ④のEMS確保分と合わせた超過を止める(確保済み行へ数量を入れると同じ現物の二重登録になる)
  const 確保=取り置き_台帳確保集計_(existingRows);
  (inputRows||[]).forEach((r,index)=>{
    const raw=r.現物取り置き数量, blank=raw==null || raw==='';
    // 表示専用行(即納・別ルート)は数量を入れさせない。入っていたら誤登録なので反映ごと止める。
    //   即納 = 全件再計算の開始前在庫を汚さない / 別ルート = 入荷日ベースの別ルート済数量と二重にしない
    const id=String(r.取置ID||'');
    if(id.indexOf('即納|')===0){
      if(!blank) errors.push('受注'+r.受注番号+' '+r.商品コード+': 即納行は表示専用です。数量は取り寄せ行へ入力してください');
      return;
    }
    if(id.indexOf('別ルート|')===0){
      if(!blank) errors.push('受注'+r.受注番号+' '+r.商品コード+': 台湾/中国(別ルート)行は表示専用です。確保は受注明細に入荷日を入れてください');
      return;
    }
    // 行単位適用(2026-07-22 仕様§8): この行のエラーは他行の反映を妨げない。
    // エラー行は inputIds へ入れない=既存の開始前在庫行も維持(誤解除防止)。
    const rowErrors=[];
    const entered=blank?0:Number(raw), ordered=Number(r.注文数量)||0;
    if(!blank && (!String(raw).trim() || !Number.isFinite(entered))) rowErrors.push('初期登録'+(index+2)+'行 / 受注'+r.受注番号+': 数量は数値で入力');
    else if(entered<0 || !Number.isInteger(entered)) rowErrors.push('初期登録'+(index+2)+'行: 数量は0以上の整数');
    if(entered>ordered) rowErrors.push('受注'+r.受注番号+': 現物'+entered+'が注文'+ordered+'を超過');
    const 台帳確保数=Number(確保[取り置き_行キー_(r)])||0;
    if(entered>0 && 台帳確保数>0 && entered+台帳確保数>ordered)
      rowErrors.push('受注'+r.受注番号+' '+r.商品コード+': 台帳確保済み'+台帳確保数+'個と合わせて注文'+ordered+'を超えます(④確保分の二重登録の疑い。数量を減らすか空にしてください)');
    if(entered>0 && lockedIds.has(String(r.取置ID||''))) rowErrors.push('受注'+r.受注番号+': 既に発送済み/解除済みの初期登録行は変更できません');
    const check=String(r.棚確認==null?'':r.棚確認).trim();
    if(entered>0 && (check==='出荷済み'||check==='未着'||check==='予約')) rowErrors.push('受注'+r.受注番号+': 棚確認が「'+check+'」なのに数量が入っています(どちらかを直してください)');
    if(rowErrors.length){ rowErrors.forEach(e=>errors.push(e)); return; }
    inputIds.add(String(r.取置ID||''));
    if(entered>0) targets[r.取置ID]=Object.assign({},r,{取り置き数量:entered});
  });
  const 解除数=(existingRows||[]).filter(r=>r.取置元種別==='開始前在庫' && r.状態===TORIOKI_STATUS.ACTIVE
    && inputIds.has(String(r.取置ID||'')) && !(String(r.取置ID||'') in targets)).length;
  const kept=(existingRows||[]).filter(r=>r.取置元種別!=='開始前在庫' || r.状態!==TORIOKI_STATUS.ACTIVE || !inputIds.has(String(r.取置ID||'')));
  Object.keys(targets).forEach(id=>{
    const r=targets[id];
    kept.push({取置ID:id,状態:TORIOKI_STATUS.ACTIVE,受注番号:r.受注番号,商品コード:r.商品コード,SKU:r.SKU,
      取り置き数量:r.取り置き数量,取置元種別:'開始前在庫',元EMS番号:'',元EMS商品コード:'',元取置ID:'',登録日時:now,更新日時:now,
      戻し処理結果:'','終了理由・メモ':String(r.メモ||'')});
  });
  return {rows:kept,errors,適用:{行:Object.keys(targets).length,
    数量:Object.keys(targets).reduce((n,id)=>n+targets[id].取り置き数量,0),解除:解除数}};
}

// 行単位適用(2026-07-22): 不正な選択・対象外IDはそのIDだけエラーにし、他の有効な選択は適用する。
function 取り置き_戻し確認計画_(inputs, ledger, now){
  const byId={}; (inputs||[]).forEach(r=>byId[String(r.取置ID||'')]=String(r.現物確認||'').trim());
  const errors=[]; let 棚へ戻した=0, 現物なし=0;
  const rows=(ledger||[]).map(r=>{
    const id=String(r.取置ID||''), choice=byId[id]; if(!choice) return Object.assign({},r);
    if(choice!=='現物あり' && choice!=='在庫なし'){ errors.push(id+': 現物確認は「現物あり」か「在庫なし」'); return Object.assign({},r); }
    if(r.状態!=='キャンセル戻し' || r.戻し処理結果!=='未確認'){ errors.push(id+': 未確認のキャンセル戻しではない'); return Object.assign({},r); }
    if(choice==='現物あり') 棚へ戻した++; else 現物なし++;
    return Object.assign({},r,{戻し処理結果:choice,更新日時:now});
  });
  return {rows,errors,適用:{棚へ戻した,現物なし}};
}

// 通常の現物取り置き数量と、赤い棚戻し待ち行の処理を一度に検証する。
// どちらか一方でもエラーなら rows=[] とし、台帳へ部分反映しない。
function 取り置き_統合反映計画_(inputs, ledger, now){
  const rows=inputs||[];
  // 棚戻し行は取置IDの形ではなく台帳の状態で見分ける。開始前在庫(INIT|系ID)が
  // キャンセルで棚戻し待ちになるケースがあり、ID接頭辞では通常行に誤分類されるため。
  const 台帳ById={};
  (ledger||[]).forEach(r=>{ 台帳ById[String(r.取置ID||'')]=r; });
  const 棚戻し行=r=>{
    const led=台帳ById[String(r.取置ID||'')];
    return !!led && led.状態===TORIOKI_STATUS.RETURN && led.戻し処理結果===TORIOKI_RETURN.UNCHECKED;
  };
  const normal=rows.filter(r=>!棚戻し行(r));
  const actions=rows.filter(棚戻し行);
  const counts={
    取り置き行:normal.filter(r=>(Number(r.現物取り置き数量)||0)>0).length,
    取り置き数量:normal.reduce((sum,r)=>sum+(Number(r.現物取り置き数量)||0),0),
    棚へ戻した:actions.filter(r=>String(r.処理||'').trim()==='棚へ戻した').length,
    現物なし:actions.filter(r=>String(r.処理||'').trim()==='現物なし').length
  };
  const initial=取り置き_初期確定計画_(normal,ledger,now);
  const returnInputs=actions.filter(r=>String(r.処理||'').trim()).map(r=>{
    const choice=String(r.処理||'').trim();
    return {取置ID:r.取置ID,現物確認:choice==='棚へ戻した'?'現物あり':choice==='現物なし'?'在庫なし':choice};
  });
  const returned=取り置き_戻し確認計画_(returnInputs,initial.rows,now);
  // 行単位適用(2026-07-22): エラーのキーだけスキップして適用可能な変更を1回で保存する。
  // countsは「実際に適用される」件数(入力意図ではなく)。
  return {rows:returned.rows,errors:initial.errors.concat(returned.errors),
    counts:{取り置き行:initial.適用.行,取り置き数量:initial.適用.数量,解除:initial.適用.解除,
      棚へ戻した:returned.適用.棚へ戻した,現物なし:returned.適用.現物なし}};
}

function EMS在庫移動_追加計画_(candidates, existing, now){
  const byId={}; (existing||[]).forEach(r=>{ byId[String(r.処理ID||'')]=r; });
  const added=[], errors=[];
  (candidates||[]).forEach(c=>{
    const id=String(c.処理ID||'');
    if(!c.処理ID || !取り置き_整数_(c.数量)) errors.push('Yahoo移動の処理IDまたは数量が不正: '+id);
    else if(byId[id]){
      // 同じ処理IDで数量が変わった=再締めで余りが増減した記録漏れの兆候。黙って捨てず締めを止める
      if(取り置き_整数_(byId[id].数量)!==取り置き_整数_(c.数量))
        errors.push('同じ処理IDで数量が一致しません(便の再締め/記録漏れの疑い): '+id+' 記録'+byId[id].数量+'／今回'+c.数量);
    }
    else {
      const row=Object.assign({},c,{移動先:'Yahoo即納',処理日時:now});
      byId[id]=row; added.push(row);
    }
  });
  return {rows:(existing||[]).map(r=>Object.assign({},r)).concat(added),added,errors};
}

function EMS在庫移動_箱計画_(surplus, existing, now){
  const candidates=(surplus||[]).filter(s=>s.qty>0).map(s=>{
    const ems=String(s.ems||'').trim(), sourceCode=String(s.sourceCode||s.code||'').trim();
    const sourceIdentity=取り置き_供給キー_('',sourceCode).slice(1);
    return {
      処理ID:ems&&sourceIdentity?'YAHOO|EMS|'+ems+'|'+sourceIdentity:'',
      EMS番号:ems,商品コード:sourceCode,数量:s.qty
    };
  });
  return EMS在庫移動_追加計画_(candidates,existing,now);
}

function EMS在庫移動_戻し計画_(returns, existing, now){
  const candidates=(returns||[])
    .filter(r=>r.状態==='キャンセル戻し'&&r.戻し処理結果==='現物あり')
    .map(r=>({
      処理ID:r.取置ID?'YAHOO|RETURN|'+r.取置ID:'',
      EMS番号:String(r.元EMS番号||''),
      商品コード:String(r.元EMS商品コード||r.商品コード||'').trim(),
      数量:r.取り置き数量
    }));
  return EMS在庫移動_追加計画_(candidates,existing,now);
}

function 取り置き_表を読む_(sheetName, headers){
  const sh=SpreadsheetApp.getActive().getSheetByName(sheetName);
  if(!sh || sh.getLastRow()<2) return [];
  const head=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(v=>String(v||'').trim());
  const index={}; headers.forEach(h=>index[h]=head.indexOf(h));
  const required=sheetName===TORIOKI_CFG.台帳 ? headers.slice(0,14) : headers;
  if(required.some(h=>index[h]<0)) throw new Error(sheetName+'の見出し不足: '+required.filter(h=>index[h]<0).join(','));
  return sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().map(row=>{
    const obj={}; headers.forEach(h=>obj[h]=index[h]<0?'':row[index[h]]); return obj;
  }).filter(obj=>String(obj[headers[0]]||'').trim());
}

function 取り置き_表を保存_(sheetName, headers, rows){
  const ss=SpreadsheetApp.getActive(); let sh=ss.getSheetByName(sheetName); if(!sh) sh=ss.insertSheet(sheetName);
  if(sh.getMaxColumns()<headers.length) sh.insertColumnsAfter(sh.getMaxColumns(),headers.length-sh.getMaxColumns());
  if(sh.getMaxRows()<rows.length+1) sh.insertRowsAfter(sh.getMaxRows(),rows.length+1-sh.getMaxRows());
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  const dataRows=Math.max(rows.length,Math.max(0,sh.getLastRow()-1)), dataCols=headers.length;
  if(dataRows>0){
    const values=Array.from({length:dataRows},(_,rowIndex)=>{
      const source=rows[rowIndex];
      return headers.map(header=>source&&source[header]!=null?source[header]:'');
    });
    sh.getRange(2,1,dataRows,dataCols).setValues(values);
  }
}

function 取り置き台帳_読む_(){ return 取り置き_表を読む_(TORIOKI_CFG.台帳,TORIOKI_CFG.台帳HDR); }
function 取り置き台帳_保存_(rows){ 取り置き_表を保存_(TORIOKI_CFG.台帳,TORIOKI_CFG.台帳HDR,rows); }
function EMS在庫移動台帳_読む_(){ return 取り置き_表を読む_(TORIOKI_CFG.移動,TORIOKI_CFG.移動HDR); }
function EMS在庫移動台帳_保存_(rows){ 取り置き_表を保存_(TORIOKI_CFG.移動,TORIOKI_CFG.移動HDR,rows); }

function 取り置き台帳_割当計画後行_(plan,existingRows,now){
  const rows=(existingRows||[]).map(r=>Object.assign({},r)), byId={};
  rows.forEach((r,index)=>byId[String(r.取置ID||'')]=index);
  (plan&&plan.returnUpdates||[]).forEach(update=>{
    const id=String(update.取置ID||''),index=byId[id]; if(index===undefined) return;
    rows[index]=Object.assign({},rows[index],update,{登録日時:rows[index].登録日時,更新日時:now});
  });
  (plan&&plan.newRows||[]).forEach(row=>{
    const id=String(row.取置ID||''),index=byId[id];
    if(index===undefined){
      const added=Object.assign({},row,{登録日時:now,更新日時:now}); byId[id]=rows.length; rows.push(added);
    }else{
      rows[index]=Object.assign({},rows[index],row,{登録日時:rows[index].登録日時,更新日時:now});
    }
  });
  return rows;
}

function 取り置き台帳_割当計画を反映_(plan,existingRows,now){
  const rows=取り置き台帳_割当計画後行_(plan,existingRows,now);
  取り置き台帳_保存_(rows);
  return rows;
}

function 取り置き_受注番号集合_(sheetName){
  // ④の一覧シートは1行目がタイムスタンプで見出しは書き出し_のstartRow(受注明細と同じ行)に入る。
  // 見出し行は固定位置にせず探す。シート未生成・見出しなし(=④未実行で空)は候補なしとして空集合。
  const sh=SpreadsheetApp.getActive().getSheetByName(sheetName), out=new Set();
  if(!sh || sh.getLastRow()<1) return out;
  const values=sh.getDataRange().getDisplayValues();
  const hr=values.findIndex(row=>row.map(v=>String(v||'').trim()).indexOf('受注番号')>=0);
  if(hr<0) return out;
  const col=values[hr].map(v=>String(v||'').trim()).indexOf('受注番号');
  for(let i=hr+1;i<values.length;i++){
    const ban=String(values[i][col]||'').trim(); if(ban) out.add(ban);
  }
  return out;
}

// 書き込み系ボタンは全て直列_(DocumentLock)で排他する(①〜④と同じ約束事)。
// 内部からの再帰呼び出しは本体_を使う(DocumentLockは再入不可のため)。
function 取り置き初期登録を作成(){ 直列_(取り置き初期登録を作成本体_); }
function 取り置き初期登録を作成本体_(options){
  const opts=options||{}, silent=opts.silent===true;
  const ss=SpreadsheetApp.getActive(), recv=ss.getSheetByName(HIKIATE_CFG.受注), ui=SpreadsheetApp.getUi();
  if(!recv){ if(!silent) ui.alert('受注明細がありません'); return {error:'受注明細がありません'}; }
  const M=列マップ_(recv), values=recv.getDataRange().getValues(), orders=[], 即納orders=[], 別ルートorders=[], 着済スタンプ=new Set(), 部分包装=new Set();
  const 受注head=values[M.hr-1].map(v=>String(v||'').trim()), cステータス=受注head.indexOf('受注ステータス');
  for(let i=M.hr;i<values.length;i++){
    const row=values[i], ban=String(row[M.番号]||'').trim(), qty=Number(row[M.個数])||0;
    if(!ban || qty<=0) continue;
    const kbn=区分_(row[M.選択肢]);
    const 別ルート=引当_別ルート判定_(row[M.選択肢], M.商品名>=0?row[M.商品名]:'');
    const ステータス=cステータス>=0?String(row[cステータス]||''):'';
    const 入荷日=M.入荷>=0?row[M.入荷]:'';
    const 氏名=M.氏名>=0?String(row[M.氏名]||''):'', code=String(row[M.コード]||''), sku=M.SKU>=0?String(row[M.SKU]||''):'';
    const 明細=取り置き_明細区分_(kbn,別ルート);
    if(明細==='別ルート'){ // 台湾/中国は表示専用。手入力入荷日=確保の証を持って別枠で集める
      別ルートorders.push({ban,氏名,code,sku,qty,ステータス,入荷日});
      continue;
    }
    if(明細==='即納'){ // 即納(指定なし含む)は表示専用(水色)で別枠で集める
      即納orders.push({ban,氏名,code,sku,qty,ステータス});
      continue;
    }
    if(明細!=='取り寄せ') continue; // 対象外(通常は発生しない)
    if(String(入荷日==null?'':入荷日).trim()!=='') 着済スタンプ.add(ban); // 旧帳簿では届いているはず=棚を確認すべき注文
    if(/部分包装/.test(ステータス)) 部分包装.add(ban); // 梱包中=現物が必ずある注文(スタンプや一覧に関係なく拾う)
    orders.push({ban,氏名,code,sku,qty,入荷日,EMS:M.EMS>=0?String(row[M.EMS]||''):'',ステータス,
      予約:取り置き_予約判定_(row[M.選択肢],M.商品名>=0?row[M.商品名]:'')});
  }
  // 出荷GO未入金(未入金滞留)・出荷可能(発送前)も棚に現物がある=候補に含める(空欄なら未取り置き扱い)。
  // どの一覧にも載っていなくても、入荷日スタンプがある注文と部分包装の注文は自動で候補に出す
  // (旧分類で引当待ちに埋もれた着済・梱包中の取りこぼし防止)。
  let candidates=取り置き_初期候補_(orders,[
    {状態:'出荷可能',bans:取り置き_受注番号集合_(HIKIATE_CFG.出荷)},
    {状態:'出荷GO未入金',bans:取り置き_受注番号集合_(HIKIATE_CFG.取置)},
    {状態:'部分在庫',bans:取り置き_受注番号集合_(HIKIATE_CFG.部分)},
    {状態:'希望日待ち',bans:取り置き_受注番号集合_(HIKIATE_CFG.希望)},
    {状態:'部分包装(要棚確認)',bans:部分包装},
    {状態:'着済スタンプ(要棚確認)',bans:着済スタンプ}
  ]);
  // 洗い替えでも手入力の数量・棚確認・メモが消えないよう、今のシート(旧名・旧レイアウト含む)から引き継ぐ
  const 読む=(name,headers)=>{ try{ return 取り置き_表を読む_(name,headers); }catch(e){ return []; } };
  let sheetRows=読む(TORIOKI_CFG.初期,TORIOKI_CFG.初期HDR);
  if(!sheetRows.length) sheetRows=読む(TORIOKI_CFG.初期,['取置ID','現物取り置き数量','棚確認','メモ']);
  if(!sheetRows.length) sheetRows=読む(TORIOKI_CFG.初期,['取置ID','現物取り置き数量','メモ']);
  if(!sheetRows.length) sheetRows=読む('取り置き初期登録',['取置ID','現物取り置き数量','メモ']); // 旧名からの一度きりの引き継ぎ
  const ledgerRows=取り置き台帳_読む_();
  candidates=取り置き_登録シート引き継ぎ_(candidates,sheetRows,ledgerRows);
  // 入力永続化(2026-07-22 仕様§8): 洗い替え前のシート入力を非表示シートへupsertし、生成行へ復元する。
  // 全数確保クリアで画面から消える数量も「未反映現物確認数量」として保存され、黙って失われない。
  const 入力保存=読む(TORIOKI_INPUT_CFG.シート,TORIOKI_INPUT_CFG.HDR);
  const 保存時刻=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd HH:mm:ss');
  const 保存マージ=取り置き_入力保存マージ_(candidates,入力保存,sheetRows,保存時刻);
  candidates=保存マージ.rows;
  取り置き_表を保存_(TORIOKI_INPUT_CFG.シート,TORIOKI_INPUT_CFG.HDR,Object.keys(保存マージ.store).sort().map(k=>保存マージ.store[k]));
  try{ ss.getSheetByName(TORIOKI_INPUT_CFG.シート).hideSheet(); }catch(e){}
  // ④が既に台帳(EMS引当等)で確保している分を重ねる(全数確保は入力欄クリア+黄色対象外)
  candidates=取り置き_台帳確保を適用_(candidates,取り置き_台帳確保集計_(ledgerRows));
  candidates=取り置き_登録絞り込み_(candidates); // 予約中・出荷GOのステータスだけで除外
  // 全行表示化(2026-07-20): 判断済みも隠さない。旧版が残した非表示スイッチの記憶は
  // 棚確認セルの既定値として一度だけ復元し、以後は空を書き戻して廃止(記憶はシートの棚確認列が担う)
  let 棚記憶={};
  try{ 棚記憶=JSON.parse(PropertiesService.getDocumentProperties().getProperty('取り置き登録_棚確認済み')||'{}'); }catch(e){}
  const 記憶適用=取り置き_棚確認記憶を適用_(candidates,棚記憶);
  candidates=記憶適用.rows;
  try{ PropertiesService.getDocumentProperties().setProperty('取り置き登録_棚確認済み',JSON.stringify(記憶適用.store)); }catch(e){}
  candidates.forEach(c=>{ c.判定=取り置き_棚確認判定_(c); });
  // 対象注文の即納行を表示専用(水色)で差し込む。数量・棚確認は入力対象外=誤入力は反映時ガードが止める
  candidates=取り置き_即納行を付与_(candidates,即納orders,sheetRows);
  // 混在注文の台湾/中国(別ルート)行を表示専用(橙)で差し込む。GoQ部分在庫=別ルート到着に気づけるように。
  candidates=取り置き_別ルート行を付与_(candidates,別ルートorders,sheetRows);
  // 数量0・注文キャンセルで確保を外す必要がある行も同じ画面へ集約する。
  const 戻し候補=取り置き_棚戻し候補_(ledgerRows,sheetRows);
  // 棚戻し待ち→要棚確認→通常の順。同じ受注番号の商品行は分断しない。
  candidates=取り置き_注文単位で並べる_(戻し候補.concat(candidates));
  取り置き_表を保存_(TORIOKI_CFG.初期,TORIOKI_CFG.初期HDR,candidates);
  const sh2=SpreadsheetApp.getActive().getSheetByName(TORIOKI_CFG.初期);
  // 列の並びが変わっても古い位置のプルダウンが残らないよう、シート全体の入力規則を消してから付け直す
  sh2.getRange(1,1,sh2.getMaxRows(),sh2.getMaxColumns()).clearDataValidations();
  if(candidates.length){
    // 通常行は棚確認、赤い棚戻し待ち行は処理のプルダウンを使う。
    const 棚確認列=TORIOKI_CFG.初期HDR.indexOf('棚確認')+1, 処理列=TORIOKI_CFG.初期HDR.indexOf('処理')+1;
    const 棚規則=SpreadsheetApp.newDataValidation().requireValueInList(TORIOKI_棚確認.slice(),true).setAllowInvalid(true).build();
    const 戻し規則=SpreadsheetApp.newDataValidation().requireValueInList(TORIOKI_戻し処理.slice(),true).setAllowInvalid(false).build();
    sh2.getRange(2,棚確認列,candidates.length,1).setDataValidation(棚規則);
    candidates.forEach((candidate,index)=>{
      if(candidate.判定==='棚戻し待ち'){
        sh2.getRange(index+2,棚確認列).clearDataValidations();
        sh2.getRange(index+2,処理列).setDataValidation(戻し規則);
      }else if(candidate.判定==='即納'||candidate.判定==='別ルート'){
        sh2.getRange(index+2,棚確認列).clearDataValidations(); // 表示専用行(即納・別ルート)に棚確認は選ばせない
      }
    });
  }
  sh2.setColumnWidth(TORIOKI_CFG.初期HDR.indexOf('要対応')+1,220);
  sh2.setColumnWidth(TORIOKI_CFG.初期HDR.indexOf('処理')+1,130);
  取り置き_登録行書式を更新_(sh2,candidates);
  const 入力済=candidates.filter(c=>String(c.現物取り置き数量)!=='').length;
  const 要確認数=candidates.filter(c=>c.判定==='要棚確認').length;
  const 戻し待ち数=candidates.filter(c=>c.判定==='棚戻し待ち').length;
  const 即納行数=candidates.filter(c=>c.判定==='即納').length;
  const 別ルート行数=candidates.filter(c=>c.判定==='別ルート').length;
  const 別ルート要対応=candidates.filter(c=>c.判定==='別ルート' && String(c.要対応||'').trim()!=='').length;
  const 全数確保=candidates.filter(c=>c.判定!=='棚戻し待ち' && (c.台帳確保数||0)>0 && c.判定!=='要棚確認').length;
  const summary={候補:candidates.length,要棚確認:要確認数,棚戻し待ち:戻し待ち数,即納表示:即納行数,別ルート表示:別ルート行数,入力済};
  if(!silent) ui.alert('取り置き登録を更新しました',
    '候補'+candidates.length+'行 ／ 🔴棚戻し待ち '+戻し待ち数+'行 ／ ⚠️要棚確認 '+要確認数+'行\n\n'
    +(戻し待ち数? '赤い行は棚を確認し、右端の「処理」で「棚へ戻した」または「現物なし」を選びます。\n':'')
    +'黄色の行は、現物があれば「現物取り置き数量」、無ければ「棚確認」を入力します。\n'
    +(即納行数? '水色の「即納」行は表示専用です(そのまま出荷する現物。数量は入れません)。\n':'')
    +(別ルート行数? '橙色の「別ルート(台湾/中国)」行は表示専用です。確保は受注明細に入荷日を入れてください'+(別ルート要対応? '（入荷日未入力 '+別ルート要対応+'行）':'')+'。\n':'')
    +'最後に「取り置き登録を反映」を1回押してください。\n'
    +(全数確保? '\n「台帳確保済み」'+全数確保+'行は④が確保済みのため数量入力不要です。':'')
    +(入力済? '\n入力済みの数量'+入力済+'行は引き継いでいます。':''),ui.ButtonSet.OK);
  return summary;
}

function 取り置き登録を反映(){ 直列_(取り置き登録を反映本体_); }
// 旧メニュー・既存の図形ボタンに割り当てた関数名も、新しい一括反映へつなぐ。
function 取り置き初期登録を確定(){ 取り置き登録を反映(); }
function 取り置き初期登録を確定本体_(){ return 取り置き登録を反映本体_(); }
function 取り置き登録を反映本体_(){
  const ui=SpreadsheetApp.getUi(), inputs=取り置き_表を読む_(TORIOKI_CFG.初期,TORIOKI_CFG.初期HDR);
  const plan=取り置き_統合反映計画_(inputs,取り置き台帳_読む_(),new Date());
  const c=plan.counts;
  // 行単位適用(2026-07-22): エラーは該当キーだけスキップ。適用できる変更が1つも無い時だけ中止
  const 適用あり=c.取り置き行>0||(c.解除||0)>0||c.棚へ戻した>0||c.現物なし>0;
  if(!適用あり){ ui.alert('取り置き登録の反映を中止しました',plan.errors.length?plan.errors.join('\n'):'適用できる入力がありません',ui.ButtonSet.OK); return; }
  const answer=ui.alert('取り置き登録を反映します',
    '通常の取り置き '+c.取り置き行+'行 / '+c.取り置き数量+'個\n'
    +'解除(空欄クリア) '+(c.解除||0)+'件\n'
    +'棚へ戻した '+c.棚へ戻した+'件\n'
    +'現物なし '+c.現物なし+'件\n'
    +(plan.errors.length?'\n⚠️ エラー'+plan.errors.length+'件はスキップして続行します(入力は保持されます):\n'+plan.errors.slice(0,6).join('\n')+(plan.errors.length>6?'\n…ほか'+(plan.errors.length-6)+'件':'')+'\n':'')
    +'\n台帳へ一括反映しますか？',ui.ButtonSet.OK_CANCEL);
  if(answer!==ui.Button.OK) return;
  取り置き台帳_保存_(plan.rows);
  // 旧管理シートは裏側の監査用として同期するが、日常操作は取り置き登録だけで完結する。
  try{ Yahoo戻し候補を更新_(); }catch(e){}
  const refreshed=取り置き初期登録を作成本体_({silent:true});
  SpreadsheetApp.getActive().toast(
    '通常 '+c.取り置き数量+'個 / 棚戻し '+c.棚へ戻した+'件 / 現物なし '+c.現物なし+'件を反映しました'
    +(refreshed&&refreshed.棚戻し待ち? '（未処理あと'+refreshed.棚戻し待ち+'件）':''),
    '取り置き登録',8);
}

function キャンセル戻し確認を更新(){ 直列_(キャンセル戻し確認を更新本体_); }
function キャンセル戻し確認を更新本体_(){
  const rows=取り置き台帳_読む_().filter(r=>r.状態==='キャンセル戻し'&&r.戻し処理結果==='未確認').map(r=>({
    取置ID:r.取置ID,受注番号:r.受注番号,商品コード:r.商品コード,数量:r.取り置き数量,元EMS番号:r.元EMS番号,現物確認:'',メモ:r['終了理由・メモ']||''
  }));
  取り置き_表を保存_(TORIOKI_CFG.戻し,戻しHDR,rows);
  const sh=SpreadsheetApp.getActive().getSheetByName(TORIOKI_CFG.戻し);
  if(rows.length) sh.getRange(2,6,rows.length,1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['現物あり','在庫なし'],true).setAllowInvalid(false).build());
  SpreadsheetApp.getActive().toast('未確認のキャンセル戻し '+rows.length+'件','取り置き台帳',6);
}

function キャンセル戻し確認を確定(){ 直列_(キャンセル戻し確認を確定本体_); }
function キャンセル戻し確認を確定本体_(){
  const ui=SpreadsheetApp.getUi(), inputs=取り置き_表を読む_(TORIOKI_CFG.戻し,戻しHDR);
  const plan=取り置き_戻し確認計画_(inputs,取り置き台帳_読む_(),new Date());
  if(plan.errors.length){ ui.alert('キャンセル戻しを確定できません',plan.errors.join('\n'),ui.ButtonSet.OK); return; }
  const answer=ui.alert('現物確認を確定します','入力済み'+inputs.filter(r=>r.現物確認).length+'件を台帳へ反映します。',ui.ButtonSet.OK_CANCEL);
  if(answer!==ui.Button.OK) return;
  取り置き台帳_保存_(plan.rows); キャンセル戻し確認を更新本体_(); Yahoo戻し候補を更新_();
}

function Yahoo戻し候補を更新_(){
  const rows=取り置き台帳_読む_().filter(r=>r.状態==='キャンセル戻し'&&r.戻し処理結果==='現物あり').map(r=>({
    取置ID:r.取置ID,商品コード:String(r.元EMS商品コード||r.商品コード||'').trim(),数量:r.取り置き数量,元EMS番号:r.元EMS番号,処理ID:'YAHOO|RETURN|'+r.取置ID,確認:''
  }));
  取り置き_表を保存_(TORIOKI_CFG.Yahoo候補,Yahoo候補HDR,rows);
  const sh=SpreadsheetApp.getActive().getSheetByName(TORIOKI_CFG.Yahoo候補);
  if(rows.length) sh.getRange(2,6,rows.length,1).insertCheckboxes();
}

function キャンセル戻しをYahoo反映済みにする(){ 直列_(キャンセル戻しをYahoo反映済みにする本体_); }
function キャンセル戻しをYahoo反映済みにする本体_(){
  const ui=SpreadsheetApp.getUi(), candidates=取り置き_表を読む_(TORIOKI_CFG.Yahoo候補,Yahoo候補HDR).filter(r=>r.確認===true||String(r.確認).toUpperCase()==='TRUE');
  if(!candidates.length){ ui.alert('Yahoo戻し候補にチェックがありません'); return; }
  const answer=ui.alert('Yahoo戻しの最終確認','Yahoo在庫へ実際に加算済みの'+candidates.length+'件だけを確定します。',ui.ButtonSet.OK_CANCEL);
  if(answer!==ui.Button.OK) return;
  const checkedIdCounts={};
  candidates.forEach(r=>{ const id=String(r.取置ID||''); checkedIdCounts[id]=(checkedIdCounts[id]||0)+1; });
  const duplicateChecked=Object.keys(checkedIdCounts).filter(id=>!id || checkedIdCounts[id]!==1);
  if(duplicateChecked.length){ ui.alert('Yahoo戻し候補が重複またはIDなしのため中止しました',duplicateChecked.join('\n'),ui.ButtonSet.OK); return; }
  const ledger=取り置き台帳_読む_(), selectedIds=new Set(candidates.map(r=>String(r.取置ID)));
  const ledgerIdCounts={};
  ledger.forEach(r=>{ const id=String(r.取置ID||''); ledgerIdCounts[id]=(ledgerIdCounts[id]||0)+1; });
  const unresolved=Array.from(selectedIds).filter(id=>!id || ledgerIdCounts[id]!==1);
  if(unresolved.length){ ui.alert('Yahoo戻し候補が現在の取り置き台帳と一致しないため中止しました',unresolved.join('\n'),ui.ButtonSet.OK); return; }
  const returns=ledger.filter(r=>selectedIds.has(String(r.取置ID)));
  const candidateById={}; candidates.forEach(r=>candidateById[String(r.取置ID)]=r);
  const tampered=returns.filter(r=>{
    const id=String(r.取置ID), candidate=candidateById[id];
    const sourceCode=String(r.元EMS商品コード||r.商品コード||'').trim();
    return String(candidate.商品コード||'').trim()!==sourceCode ||
      取り置き_整数_(candidate.数量)!==取り置き_整数_(r.取り置き数量) ||
      String(candidate.元EMS番号||'').trim()!==String(r.元EMS番号||'').trim() ||
      String(candidate.処理ID||'').trim()!=='YAHOO|RETURN|'+id;
  });
  if(tampered.length){ ui.alert('Yahoo戻し候補が現在の取り置き台帳と一致しないため中止しました',tampered.map(r=>r.取置ID).join('\n'),ui.ButtonSet.OK); return; }
  if(returns.some(r=>r.状態!=='キャンセル戻し'||r.戻し処理結果!=='現物あり')){ ui.alert('現物ありでない候補が含まれるため中止しました'); return; }
  const existingMoves=EMS在庫移動台帳_読む_(), plan=EMS在庫移動_戻し計画_(returns,existingMoves,new Date());
  if(plan.errors.length){ ui.alert('Yahoo移動を中止しました',plan.errors.join('\n'),ui.ButtonSet.OK); return; }
  const updatedLedger=ledger.map(r=>selectedIds.has(String(r.取置ID))?Object.assign({},r,{戻し処理結果:'Yahoo反映済み',更新日時:new Date()}):r);
  try{ EMS在庫移動台帳_保存_(plan.rows); }
  catch(error){ ui.alert('Yahoo移動台帳の保存に失敗しました',error.message,ui.ButtonSet.OK); return; }
  try{ 取り置き台帳_保存_(updatedLedger); }
  catch(ledgerError){
    try{ EMS在庫移動台帳_保存_(existingMoves); }
    catch(rollbackError){
      ui.alert('取り置き台帳の保存とYahoo移動台帳の復旧に失敗しました',
        '取り置き台帳: '+ledgerError.message+'\nYahoo移動台帳: '+rollbackError.message,ui.ButtonSet.OK);
      return;
    }
    ui.alert('取り置き台帳の保存に失敗したためYahoo移動台帳を元へ戻しました',ledgerError.message,ui.ButtonSet.OK);
    return;
  }
  Yahoo戻し候補を更新_();
}

// 孤児取り置き一括解除(2026-07-20): 品切れ→代替発送などで受注明細から注文が消えたのに残る取り置きをまとめて掃除する。
// 孤児 = 状態が取り置き中なのに行キー(受注番号|商品コード|SKU)が今の受注明細に無い取り置き(既存のタグ注意と同じ基準)。
// 出荷形跡 = その受注番号が消込台帳の出荷済み行に載っている(代替でも受注番号一致で形跡ありと判定=既定チェックの根拠)。
function 取り置き_孤児取り置き一覧_(ledgerRows, 注文キー集合, 出荷済み受注番号集合){
  const orderKeys=注文キー集合||new Set(), shippedBans=出荷済み受注番号集合||new Set();
  return (ledgerRows||[])
    .filter(r=>String(r.状態||'')===TORIOKI_STATUS.ACTIVE && !orderKeys.has(取り置き_行キー_(r)))
    .map(r=>({取置ID:String(r.取置ID||''),受注番号:String(r.受注番号||''),商品コード:String(r.商品コード||''),
      SKU:String(r.SKU||''),取り置き数量:取り置き_整数_(r.取り置き数量),出荷形跡:shippedBans.has(String(r.受注番号||''))}));
}

// 指定した取置IDの「取り置き中」行だけを手動解除にする(冪等・非指定/取り置き中以外は不変)。
function 取り置き_一括解除計画_(ledgerRows, 解除ID集合, 理由, now){
  const ids=解除ID集合||new Set(), reason=String(理由||'').trim();
  return (ledgerRows||[]).map(r=>{
    if(!ids.has(String(r.取置ID||'')) || String(r.状態||'')!==TORIOKI_STATUS.ACTIVE) return r;
    return Object.assign({},r,{状態:TORIOKI_STATUS.RELEASED,'終了理由・メモ':reason,更新日時:now});
  });
}

function 選択した取り置きを手動解除(){ 直列_(選択した取り置きを手動解除本体_); }
function 選択した取り置きを手動解除本体_(){
  const ss=SpreadsheetApp.getActive(), sh=ss.getActiveSheet(), ui=SpreadsheetApp.getUi(), row=sh.getActiveRange().getRow();
  if(sh.getName()!==TORIOKI_CFG.台帳 || row<2){ ui.alert('取り置き台帳の解除する行を選択してください'); return; }
  // 行番号ではなく選択行の取置IDで台帳の対象を特定する(空行・手動の行操作でズレても別の行を解除しない)
  const selectedId=String(sh.getRange(row,1).getDisplayValue()||'').trim();
  if(!selectedId){ ui.alert('選択した行に取置IDがありません'); return; }
  const ledger=取り置き台帳_読む_(), matches=ledger.filter(r=>String(r.取置ID||'')===selectedId);
  if(matches.length!==1){ ui.alert('取置IDが台帳で一意に特定できません: '+selectedId); return; }
  const target=matches[0];
  if(target.状態!=='取り置き中'){ ui.alert('取り置き中の行だけ手動解除できます'); return; }
  const response=ui.prompt('手動解除の理由','登録間違い、現物不足などの理由を入力してください。',ui.ButtonSet.OK_CANCEL);
  if(response.getSelectedButton()!==ui.Button.OK) return;
  const reason=String(response.getResponseText()||'').trim(); if(!reason){ ui.alert('解除理由は必須です'); return; }
  取り置き台帳_保存_(ledger.map(r=>String(r.取置ID||'')===selectedId
    ? Object.assign({},r,{状態:'手動解除','終了理由・メモ':reason,更新日時:new Date()}) : r));
}

// 🧹 孤児取り置きをまとめて解除: 孤児を一覧化してチェックで選び、まとめて手動解除する。
// 読み取り＋ダイアログ表示のみ(書き込みは orphanBulkRelease 側で直列_)。ダニエル引当のダイアログ方式を踏襲。
function 孤児取り置きをまとめて解除(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  const recv=ss.getSheetByName(HIKIATE_CFG.受注);
  if(!recv){ ui.alert('受注明細がありません'); return; }
  const M=列マップ_(recv), values=recv.getDataRange().getValues(), 注文キー集合=new Set();
  for(let i=M.hr;i<values.length;i++){
    const row=values[i], ban=String(row[M.番号]||'').trim(); if(!ban) continue;
    注文キー集合.add(取り置き_行キー_({受注番号:ban,商品コード:M.コード>=0?row[M.コード]:'',SKU:M.SKU>=0?row[M.SKU]:''}));
  }
  let shippedBans=new Set();
  try{ shippedBans=new Set(消込台帳_出荷済み行_().map(r=>String(r.ban||''))); }catch(e){}
  const orphans=取り置き_孤児取り置き一覧_(取り置き台帳_読む_(),注文キー集合,shippedBans);
  if(!orphans.length){ ui.alert('孤児の取り置きはありません','受注明細に注文が無い「取り置き中」は見つかりませんでした。',ui.ButtonSet.OK); return; }
  let rowsHtml='';
  orphans.forEach(o=>{
    rowsHtml+='<label style="display:block;margin:5px 0;padding:5px 6px;border:1px solid #eee;border-radius:4px;cursor:pointer;">'
      +'<input type="checkbox" class="orphan" value="'+o.取置ID+'"'+(o.出荷形跡?' checked':'')+'> '
      +'<b>'+o.受注番号+'</b> '+o.商品コード+' <span style="color:#777">('+o.取り置き数量+'個)</span> '
      +(o.出荷形跡?'<span style="color:#093;font-size:12px;">✔出荷形跡あり</span>':'<span style="color:#999;font-size:12px;">形跡なし</span>')+'</label>';
  });
  const html='<div style="font-family:sans-serif;font-size:13px;line-height:1.4;">'
    +'<p style="margin:0 0 8px;">解除する孤児取り置きにチェック（<b>出荷形跡あり</b>は既定でオン）。品切れ→代替発送などで注文が消えた分を掃除します。</p>'
    +rowsHtml
    +'<div style="margin-top:8px;">解除理由: <input type="text" id="reason" value="品切れ・代替品で発送" style="width:60%;padding:3px;"></div>'
    +'<div style="margin-top:6px;"><button type="button" onclick="all(true)">全部</button> <button type="button" onclick="all(false)">全部外す</button></div>'
    +'<div style="margin-top:14px;text-align:right;"><button type="button" id="go" style="padding:7px 16px;font-weight:bold;background:#4472c4;color:#fff;border:none;border-radius:4px;cursor:pointer;" onclick="go(this)">選んだ分を解除</button></div>'
    +'<script>'
    +'function all(v){var cs=document.querySelectorAll(".orphan");for(var i=0;i<cs.length;i++)cs[i].checked=v;}'
    +'function go(b){var a=[],cs=document.querySelectorAll(".orphan:checked");for(var i=0;i<cs.length;i++)a.push(cs[i].value);'
    +'var reason=document.getElementById("reason").value;'
    +'if(!a.length){alert("解除する行を1つ以上選んでください");return;}'
    +'if(!reason.trim()){alert("解除理由を入力してください");return;}'
    +'b.disabled=true;b.textContent="解除中...";'
    +'google.script.run.withSuccessHandler(function(n){google.script.host.close();})'
    +'.withFailureHandler(function(e){b.disabled=false;b.textContent="選んだ分を解除";alert("エラー: "+e.message);}).orphanBulkRelease(a,reason);}'
    +'</script></div>';
  const out=HtmlService.createHtmlOutput(html).setWidth(480).setHeight(Math.min(150+orphans.length*40,560));
  ui.showModalDialog(out,'🧹 孤児取り置きをまとめて解除');
}

// ダイアログから呼ぶ実体(google.script.run対象なので末尾_を付けない=ASCII名で安全に)。選ばれた取置IDの取り置き中だけ手動解除。
function orphanBulkRelease(ids, reason){
  return 直列_(()=>{
    const 理由=String(reason||'').trim(); if(!理由) throw new Error('解除理由が空です');
    const idSet=new Set((ids||[]).map(x=>String(x))); if(!idSet.size) throw new Error('解除対象がありません');
    const ledger=取り置き台帳_読む_();
    const 対象数=ledger.filter(r=>idSet.has(String(r.取置ID||'')) && String(r.状態||'')===TORIOKI_STATUS.ACTIVE).length;
    取り置き台帳_保存_(取り置き_一括解除計画_(ledger,idSet,理由,new Date()));
    SpreadsheetApp.getActive().toast(対象数+'件の孤児取り置きを手動解除しました','🧹 孤児解除',6);
    return 対象数;
  });
}

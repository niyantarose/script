const TORIOKI_CFG = Object.freeze({
  台帳:'取り置き台帳', 初期:'取り置き登録', 要確認:'取り置き要確認', 戻し:'キャンセル戻し確認', Yahoo候補:'Yahoo戻し候補', 移動:'EMS在庫移動台帳',
  台帳HDR:['取置ID','状態','受注番号','商品コード','SKU','取り置き数量','取置元種別','元EMS番号','元EMS商品コード','元取置ID','登録日時','更新日時','戻し処理結果','終了理由・メモ','引当段階','EMS到着予定日','現物確認日時','現物確認メモ','供給控除EMS','引当系譜ID','引当系譜数量','供給処理'],
  // 作業者目線の17列(2026-07-23): 確保済み/不足の数字列+確保内訳(自動=④②/棚=自分の登録/入荷日)+差分入力(追加数量/マイナス数量)。
  // 現物取り置き数量(絶対値入力)と空欄=解除は廃止。空欄=何もしない。
  // 現在の状態・旧入荷日・旧EMSは内部データとしてだけ持ち表示しない。取置IDは列を残して非表示
  初期HDR:['取置ID','受注番号','氏名','商品コード','SKU','注文数量','確保済み','不足','確保内訳','受注ステータス','追加数量','マイナス数量','棚確認','メモ','判定','要対応','処理'],
  // 取り置き登録のレイアウト(2026-07-23 ユーザー要望「ボタンを置きたい」): 1〜2行目=ボタン置き場、
  // 3行目=見出し、4行目〜=データ。ボタンは図形で「取り置き初期登録を作成」「取り置き登録を反映」を割り当てる
  初期HDR行:3,
  要確認HDR:['取置ID','受注番号','商品コード','理由'],
  移動HDR:['処理ID','EMS番号','商品コード','数量','移動先','処理日時']
});
const 戻しHDR=['取置ID','受注番号','商品コード','数量','元EMS番号','現物確認','メモ'];
// 手入力の永続保存(非表示シート)。GoQ取込・更新・引当・全件再計算で入力を消さないための内部保存。仕様§8
const TORIOKI_INPUT_CFG=Object.freeze({シート:'取り置き入力保存',履歴:'取り置き入力履歴',
  HDR:['入力キー','受注番号','SKU','商品コード','棚確認','取り置きメモ','確認メモ','注文作業メモ','未反映追加数量','未反映マイナス数量','入力エラー','最終表示日時','更新日時'],
  旧HDR:['入力キー','受注番号','SKU','商品コード','棚確認','取り置きメモ','確認メモ','注文作業メモ','未反映現物確認数量','入力エラー','最終表示日時','更新日時']});

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
    const add=r.追加数量; if(add!=null&&String(add).trim()!=='') cur.未反映追加数量=add;
    const sub=r.マイナス数量; if(sub!=null&&String(sub).trim()!=='') cur.未反映マイナス数量=sub;
    cur.更新日時=now||'';
  });
  const rows=(generatedRows||[]).map(g=>{
    const k=取り置き_入力キー_(g), s=store[k];
    if(!s) return g;
    s.最終表示日時=now||'';
    return Object.assign({},g,{
      棚確認:String(g.棚確認||'').trim()||String(s.棚確認||''),
      メモ:String(g.メモ||'').trim()||String(s.取り置きメモ||''),
      追加数量:(g.追加数量!=null&&String(g.追加数量).trim()!=='')?g.追加数量
        :(s.未反映追加数量!=null&&String(s.未反映追加数量).trim()!==''?s.未反映追加数量:''),
      マイナス数量:(g.マイナス数量!=null&&String(g.マイナス数量).trim()!=='')?g.マイナス数量
        :(s.未反映マイナス数量!=null&&String(s.未反映マイナス数量).trim()!==''?s.未反映マイナス数量:'')
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
  const 現物あり=has(r=>(Number(r.確保済み)||0)>0||(Number(r.台帳確保数)||0)>0||TORIOKI_部分在庫系.indexOf(String(r.棚確認||''))>=0);
  const 部分=has(r=>String(r.現在の状態||'')==='部分在庫');
  const 希望=has(r=>String(r.現在の状態||'')==='希望日待ち');
  if(mode==='部分在庫') return 部分;
  if(mode==='希望日待ち・現物あり') return 希望&&現物あり;
  if(mode==='先行引当') return has(r=>/先行/.test(String(r.台帳確保||'')+String(r.状態の理由||'')));
  const 要作業=has(r=>String(r.判定||'')==='要棚確認'||String(r.判定||'')==='棚戻し待ち'||String(r.要対応||'').trim()!=='');
  // 現物ありの希望日待ちは「棚確認が未記入の行」がある間だけ要作業に出す。
  // 棚確認を付けた(発送待ち・未着など判断済み)注文は片付いた扱いで消える(2026-07-23)
  const 未判断あり=has(r=>((Number(r.確保済み)||0)>0||(Number(r.台帳確保数)||0)>0)
    &&String(r.棚確認||'').trim()==='');
  return 部分||要作業||(希望&&現物あり&&未判断あり);
}
// 取り置き登録の「棚確認」プルダウン。出荷済み/未着/予約は数量なし(登録しない)の目印。
// 部分在庫は確保の時期で3種類(2026-07-23 ユーザー設計): 過去部分在庫=今日より前からの確保(緑)、
// 当日部分在庫=今日届いた箱からの確保(青緑・箱を開ける)、先行部分在庫=先行引当のみ(薄紫・現物まだ無い)
const TORIOKI_棚確認=Object.freeze(['発送待ち','過去部分在庫','当日部分在庫','先行部分在庫','出荷済み','未着','予約']);
// 旧名「部分在庫」も系として認識し、更新時に確保表示整合_が現行名へ貼り替える
const TORIOKI_部分在庫系=Object.freeze(['部分在庫','過去部分在庫','当日部分在庫','先行部分在庫']);
const TORIOKI_戻し処理=Object.freeze(['棚へ戻した','現物なし']);

function 取り置き_棚確認書式定義_(棚確認列, 開始行){
  let n=棚確認列, 列記号='';
  while(n>0){ n--; 列記号=String.fromCharCode(65+n%26)+列記号; n=Math.floor(n/26); }
  return [
    {値:'発送待ち',背景:'#cfe2f3'},
    {値:'過去部分在庫',背景:'#d9ead3'},    // 今日より前からの確保(緑)
    {値:'当日部分在庫',背景:'#f6b26b'},    // 今日届いた箱からの確保(オレンジ・箱を開ける)
    {値:'先行部分在庫',背景:'#ead1dc'},    // 先行引当のみ(薄紫・現物まだ無い)
    {値:'出荷済み',背景:'#f4cccc',文字色:'#990000',太字:true},
    {値:'未着',背景:'#d9d9d9'},
    {値:'予約',背景:'#d9d2e9'}
  ].map(def=>Object.assign({条件:'=$'+列記号+開始行+'="'+def.値+'"'},def));
}

// 確保内訳(I列)のセル色(2026-07-23 ユーザー要望): 誰が確保したか色でも見分ける。
// 先行を含む=薄紫(現物まだ無い)、自動を含む=薄オレンジ(俺が数えたものじゃない)、現物だけ=緑(俺が数えた)。
// 行全体の棚確認色より先に並べてI列のセルだけ内訳色が勝つようにする。
function 取り置き_内訳書式定義_(内訳列, 開始行){
  let n=内訳列, 列記号='';
  while(n>0){ n--; 列記号=String.fromCharCode(65+n%26)+列記号; n=Math.floor(n/26); }
  const cell='$'+列記号+開始行;
  return [
    {条件:'=ISNUMBER(SEARCH("先行",'+cell+'))',背景:'#ead1dc'},
    {条件:'=ISNUMBER(SEARCH("自動",'+cell+'))',背景:'#fce5cd'},
    {条件:'=ISNUMBER(SEARCH("現物",'+cell+'))',背景:'#d9ead3'}
  ];
}

function 取り置き_棚確認書式を設定_(sh, rowCount){
  if(rowCount<=0){ sh.setConditionalFormatRules([]); return; }
  const 棚確認列=TORIOKI_CFG.初期HDR.indexOf('棚確認')+1, 幅=TORIOKI_CFG.初期HDR.length;
  const 内訳列=TORIOKI_CFG.初期HDR.indexOf('確保内訳')+1;
  const データ行=TORIOKI_CFG.初期HDR行+1;
  const target=sh.getRange(データ行,1,rowCount,幅);
  const 内訳target=sh.getRange(データ行,内訳列,rowCount,1);
  const 内訳rules=取り置き_内訳書式定義_(内訳列,データ行).map(def=>
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(def.条件)
      .setBackground(def.背景)
      .setRanges([内訳target])
      .build());
  const rules=取り置き_棚確認書式定義_(棚確認列,データ行).map(def=>{
    let b=SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(def.条件)
      .setBackground(def.背景)
      .setRanges([target]);
    if(def.文字色) b=b.setFontColor(def.文字色);
    if(def.太字) b=b.setBold(true);
    return b.build();
  });
  sh.setConditionalFormatRules(内訳rules.concat(rules));
}

function 取り置き_登録行書式を更新_(sh, candidates){
  const rows=candidates||[], 幅=TORIOKI_CFG.初期HDR.length;
  const データ行=TORIOKI_CFG.初期HDR行+1; // 1〜2行目=ボタン置き場、3行目=見出し
  const 管理行数=Math.max(0,sh.getMaxRows()-データ行+1);
  if(管理行数>0){
    const 管理範囲=sh.getRange(データ行,1,管理行数,幅);
    管理範囲.setBackground(null);
    管理範囲.setBorder(false,false,false,false,false,false);
  }
  注文罫線_(sh,データ行,1);
  if(rows.length){
    sh.getRange(データ行,1,rows.length,幅).setBackgrounds(
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
  const qtyText=String(r.確保済み==null?'':r.確保済み).trim(); // 差分方式: 現物証拠は確保済み(台帳)で見る
  return String(r.旧入荷日==null?'':r.旧入荷日).trim()!=='' ||
    String(r.旧EMS==null?'':r.旧EMS).trim()!=='' ||
    Number(r.台帳確保数)>0 ||
    (qtyText!=='' && Number(qtyText)>0);
}
const Yahoo候補HDR=['取置ID','商品コード','数量','元EMS番号','処理ID','確認'];
// 取り置き登録と同じく1〜2行目をボタン置き場にする(2026-07-23)。図形へ
// 「Yahoo戻し候補を更新」「キャンセル戻しをYahoo反映済みにする」を割り当てる
const Yahoo候補HDR行=3;
// 「日本在庫へ移す」で確定した戻り現物の控え(非表示)。日本在庫シートは②のたびに作り直されるため、
// CSVを作るまでリストに残す持ち場が要る。CSV作成で出力日時が入り、リストから外れる(2026-07-24)
const 戻り待ちCFG=Object.freeze({シート:'日本在庫戻り待ち',
  HDR:['処理ID','商品コード','数量','確定日時','出力日時']});
function 日本在庫_戻り待ち_読む_(){ try{ return 取り置き_表を読む_(戻り待ちCFG.シート,戻り待ちCFG.HDR); }catch(e){ return []; } }
function 日本在庫_戻り待ち_保存_(rows){
  取り置き_表を保存_(戻り待ちCFG.シート,戻り待ちCFG.HDR,rows);
  try{ SpreadsheetApp.getActive().getSheetByName(戻り待ちCFG.シート).hideSheet(); }catch(e){}
}
// 日本在庫へ出す「戻り」行(未出力ぶんだけ)
function 日本在庫_戻り待ち行_(){ return 日本在庫_戻り行_(日本在庫_戻り待ち_読む_()); }

function CSV行を受注行オブジェクトへ_(header, rows){
  const head=(header||[]).map(v=>String(v||'').trim()), index=name=>head.indexOf(name);
  const cBan=index('受注番号'), cStatus=index('受注ステータス'), cCode=index('商品コード'), cQty=index('個数');
  const cSku=index('商品SKU')>=0?index('商品SKU'):index('SKU');
  const cShipU=index('出荷日');
  const cShipW=index('出荷日(複数時には配送先毎)')>=0?index('出荷日(複数時には配送先毎)'):index('出荷日（複数時には配送先毎）');
  const missing=[];
  if(cBan<0) missing.push('受注番号'); if(cStatus<0) missing.push('受注ステータス');
  if(cCode<0) missing.push('商品コード'); if(cQty<0) missing.push('個数'); if(cSku<0) missing.push('商品SKU/SKU');
  if(missing.length) throw new Error('全ステータスCSVの見出し不足: '+missing.join(','));
  return (rows||[]).map(row=>({受注番号:String(row[cBan]||'').replace(/^niyantarose-/i,''),受注ステータス:String(row[cStatus]||''),
    商品コード:String(row[cCode]||''),SKU:String(row[cSku]||''),個数:Number(row[cQty])||0,
    // 分割出荷済み行(出荷日あり)。ステータスが注文単位のままでも行単位で「もう送った」が分かる
    行出荷済み:(cShipU>=0&&String(row[cShipU]==null?'':row[cShipU]).trim()!=='')||(cShipW>=0&&String(row[cShipW]==null?'':row[cShipW]).trim()!=='')}));
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
    // 分割出荷済み(行の出荷日あり)は処理済み扱いで需要から外す。ただし注文全体の発送済み遷移には使わない
    const 行出荷=!!r.行出荷済み && !shipped && !cancelled;
    const rawQty=r.個数, qtyKnown=rawQty!==undefined && rawQty!==null && String(rawQty).trim()!=='' && isFinite(Number(rawQty));
    const qty=qtyKnown?Number(rawQty):0;
    group.発送済み=group.発送済み||shipped;
    group.キャンセル=group.キャンセル||cancelled;
    group.数量既知=group.数量既知||qtyKnown;
    // キャンセル行の正の数量を注文数へ混ぜない(有効・処理済み・キャンセルを別々に集計)
    if(qtyKnown){
      if(cancelled) group.キャンセル数量+=qty;
      else if(shipped || 行出荷) group.処理済数量+=qty;
      else group.有効数量+=qty;
    }
    // 同じ受注商品が分割され、数量0のキャンセル行と数量1以上の生きた行が共存する場合、
    // 商品全体の取り置きは解除しない。全分割行が0になったときだけ棚戻しへ進める。
    if(qtyKnown && qty>0 && !shipped && !cancelled && !行出荷) group.生存数量=true;
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
        棚確認:自動予約?'予約':'',自動予約,追加数量:'',マイナス数量:'',メモ:'',判定:''};
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
      台帳確保:'',台帳確保数:0,棚確認:'',追加数量:'',マイナス数量:'',メモ:prev?String(prev.メモ==null?'':prev.メモ).trim():'',判定:'即納',要対応:'',処理:''};
  }));
}

// 全行表示化v2(2026-07-20): 台湾/中国(別ルート)行を橙で差し込む。GoQで部分在庫=別ルートで
// 現物が到着した合図なのに候補から消えて気づけない、を防ぐ。2026-07-23から差分入力(+−)可能。
// 確保済み/不足は②の別ルート二重控除_と同じ実効確保=max(台帳の棚登録, 受注明細入荷日ベース)で出す
// (両方式の併用時に足し算すると二重計上に見えるため大きい方)。
function 取り置き_別ルート行を付与_(candidates, betsuOrders, sheetRows, securedByKey, 現物登録ByKey){
  const list=candidates||[];
  const bans=new Set(list.map(c=>String(c.受注番号)));
  const bySheet={};
  (sheetRows||[]).forEach(r=>{ const id=String(r&&r.取置ID||''); if(id) bySheet[id]=r; });
  const byKey={}, keys=[];
  (betsuOrders||[]).forEach(o=>{
    if(!bans.has(String(o.ban))) return;
    const key=取り置き_ID部_(o); // 取置ID(別ルート|…)の互換のため従来3部形式で束ねる
    if(!byKey[key]){ byKey[key]={o,qty:0,入荷日:'',入荷日確保:0,ステータス一覧:[]}; keys.push(key); }
    byKey[key].qty+=Number(o.qty)||0;
    if(!byKey[key].入荷日){ const d=ymd_(o.入荷日); if(d) byKey[key].入荷日=d; }
    // ②と同じ判定(引当実行_本体_: trimが空でなければ入荷=確保)。行ごとに数える(束の一部だけ入荷日ありに対応)
    if(String(o.入荷日==null?'':o.入荷日).trim()!=='') byKey[key].入荷日確保+=Number(o.qty)||0;
    const status=String(o.ステータス||'').trim();
    if(status && byKey[key].ステータス一覧.indexOf(status)<0) byKey[key].ステータス一覧.push(status);
  });
  return list.concat(keys.map(key=>{
    const c=byKey[key], o=c.o, id='別ルート|'+key, prev=bySheet[id];
    const rowKey=取り置き_行キー_({ban:o.ban,sku:o.sku,code:o.code});
    const 台帳確保数=Number(securedByKey&&securedByKey[rowKey])||0;
    const shelf=Number(現物登録ByKey&&現物登録ByKey[rowKey])||0;
    const held=台帳確保数+shelf;
    const secured=Math.max(held,c.入荷日確保);
    const shortage=Math.max(0,c.qty-secured);
    // 内訳は事実の列挙(・区切り)。確保済みは大きい方なので足し算に見える+は使わない
    const 内訳=[c.入荷日確保?'入荷日'+c.入荷日確保:'',shelf?'現物'+shelf:'',台帳確保数?'自動'+台帳確保数:''].filter(String).join('・');
    return {取置ID:id,受注番号:String(o.ban),氏名:String(o.氏名||''),商品コード:取り置き_商品コード_(o.sku,o.code),SKU:String(o.sku||''),
      注文数量:c.qty,確保済み:secured,不足:shortage,確保内訳:内訳,現在の状態:'別ルート',受注ステータス:c.ステータス一覧.join(' / '),旧入荷日:c.入荷日,旧EMS:'',
      台帳確保:'',台帳確保数,棚確認:'',追加数量:'',マイナス数量:'',メモ:prev?String(prev.メモ==null?'':prev.メモ).trim():'',
      判定:'別ルート',要対応:shortage>0?'追加数量で棚登録するか、受注明細に入荷日を入れて確保':'',処理:''};
  }));
}

// 棚まで見に行くべき行の自動判定。
// 旧帳簿が「着いているはず」(旧入荷日/旧EMSあり)なのに、数量も棚確認も未入力の行だけ「要棚確認」。
// 数量を入れる or 棚確認(出荷済み/未着など)を選べば解決扱い。
// 台帳(EMS引当等)が注文の全数を確保済みの行は、棚の現物が④の管理下にあるため対象外。
// 一部だけ確保済みの行は残りの現物を確かめるべきなので、旧情報が無くても要棚確認に残す。
function 取り置き_棚確認判定_(c){
  const qty=String(c&&c.追加数量==null?'':c.追加数量).trim()+String(c&&c.マイナス数量==null?'':c.マイナス数量).trim();
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
    // 差分入力方式(2026-07-23)では入力欄の自動クリアは不要(追加/マイナスは行為であって状態ではない)
    return Object.assign({},c,{
      台帳確保数:secured,
      台帳確保: full? '台帳確保済み'+secured+'個'
                    : '台帳確保済み'+secured+'個／残り'+Math.max(0,ordered-secured)+'個要確認'
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

// 確保内訳の表記(2026-07-23 ユーザー要望「自動ってどこの箱からとっているかわからん」)。
// ④/②の自動確保は元EMS番号ごとに「自動N(箱番号)」で出し、どの箱を開ければ現物があるか分かるようにする。
// 引当段階=現物確認済み(現物確認移行・復元を含む)は棚で数えた現物なので、元EMS番号が記録に残っていても
// 箱番号は出さず「現物N」(=棚にある・箱は開けない)。引当段階=先行は帳簿のみ(現物なし)で「先行N(箱番号)」。
// この画面の追加数量で登録した分も「現物N」へ合算する(現物確認移行も棚登録も「俺が数えた棚の現物」で同じ)。
function 取り置き_確保内訳表記_(autoRows, shelfQty){
  const byLabel={}, order=[];
  (autoRows||[]).forEach(r=>{
    const stage=String(r&&r.引当段階||'').trim();
    const ems=String(r&&r.元EMS番号||'').trim();
    let label;
    if(stage===TORIOKI_STAGE.PLANNED) label='先行|'+(ems||'?');
    else if(stage===TORIOKI_STAGE.PHYSICAL || !ems) label='現物|';
    else label='自動|'+ems;
    if(!(label in byLabel)){ byLabel[label]=0; order.push(label); }
    byLabel[label]+=取り置き_整数_(r&&r.取り置き数量);
  });
  if(shelfQty>0){
    const l='現物|';
    if(!(l in byLabel)){ byLabel[l]=0; order.push(l); }
    byLabel[l]+=shelfQty;
  }
  const parts=order.filter(l=>byLabel[l]>0).map(l=>{
    const kind=l.slice(0,2), ems=l.slice(3);
    return kind==='現物'? kind+byLabel[l] : kind+byLabel[l]+'('+ems+')';
  });
  return parts.join('+');
}

// 自動確保の時期分類(2026-07-23 ユーザー設計): 先行行があれば「先行」、今日②が登録した行があれば「当日」、
// それ以外(以前の箱・現物確認済み・移行復元)は「過去」。混在時は注意が要る方(先行>当日>過去)を返す。
function 取り置き_確保時期_(autoRows, todayYmd){
  let has先行=false, has当日=false;
  (autoRows||[]).forEach(r=>{
    if(String(r&&r.引当段階||'').trim()===TORIOKI_STAGE.PLANNED){ has先行=true; return; }
    const reg=ymd_(r&&r.登録日時);
    if(reg && todayYmd && reg===todayYmd) has当日=true;
  });
  return has先行?'先行':has当日?'当日':'過去';
}

// 確保の証拠と矛盾する古い棚確認の自己修復＋確保状態の自動表示(2026-07-23 ユーザー要望
// 「確保になっているやつは部分在庫になってくれないとわけがわからん」「時期で色を分けよう」)。
// ・確保済み>0の行(全数・一部どちらも)は確保時期に応じて「過去部分在庫／当日部分在庫／先行部分在庫」を
//   自動表示。空欄・未着・古い部分在庫系(旧名「部分在庫」含む)はその日の分類へ貼り替える
//   (日をまたげば当日→過去部分在庫へ自動で落ち着く)。残りが要るかは不足列の数字で見る。
// ・確保済み0の行は「未着」を自動表示。ただし旧入荷日/旧EMSの到着証拠がある行は空欄のまま=要棚確認(黄色)
//   (未着を貼ると要棚確認が消え、着いているはずの現物の登録漏れを見逃すため)。
// 発送待ち・出荷済み・予約など他の判断は上書きしない。
function 取り置き_確保表示整合_(candidates){
  return (candidates||[]).map(c=>{
    const check=String(c.棚確認==null?'':c.棚確認).trim();
    const secured=(Number(c.確保済み)||0)>0;
    const 旧証拠=String(c.旧入荷日==null?'':c.旧入荷日).trim()!=='' || String(c.旧EMS==null?'':c.旧EMS).trim()!=='';
    const 系=TORIOKI_部分在庫系.indexOf(check)>=0;
    if(secured){
      const want=c.確保時期==='先行'?'先行部分在庫':c.確保時期==='当日'?'当日部分在庫':'過去部分在庫';
      if(check===''||check==='未着'||系) return check===want? c : Object.assign({},c,{棚確認:want});
      return c;
    }
    // 自動で貼った「未着」も毎回貼り直す/剥がす。到着証拠(旧入荷日/旧EMS)が付いたら空へ戻して
    // 要棚確認(黄色)に出す。貼る瞬間だけの判定にすると未着が焼き付き、棚を見に行く合図が二度と出ない
    if(系 || check==='' || check==='未着') return 旧証拠
      ? (check===''? c : Object.assign({},c,{棚確認:''}))
      : (check==='未着'? c : Object.assign({},c,{棚確認:'未着'}));
    return c;
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
    let check=String(prev.棚確認==null?'':prev.棚確認).trim();
    let memo=String(prev.メモ==null?'':prev.メモ).trim();
    if(!check && TORIOKI_棚確認.indexOf(memo)>=0){ check=memo; memo=''; } // 旧メモの分類語はプルダウンへ移す
    const add=prev.追加数量, sub=prev.マイナス数量;
    return Object.assign({},c,{
      追加数量:(add!=null&&String(add).trim()!=='')?add:(c.追加数量!=null?c.追加数量:''),
      マイナス数量:(sub!=null&&String(sub).trim()!=='')?sub:(c.マイナス数量!=null?c.マイナス数量:''),
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
        台帳確保:'確保済み'+qty+'個',台帳確保数:qty,棚確認:'',追加数量:'',マイナス数量:'',
        メモ:String(prev.メモ||'').trim()||reason,判定:'棚戻し待ち',要対応:'棚へ戻す '+qty+'個',処理:String(prev.処理||'').trim()
      };
    });
}

function 取り置き_初期確定計画_(inputRows, existingRows, now, opts){
  const 設定=opts||{};
  const errors=[], targets={}, inputIds=new Set(), 自動解除要求=[], マイナス適用=[];
  // 発送済み・解除済みになった開始前在庫行は履歴。再確定で復活も消滅もさせない(誤操作ガード)
  const lockedIds=new Set((existingRows||[]).filter(r=>r.取置元種別==='開始前在庫' && r.状態!==TORIOKI_STATUS.ACTIVE).map(r=>String(r.取置ID||'')));
  // 棚登録優先(2026-07-23 ユーザー要望「取り置き登録だけですませたい」): 棚の現物を登録して注文が
  // 埋まるなら、④/②の自動確保は同じ反映の中で自動解除して箱へ返す(手動解除→棚登録の二段運用を廃止)。
  // 棚登録単独で注文を超える入力だけは真の誤入力として従来どおり止める。
  const 確保=取り置き_台帳確保集計_(existingRows);
  // 既存の開始前在庫(自分の棚登録)の数量。差分入力(追加/マイナス)の起点になる
  const initQtyById={};
  (existingRows||[]).forEach(r=>{
    if(r.取置元種別!=='開始前在庫' || r.状態!==TORIOKI_STATUS.ACTIVE) return;
    initQtyById[String(r.取置ID||'')]=(initQtyById[String(r.取置ID||'')]||0)+取り置き_整数_(r.取り置き数量);
  });
  (inputRows||[]).forEach((r,index)=>{
    const addRaw=r.追加数量, subRaw=r.マイナス数量;
    const addBlank=addRaw==null||String(addRaw).trim()==='', subBlank=subRaw==null||String(subRaw).trim()==='';
    // 表示専用行(即納・別ルート)は数量を入れさせない。入っていたら誤登録なので止める。
    const id=String(r.取置ID||'');
    if(id.indexOf('即納|')===0){
      if(!addBlank||!subBlank) errors.push('受注'+r.受注番号+' '+r.商品コード+': 即納行は表示専用です。数量は取り寄せ行へ入力してください');
      return;
    }
    // 別ルート(台湾/中国)もこの画面で+−できる(2026-07-23)。台帳へ現物確認済みとして登録され、
    // 受注明細の入荷日方式と二重計上しないよう④側で自動控除する(別ルート二重控除_)
    if(addBlank && subBlank) return; // 空欄=何もしない(空欄クリアによる解除は廃止 2026-07-23)
    // 行単位適用(仕様§8): この行のエラーは他行の反映を妨げない。エラー行はinputIdsへ入れない
    const rowErrors=[];
    const add=addBlank?0:Number(addRaw), sub=subBlank?0:Number(subRaw);
    const シート行=index+TORIOKI_CFG.初期HDR行+1; // 見出し3行目・データ4行目〜(エラー文言の行番号を実シートに合わせる)
    if(!addBlank && (!Number.isFinite(add)||!Number.isInteger(add)||add<=0)) rowErrors.push('初期登録'+シート行+'行 / 受注'+r.受注番号+': 追加数量は正の整数で入力');
    if(!subBlank && (!Number.isFinite(sub)||!Number.isInteger(sub)||sub<=0)) rowErrors.push('初期登録'+シート行+'行 / 受注'+r.受注番号+': マイナス数量は正の整数で入力');
    if(!addBlank && !subBlank) rowErrors.push('受注'+r.受注番号+' '+r.商品コード+': 追加とマイナスは同時に入れられません');
    const cur=initQtyById[id]||0, ordered=Number(r.注文数量)||0;
    if(!rowErrors.length){
      if(sub>cur) rowErrors.push('受注'+r.受注番号+' '+r.商品コード+': マイナス'+sub+'は自分の棚登録'+cur+'個を超えます(④確保分の解除は「選択した取り置きを手動解除」で)');
      const next=cur+add-sub;
      const 台帳確保数=Number(確保[取り置き_行キー_(r)])||0;
      if(next>ordered)
        rowErrors.push('受注'+r.受注番号+' '+r.商品コード+': 棚登録'+next+'個が注文'+ordered+'を超えます');
      const 解除必要=Math.max(0,Math.min(next,ordered)+台帳確保数-ordered); // 棚登録で押し出される④確保分
      if(lockedIds.has(id)) rowErrors.push('受注'+r.受注番号+': 既に発送済み/解除済みの初期登録行は変更できません');
      const check=String(r.棚確認==null?'':r.棚確認).trim();
      // 「未着」は確保0のとき機械が自動で貼るラベル。棚で現物を見つけた人の入力の方が新しく強い情報なので
      // 追加数量を止めない(反映後は確保>0になり次の更新で部分在庫へ貼り替わる)。人が選ぶ出荷済み/予約だけ止める
      if(add>0 && (check==='出荷済み'||check==='予約')) rowErrors.push('受注'+r.受注番号+': 棚確認が「'+check+'」なのに追加数量が入っています(どちらかを直してください)');
      if(rowErrors.length){ rowErrors.forEach(e=>errors.push(e)); return; }
      inputIds.add(id);
      if(解除必要>0) 自動解除要求.push({key:取り置き_行キー_(r),数量:解除必要,受注番号:String(r.受注番号||''),商品コード:String(r.商品コード||'')});
      if(sub>0 && cur>0) マイナス適用.push({id,数量:Math.min(sub,cur),受注番号:String(r.受注番号||''),商品コード:String(r.商品コード||''),SKU:String(r.SKU||'')});
      if(next>0) targets[id]=Object.assign({},r,{取り置き数量:next});
      // next===0 は解除(inputIdsに入りtargets無し=既存行が外れる)
      return;
    }
    rowErrors.forEach(e=>errors.push(e));
  });
  const 解除数=(existingRows||[]).filter(r=>r.取置元種別==='開始前在庫' && r.状態===TORIOKI_STATUS.ACTIVE
    && inputIds.has(String(r.取置ID||'')) && !(String(r.取置ID||'') in targets)).length;
  const kept=(existingRows||[]).filter(r=>r.取置元種別!=='開始前在庫' || r.状態!==TORIOKI_STATUS.ACTIVE || !inputIds.has(String(r.取置ID||'')));
  Object.keys(targets).forEach(id=>{
    const r=targets[id];
    // 新規登録は最初から現物確認済み段階(棚で数えた現物=供給解放)。要移行に戻らず全件反映も塞がない
    kept.push({取置ID:id,状態:TORIOKI_STATUS.ACTIVE,受注番号:r.受注番号,商品コード:r.商品コード,SKU:r.SKU,
      取り置き数量:r.取り置き数量,取置元種別:'開始前在庫',引当段階:'現物確認済み',供給処理:'供給解放',現物確認日時:now,
      元EMS番号:'',元EMS商品コード:'',元取置ID:'',登録日時:now,更新日時:now,
      戻し処理結果:'','終了理由・メモ':String(r.メモ||'')});
  });
  // 棚登録優先の自動解除: 押し出された④確保を新しい登録から順に解除(棚の現物が正・箱の分は余りへ返す)。
  // 行の途中までなら分割し、解除分は別行(取置ID+|棚解除)で履歴に残す。
  const 自動解除=[];
  自動解除要求.forEach(req=>{
    let 残り=req.数量;
    kept.map((row,i)=>({row,i}))
      .filter(x=>x.row.状態===TORIOKI_STATUS.ACTIVE && String(x.row.取置元種別||'')!=='開始前在庫' && 取り置き_行キー_(x.row)===req.key)
      .sort((a,b)=>String(b.row.登録日時||'').localeCompare(String(a.row.登録日時||'')))
      .forEach(x=>{
        if(残り<=0) return;
        const qty=取り置き_整数_(x.row.取り置き数量), take=Math.min(qty,残り);
        if(take<=0) return;
        残り-=take;
        const memo='棚登録優先で自動解除';
        if(take===qty){
          const 既存メモ=String(x.row['終了理由・メモ']||'').trim();
          kept[x.i]=Object.assign({},x.row,{状態:TORIOKI_STATUS.RELEASED,更新日時:now,
            '終了理由・メモ':既存メモ?既存メモ+' / '+memo:memo});
        }else{
          kept[x.i]=Object.assign({},x.row,{取り置き数量:qty-take,
            引当系譜数量:Math.max(0,(取り置き_整数_(x.row.引当系譜数量)||qty)-take),更新日時:now});
          kept.push(Object.assign({},x.row,{取置ID:String(x.row.取置ID||'')+'|棚解除',取り置き数量:take,
            引当系譜数量:take,状態:TORIOKI_STATUS.RELEASED,更新日時:now,'終了理由・メモ':memo+'(分割)'}));
        }
        自動解除.push({受注番号:req.受注番号,商品コード:req.商品コード,元EMS番号:String(x.row.元EMS番号||''),数量:take});
      });
    if(残り>0) errors.push('受注'+req.受注番号+' '+req.商品コード+': ④確保の自動解除が'+残り+'個分見つかりませんでした(台帳を確認してください)');
  });
  // マイナス解除の現物を在庫へ戻す(2026-07-23 幽霊現物防止): 反映側で「現物あり」を選んだ時だけ、
  // 外した数量をキャンセル戻し(現物あり)として台帳へ残す→②の再引当→残ればYahoo戻し候補、の
  // 既存配管に乗る(手動の★Yahoo在庫変更が不要になる)。登録間違いの時は従来どおり帳簿から消すだけ。
  const 在庫戻し=[];
  if(設定.マイナス現物あり && マイナス適用.length){
    const 既存ID2=new Set(kept.map(r=>String(r.取置ID||'')));
    マイナス適用.forEach(m=>{
      if(m.数量<=0) return;
      let id=String(m.id)+'|棚外し', n=1;
      while(既存ID2.has(id)){ n++; id=String(m.id)+'|棚外し'+n; }
      既存ID2.add(id);
      kept.push({取置ID:id,状態:TORIOKI_STATUS.RETURN,受注番号:m.受注番号,商品コード:m.商品コード,SKU:m.SKU,
        取り置き数量:m.数量,取置元種別:'開始前在庫',引当段階:'現物確認済み',元EMS番号:'',元EMS商品コード:'',
        元取置ID:String(m.id),登録日時:now,更新日時:now,戻し処理結果:TORIOKI_RETURN.PRESENT,
        '終了理由・メモ':'マイナス解除(現物あり→在庫へ戻す)'});
      在庫戻し.push(m);
    });
  }
  return {rows:kept,errors,適用:{行:Object.keys(targets).length,
    数量:Object.keys(targets).reduce((n,id)=>n+targets[id].取り置き数量,0),解除:解除数,
    自動解除数量:自動解除.reduce((n,d)=>n+d.数量,0),自動解除:自動解除,
    マイナス数量:マイナス適用.reduce((n,m)=>n+m.数量,0),
    在庫戻し数量:在庫戻し.reduce((n,m)=>n+m.数量,0)}};
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
function 取り置き_統合反映計画_(inputs, ledger, now, opts){
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
    取り置き行:normal.filter(r=>(Number(r.追加数量)||0)>0||(Number(r.マイナス数量)||0)>0).length,
    取り置き数量:normal.reduce((sum,r)=>sum+(Number(r.追加数量)||0),0),
    棚へ戻した:actions.filter(r=>String(r.処理||'').trim()==='棚へ戻した').length,
    現物なし:actions.filter(r=>String(r.処理||'').trim()==='現物なし').length
  };
  const initial=取り置き_初期確定計画_(normal,ledger,now,opts);
  const returnInputs=actions.filter(r=>String(r.処理||'').trim()).map(r=>{
    const choice=String(r.処理||'').trim();
    return {取置ID:r.取置ID,現物確認:choice==='棚へ戻した'?'現物あり':choice==='現物なし'?'在庫なし':choice};
  });
  const returned=取り置き_戻し確認計画_(returnInputs,initial.rows,now);
  // 行単位適用(2026-07-22): エラーのキーだけスキップして適用可能な変更を1回で保存する。
  // countsは「実際に適用される」件数(入力意図ではなく)。
  return {rows:returned.rows,errors:initial.errors.concat(returned.errors),
    counts:{取り置き行:initial.適用.行,取り置き数量:initial.適用.数量,解除:initial.適用.解除,
      自動解除数量:initial.適用.自動解除数量||0,自動解除:initial.適用.自動解除||[],
      マイナス数量:initial.適用.マイナス数量||0,在庫戻し数量:initial.適用.在庫戻し数量||0,
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
  // 見出し行は先頭5行から自動検出(取り置き登録は3行目・台帳等は1行目。移行途中の旧レイアウトも読める)
  const 探索行数=Math.min(5,sh.getLastRow());
  const 先頭=sh.getRange(1,1,探索行数,sh.getLastColumn()).getDisplayValues()
    .map(row=>row.map(v=>String(v||'').trim()));
  let hr=先頭.findIndex(row=>row.indexOf(String(headers[0]))>=0 && row.indexOf(String(headers[1]))>=0);
  if(hr<0) hr=0;
  const head=先頭[hr];
  const index={}; headers.forEach(h=>index[h]=head.indexOf(h));
  // 台帳本体と台帳のコピー(取り置き台帳_全件再計算前_/移行前_等の監査・基準シート)は
  // 三段階21列より前の旧14列でも読める(新列は空として返す)
  const 台帳系=sheetName===TORIOKI_CFG.台帳 || String(sheetName).indexOf(TORIOKI_CFG.台帳+'_')===0;
  const required=台帳系 ? headers.slice(0,14) : headers;
  if(required.some(h=>index[h]<0)) throw new Error(sheetName+'の見出し不足: '+required.filter(h=>index[h]<0).join(','));
  const dataStart=hr+2; // 見出しの次の行から(1始まり)
  if(sh.getLastRow()<dataStart) return [];
  return sh.getRange(dataStart,1,sh.getLastRow()-dataStart+1,sh.getLastColumn()).getValues().map(row=>{
    const obj={}; headers.forEach(h=>obj[h]=index[h]<0?'':row[index[h]]); return obj;
  }).filter(obj=>String(obj[headers[0]]||'').trim());
}

function 取り置き_表を保存_(sheetName, headers, rows, headerRow){
  const hr=Math.max(1,Number(headerRow)||1); // 取り置き登録は3(上2行=ボタン置き場)。台帳等は従来どおり1
  const ss=SpreadsheetApp.getActive(); let sh=ss.getSheetByName(sheetName); if(!sh) sh=ss.insertSheet(sheetName);
  if(sh.getMaxColumns()<headers.length) sh.insertColumnsAfter(sh.getMaxColumns(),headers.length-sh.getMaxColumns());
  if(sh.getMaxRows()<rows.length+hr) sh.insertRowsAfter(sh.getMaxRows(),rows.length+hr-sh.getMaxRows());
  sh.getRange(hr,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
  sh.setFrozenRows(hr);
  const dataRows=Math.max(rows.length,Math.max(0,sh.getLastRow()-hr)), dataCols=headers.length;
  if(dataRows>0){
    const values=Array.from({length:dataRows},(_,rowIndex)=>{
      const source=rows[rowIndex];
      return headers.map(header=>source&&source[header]!=null?source[header]:'');
    });
    sh.getRange(hr+1,1,dataRows,dataCols).setValues(values);
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
    if(typeof 引当_行出荷済み_==='function' && 引当_行出荷済み_(row,M)) continue; // 分割出荷済みの行は取り置き対象外
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
  if(!sheetRows.length) sheetRows=読む(TORIOKI_CFG.初期,['取置ID','棚確認','メモ']); // 旧レイアウト(絶対値入力時代)からは棚確認とメモだけ引き継ぐ
  const ledgerRows=取り置き台帳_読む_();
  candidates=取り置き_登録シート引き継ぎ_(candidates,sheetRows,ledgerRows);
  // 入力永続化(2026-07-22 仕様§8): 洗い替え前のシート入力を非表示シートへupsertし、生成行へ復元する。
  // 全数確保クリアで画面から消える数量も「未反映現物確認数量」として保存され、黙って失われない。
  let 入力保存=読む(TORIOKI_INPUT_CFG.シート,TORIOKI_INPUT_CFG.HDR);
  if(!入力保存.length) 入力保存=読む(TORIOKI_INPUT_CFG.シート,TORIOKI_INPUT_CFG.旧HDR); // 旧レイアウトからは棚確認・メモだけ引き継ぐ(絶対値の数量は差分へ持ち込まない)
  const 保存時刻=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd HH:mm:ss');
  const 保存マージ=取り置き_入力保存マージ_(candidates,入力保存,sheetRows,保存時刻);
  candidates=保存マージ.rows;
  取り置き_表を保存_(TORIOKI_INPUT_CFG.シート,TORIOKI_INPUT_CFG.HDR,Object.keys(保存マージ.store).sort().map(k=>保存マージ.store[k]));
  try{ ss.getSheetByName(TORIOKI_INPUT_CFG.シート).hideSheet(); }catch(e){}
  // ④が既に台帳(EMS引当等)で確保している分を重ねる(全数確保は入力欄クリア+黄色対象外)
  const 台帳確保マップ=取り置き_台帳確保集計_(ledgerRows);
  candidates=取り置き_台帳確保を適用_(candidates,台帳確保マップ);
  // 確保済み/不足の数字列(2026-07-23): 確保済み=④等の台帳確保+自分が登録済みの棚の現物。
  // 入力中でまだ反映していない数量は含めない(反映して初めて確保になる)
  const 現物登録={}, 自動確保行={};
  ledgerRows.forEach(r=>{
    if(!r || r.状態!==TORIOKI_STATUS.ACTIVE) return;
    const k=取り置き_行キー_(r);
    if(String(r.取置元種別||'')==='開始前在庫'){ 現物登録[k]=(現物登録[k]||0)+取り置き_整数_(r.取り置き数量); return; }
    (自動確保行[k]=自動確保行[k]||[]).push(r); // ④/②の自動確保。確保内訳で元EMS番号(どの箱か)を出す
  });
  candidates.forEach(c=>{
    const key=取り置き_行キー_(c);
    const auto=Number(c.台帳確保数)||0, shelf=現物登録[key]||0, secured=auto+shelf;
    c.確保済み=secured;
    c.不足=Math.max(0,(Number(c.注文数量)||0)-secured);
    // 「俺が確保したものじゃない」を見分ける列: 自動N(箱番号)=④/②が台帳で確保した分(その箱にある)、
    // 現物N=現物確認済み(移行・復元)+この画面の追加数量で登録した分(どちらも俺が数えた棚の現物)
    c.確保内訳=取り置き_確保内訳表記_(自動確保行[key],shelf);
    // 確保の時期(過去/当日/先行)。確保表示整合_が部分在庫の3色ステータスに使う
    c.確保時期=取り置き_確保時期_(自動確保行[key],保存時刻.slice(0,10));
  });
  candidates=取り置き_登録絞り込み_(candidates); // 予約中・出荷GOのステータスだけで除外
  // 全行表示化(2026-07-20): 判断済みも隠さない。旧版が残した非表示スイッチの記憶は
  // 棚確認セルの既定値として一度だけ復元し、以後は空を書き戻して廃止(記憶はシートの棚確認列が担う)
  let 棚記憶={};
  try{ 棚記憶=JSON.parse(PropertiesService.getDocumentProperties().getProperty('取り置き登録_棚確認済み')||'{}'); }catch(e){}
  const 記憶適用=取り置き_棚確認記憶を適用_(candidates,棚記憶);
  candidates=記憶適用.rows;
  try{ PropertiesService.getDocumentProperties().setProperty('取り置き登録_棚確認済み',JSON.stringify(記憶適用.store)); }catch(e){}
  // 確保済みの行へ「部分在庫」を自動表示し、矛盾した「未着」を直す(判定の前に走らせて要棚確認と整合させる)
  candidates=取り置き_確保表示整合_(candidates);
  candidates.forEach(c=>{ c.判定=取り置き_棚確認判定_(c); });
  // 対象注文の即納行を表示専用(水色)で差し込む。数量・棚確認は入力対象外=誤入力は反映時ガードが止める
  candidates=取り置き_即納行を付与_(candidates,即納orders,sheetRows);
  // 混在注文の台湾/中国(別ルート)行を橙で差し込む(+−入力可)。GoQ部分在庫=別ルート到着に気づけるように。
  // 確保済み/不足も出すため④確保と棚登録(開始前在庫)のマップを渡す(実効確保=入荷日方式との大きい方)
  candidates=取り置き_別ルート行を付与_(candidates,別ルートorders,sheetRows,台帳確保マップ,現物登録);
  // 数量0・注文キャンセルで確保を外す必要がある行も同じ画面へ集約する。
  const 戻し候補=取り置き_棚戻し候補_(ledgerRows,sheetRows);
  // 棚戻し待ち→要棚確認→通常の順。同じ受注番号の商品行は分断しない。
  candidates=取り置き_注文単位で並べる_(戻し候補.concat(candidates));
  // 受注明細にも確保済み/不足/確保内訳を書き戻す(個数の隣。表示モードで絞る前の全候補で)
  try{ if(typeof 受注明細_確保列を書く_==='function') 受注明細_確保列を書く_(candidates); }catch(e){}
  // 表示モード(2026-07-22 仕様§10): 既定「要作業」は部分在庫・要棚確認/棚戻し/要対応・現物ありの
  // 希望日待ちだけを表示。非表示になった注文の入力は入力保存シートが保持し、再表示時に復元される。
  let 表示モード='要作業';
  try{ 表示モード=PropertiesService.getDocumentProperties().getProperty('取り置き登録_表示モード')||'要作業'; }catch(e){}
  const 全注文数=new Set(candidates.map(c=>String(c.受注番号))).size;
  if(表示モード!=='すべて'){
    const byBan={}; candidates.forEach(c=>{(byBan[String(c.受注番号)]=byBan[String(c.受注番号)]||[]).push(c);});
    candidates=candidates.filter(c=>取り置き_作業対象判定_(byBan[String(c.受注番号)],表示モード));
  }
  const 表示注文数=new Set(candidates.map(c=>String(c.受注番号))).size;
  // 旧レイアウト(1行目見出し)からの一度きりの移行: ボタン置き場になる1〜2行目を掃除してから保存する
  {
    const sh0=ss.getSheetByName(TORIOKI_CFG.初期);
    if(sh0 && TORIOKI_CFG.初期HDR行>1 && sh0.getLastRow()>=1){
      const r1=sh0.getRange(1,1,1,Math.max(1,sh0.getLastColumn())).getDisplayValues()[0].map(v=>String(v||'').trim());
      if(r1.indexOf('取置ID')>=0 || r1.indexOf('受注番号')>=0)
        sh0.getRange(1,1,TORIOKI_CFG.初期HDR行-1,sh0.getMaxColumns()).clearContent().clearFormat();
    }
  }
  取り置き_表を保存_(TORIOKI_CFG.初期,TORIOKI_CFG.初期HDR,candidates,TORIOKI_CFG.初期HDR行);
  const 初期データ行=TORIOKI_CFG.初期HDR行+1;
  const sh2=SpreadsheetApp.getActive().getSheetByName(TORIOKI_CFG.初期);
  // 列数が減った時の旧列残骸(重複した判定/要対応/処理等)を掃除する(表を保存_はatomic契約のため触らない)
  if(sh2.getLastColumn()>TORIOKI_CFG.初期HDR.length){
    sh2.getRange(1,TORIOKI_CFG.初期HDR.length+1,Math.max(1,sh2.getLastRow()),sh2.getLastColumn()-TORIOKI_CFG.初期HDR.length)
      .clearContent().clearFormat();
  }
  // 列の並びが変わっても古い位置のプルダウンが残らないよう、シート全体の入力規則を消してから付け直す
  sh2.getRange(1,1,sh2.getMaxRows(),sh2.getMaxColumns()).clearDataValidations();
  if(candidates.length){
    // 通常行は棚確認、赤い棚戻し待ち行は処理のプルダウンを使う。
    const 棚確認列=TORIOKI_CFG.初期HDR.indexOf('棚確認')+1, 処理列=TORIOKI_CFG.初期HDR.indexOf('処理')+1;
    const 棚規則=SpreadsheetApp.newDataValidation().requireValueInList(TORIOKI_棚確認.slice(),true).setAllowInvalid(true).build();
    const 戻し規則=SpreadsheetApp.newDataValidation().requireValueInList(TORIOKI_戻し処理.slice(),true).setAllowInvalid(false).build();
    sh2.getRange(初期データ行,棚確認列,candidates.length,1).setDataValidation(棚規則);
    candidates.forEach((candidate,index)=>{
      if(candidate.判定==='棚戻し待ち'){
        sh2.getRange(index+初期データ行,棚確認列).clearDataValidations();
        sh2.getRange(index+初期データ行,処理列).setDataValidation(戻し規則);
      }else if(candidate.判定==='即納'){
        sh2.getRange(index+初期データ行,棚確認列).clearDataValidations(); // 表示専用行(即納)に棚確認は選ばせない
      }
    });
  }
  // 列幅は手動調整を尊重し、列構成が変わった時だけ既定幅を当てる(2026-07-23「列幅が自動で戻るのやめて」)
  try{
    const 幅props=PropertiesService.getDocumentProperties(), 幅key='取り置き登録_列幅適用済み';
    if(幅props.getProperty(幅key)!==String(TORIOKI_CFG.初期HDR.length)){
      sh2.setColumnWidth(TORIOKI_CFG.初期HDR.indexOf('要対応')+1,220);
      sh2.setColumnWidth(TORIOKI_CFG.初期HDR.indexOf('処理')+1,130);
      sh2.setColumnWidth(TORIOKI_CFG.初期HDR.indexOf('確保内訳')+1,170); // 自動N(EMS番号)が切れない幅
      幅props.setProperty(幅key,String(TORIOKI_CFG.初期HDR.length));
    }
  }catch(e){}
  sh2.hideColumns(1); // 取置IDは反映処理用の内部キー。表示はしない(列としては保持)
  取り置き_登録行書式を更新_(sh2,candidates);
  const 入力済=candidates.filter(c=>String(c.追加数量||'')!==''||String(c.マイナス数量||'')!=='').length;
  const 要確認数=candidates.filter(c=>c.判定==='要棚確認').length;
  const 戻し待ち数=candidates.filter(c=>c.判定==='棚戻し待ち').length;
  const 即納行数=candidates.filter(c=>c.判定==='即納').length;
  const 別ルート行数=candidates.filter(c=>c.判定==='別ルート').length;
  const 別ルート要対応=candidates.filter(c=>c.判定==='別ルート' && String(c.要対応||'').trim()!=='').length;
  const 全数確保=candidates.filter(c=>c.判定!=='棚戻し待ち' && (c.台帳確保数||0)>0 && c.判定!=='要棚確認').length;
  const summary={候補:candidates.length,要棚確認:要確認数,棚戻し待ち:戻し待ち数,即納表示:即納行数,別ルート表示:別ルート行数,入力済};
  if(!silent) ui.alert('取り置き登録を更新しました',
    '表示モード: '+表示モード+'（全'+全注文数+'注文中 '+表示注文数+'注文を表示。切替はメニュー👀）\n'
    +'候補'+candidates.length+'行 ／ 🔴棚戻し待ち '+戻し待ち数+'行 ／ ⚠️要棚確認 '+要確認数+'行\n\n'
    +(戻し待ち数? '赤い行は棚を確認し、右端の「処理」で「棚へ戻した」または「現物なし」を選びます。\n':'')
    +'棚で現物を見つけたら「追加数量」に＋、登録の取り消しは「マイナス数量」に−の個数を入れます。\n'
    +'空欄は「何もしない」。現物が無い行は「棚確認」で理由を選びます。\n'
    +(即納行数? '水色の「即納」行は表示専用です(そのまま出荷する現物。数量は入れません)。\n':'')
    +(別ルート行数? '橙色の「別ルート(台湾/中国)」行もここで追加数量・棚確認ができます(入荷日方式と併用可・二重計上は自動控除)'+(別ルート要対応? '（未確保 '+別ルート要対応+'行）':'')+'。\n':'')
    +'最後に「取り置き登録を反映」を1回押してください。\n'
    +(全数確保? '\n「台帳確保済み」'+全数確保+'行は④が確保済みのため数量入力不要です。':'')
    +(入力済? '\n入力済みの数量'+入力済+'行は引き継いでいます。':''),ui.ButtonSet.OK);
  return summary;
}

// 表示モード切替(仕様§10)。設定して即座に更新をかけ直す
function 取り置き_表示モード設定_(mode){
  PropertiesService.getDocumentProperties().setProperty('取り置き登録_表示モード',mode);
  取り置き初期登録を作成();
}
function 取り置き表示_要作業(){ 取り置き_表示モード設定_('要作業'); }
function 取り置き表示_部分在庫(){ 取り置き_表示モード設定_('部分在庫'); }
function 取り置き表示_希望日現物(){ 取り置き_表示モード設定_('希望日待ち・現物あり'); }
function 取り置き表示_先行引当(){ 取り置き_表示モード設定_('先行引当'); }
function 取り置き表示_すべて(){ 取り置き_表示モード設定_('すべて'); }

function 取り置き登録を反映(){ 直列_(取り置き登録を反映本体_); }
// 旧メニュー・既存の図形ボタンに割り当てた関数名も、新しい一括反映へつなぐ。
function 取り置き初期登録を確定(){ 取り置き登録を反映(); }
function 取り置き初期登録を確定本体_(){ return 取り置き登録を反映本体_(); }
function 取り置き登録を反映本体_(){
  const ui=SpreadsheetApp.getUi(), inputs=取り置き_表を読む_(TORIOKI_CFG.初期,TORIOKI_CFG.初期HDR);
  let plan=取り置き_統合反映計画_(inputs,取り置き台帳_読む_(),new Date());
  // マイナスで外す現物の行き先を確認(2026-07-23): 現物あり=キャンセル戻し(現物あり)として台帳へ残し、
  // ②の再引当→残ればYahoo戻し候補で在庫へ戻る。登録間違い=帳簿から消すだけ(現物が無いので戻す物もない)
  if((plan.counts.マイナス数量||0)>0){
    const ans=ui.alert('マイナスで外す現物 '+plan.counts.マイナス数量+'個',
      '棚に現物はありますか？\n\n'
      +'「はい」= 在庫へ戻す（②の再引当→残ればYahoo戻し候補に載る）\n'
      +'「いいえ」= 登録間違い（帳簿から消すだけ。現物は無い前提）',ui.ButtonSet.YES_NO);
    if(ans===ui.Button.YES) plan=取り置き_統合反映計画_(inputs,取り置き台帳_読む_(),new Date(),{マイナス現物あり:true});
  }
  const c=plan.counts;
  // 行単位適用(2026-07-22): エラーは該当キーだけスキップ。適用できる変更が1つも無い時だけ中止
  const 適用あり=c.取り置き行>0||(c.解除||0)>0||c.棚へ戻した>0||c.現物なし>0;
  if(!適用あり){ ui.alert('取り置き登録の反映を中止しました',plan.errors.length?plan.errors.join('\n'):'適用できる入力がありません',ui.ButtonSet.OK); return; }
  const answer=ui.alert('取り置き登録を反映します',
    '通常の取り置き '+c.取り置き行+'行 / '+c.取り置き数量+'個\n'
    +((c.自動解除数量||0)>0? '⤴️ ④の自動確保 '+c.自動解除数量+'個を解除して箱へ返します(棚登録優先: '
      +c.自動解除.map(d=>d.元EMS番号||'現物').filter((v,i,a)=>a.indexOf(v)===i).join(', ')+')\n':'')
    +((c.在庫戻し数量||0)>0? '♻️ マイナス分 '+c.在庫戻し数量+'個を在庫へ戻します(現物あり→②再引当→残ればYahoo戻し候補)\n':'')
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
    +((c.自動解除数量||0)>0? '（④確保'+c.自動解除数量+'個を棚登録優先で自動解除）':'')
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

// ボタン用の公開版(図形へ割り当てる)。内部からは従来どおり Yahoo戻し候補を更新_ を呼ぶ
function Yahoo戻し候補を更新(){ 直列_(()=>{ Yahoo戻し候補を更新_(); SpreadsheetApp.getActive().toast('Yahoo戻し候補を更新しました','♻️Yahoo戻し',5); }); }
function Yahoo戻し候補を更新_(){
  const rows=取り置き台帳_読む_().filter(r=>r.状態==='キャンセル戻し'&&r.戻し処理結果==='現物あり').map(r=>({
    取置ID:r.取置ID,商品コード:String(r.元EMS商品コード||r.商品コード||'').trim(),数量:r.取り置き数量,元EMS番号:r.元EMS番号,処理ID:'YAHOO|RETURN|'+r.取置ID,確認:''
  }));
  // 旧レイアウト(1行目見出し)からの一度きりの移行: ボタン置き場になる1〜2行目を掃除する
  // (補助処理。失敗しても候補の保存は続ける)
  try{
    const sh0=SpreadsheetApp.getActive().getSheetByName(TORIOKI_CFG.Yahoo候補);
    if(sh0 && Yahoo候補HDR行>1 && sh0.getLastRow()>=1){
      const r1=sh0.getRange(1,1,1,Math.max(1,sh0.getLastColumn())).getDisplayValues()[0].map(v=>String(v||'').trim());
      if(r1.indexOf('取置ID')>=0) sh0.getRange(1,1,Yahoo候補HDR行-1,sh0.getMaxColumns()).clearContent().clearDataValidations().clearFormat();
    }
  }catch(e){}
  取り置き_表を保存_(TORIOKI_CFG.Yahoo候補,Yahoo候補HDR,rows,Yahoo候補HDR行);
  const sh=SpreadsheetApp.getActive().getSheetByName(TORIOKI_CFG.Yahoo候補);
  const 確認列=Yahoo候補HDR.indexOf('確認')+1;
  // 見出し行や件数が減った後の空行に旧チェックボックスが残らないよう、確認列の入力規則を一度消す
  try{ sh.getRange(1,確認列,sh.getMaxRows(),1).clearDataValidations(); }catch(e){}
  if(rows.length) sh.getRange(Yahoo候補HDR行+1,確認列,rows.length,1).insertCheckboxes();
  日本在庫_戻り行を更新_(); // ②を待たずに日本在庫の「戻り」行を最新へ
}

// 日本在庫シートの「戻り」行だけを今の台帳へ合わせる(箱の余り行=状態'到着済'は触らない)。
// ②未実行・シート未作成なら何もしない。補助処理なので失敗しても呼び出し元を止めない
function 日本在庫_戻り行を更新_(){
  try{
    const sh=SpreadsheetApp.getActive().getSheetByName(HIKIATE_CFG.純在庫);
    if(!sh || sh.getLastRow()<3) return;
    const 幅=5, 件数=sh.getLastRow()-2;
    const 既存=sh.getRange(3,1,件数,幅).getDisplayValues()
      .filter(r=>String(r[0]||'').trim()!=='' && String(r[0]||'').trim()!=='戻り' && String(r[2]||'').trim()!=='');
    const rows=既存.concat(日本在庫_戻り待ち行_());
    sh.getRange(3,1,件数,幅).clearContent();
    if(rows.length) sh.getRange(3,1,rows.length,幅).setValues(rows);
    else sh.getRange(3,1).setValue('(日本在庫なし)');
  }catch(e){}
}

function キャンセル戻しをYahoo反映済みにする(){ 直列_(キャンセル戻しをYahoo反映済みにする本体_); }
function キャンセル戻しをYahoo反映済みにする本体_(){
  const ui=SpreadsheetApp.getUi(), candidates=取り置き_表を読む_(TORIOKI_CFG.Yahoo候補,Yahoo候補HDR).filter(r=>r.確認===true||String(r.確認).toUpperCase()==='TRUE');
  if(!candidates.length){ ui.alert('Yahoo戻し候補にチェックがありません'); return; }
  const answer=ui.alert('棚へ戻した現物を日本在庫へ移します',
    'チェックした'+candidates.length+'件を「日本在庫(Yahooへ足すものリスト)」へ移します。\n\n'
    +'このあと日本在庫シートの「📤 CSVデータ作成」でCSVを作り、Yahooへ貼ってください。',ui.ButtonSet.OK_CANCEL);
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
  // 箱由来の戻し(元EMS番号あり)だけEMS在庫移動台帳へ記録する。棚登録のマイナス解除など供給が無い戻しを
  // 記録すると「供給に無い使用」が生まれ全件検算の突合が崩れるため除く(2026-07-24)
  const 箱戻し=returns.filter(r=>String(r.元EMS番号||'').trim());
  const existingMoves=EMS在庫移動台帳_読む_(), plan=EMS在庫移動_戻し計画_(箱戻し,existingMoves,new Date());
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
  // 確定分を戻り待ちへ積む(日本在庫に出続け、CSVを作るまで消えない)
  try{
    const 待ち=日本在庫_戻り待ち_読む_(), 既存ID=new Set(待ち.map(r=>String(r.処理ID||'')));
    const 確定日時=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd HH:mm:ss');
    returns.forEach(r=>{
      const id='YAHOO|RETURN|'+String(r.取置ID||'');
      if(既存ID.has(id)) return; // 同じ戻しを二度積まない
      待ち.push({処理ID:id,商品コード:String(r.元EMS商品コード||r.商品コード||'').trim(),
        数量:取り置き_整数_(r.取り置き数量),確定日時,出力日時:''});
    });
    日本在庫_戻り待ち_保存_(待ち);
  }catch(e){ ui.alert('日本在庫への追加に失敗しました',e.message+'\n(台帳の確定は完了しています)',ui.ButtonSet.OK); }
  Yahoo戻し候補を更新_();
  const 残=日本在庫_戻り待ち_読む_().filter(r=>!String(r.出力日時||'').trim()).length;
  SpreadsheetApp.getActive().toast(candidates.length+'件を日本在庫へ移しました（Yahoo待ち '+残+'件）'
    +'／日本在庫シートの「📤 CSVデータ作成」でCSVを作ってください','♻️Yahoo戻し',8);
}

// 日本在庫シートの「📤 CSVデータ作成」ボタン。戻り分(Yahooへ足す待ち)をCSVにして出力日時を入れる。
// 箱の余りは⑤便締めが自動で出力するため、ここでは扱わない(便ごとの二重加算を防ぐ)
function 日本在庫CSVを作成(){ 直列_(日本在庫CSVを作成本体_); }
function 日本在庫CSVを作成本体_(){
  const ui=SpreadsheetApp.getUi();
  const 待ち=日本在庫_戻り待ち_読む_(), 未出力=待ち.filter(r=>!String(r.出力日時||'').trim());
  if(!未出力.length){ ui.alert('CSVを作る戻り分がありません',
    '日本在庫に「戻り」行がありません。\n(Yahoo戻し候補でチェック→「日本在庫へ移す」を先に行ってください。\n 箱の余りは📦⑤の便締めで自動出力されます)',ui.ButtonSet.OK); return; }
  const 結果=Yahoo在庫変更を出力本体_(null,{戻りのみ:true});
  if(!結果 || !結果.ok) return; // 出力を中止した場合は出力日時を入れない(次回も出せる)
  const 出力日時=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd HH:mm:ss');
  const 未出力ID=new Set(未出力.map(r=>String(r.処理ID||'')));
  日本在庫_戻り待ち_保存_(待ち.map(r=>未出力ID.has(String(r.処理ID||''))?Object.assign({},r,{出力日時}):r));
  日本在庫_戻り行を更新_(); // CSVにした分は日本在庫から外れる
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

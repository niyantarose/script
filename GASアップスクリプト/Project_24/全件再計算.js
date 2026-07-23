/**
 * 引当全件再構築 v3
 *
 * このファイルの前半はGASサービスへ触れない純粋計算にする。
 * EMS・発送済み・現在受注・Yahoo自由在庫から、現行の
 * 取り置き台帳 / EMS在庫移動台帳へ置換できる計画を作る。
 */

function 全件再計算_SKU正規化_(value, source){
  let code=String(value==null?'':value).trim()
    .replace(/[\s　]*(?:[（(]\s*\d{7,}\s*[）)]|[/／]\s*\d{7,})\s*$/,'')
    .toUpperCase().replace(/_/g,'-');
  const kind=String(source||'').toLowerCase();
  // a/bはYahoo・GoQ・受注のSKUにだけ付く在庫区分。EMS商品コードの実在末尾は守る。
  if(kind==='yahoo' || kind==='goq' || kind==='受注' || kind==='order') code=code.replace(/[AB]$/,'');
  // 旧コード別名(例 AISTALT01S→AISTALT01S-0)は現行コードへ寄せ、供給と需要を同じバケットで出会わせる。
  // 台帳へ書く元EMS商品コードはsourceCode(raw)のままなので、②の箱残・供給キー照合とはズレない。
  if(typeof 引当_コード別名マップ_==='function'){
    try{ const to=引当_コード別名マップ_()[code]; if(to) code=to; }catch(e){}
  }
  return code;
}

function 全件再計算_タグ受注番号_(value){
  const match=String(value==null?'':value).trim().match(/(?:[（(]\s*(\d{7,})\s*[）)]|[/／]\s*(\d{7,}))\s*$/);
  return match?String(match[1]||match[2]):'';
}

function 全件再計算_日付_(value){
  if(value instanceof Date && !isNaN(value.getTime())){
    return [value.getFullYear(),String(value.getMonth()+1).padStart(2,'0'),String(value.getDate()).padStart(2,'0')].join('-');
  }
  const s=String(value==null?'':value).trim();
  if(!s) return '';
  const m=s.match(/(20\d{2})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})/);
  if(!m) return '';
  const y=Number(m[1]),mo=Number(m[2]),d=Number(m[3]),date=new Date(y,mo-1,d);
  if(date.getFullYear()!==y || date.getMonth()!==mo-1 || date.getDate()!==d) return '';
  return [y,String(mo).padStart(2,'0'),String(d).padStart(2,'0')].join('-');
}

function 全件再計算_時刻_(value){
  if(value instanceof Date) return isNaN(value.getTime())?NaN:value.getTime();
  const s=String(value==null?'':value).trim();
  if(!s) return NaN;
  const n=Date.parse(s.replace(/\//g,'-'));
  return isNaN(n)?NaN:n;
}

function 全件再計算_正整数_(value){
  const n=Number(value);
  return Number.isInteger(n) && n>0?n:0;
}

function 全件再計算_実EMS番号_(value){
  const ems=String(value==null?'':value).trim();
  return !!ems && !/^棚卸/i.test(ems);
}

// 在庫管理の対象外コード(付属ポスター印・贈呈品・マスタ除外コード)。供給・余り・ブロックのどれにも数えない
function 全件再計算_在庫対象外コード_(code, excludeSet){
  const c=String(code==null?'':code).trim();
  if(c==='★コピペ' || /promotional/i.test(c)) return true;
  return !!(excludeSet && excludeSet.size && excludeSet.has(normCode_(c)));
}

// 発注共有「マスタ除外コード」シート(A列2行目以降・P19の商品マスタ除外と共用)を
// 在庫対象外リストとして読む。人名コード等をユーザーがここに足せば集計から消える
function 全件再計算_マスタ除外集合_(){
  try{
    const sh=発注共有を開く_().getSheetByName('マスタ除外コード');
    if(!sh || sh.getLastRow()<2) return new Set();
    return new Set(sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues()
      .map(r=>normCode_(r[0])).filter(Boolean));
  }catch(e){ return new Set(); }
}

function 全件再計算_実EMS行_(row){
  const r=row||{},ems=String(r.ems==null?r.EMS番号:r.ems).trim();
  const code=String(r.code==null?r.商品コード:r.code).trim();
  const qty=全件再計算_正整数_(r.qty==null?r.数量:r.qty);
  const arrival=全件再計算_日付_(r.arrival==null?r.到着日:r.arrival);
  const status=String(r.status==null?r.ステータス:r.status).trim();
  if(!全件再計算_実EMS番号_(ems) || !code || !qty || !arrival) return null;
  // 未着箱も先行引当の供給として受け入れる(段階は全件再計算_供給段階_が決める。2026-07-22)
  return {
    ems,code,sourceCode:code,sku:全件再計算_SKU正規化_(code,'EMS'),qty,arrival,
    row:Number(r.row==null?r.行:r.row)||0,purchaseNo:String(r.purchaseNo==null?r.購入No:r.purchaseNo||''),status,
    directBan:全件再計算_タグ受注番号_(code)
  };
}

function 全件再計算_発送内容キー_(row){
  const r=row||{};
  return [Number(r.qty)||0,String(r.status||''),全件再計算_日付_(r.shipDate),String(r.sku||''),String(r.code||''),String(r.choice||'')].join('|');
}

function 全件再計算_発送最新化_(inputRows){
  const groups={},issues=[];
  (inputRows||[]).forEach((source,index)=>{
    const r=Object.assign({},source),ban=String(r.ban||r.受注番号||'').trim(),itemId=String(r.itemId||r.商品ID||'').trim();
    if(!ban || !itemId){
      issues.push({severity:'重要',type:'発送キー不足',ban,itemId,row:index+1}); return;
    }
    const key=ban+'|'+itemId;
    r.ban=ban; r.itemId=itemId; r._index=index;
    (groups[key]=groups[key]||[]).push(r);
  });
  const rows=[];
  Object.keys(groups).sort().forEach(key=>{
    const group=groups[key], times=group.map(r=>全件再計算_時刻_(r.importedAt||r.取込日時));
    if(times.some(t=>isNaN(t))){ issues.push({severity:'重要',type:'発送取込日時不正',key}); return; }
    const latest=Math.max.apply(null,times),candidates=group.filter((r,i)=>times[i]===latest);
    const contents=Array.from(new Set(candidates.map(全件再計算_発送内容キー_)));
    if(contents.length!==1){ issues.push({severity:'重要',type:'発送最新競合',key}); return; }
    const r=Object.assign({},candidates.sort((a,b)=>a._index-b._index)[0]);
    delete r._index;
    const qty=Number(r.qty==null?r.個数:r.qty);
    if(!Number.isInteger(qty) || qty<0){ issues.push({severity:'重要',type:'発送数量不正',key}); return; }
    if(qty===0) return;
    const status=String(r.status||r.受注ステータス||'');
    if(!/処理済|発送済|出荷済/.test(status)) return;
    const shipDate=全件再計算_日付_(r.shipDate||r.出荷日);
    if(!shipDate){ issues.push({severity:'重要',type:'発送日不正',key}); return; }
    r.qty=qty; r.shipDate=shipDate; rows.push(r);
  });
  return {rows,issues};
}

function 全件再計算_CSV行_(line){
  const out=[]; let current='',quoted=false;
  for(let i=0;i<String(line||'').length;i++){
    const ch=String(line||'')[i];
    if(quoted){
      if(ch==='"' && String(line||'')[i+1]==='"'){ current+='"'; i++; }
      else if(ch==='"') quoted=false;
      else current+=ch;
    }else if(ch==='"') quoted=true;
    else if(ch===','){ out.push(current); current=''; }
    else current+=ch;
  }
  out.push(current);
  return out;
}

function 全件再計算_YahooCSV厳密集計_(text){
  const lines=String(text||'').split(/\r?\n/);
  if(lines.length<2) return {error:'Yahoo CSVが空です',a在庫:{},商品名:{}};
  const header=全件再計算_CSV行_(lines[0]).map((v,i)=>(i===0?String(v||'').replace(/^\uFEFF/,''):String(v||'')).trim());
  const required=['code','name','sub-code','quantity','allow-overdraft','stock-close'];
  const missing=required.filter(name=>header.indexOf(name)<0);
  if(missing.length) return {error:'Yahoo CSVの見出し不足: '+missing.join(','),a在庫:{},商品名:{}};
  const col={code:header.indexOf('code'),name:header.indexOf('name'),sub:header.indexOf('sub-code'),qty:header.indexOf('quantity')};
  const seen=new Set(),a在庫={},商品名={}; let subなし件数=0;
  for(let i=1;i<lines.length;i++){
    if(!String(lines[i]||'').trim()) continue;
    const row=全件再計算_CSV行_(lines[i]);
    const code=String(row[col.code]||'').trim(),sub=String(row[col.sub]||'').trim(),rawQty=String(row[col.qty]||'').trim();
    if(!code && !sub && !rawQty) continue;
    const qty=Number(rawQty);
    // 負数は在庫切れ承諾(overdraft)で正常発生するため中止しない。数値の妥当性はSKU単位で再構築側(Yahoo数量不正)が判定する
    if(!Number.isInteger(qty)) return {error:'Yahoo CSV '+(i+1)+'行: quantityが整数ではありません',a在庫:{},商品名:{}};
    if(!code){
      if(qty>0) return {error:'Yahoo CSV '+(i+1)+'行: codeが空です',a在庫:{},商品名:{}};
      continue;
    }
    // sub-code空はa/b運用外の基本商品(親行・直在庫)で正常形。a在庫に数えず件数だけ返す(棚卸のYahooCSV集計_と同じ扱い)
    if(!sub){ if(qty>0) subなし件数++; continue; }
    const unique=code+'\u0000'+sub;
    if(seen.has(unique)) return {error:'Yahoo CSV '+(i+1)+'行: code+sub-codeが重複しています: '+code+' / '+sub,a在庫:{},商品名:{}};
    seen.add(unique);
    const match=sub.match(/^(.+)([abAB])$/); if(!match || match[2].toLowerCase()!=='a') continue;
    const sku=全件再計算_SKU正規化_(sub,'Yahoo'); if(!sku) continue;
    a在庫[sku]=(a在庫[sku]||0)+qty;
    if(!商品名[sku]) 商品名[sku]=String(row[col.name]||'').trim();
  }
  return {error:'',a在庫,商品名,rowCount:seen.size,subなし件数};
}

function 全件再計算_未解決台帳作業_(rows){
  return (rows||[]).filter(row=>String(row&&row.状態||'')==='キャンセル戻し' &&
    (String(row&&row.戻し処理結果||'')==='未確認' || String(row&&row.戻し処理結果||'')==='現物あり'));
}

function 全件再計算_韓国派生クリア対象_(choice, productName){
  const text=String(choice||''),name=String(productName||'');
  if(/台湾|中国/.test(text) || /台湾|中国/.test(name)) return false;
  return /取り寄せ|取寄/.test(text);
}

function 全件再計算_ブロック供給_(supplies, blocked){
  const set=blocked instanceof Set?blocked:new Set(blocked||[]);
  return (supplies||[]).filter(row=>!set.has(全件再計算_SKU正規化_(row&&row.sourceCode||row&&row.code||'','EMS')));
}

function 全件再計算_需要SKU_(row, source){
  const r=row||{},sku=String(r.sku==null?r.SKU||r.商品SKU:r.sku).trim();
  if(sku) return 全件再計算_SKU正規化_(sku,source);
  // 商品コードは在庫枝番とは限らないため、SKUが無い場合はEMSと同じ厳格な扱いにする。
  return 全件再計算_SKU正規化_(r.code==null?r.商品コード:r.code,'EMS');
}

function 全件再計算_在庫枝番_(row){
  const r=row||{},sku=String(r.sku==null?r.SKU||r.商品SKU:r.sku).trim(),match=sku.match(/([ab])$/i);
  if(match) return match[1].toLowerCase();
  const choice=String(r.choice||r.選択肢||r['項目・選択肢']||'');
  if(/即納/.test(choice)) return 'a';
  if(/取寄せ|取り寄せ/.test(choice)) return 'b';
  return '';
}

// 供給の段階: 到着済=現物 / 在庫反映済み=過去締め済み(現役に使わない) /
// 先行=ユーザーが箱単位で状態を「先行」にした時だけ(帳簿のみ) / それ以外(未着等)=対象外。
// 未着を自動で拾わないのはユーザー指示(2026-07-22「あくまで自分で選びたい」)。
function 全件再計算_供給段階_(status){
  const text=String(status==null?'':status).trim();
  if(/到着済/.test(text)) return '到着済';
  if(/在庫反映済/.test(text)) return '過去締め済み';
  if(/先行/.test(text)) return '先行';
  return '未着';
}

// 取置IDから末尾の連番を外した素性。連番は他SKUの行数で揺れるため、昇格の同一性判定に使わない
function 全件再計算_ID素性_(id){
  return String(id||'').replace(/\|\d+$/,'');
}

function 全件再計算_台帳行_(allocation, state, source, sequence, stage){
  const a=allocation||{},d=a.demand||{},s=a.supply||{};
  const kind=state==='発送済み'?'SHIP':'ACTIVE';
  const row={
    取置ID:['REBUILD',kind,String(d.ban||''),String(d.itemId||d.row||''),String(s.ems||''),String(s.row||''),String(sequence||0)].join('|'),
    // 商品コードは受注側の親コード。行キー(受注番号|商品コード|SKU)の照合先(受注明細/GoQ CSV/通常登録の台帳行)が
    // 全て親コードなので、正規化SKUを入れると多バリエーション商品(JMEE-SPKZ等)だけ孤児/タグ注意/必要数が外れる
    状態:state,受注番号:String(d.ban||''),商品コード:String(d.code||a.sku||''),SKU:String(d.sku||''),
    取り置き数量:a.qty,取置元種別:source,元EMS番号:String(s.ems||''),元EMS商品コード:String(s.sourceCode||s.code||''),元取置ID:'',
    登録日時:'',更新日時:'',戻し処理結果:'','終了理由・メモ':'全件再計算v3 / 元EMS到着日='+String(s.arrival||'')
  };
  if(stage){
    row.引当段階=stage;
    if(stage==='先行') row.EMS到着予定日=String(s.arrival||'');
  }
  return row;
}

// 確保行の生成。既存先行行と素性が一致すれば取置IDを引き継ぎ(置換ではなく段階更新)、
// 先行→到着済になった行は昇格として報告する。
function 全件再計算_確保台帳行_(allocation, stage, sequence, plannedById, promotionRows){
  const row=全件再計算_台帳行_(allocation,'取り置き中','EMS',sequence,stage);
  const prev=plannedById?plannedById[全件再計算_ID素性_(row.取置ID)]:null;
  if(prev){
    row.取置ID=String(prev.取置ID);
    if(stage==='到着済' && String(prev.引当段階||'')==='先行' && promotionRows) promotionRows.push(Object.assign({},row));
  }
  return row;
}

// ブロックSKUの取り置き中行を落とす(開始前在庫は現物事実なので残す)。
// 商品コードは親コード保存になったため、ブロックSKU(バリエーション水準)とはSKUの正規化で突き合わせる
function 全件再計算_ブロック台帳行除外_(ledgerRows, blockedSkus){
  const list=blockedSkus||[];
  return (ledgerRows||[]).filter(row=>row.状態!=='取り置き中' || row.取置元種別==='開始前在庫' ||
    list.indexOf(全件再計算_SKU正規化_(row.SKU||row.商品コード,'GoQ'))<0);
}

function 全件再計算_台帳到着日_(row){
  const match=String(row&&row['終了理由・メモ']||'').match(/元EMS到着日=(20\d{2}-\d{2}-\d{2})/);
  return match?全件再計算_日付_(match[1]):'';
}

function 全件再計算_移動行_(allocation, destination, sequence){
  const a=allocation||{},d=a.demand||{},s=a.supply||{};
  return {
    処理ID:['REBUILD',destination,String(d.ban||''),String(d.itemId||d.row||''),String(s.ems||''),String(s.row||''),String(sequence||0)].join('|'),
    EMS番号:String(s.ems||''),商品コード:String(s.sourceCode||s.code||''),数量:a.qty,移動先:destination,処理日時:''
  };
}

function 全件再計算_需要並び_(a,b){
  const at=Number(a&&a._orderTime),bt=Number(b&&b._orderTime);
  const timeOrder=(isFinite(at)?at:Number.MAX_SAFE_INTEGER)-(isFinite(bt)?bt:Number.MAX_SAFE_INTEGER);
  return timeOrder || String(a.orderDate||'').localeCompare(String(b.orderDate||'')) ||
    String(a.ban||'').localeCompare(String(b.ban||''),undefined,{numeric:true}) ||
    (Number(a.row)||0)-(Number(b.row)||0) || String(a.itemId||'').localeCompare(String(b.itemId||''));
}

function 全件再計算_再構築_(input){
  input=input||{};
  const issues=[],allocations=[],ledgerRows=[],movementRows=[],blocked=new Set();
  const invariantErrors=[],futureReservations=[],promotionRows=[],stageSummary={};
  // 既存の先行行(現在台帳)。素性一致で取置IDを引き継ぎ、箱到着時の昇格を検出する
  const plannedById={};
  (input.currentPlanned||[]).forEach(row=>{ if(row && row.取置ID) plannedById[全件再計算_ID素性_(row.取置ID)]=row; });
  const excludeSet=input.excludeCodes instanceof Set?input.excludeCodes:new Set((input.excludeCodes||[]).map(c=>normCode_(c)));
  const blockedReason={}; // SKU→最初のブロック理由(内訳レポート用)
  const block=(sku,type)=>{ if(!sku) return; blocked.add(sku); if(!(sku in blockedReason)) blockedReason[sku]=type; };
  const suppliesBySku={},shippedBySku={},shippedImmediateBySku={},immediateBySku={},backorderBySku={},yahoo={};
  const knownEms=new Set(); // 発注共有EMSリストに載る韓国便のEMS番号(壊れた行も番号は韓国便として扱う)

  (input.supplies||[]).forEach((source,index)=>{
    // 付属ポスター印(★コピペ)・贈呈品(PromotionalItem)・マスタ除外コードは在庫管理対象外。供給にも不正にも数えない
    if(全件再計算_在庫対象外コード_(source&&source.code||source&&source.商品コード, excludeSet)){
      const 対象外ems=String(source&&source.ems||source&&source.EMS番号||'').trim();
      if(全件再計算_実EMS番号_(対象外ems)) knownEms.add(対象外ems);
      return;
    }
    const normalized=全件再計算_実EMS行_(source);
    if(!normalized){
      const ems=String(source&&source.ems||source&&source.EMS番号||'').trim();
      if(全件再計算_実EMS番号_(ems)){
        knownEms.add(ems);
        const sku=全件再計算_SKU正規化_(source&&source.code||source&&source.商品コード||'','EMS');
        issues.push({severity:'重要',type:'EMS行不正',sku,row:Number(source&&source.row)||index+1});
        block(sku,'EMS行不正');
      }
      return;
    }
    knownEms.add(normalized.ems);
    normalized._remaining=normalized.qty;
    (suppliesBySku[normalized.sku]=suppliesBySku[normalized.sku]||[]).push(normalized);
  });
  Object.keys(suppliesBySku).forEach(sku=>suppliesBySku[sku].sort((a,b)=>
    a.arrival.localeCompare(b.arrival)||(a.row-b.row)||a.ems.localeCompare(b.ems)||a.sourceCode.localeCompare(b.sourceCode)));

  (input.shipped||[]).forEach((source,index)=>{
    const qty=全件再計算_正整数_(source&&source.qty==null?source&&source.個数:source&&source.qty);
    const shipDate=全件再計算_日付_(source&&source.shipDate||source&&source.出荷日);
    const sku=全件再計算_需要SKU_(source,'GoQ');
    if(!sku || !qty || !shipDate){
      issues.push({severity:'重要',type:'発送行不正',sku,row:index+1}); block(sku,'発送行不正'); return;
    }
    const orderValue=source.orderDate||source.注文日時||'';
    const row=Object.assign({},source,{qty,shipDate,sku:String(source.sku||source.商品SKU||''),_sku:sku,
      ban:String(source.ban||source.受注番号||''),itemId:String(source.itemId||source.商品ID||''),
      orderDate:String(orderValue||''),_orderTime:全件再計算_時刻_(orderValue)});
    const branch=全件再計算_在庫枝番_(source);
    if(branch==='a') (shippedImmediateBySku[sku]=shippedImmediateBySku[sku]||[]).push(row);
    else if(branch==='b') (shippedBySku[sku]=shippedBySku[sku]||[]).push(row);
    else { issues.push({severity:'重要',type:'発送在庫区分不明',sku,ban:row.ban,itemId:row.itemId}); block(sku,'発送在庫区分不明'); }
  });

  (input.currentOrders||[]).forEach((source,index)=>{
    const rawQty=source&&source.qty==null?source&&source.個数:source&&source.qty;
    const numericQty=Number(rawQty),qty=全件再計算_正整数_(rawQty);
    const sku=全件再計算_需要SKU_(source,'受注');
    const route=String(source&&source.route||source&&source.区分||source&&source.kbn||''),orderValue=source&&source.orderDate||source&&source.注文日時||'';
    const isImmediate=/即納/.test(route),isBackorder=/取寄せ|取り寄せ/.test(route) && !/台湾|中国/.test(route);
    // ダニエル便など、発注共有EMSリストに無いEMS番号が付いた取寄せ行は別便で供給済み。
    // 韓国便の需要に数えると箱を横取りするため除外する(その便の帳簿はダニエル余り側が持つ)。
    const boxEms=String(source&&source.boxEms||source&&source.箱EMS||'').trim();
    if(isBackorder && boxEms && !knownEms.has(boxEms)){
      issues.push({severity:'情報',type:'別便供給の注文(ダニエル等)',sku,ban:String(source&&source.ban||''),detail:boxEms});
      return;
    }
    if((isImmediate||isBackorder) && String(rawQty==null?'':rawQty).trim()!=='' && (!Number.isInteger(numericQty)||numericQty<0)){
      issues.push({severity:'重要',type:'受注数量不正',sku,ban:String(source&&source.ban||''),row:Number(source&&source.row)||index+1});
      block(sku,'受注数量不正'); return;
    }
    if(!sku || !qty) return;
    const row=Object.assign({},source,{
      qty,_sku:sku,sku:String(source.sku||source.SKU||source.商品SKU||''),ban:String(source.ban||source.受注番号||''),
      itemId:String(source.itemId||source.商品ID||''),row:Number(source.row==null?index+1:source.row)||index+1,
      orderDate:String(orderValue||''),_orderTime:全件再計算_時刻_(orderValue)
    });
    if((isImmediate||isBackorder) && !isFinite(row._orderTime)){
      issues.push({severity:'重要',type:'受注日時不正',sku,ban:row.ban,row:row.row}); block(sku,'受注日時不正');
    }
    if(isImmediate) (immediateBySku[sku]=immediateBySku[sku]||[]).push(row);
    else if(isBackorder) (backorderBySku[sku]=backorderBySku[sku]||[]).push(row);
  });

  Object.keys(input.yahooA||{}).forEach(raw=>{
    const sku=全件再計算_SKU正規化_(raw,'Yahoo'),qty=Number(input.yahooA[raw]);
    if(!sku) return;
    if(!Number.isInteger(qty) || qty<0){ issues.push({severity:'重要',type:'Yahoo数量不正',sku}); block(sku,'Yahoo数量不正'); return; }
    yahoo[sku]=(yahoo[sku]||0)+qty;
  });

  // 現物確認済み(棚で数えた現物)は最優先の需要固定。行は削除・移動せずそのまま引き継ぐ。
  const physRows=[],physQtyByKey={},physBySku={};
  (input.physicalRows||[]).forEach(row=>{
    if(!row || row.状態!=='取り置き中' || String(row.引当段階||'')!=='現物確認済み') return;
    const qty=取り置き_整数_(row.取り置き数量); if(!qty) return;
    const sku=全件再計算_需要SKU_(row,'受注'); if(!sku) return;
    const copy=Object.assign({},row);
    physRows.push(copy);
    const key=取り置き_行キー_(row);
    physQtyByKey[key]=(physQtyByKey[key]||0)+qty;
    (physBySku[sku]=physBySku[sku]||[]).push(copy);
  });

  // ユーザーが棚確認で確定した開始前在庫(取り置き中)は事実データと同格に扱う。
  // 該当注文の需要から先に差し引き、同数のEMS供給も消費して「説明不能余り」への誤ブロックを防ぐ。
  const holdRows=[],holdQtyByKey={},holdBySku={};
  (input.initialHolds||[]).forEach(row=>{
    if(!row || row.状態!=='取り置き中' || row.取置元種別!=='開始前在庫') return;
    // 現物確認済み段階を持つ行はphysicalRows側で最優先固定される(二重計上・二重引き継ぎ防止)
    if(String(row.引当段階||'')==='現物確認済み') return;
    const qty=取り置き_整数_(row.取り置き数量); if(!qty) return;
    const sku=全件再計算_需要SKU_(row,'受注'); if(!sku) return;
    holdRows.push(row);
    const key=取り置き_行キー_(row);
    holdQtyByKey[key]=(holdQtyByKey[key]||0)+qty;
    holdBySku[sku]=(holdBySku[sku]||0)+qty;
  });

  const skuSet=new Set([].concat(Object.keys(suppliesBySku),Object.keys(shippedBySku),Object.keys(shippedImmediateBySku),Object.keys(immediateBySku),Object.keys(backorderBySku),Object.keys(yahoo),Object.keys(holdBySku),Object.keys(physBySku)));
  const summary=[];
  let ledgerSequence=0,movementSequence=0;

  Array.from(skuSet).sort().forEach(sku=>{
    const supplies=suppliesBySku[sku]||[],shipped=(shippedBySku[sku]||[]).slice().sort((a,b)=>
      a.shipDate.localeCompare(b.shipDate)||全件再計算_需要並び_(a,b));
    const shippedImmediate=(shippedImmediateBySku[sku]||[]).slice().sort((a,b)=>a.shipDate.localeCompare(b.shipDate)||全件再計算_需要並び_(a,b));
    const immediate=(immediateBySku[sku]||[]).slice().sort(全件再計算_需要並び_);
    const backorders=(backorderBySku[sku]||[]).slice().sort(全件再計算_需要並び_);
    const skuAllocations=[];
    const take=(demand,kind,cutoff,stageFilter)=>{
      let left=demand.qty;
      const ordered=supplies.slice().sort((a,b)=>{
        const ap=a.directBan===String(demand.ban||'')?0:a.directBan?2:1;
        const bp=b.directBan===String(demand.ban||'')?0:b.directBan?2:1;
        return ap-bp || a.arrival.localeCompare(b.arrival) || (a.row-b.row);
      });
      const stages=stageFilter?(Array.isArray(stageFilter)?stageFilter:[stageFilter]):null;
      for(const supply of ordered){
        if(left<=0) break;
        // 現役の物理確保は「到着済」からだけ(締め済み箱の穴埋め=幽霊確保の防止。2026-07-21 実例10117699)。
        // 歴史説明(発送済み/即納/Yahoo/開始前在庫)は到着済+過去締め済み、先行引当は「先行」だけを使う
        if(stages && stages.indexOf(全件再計算_供給段階_(supply.status||'到着済'))<0) continue;
        if(supply.directBan && supply.directBan!==String(demand.ban||'')) continue;
        if(cutoff && supply.arrival>cutoff) continue;
        const qty=Math.min(left,supply._remaining||0); if(qty<=0) continue;
        supply._remaining-=qty; left-=qty;
        const allocation={kind,sku,qty,supply,demand};
        allocations.push(allocation); skuAllocations.push(allocation);
      }
      return left;
    };

    shipped.forEach(demand=>{
      const before=skuAllocations.length,left=take(demand,'発送済み',demand.shipDate,['到着済','過去締め済み']);
      skuAllocations.slice(before).forEach(a=>ledgerRows.push(全件再計算_台帳行_(a,'発送済み','全件再計算_出荷実績',++ledgerSequence)));
      // 供給の記録が無い過去出荷(EMSリストの古い行削除等)は歴史ギャップ。現物は出荷済みで現在の
      // 出荷リスクは無いため情報に留め、SKU(=今の到着済在庫)をブロックしない(2026-07-21 計測1532件)
      if(left>0) issues.push({severity:'情報',type:'発送供給不足',sku,qty:left,ban:demand.ban,itemId:demand.itemId});
    });

    immediate.forEach(demand=>{
      const before=skuAllocations.length,left=take(demand,'現在即納','',['到着済','過去締め済み']);
      skuAllocations.slice(before).forEach(a=>movementRows.push(全件再計算_移動行_(a,'現在即納受注',++movementSequence)));
      if(left>0) issues.push({severity:'情報',type:'即納EMS外在庫',sku,qty:left,ban:demand.ban});
    });

    // 現物確認済みの供給控除(Yahoo保護より先=§5の順)。供給解放行は箱を消費しない。
    let physicalTotal=0;
    (physBySku[sku]||[]).forEach(row=>{
      const qty=取り置き_整数_(row.取り置き数量); physicalTotal+=qty;
      if(String(row.供給処理||'')==='供給解放') return;
      const explicit=String(row.供給控除EMS||row.元EMS番号||'').trim();
      if(explicit){
        const box=supplies.find(s=>s.ems===explicit && (!String(row.元EMS商品コード||'').trim() || s.sourceCode===String(row.元EMS商品コード).trim()));
        if(!box || (box._remaining||0)<qty){
          issues.push({severity:'重要',type:'現物供給不一致',sku,qty,ban:String(row.受注番号||''),detail:explicit});
          invariantErrors.push('現物供給不一致: '+sku+' 受注'+String(row.受注番号||'')+' '+qty+'個 (箱'+explicit+')');
          return; // 現物行自体は維持する(削除・移動しない)
        }
        box._remaining-=qty;
        return;
      }
      // 元EMS不明: 同SKUの到着済箱から古い順に控除し、控除先を証跡に残す。控除不能は推測せず停止対象
      const arrivedBoxes=supplies.filter(s=>全件再計算_供給段階_(s.status||'到着済')==='到着済');
      const available=arrivedBoxes.reduce((n,s)=>n+(s._remaining||0),0);
      if(available<qty){
        issues.push({severity:'重要',type:'現物控除不能',sku,qty,ban:String(row.受注番号||'')});
        invariantErrors.push('現物控除不能: '+sku+' 受注'+String(row.受注番号||'')+' '+qty+'個(到着済供給'+available+'個)');
        return;
      }
      let left=qty; const sources=[];
      for(const s of arrivedBoxes){ if(left<=0) break; const n=Math.min(left,s._remaining||0); if(n>0){ s._remaining-=n; left-=n; sources.push(s.ems); } }
      row.供給控除EMS=sources.join('/');
    });

    let yahooLeft=yahoo[sku]||0;
    if(yahooLeft>0){
      const demand={qty:yahooLeft,ban:'',itemId:'YAHOO-A',sku:sku+'a'},before=skuAllocations.length;
      yahooLeft=take(demand,'Yahoo自由在庫','',['到着済','過去締め済み']);
      skuAllocations.slice(before).forEach(a=>movementRows.push(全件再計算_移動行_(a,'Yahoo自由在庫',++movementSequence)));
      if(yahooLeft>0) issues.push({severity:'情報',type:'Yahoo自由在庫EMS外',sku,qty:yahooLeft});
    }

    // 棚の現物(開始前在庫)ぶんの供給を、後続の取寄せ割当より先に確保する。
    // 棚の現物はEMS履歴で説明できなくてもブロックしない(現物がユーザー確認済みの事実のため)。
    let holdLeft=holdBySku[sku]||0;
    if(holdLeft>0){
      const demand={qty:holdLeft,ban:'',itemId:'INITIAL-HOLD',sku},before=skuAllocations.length;
      holdLeft=take(demand,'開始前在庫','',['到着済','過去締め済み']);
      skuAllocations.slice(before).forEach(a=>movementRows.push(全件再計算_移動行_(a,'開始前在庫充当',++movementSequence)));
      if(holdLeft>0) issues.push({severity:'情報',type:'開始前在庫EMS外',sku,qty:holdLeft});
    }

    let arrivedAllocated=0,plannedAllocated=0,unallocated=0;
    backorders.forEach(demand=>{
      // 現物確認済み・開始前在庫が確保している数量は、この再計算では新たに割り当てない(同じ現物の二重確保防止)。
      const rowKey=取り置き_行キー_(demand);
      const phys=Math.min(demand.qty,physQtyByKey[rowKey]||0);
      if(phys>0){ physQtyByKey[rowKey]-=phys; demand=Object.assign({},demand,{qty:demand.qty-phys}); }
      const held=Math.min(demand.qty,holdQtyByKey[rowKey]||0);
      if(held>0){ holdQtyByKey[rowKey]-=held; demand=Object.assign({},demand,{qty:demand.qty-held}); }
      if(demand.qty<=0) return;
      // 到着済(物理)→未着(先行=帳簿のみ)の2パスFIFO
      let before=skuAllocations.length,left=take(demand,'現在取寄せ','','到着済');
      skuAllocations.slice(before).forEach(a=>{arrivedAllocated+=a.qty;ledgerRows.push(全件再計算_確保台帳行_(a,'到着済',++ledgerSequence,plannedById,promotionRows));});
      if(left>0){
        const rest=Object.assign({},demand,{qty:left});
        before=skuAllocations.length;
        left=take(rest,'現在取寄せ先行','','先行');
        skuAllocations.slice(before).forEach(a=>{plannedAllocated+=a.qty;const row=全件再計算_確保台帳行_(a,'先行',++ledgerSequence,plannedById,promotionRows);ledgerRows.push(row);futureReservations.push(row);});
      }
      if(left>0){ unallocated+=left; issues.push({severity:'情報',type:'現在取寄せ未引当',sku,qty:left,ban:demand.ban,row:demand.row}); }
    });

    const excessByStage={'到着済':0,'過去締め済み':0,'先行':0,'未着':0};
    supplies.forEach(s=>{ excessByStage[全件再計算_供給段階_(s.status||'到着済')]+=(s._remaining||0); });
    const excess=excessByStage['到着済']+excessByStage['過去締め済み']+excessByStage['先行']+excessByStage['未着'];
    // 到着済の余り=⑤でYahooへ移す前の正常な状態(検品済みの現物)。ブロックすると④が供給を隠し
    // ⑤のYahoo移動からも漏れるため情報に留める(2026-07-21)。ブロックは数量不正等のデータ異常だけ
    if(excessByStage['到着済']>0) issues.push({severity:'情報',type:'到着済の余り',sku,qty:excessByStage['到着済']});
    // 締め済み箱の歴史ギャップは現物の出荷リスクが無い(④の供給は到着済のみ)ため情報に留めブロックしない
    if(excessByStage['過去締め済み']>0) issues.push({severity:'情報',type:'締め済み箱の説明不能余り',sku,qty:excessByStage['過去締め済み']});
    // 未着・先行の余りは現物ではないためYahoo追加候補にしない(箱が着いてから到着済の余りとして扱う)
    if(excessByStage['先行']>0) issues.push({severity:'情報',type:'先行の余り',sku,qty:excessByStage['先行']});
    if(excessByStage['未着']>0) issues.push({severity:'情報',type:'未着の余り',sku,qty:excessByStage['未着']});
    stageSummary[sku]={現物確認済み:physicalTotal,到着済引当:arrivedAllocated,先行引当:plannedAllocated,未引当:unallocated};
    summary.push({
      sku,EMS到着済:supplies.reduce((n,s)=>n+s.qty,0),発送済:shipped.reduce((n,r)=>n+r.qty,0),発送済即納:shippedImmediate.reduce((n,r)=>n+r.qty,0),
      現在即納:immediate.reduce((n,r)=>n+r.qty,0),Yahoo自由在庫:yahoo[sku]||0,開始前在庫:holdBySku[sku]||0,
      現在取寄せ:backorders.reduce((n,r)=>n+r.qty,0),再計算後余り:excess,判定:blocked.has(sku)?'要確認':'OK'
    });
  });

  const blockedSkus=Array.from(blocked).sort();
  // ブロックSKUに現在取り置きを作ると通常④が供給を止めてもラベンダー表示が残るため除外する。
  const safeLedger=ledgerRows.filter(r=>r.状態!=='取り置き中' || blockedSkus.indexOf(r.商品コード)<0);
  // 開始前在庫・現物確認済みはユーザー確定の事実なので、ブロックSKUでも消さずにそのまま引き継ぐ。
  holdRows.forEach(row=>safeLedger.push(Object.assign({},row)));
  physRows.forEach(row=>safeLedger.push(Object.assign({},row)));
  return {ledgerRows:safeLedger,movementRows,allocations,issues,blockedSkus,blockedReason,summary,
    stageSummary,futureReservations,promotionRows,invariantErrors};
}

// ===== GASアダプター / プレビュー =====

const ZENKEN_REBUILD_CFG = Object.freeze({
  VERSION:'3-ledger-20260719',
  サマリ:'全件再計算_サマリ',明細:'全件再計算_割当明細',要確認:'全件再計算_要確認',内部:'全件再計算_内部',
  内部チャンク:40000,ガードキー:'FULL_REBUILD_V3_GUARD',ブロックキー:'FULL_REBUILD_V3_BLOCKED_SKUS'
});

function 全件再計算_バイト16進_(bytes){
  return (bytes||[]).map(value=>('0'+((Number(value)+256)%256).toString(16)).slice(-2)).join('');
}

function 全件再計算_安定化_(value){
  if(value instanceof Date) return isNaN(value.getTime())?'':value.toISOString();
  if(Array.isArray(value)) return value.map(全件再計算_安定化_);
  if(value && typeof value==='object'){
    const out={}; Object.keys(value).sort().forEach(key=>out[key]=全件再計算_安定化_(value[key])); return out;
  }
  return value==null?'':value;
}

function 全件再計算_署名_(value){
  const text=JSON.stringify(全件再計算_安定化_(value));
  return 全件再計算_バイト16進_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,text,Utilities.Charset.UTF_8));
}

function 全件再計算_Yahoo在庫を読む_(){
  const folders=DriveApp.getFoldersByName(TANAOROSHI_CFG.フォルダ名);
  if(!folders.hasNext()) throw new Error('ドライブにフォルダ「'+TANAOROSHI_CFG.フォルダ名+'」が見つかりません');
  const folder=folders.next(),files=folder.getFiles(); let newest=null;
  while(files.hasNext()){
    const file=files.next(); if(!/\.csv$/i.test(file.getName())) continue;
    if(!newest || file.getLastUpdated().getTime()>newest.getLastUpdated().getTime()) newest=file;
  }
  if(!newest) throw new Error('フォルダ「'+TANAOROSHI_CFG.フォルダ名+'」にYahoo在庫CSVがありません');
  const blob=newest.getBlob(),bytes=blob.getBytes(),parsed=全件再計算_YahooCSV厳密集計_(blob.getDataAsString('Shift_JIS'));
  if(parsed.error) throw new Error(parsed.error+'\nファイル: '+newest.getName());
  return Object.assign({},parsed,{fileId:newest.getId(),fileName:newest.getName(),updatedAt:newest.getLastUpdated(),
    md5:全件再計算_バイト16進_(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,bytes))});
}

function 全件再計算_EMS入力を読む_(){
  const sh=発注共有を開く_().getSheetByName(P_KAKUTEI_CFG.シート);
  if(!sh) throw new Error('発注共有に「'+P_KAKUTEI_CFG.シート+'」がありません');
  const hr=P_KAKUTEI_CFG.ヘッダー行,last=sh.getLastRow(),width=sh.getLastColumn();
  if(last<=hr) throw new Error('発注共有EMSリストにデータがありません');
  const header=sh.getRange(hr,1,1,width).getDisplayValues()[0].map(v=>String(v||'').trim());
  const find=(...names)=>{ for(const name of names){ const index=header.indexOf(name); if(index>=0) return index; } return -1; };
  const c={status:find('ステータス列','ステータス'),arrival:find('EMS到着日','到着日'),code:find('商品コード'),qty:find('数量'),ems:find('EMS番号'),purchaseNo:find('購入No.','購入No')};
  const missing=[]; ['status','arrival','code','qty','ems'].forEach(key=>{ if(c[key]<0) missing.push(key); });
  if(missing.length) throw new Error('発注共有EMSリストの見出し不足: '+missing.join(','));
  const values=sh.getRange(hr+1,1,last-hr,width).getValues(),display=sh.getRange(hr+1,1,last-hr,width).getDisplayValues();
  const signatureRows=[],supplies=[];
  values.forEach((row,index)=>{
    const shown=display[index],source={row:hr+1+index,status:shown[c.status],arrival:row[c.arrival],code:shown[c.code],qty:row[c.qty],ems:shown[c.ems],purchaseNo:c.purchaseNo>=0?shown[c.purchaseNo]:''};
    signatureRows.push({row:source.row,status:source.status,arrival:全件再計算_日付_(source.arrival)||String(shown[c.arrival]||''),code:source.code,qty:source.qty,ems:source.ems,purchaseNo:source.purchaseNo});
    if(/到着済|在庫反映済み/.test(source.status)) supplies.push(source);
  });
  return {supplies,signatureRows,spreadsheetId:発注共有を開く_().getId(),sheetId:sh.getSheetId()};
}

function 全件再計算_発送入力を読む_(){
  const sh=SpreadsheetApp.getActive().getSheetByName('発送済み');
  if(!sh || sh.getLastRow()<2) throw new Error('「発送済み」シートにデータがありません');
  const width=sh.getLastColumn(),header=sh.getRange(1,1,1,width).getDisplayValues()[0].map(v=>String(v||'').trim());
  const find=(...names)=>{ for(const name of names){ const index=header.indexOf(name); if(index>=0) return index; } return -1; };
  const c={importedAt:find('取込日時'),ban:find('受注番号'),itemId:find('商品ID'),status:find('受注ステータス'),code:find('商品コード'),sku:find('商品SKU','SKU'),qty:find('個数'),shipDate:find('出荷日'),orderDate:find('注文日時'),choice:find('項目・選択肢','項目選択肢')};
  const missing=[]; ['importedAt','ban','itemId','status','code','sku','qty','shipDate'].forEach(key=>{ if(c[key]<0) missing.push(key); });
  if(missing.length) throw new Error('発送済みシートの見出し不足: '+missing.join(','));
  const values=sh.getRange(2,1,sh.getLastRow()-1,width).getValues(),display=sh.getRange(2,1,sh.getLastRow()-1,width).getDisplayValues();
  const raw=values.map((row,index)=>({
    sourceRow:index+2,importedAt:row[c.importedAt]||display[index][c.importedAt],
    ban:String(display[index][c.ban]||'').replace(/^niyantarose-/i,''),itemId:String(display[index][c.itemId]||''),
    status:String(display[index][c.status]||''),code:String(display[index][c.code]||''),sku:String(display[index][c.sku]||''),
    qty:row[c.qty],shipDate:row[c.shipDate]||display[index][c.shipDate],orderDate:c.orderDate>=0?(row[c.orderDate]||display[index][c.orderDate]):'',
    choice:c.choice>=0?String(display[index][c.choice]||''):''
  }));
  const reduced=全件再計算_発送最新化_(raw);
  reduced.issues.forEach(issue=>{
    const candidate=raw.find(r=>(r.ban+'|'+r.itemId)===issue.key);
    issue.sku=candidate?全件再計算_需要SKU_(candidate,'GoQ'):'';
  });
  return {rows:reduced.rows,issues:reduced.issues,signatureRows:raw};
}

function 全件再計算_現在受注を読む_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(HIKIATE_CFG.受注);
  if(!sh) throw new Error('「'+HIKIATE_CFG.受注+'」シートがありません');
  const M=列マップ_(sh),width=sh.getLastColumn(),values=sh.getDataRange().getValues();
  const header=sh.getRange(M.hr,1,1,width).getDisplayValues()[0].map(v=>String(v||'').trim());
  const cOrder=header.indexOf('注文日時'),cItem=header.indexOf('商品ID');
  const cShipU=header.indexOf('出荷日');
  let cShipW=header.indexOf('出荷日(複数時には配送先毎)'); if(cShipW<0) cShipW=header.indexOf('出荷日（複数時には配送先毎）');
  if(M.番号<0 || M.コード<0 || M.個数<0 || M.選択肢<0) throw new Error('受注明細の必須見出しが不足しています');
  const rows=[],signatureRows=[];
  for(let i=M.hr;i<values.length;i++){
    const row=values[i],ban=String(row[M.番号]||'').replace(/^niyantarose-/i,'').trim(),code=String(row[M.コード]||'').trim();
    if(!ban && !code) continue;
    const rawQty=row[M.個数],qty=Number(rawQty)||0,choice=String(row[M.選択肢]||''),name=M.商品名>=0?String(row[M.商品名]||''):'';
    const kind=区分_(choice),other=引当_別ルート判定_(choice,name);
    const route=kind==='即納'?'即納':kind==='取り寄せ'?(other?'台湾中国取寄せ':'韓国取寄せ'):kind;
    // 分割出荷済み(行の出荷日あり)はもう送った分=需要に数えない(②と同じ判定 2026-07-23)
    const shippedSplit=(cShipU>=0&&String(row[cShipU]==null?'':row[cShipU]).trim()!=='')
      ||(cShipW>=0&&String(row[cShipW]==null?'':row[cShipW]).trim()!=='');
    const source={row:i+1,ban,itemId:cItem>=0?String(row[cItem]||''):'',code,sku:M.SKU>=0?String(row[M.SKU]||''):'',qty:rawQty,
      orderDate:cOrder>=0?row[cOrder]:'',route,choice,name,
      boxEms:M.EMS>=0?String(row[M.EMS]||'').trim():'',shippedSplit};
    signatureRows.push(source);
    if(qty>0 && !shippedSplit && (route==='即納' || route==='韓国取寄せ')) rows.push(source);
  }
  return {rows,signatureRows};
}

function 全件再計算_入力を読む_(){
  const yahoo=全件再計算_Yahoo在庫を読む_(),ems=全件再計算_EMS入力を読む_(),shipped=全件再計算_発送入力を読む_(),orders=全件再計算_現在受注を読む_();
  const ledger=取り置き台帳_読む_(),movements=EMS在庫移動台帳_読む_(),unresolved=全件再計算_未解決台帳作業_(ledger);
  const signatureSource={
    yahoo:{fileId:yahoo.fileId,fileName:yahoo.fileName,updatedAt:yahoo.updatedAt,md5:yahoo.md5,a在庫:yahoo.a在庫},
    ems:ems.signatureRows,shipped:shipped.signatureRows,orders:orders.signatureRows,ledger,movements
  };
  return {yahoo,ems,shipped,orders,ledger,movements,unresolved,signature:全件再計算_署名_(signatureSource)};
}

function 全件再計算_計画を作る_(){
  const source=全件再計算_入力を読む_();
  // 現在台帳の現物確認済み行は最優先で固定、取り置き中行は先行昇格のID引き継ぎ素性として渡す
  const activeLedger=(source.ledger||[]).filter(r=>r&&r.状態==='取り置き中');
  const result=全件再計算_再構築_({supplies:source.ems.supplies,shipped:source.shipped.rows,currentOrders:source.orders.rows,yahooA:source.yahoo.a在庫,
    initialHolds:source.ledger,physicalRows:activeLedger,currentPlanned:activeLedger,excludeCodes:全件再計算_マスタ除外集合_()});
  source.shipped.issues.forEach(issue=>{
    result.issues.push(issue);
    if(issue.sku && result.blockedSkus.indexOf(issue.sku)<0){
      result.blockedSkus.push(issue.sku);
      if(!(issue.sku in result.blockedReason)) result.blockedReason[issue.sku]=String(issue.type||'発送取込エラー');
    }
  });
  result.blockedSkus.sort();
  result.ledgerRows=全件再計算_ブロック台帳行除外_(result.ledgerRows,result.blockedSkus);
  source.unresolved.forEach(row=>result.issues.push({severity:'停止',type:'未解決キャンセル戻し',sku:'',qty:Number(row.取り置き数量)||0,ban:String(row.受注番号||''),detail:String(row.取置ID||'')}));
  // 旧開始前在庫(段階未設定=要移行)が残っている間は全件反映を止める(移行=🔄現物確認移行で仕分けしてから)。
  // 反映が新規に作る開始前在庫行は最初から現物確認済み段階を持つため対象外。仕様§13
  const 未移行=activeLedger.filter(r=>String(r.取置元種別||'')==='開始前在庫' && String(r.引当段階||'')!=='現物確認済み').length;
  if(未移行>0) result.issues.push({severity:'停止',type:'現物確認移行が未完了',sku:'',qty:未移行,detail:'🔄現物確認移行を作成→仕分け→反映してから全件反映してください'});
  const unidentifiedCritical=result.issues.filter(issue=>(issue.severity==='重要'||issue.severity==='停止') && !String(issue.sku||'').trim());
  // ブロック内訳(理由→件数)。プレビューのダイアログとサマリのメタ欄に表示する
  const ブロック内訳={};
  result.blockedSkus.forEach(sku=>{ const 理由=result.blockedReason[sku]||'その他'; ブロック内訳[理由]=(ブロック内訳[理由]||0)+1; });
  return {
    version:ZENKEN_REBUILD_CFG.VERSION,createdAt:new Date(),signature:source.signature,
    yahooMeta:{fileId:source.yahoo.fileId,fileName:source.yahoo.fileName,updatedAt:source.yahoo.updatedAt,md5:source.yahoo.md5,rowCount:source.yahoo.rowCount},
    ledgerRows:result.ledgerRows,movementRows:result.movementRows,blockedSkus:result.blockedSkus,ブロック内訳,
    summary:result.summary,issues:result.issues,allocations:result.allocations,
    stageSummary:result.stageSummary,futureReservations:result.futureReservations,
    promotionRows:result.promotionRows,invariantErrors:result.invariantErrors,
    // 不変条件エラー(現物供給不一致・現物控除不能等)が1件でもあればプレビューのみ=運用台帳を置換しない
    applyBlocked:source.unresolved.length>0 || unidentifiedCritical.length>0 || (result.invariantErrors||[]).length>0
  };
}

function 全件再計算_表を書く_(name,headers,rows){
  const ss=SpreadsheetApp.getActive(); let sh=ss.getSheetByName(name); if(!sh) sh=ss.insertSheet(name);
  sh.clearContents();
  if(sh.getMaxColumns()<headers.length) sh.insertColumnsAfter(sh.getMaxColumns(),headers.length-sh.getMaxColumns());
  if(sh.getMaxRows()<rows.length+1) sh.insertRowsAfter(sh.getMaxRows(),rows.length+1-sh.getMaxRows());
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#4472c4').setFontColor('#ffffff');
  if(rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows);
  sh.setFrozenRows(1); return sh;
}

function 全件再計算_内部保存_(plan){
  const compact={version:plan.version,createdAt:plan.createdAt,signature:plan.signature,yahooMeta:plan.yahooMeta,
    ledgerRows:plan.ledgerRows,movementRows:plan.movementRows,blockedSkus:plan.blockedSkus,applyBlocked:plan.applyBlocked};
  const json=JSON.stringify(全件再計算_安定化_(compact)),rows=[];
  for(let offset=0,index=1;offset<json.length;offset+=ZENKEN_REBUILD_CFG.内部チャンク,index++) rows.push(['PLAN',index,json.slice(offset,offset+ZENKEN_REBUILD_CFG.内部チャンク)]);
  const sh=全件再計算_表を書く_(ZENKEN_REBUILD_CFG.内部,['キー','連番','内容'],rows); sh.hideSheet();
}

function 全件再計算_プレビューを書く_(plan){
  const summaryHeaders=['SKU','EMS到着済','GoQ発送済b','GoQ発送済a(EMS対象外)','現在即納','Yahoo a自由在庫','開始前在庫','現在取寄せ','再計算後余り','判定'];
  const summaryRows=(plan.summary||[]).map(r=>[r.sku,r.EMS到着済,r.発送済,r.発送済即納,r.現在即納,r.Yahoo自由在庫,r.開始前在庫||0,r.現在取寄せ,r.再計算後余り,r.判定]);
  const summary=全件再計算_表を書く_(ZENKEN_REBUILD_CFG.サマリ,summaryHeaders,summaryRows);
  const metaCol=summaryHeaders.length+2;
  if(summary.getMaxColumns()<metaCol) summary.insertColumnsAfter(summary.getMaxColumns(),metaCol-summary.getMaxColumns());
  const ブロック内訳text=Object.keys(plan.ブロック内訳||{}).sort((a,b)=>plan.ブロック内訳[b]-plan.ブロック内訳[a])
    .map(k=>k+' '+plan.ブロック内訳[k]).join(' / ');
  summary.getRange(1,metaCol).setValue('作成: '+Utilities.formatDate(plan.createdAt,'Asia/Tokyo','yyyy-MM-dd HH:mm:ss')+' / Yahoo: '+plan.yahooMeta.fileName+' / ブロックSKU: '+plan.blockedSkus.length+(ブロック内訳text?'（'+ブロック内訳text+'）':'')+(plan.applyBlocked?' / ⛔反映停止条件あり':''));
  summary.getRange(2,metaCol).setValue('入力署名: '+plan.signature+' / Yahoo MD5: '+plan.yahooMeta.md5);
  const detailHeaders=['区分','SKU','EMS番号','EMS到着日','EMS元行','EMS商品コード','受注番号','商品ID/受注行','数量'];
  const detailRows=(plan.allocations||[]).map(a=>[a.kind,a.sku,a.supply.ems,a.supply.arrival,a.supply.row,a.supply.sourceCode||a.supply.code,a.demand.ban||'',a.demand.itemId||a.demand.row||'',a.qty]);
  全件再計算_表を書く_(ZENKEN_REBUILD_CFG.明細,detailHeaders,detailRows);
  const issueHeaders=['重要度','SKU','種別','数量','受注番号','商品ID/行/詳細'];
  const issueRows=(plan.issues||[]).map(r=>[r.severity||'',r.sku||'',r.type||'',r.qty||'',r.ban||'',r.itemId||r.row||r.detail||r.key||'']);
  全件再計算_表を書く_(ZENKEN_REBUILD_CFG.要確認,issueHeaders,issueRows);
  全件再計算_内部保存_(plan);
  SpreadsheetApp.getActive().setActiveSheet(summary);
}

function 全件再計算プレビュー(){ return 直列_(全件再計算プレビュー本体_); } // 書き込み系は直列_で排他(e591f3c)
function 全件再計算プレビュー本体_(){
  const ui=SpreadsheetApp.getUi();
  try{
    const plan=全件再計算_計画を作る_();
    全件再計算_プレビューを書く_(plan); SpreadsheetApp.flush();
    const 内訳=Object.keys(plan.ブロック内訳||{}).sort((a,b)=>plan.ブロック内訳[b]-plan.ブロック内訳[a])
      .map(k=>'・'+k+': '+plan.ブロック内訳[k]+'件').join('\n');
    ui.alert('全件再計算プレビューを作成しました',
      'Yahoo: '+plan.yahooMeta.fileName+'\nブロックSKU: '+plan.blockedSkus.length+'件'+(内訳?'\n'+内訳:'')+'\n要確認: '+plan.issues.length+'件'+
      (plan.applyBlocked?'\n\n⛔ 未解決作業またはSKUを特定できない重要エラーがあるため、現在は反映できません。':''),ui.ButtonSet.OK);
    return plan;
  }catch(error){ ui.alert('全件再計算プレビューを中止しました',error.message,ui.ButtonSet.OK); throw error; }
}

// ===== ガード付き反映 =====

function 全件再計算_ブロックSKU集合_(){
  const raw=PropertiesService.getDocumentProperties().getProperty(ZENKEN_REBUILD_CFG.ブロックキー);
  if(!raw) return new Set();
  try{ const parsed=JSON.parse(raw); return new Set(Array.isArray(parsed)?parsed:[]); }
  catch(error){ throw new Error('全件再計算のブロックSKU設定が壊れています。通常引当を中止しました。'); }
}

function 全件再計算_通常処理ガード_(){
  const raw=PropertiesService.getDocumentProperties().getProperty(ZENKEN_REBUILD_CFG.ガードキー);
  if(!raw) return null;
  try{ return JSON.parse(raw); }catch(error){ return {stage:'不明',message:raw}; }
}

function 全件再計算_ガード更新_(value){
  PropertiesService.getDocumentProperties().setProperty(ZENKEN_REBUILD_CFG.ガードキー,JSON.stringify(全件再計算_安定化_(value)));
}

function 全件再計算_内部読込_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(ZENKEN_REBUILD_CFG.内部);
  if(!sh || sh.getLastRow()<2) throw new Error('全件再計算の承認済みプレビューがありません。先にプレビューを実行してください。');
  const rows=sh.getRange(2,1,sh.getLastRow()-1,3).getValues().filter(row=>String(row[0])==='PLAN')
    .sort((a,b)=>(Number(a[1])||0)-(Number(b[1])||0));
  if(!rows.length) throw new Error('全件再計算の内部計画が空です。プレビューを作り直してください。');
  try{ return JSON.parse(rows.map(row=>String(row[2]||'')).join('')); }
  catch(error){ throw new Error('全件再計算の内部計画を復元できません。プレビューを作り直してください。'); }
}

function 全件再計算_ファイルをバックアップ_(fileId,name){
  const file=DriveApp.getFileById(fileId),parents=file.getParents();
  const copy=parents.hasNext()?file.makeCopy(name,parents.next()):file.makeCopy(name);
  if(!copy || !copy.getId()) throw new Error('バックアップIDを取得できません: '+name);
  DriveApp.getFileById(copy.getId()).getName(); // 作成直後に読み直して実在確認
  return {id:copy.getId(),name:copy.getName(),url:copy.getUrl()};
}

function 全件再計算_監査シート名_(base,timestamp){
  const ss=SpreadsheetApp.getActive(); let name=(base+'_全件再計算前_'+timestamp).slice(0,99),n=1;
  while(ss.getSheetByName(name)){ n++; name=(base+'_全件再計算前_'+timestamp+'_'+n).slice(0,99); }
  return name;
}

function 全件再計算_台帳を監査保存_(timestamp){
  const ss=SpreadsheetApp.getActive(),saved=[];
  [TORIOKI_CFG.台帳,TORIOKI_CFG.移動].forEach(base=>{
    const source=ss.getSheetByName(base); if(!source) return;
    const copy=source.copyTo(ss),name=全件再計算_監査シート名_(base,timestamp);
    copy.setName(name); copy.hideSheet(); saved.push(name);
  });
  return saved;
}

function 全件再計算_反映日時を設定_(plan,now){
  // 引き継いだ開始前在庫の登録日時は棚確認した時点の記録なので上書きしない。
  const ledger=(plan.ledgerRows||[]).map(row=>Object.assign({},row,{登録日時:row.登録日時||now,更新日時:now}));
  const movements=(plan.movementRows||[]).map(row=>Object.assign({},row,{処理日時:now}));
  return {ledger,movements};
}

function 全件再計算_韓国派生欄をクリア_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(HIKIATE_CFG.受注);
  if(!sh) throw new Error('「'+HIKIATE_CFG.受注+'」シートがありません');
  const M=列マップ_(sh),start=M.hr+1,last=sh.getLastRow();
  if(last<start) return {rows:0,入荷日:0,EMS:0};
  const values=sh.getDataRange().getValues(),count=last-start+1;
  const arrivals=M.入荷>=0?sh.getRange(start,M.入荷+1,count,1).getValues():null;
  const emsValues=M.EMS>=0?sh.getRange(start,M.EMS+1,count,1).getValues():null;
  let rows=0,arrivalCount=0,emsCount=0;
  for(let sheetRow=start;sheetRow<=last;sheetRow++){
    const row=values[sheetRow-1],index=sheetRow-start;
    const choice=M.選択肢>=0?row[M.選択肢]:'',name=M.商品名>=0?row[M.商品名]:'';
    if(!全件再計算_韓国派生クリア対象_(choice,name)) continue;
    rows++;
    if(arrivals && String(arrivals[index][0]==null?'':arrivals[index][0]).trim()){ arrivals[index][0]=''; arrivalCount++; }
    if(emsValues && String(emsValues[index][0]==null?'':emsValues[index][0]).trim()){ emsValues[index][0]=''; emsCount++; }
  }
  if(arrivals) sh.getRange(start,M.入荷+1,count,1).setValues(arrivals);
  if(emsValues) sh.getRange(start,M.EMS+1,count,1).setValues(emsValues);
  return {rows,入荷日:arrivalCount,EMS:emsCount};
}

function 全件再計算を反映(){
  const ui=SpreadsheetApp.getUi(),response=ui.prompt('全件再計算を反映',
    '先に作成したプレビューの内容で台帳・P列・韓国EMS入荷情報を作り直します。\n'+
    '2ファイルをバックアップしてから実行します。続ける場合は「全件再計算を反映」と入力してください。',ui.ButtonSet.OK_CANCEL);
  if(response.getSelectedButton()!==ui.Button.OK) return;
  if(String(response.getResponseText()||'').trim()!=='全件再計算を反映'){ ui.alert('入力が一致しないため中止しました。'); return; }
  return 直列_(全件再計算を反映本体_); // 書き込み系は直列_で排他(e591f3c)
}

function 全件再計算を反映本体_(){
  const ui=SpreadsheetApp.getUi();
  let backups=null,stage='開始前';
  try{
    const plan=全件再計算_内部読込_();
    if(plan.version!==ZENKEN_REBUILD_CFG.VERSION) throw new Error('プレビューの版が古いため反映できません。プレビューを作り直してください。');
    if(plan.applyBlocked) throw new Error('プレビューに反映停止条件があります。「全件再計算_要確認」を解消してプレビューを作り直してください。');
    stage='入力署名の再確認';
    const current=全件再計算_入力を読む_();
    if(current.signature!==plan.signature) throw new Error('プレビュー後にYahoo・EMS・発送済み・受注明細・取り置き台帳・移動台帳のいずれかが変わりました。プレビューを作り直してください。');
    const unresolved=全件再計算_未解決台帳作業_(current.ledger);
    if(unresolved.length) throw new Error('未確認のキャンセル戻しが'+unresolved.length+'件あります。物理作業を完了してから再プレビューしてください。');

    stage='Driveバックアップ';
    const timestamp=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyyMMdd_HHmmss'),active=SpreadsheetApp.getActive();
    backups={
      引当:全件再計算_ファイルをバックアップ_(active.getId(),'引当ファイル_全件再計算前_'+timestamp),
      発注共有:全件再計算_ファイルをバックアップ_(P_KAKUTEI_CFG.発注共有ID,'発注共有ファイル_全件再計算前_'+timestamp)
    };
    全件再計算_ガード更新_({startedAt:new Date(),stage,backups});

    stage='旧台帳の監査保存'; 全件再計算_ガード更新_({startedAt:new Date(),stage,backups});
    const archives=全件再計算_台帳を監査保存_(timestamp);

    stage='台帳の一括置換'; 全件再計算_ガード更新_({startedAt:new Date(),stage,backups,archives});
    const now=new Date(),dated=全件再計算_反映日時を設定_(plan,now);
    取り置き台帳_保存_(dated.ledger); EMS在庫移動台帳_保存_(dated.movements);
    PropertiesService.getDocumentProperties().setProperty(ZENKEN_REBUILD_CFG.ブロックキー,JSON.stringify(plan.blockedSkus||[]));

    stage='旧韓国派生欄のクリア'; 全件再計算_ガード更新_({startedAt:new Date(),stage,backups,archives});
    const cleared=全件再計算_韓国派生欄をクリア_();

    stage='EMS在庫更新'; 全件再計算_ガード更新_({startedAt:new Date(),stage,backups,archives});
    _発注共有SS_キャッシュ=null; EMS在庫を更新_本体_();

    stage='既存④で派生データ再生成'; 全件再計算_ガード更新_({startedAt:new Date(),stage,backups,archives});
    const allocation=引当実行_本体_({preview:false,ignoreRebuildGuard:true,clearCurrentP:true,silentSummary:true});
    if(!allocation || allocation.success!==true) throw new Error('既存④が完了結果を返しませんでした。ガードを残して停止します。');

    PropertiesService.getDocumentProperties().deleteProperty(ZENKEN_REBUILD_CFG.ガードキー);
    ui.alert('全件再計算の反映が完了しました',
      '取り置き台帳: '+dated.ledger.length+'行\nEMS在庫移動台帳: '+dated.movements.length+'行\nブロックSKU: '+(plan.blockedSkus||[]).length+'件\n\n'+
      '引当バックアップ: '+backups.引当.url+'\n発注共有バックアップ: '+backups.発注共有.url,ui.ButtonSet.OK);
    return {success:true,backups,archives,cleared};
  }catch(error){
    if(backups) 全件再計算_ガード更新_({failedAt:new Date(),stage,error:error.message,backups});
    ui.alert('全件再計算を停止しました',
      '停止段階: '+stage+'\n'+error.message+(backups?'\n\nバックアップは作成済みです。通常④は安全のため停止状態です。':''),ui.ButtonSet.OK);
    throw error;
  }
}

// 反映が途中で失敗した後の復旧用。監査シート・バックアップで台帳の状態を確認してから使う。
function 全件再計算ガードを解除(){
  const ui=SpreadsheetApp.getUi(),guard=全件再計算_通常処理ガード_();
  if(!guard){ ui.alert('全件再計算のガードは設定されていません。通常④はそのまま実行できます。'); return; }
  const response=ui.prompt('🔓 全件再計算ガードを解除(復旧用)',
    '通常④の停止を解除します。\n停止段階: '+String(guard.stage||'不明')+'\n\n'+
    '台帳が置換途中の可能性があります。監査シート(台帳_全件再計算前_…)とDriveバックアップを確認してから、「解除」と入力してください。',ui.ButtonSet.OK_CANCEL);
  if(response.getSelectedButton()!==ui.Button.OK || String(response.getResponseText()||'').trim()!=='解除'){ ui.alert('中止しました。'); return; }
  直列_(()=>PropertiesService.getDocumentProperties().deleteProperty(ZENKEN_REBUILD_CFG.ガードキー));
  ui.alert('ガードを解除しました。通常④を実行できます。');
}

// ブロックSKUの供給0扱いを解除する(元データを直して再プレビューする前提の運用ボタン)。
function 全件再計算ブロックSKUをクリア(){
  const ui=SpreadsheetApp.getUi(),props=PropertiesService.getDocumentProperties();
  const raw=props.getProperty(ZENKEN_REBUILD_CFG.ブロックキー);
  if(!raw){ ui.alert('ブロック中のSKUはありません。'); return; }
  let list=[]; try{ const parsed=JSON.parse(raw); list=Array.isArray(parsed)?parsed.slice().sort():[]; }catch(e){ list=['(設定が壊れています)']; }
  const response=ui.alert('🧹 全件再計算のブロックSKUをクリア',
    'ブロック中 '+list.length+'件の供給0扱いを解除します:\n'+list.slice(0,15).join(', ')+(list.length>15?' …他'+(list.length-15)+'件':'')+
    '\n\n解除すると通常④がこれらのSKUへ再び引き当てます。元データの修正が済んでいることを確認してください。',ui.ButtonSet.OK_CANCEL);
  if(response!==ui.Button.OK){ ui.alert('中止しました。'); return; }
  直列_(()=>props.deleteProperty(ZENKEN_REBUILD_CFG.ブロックキー));
  ui.alert('ブロックSKUをクリアしました。');
}

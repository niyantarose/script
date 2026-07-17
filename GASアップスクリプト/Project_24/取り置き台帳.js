const TORIOKI_CFG = Object.freeze({
  台帳:'取り置き台帳', 初期:'取り置き登録', 要確認:'取り置き要確認', 戻し:'キャンセル戻し確認', Yahoo候補:'Yahoo戻し候補', 移動:'EMS在庫移動台帳',
  台帳HDR:['取置ID','状態','受注番号','商品コード','SKU','取り置き数量','取置元種別','元EMS番号','元EMS商品コード','元取置ID','登録日時','更新日時','戻し処理結果','終了理由・メモ'],
  初期HDR:['取置ID','受注番号','氏名','商品コード','SKU','注文数量','現在の状態','受注ステータス','旧入荷日','旧EMS','棚確認','現物取り置き数量','メモ','判定'],
  要確認HDR:['取置ID','受注番号','商品コード','理由'],
  移動HDR:['処理ID','EMS番号','商品コード','数量','移動先','処理日時']
});
const 戻しHDR=['取置ID','受注番号','商品コード','数量','元EMS番号','現物確認','メモ'];
// 取り置き登録の「棚確認」プルダウン。出荷済み/未着/予約は数量なし(登録しない)の目印
const TORIOKI_棚確認=Object.freeze(['発送待ち','部分在庫','出荷済み','未着','予約']);

// 予約商品の判定。「予約」表記だけでは決めない:
//   ・発売予定日が読めて未来 → 予約(まだ来ないのが正常)として棚確認へ自動セット
//   ・発売予定日が過去/日付が読めない(「予約早期完売」等) → 自動では付けない(発売済みで
//     入荷している可能性があるため、人が棚確認で判断する)
function 取り置き_予約判定_(選択肢, 商品名, today){
  const text=String(選択肢||'')+' '+String(商品名||'');
  if(!/予約/.test(text)) return false;
  const base=today instanceof Date && !isNaN(today.getTime())? today : new Date();
  const full=text.match(/(20\d{2})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/); // 2026/4/29・2026年4月29日 形式
  if(full){
    return new Date(Number(full[1]),Number(full[2])-1,Number(full[3])).getTime()>base.getTime();
  }
  const m=text.match(/(\d{1,2})月/); // 「9月」「7月末」形式(年なし)。今月以降なら未来扱い
  if(m){
    const month=Number(m[1]);
    if(month>=1&&month<=12) return month>=(base.getMonth()+1);
  }
  return false;
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

function 取り置き_CSV遷移計画_(csvRows, ledgerRows, now){
  const statuses={}, errors=[], review=[];
  (csvRows||[]).forEach(r=>{
    const ban=String(r.受注番号||'').replace(/^niyantarose-/i,'').trim();
    const key=取り置き_行キー_({ban,code:r.商品コード,sku:r.SKU});
    if(!ban || !取り置き_商品コード_(r.SKU,r.商品コード)) return;
    const next=/キャンセル/.test(String(r.受注ステータス||''))?'キャンセル':/処理済|発送済|出荷済/.test(String(r.受注ステータス||''))?'発送済み':'継続';
    if(statuses[key] && statuses[key]!==next) errors.push('同じ受注行にステータス競合: '+key);
    statuses[key]=next;
  });
  if(errors.length) return {rows:[],review,errors};
  const rows=(ledgerRows||[]).map(r=>{
    const copy=Object.assign({},r); if(copy.状態!==TORIOKI_STATUS.ACTIVE) return copy;
    const key=取り置き_行キー_(copy), next=statuses[key];
    if(next==='発送済み'){ copy.状態=TORIOKI_STATUS.SHIPPED; copy.更新日時=now; }
    else if(next==='キャンセル'){ copy.状態=TORIOKI_STATUS.RETURN; copy.戻し処理結果=TORIOKI_RETURN.UNCHECKED; copy.更新日時=now; }
    else if(!next) review.push({取置ID:copy.取置ID,受注番号:copy.受注番号,商品コード:copy.商品コード,理由:'最新CSVに受注行なし'});
    return copy;
  });
  return {rows,review,errors:[]};
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
    const key=取り置き_行キー_(o);
    if(!byKey[key]){ byKey[key]={o,qty:0,state,入荷日:'',EMS:'',予約:false,ステータス:''}; keys.push(key); }
    byKey[key].qty+=Number(o.qty)||0;
    // 旧帳簿の着済情報(入荷日スタンプ/EMS番号)。棚を確認すべき行の目印として表示する
    if(!byKey[key].入荷日){ const d=ymd_(o.入荷日); if(d) byKey[key].入荷日=d; }
    if(!byKey[key].EMS && String(o.EMS||'').trim()) byKey[key].EMS=String(o.EMS).trim();
    if(o.予約) byKey[key].予約=true;
    if(!byKey[key].ステータス && String(o.ステータス||'').trim()) byKey[key].ステータス=String(o.ステータス).trim();
  });
  const rank={}; list.forEach((s,i)=>{ rank[String(s.状態||'')]=i; });
  return keys
    .sort((a,b)=>(rank[byKey[a].state]-rank[byKey[b].state]) || (a<b?-1:a>b?1:0))
    .map(key=>{
      const c=byKey[key], o=c.o;
      return {取置ID:'INIT|'+key,受注番号:String(o.ban),氏名:String(o.氏名||''),商品コード:取り置き_商品コード_(o.sku,o.code),SKU:String(o.sku||''),
        注文数量:c.qty,現在の状態:c.state,受注ステータス:c.ステータス,旧入荷日:c.入荷日,旧EMS:c.EMS,
        // 入荷日スタンプ・部分包装(梱包中=現物あり)は「予約」表記より現物の形跡が勝つ
        棚確認:(c.予約 && !c.入荷日 && !/部分包装/.test(String(c.ステータス||'')))?'予約':'',現物取り置き数量:'',メモ:'',判定:''};
    });
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

// 棚まで見に行くべき行の自動判定。
// 旧帳簿が「着いているはず」(旧入荷日/旧EMSあり)なのに、数量も棚確認も未入力の行だけ「要棚確認」。
// 数量を入れる or 棚確認(出荷済み/未着など)を選べば解決扱い。
function 取り置き_棚確認判定_(c){
  const qty=String(c&&c.現物取り置き数量==null?'':c.現物取り置き数量).trim();
  const check=String(c&&c.棚確認==null?'':c.棚確認).trim();
  if(qty!=='' || check!=='') return '';
  const 旧あり=String(c&&c.旧入荷日==null?'':c.旧入荷日).trim()!=='' || String(c&&c.旧EMS==null?'':c.旧EMS).trim()!=='';
  return 旧あり? '要棚確認' : '';
}

// 棚チェックの手間を最小にする絞り込み: 「物がある可能性が高い行」だけを残す。
//   残す = 確定済み/入力済みの数量がある行、または 旧入荷日あり(帳簿上は届いているはず)で
//          予約中・未来の予約でない行
//   落とす = 予約中の幽霊スタンプ、未来発売の予約、帳簿上まだ届いていない行(どうせ棚に無い)
function 取り置き_登録絞り込み_(rows){
  return (rows||[]).filter(c=>{
    const st=String(c.受注ステータス||'');
    // ステータス除外は入力済みでも最優先(登録済みの取り置きはCSV取込が出荷時に自動で発送済みへ落とす)
    if(/予約/.test(st)) return false;   // 予約中=幽霊スタンプ。来ないものは来ない
    if(/出荷GO/.test(st)) return false; // 揃って出荷作業に入る注文=棚チェック不要
    if(/部分包装/.test(st)) return true; // 梱包が始まっている=現物が事務所に必ずある。スタンプ無しでも出す
    if(String(c.現物取り置き数量==null?'':c.現物取り置き数量).trim()!=='') return true; // 確定/入力済みは表示
    if(String(c.旧入荷日||'').trim()==='') return false;
    if(String(c.棚確認||'')==='予約') return false;
    return true;
  });
}

// 「出荷済み/未着」と判断済みの行は次回から表示しない。判断はDocumentPropertiesに取置IDで記憶し、
// 注文が候補から消えたら記憶も自動で消える(store=次回保存する記憶)。数量入りの行は対象外(表示)。
function 取り置き_棚確認記憶を適用_(candidates, store){
  const memo=store||{}, out=[], next={};
  (candidates||[]).forEach(c=>{
    const id=String(c.取置ID||'');
    const qty=String(c.現物取り置き数量==null?'':c.現物取り置き数量).trim();
    let check=String(c.棚確認||'').trim();
    if(!check && memo[id]) check=String(memo[id]); // 過去の判断を復元
    if(qty==='' && (check==='出荷済み'||check==='未着')){ next[id]=check; return; } // 判断済み=非表示で記憶
    out.push(Object.assign({},c,{棚確認:check}));
  });
  return {rows:out, store:next};
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
      メモ: memo!==''? memo : c.メモ
    });
  });
}

function 取り置き_初期確定計画_(inputRows, existingRows, now){
  const errors=[], targets={}, inputIds=new Set();
  // 発送済み・解除済みになった開始前在庫行は履歴。再確定で復活も消滅もさせない(誤操作ガード)
  const lockedIds=new Set((existingRows||[]).filter(r=>r.取置元種別==='開始前在庫' && r.状態!==TORIOKI_STATUS.ACTIVE).map(r=>String(r.取置ID||'')));
  (inputRows||[]).forEach((r,index)=>{
    inputIds.add(String(r.取置ID||''));
    const raw=r.現物取り置き数量, blank=raw==null || raw==='', entered=blank?0:Number(raw), ordered=Number(r.注文数量)||0;
    if(!blank && (!String(raw).trim() || !Number.isFinite(entered))) errors.push('初期登録'+(index+2)+'行 / 受注'+r.受注番号+': 数量は数値で入力');
    else if(entered<0 || !Number.isInteger(entered)) errors.push('初期登録'+(index+2)+'行: 数量は0以上の整数');
    if(entered>ordered) errors.push('受注'+r.受注番号+': 現物'+entered+'が注文'+ordered+'を超過');
    if(entered>0 && lockedIds.has(String(r.取置ID||''))) errors.push('受注'+r.受注番号+': 既に発送済み/解除済みの初期登録行は変更できません');
    const check=String(r.棚確認==null?'':r.棚確認).trim();
    if(entered>0 && (check==='出荷済み'||check==='未着'||check==='予約')) errors.push('受注'+r.受注番号+': 棚確認が「'+check+'」なのに数量が入っています(どちらかを直してください)');
    if(entered>0) targets[r.取置ID]=Object.assign({},r,{取り置き数量:entered});
  });
  if(errors.length) return {rows:[],errors};
  const kept=(existingRows||[]).filter(r=>r.取置元種別!=='開始前在庫' || r.状態!==TORIOKI_STATUS.ACTIVE || !inputIds.has(String(r.取置ID||'')));
  Object.keys(targets).forEach(id=>{
    const r=targets[id];
    kept.push({取置ID:id,状態:TORIOKI_STATUS.ACTIVE,受注番号:r.受注番号,商品コード:r.商品コード,SKU:r.SKU,
      取り置き数量:r.取り置き数量,取置元種別:'開始前在庫',元EMS番号:'',元EMS商品コード:'',元取置ID:'',登録日時:now,更新日時:now,
      戻し処理結果:'','終了理由・メモ':String(r.メモ||'')});
  });
  return {rows:kept,errors:[]};
}

function 取り置き_戻し確認計画_(inputs, ledger, now){
  const byId={}; (inputs||[]).forEach(r=>byId[String(r.取置ID||'')]=String(r.現物確認||'').trim());
  const errors=[];
  Object.keys(byId).forEach(id=>{ if(byId[id] && byId[id]!=='現物あり' && byId[id]!=='在庫なし') errors.push(id+': 現物確認は「現物あり」か「在庫なし」'); });
  if(errors.length) return {rows:[],errors};
  return {rows:(ledger||[]).map(r=>{
    const choice=byId[String(r.取置ID||'')]; if(!choice) return Object.assign({},r);
    if(r.状態!=='キャンセル戻し' || r.戻し処理結果!=='未確認'){ errors.push(r.取置ID+': 未確認のキャンセル戻しではない'); return Object.assign({},r); }
    return Object.assign({},r,{戻し処理結果:choice,更新日時:now});
  }),errors};
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
  if(headers.some(h=>index[h]<0)) throw new Error(sheetName+'の見出し不足: '+headers.filter(h=>index[h]<0).join(','));
  return sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues().map(row=>{
    const obj={}; headers.forEach(h=>obj[h]=row[index[h]]); return obj;
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
function 取り置き初期登録を作成本体_(){
  const ss=SpreadsheetApp.getActive(), recv=ss.getSheetByName(HIKIATE_CFG.受注), ui=SpreadsheetApp.getUi();
  if(!recv){ ui.alert('受注明細がありません'); return; }
  const M=列マップ_(recv), values=recv.getDataRange().getValues(), orders=[], 着済スタンプ=new Set(), 部分包装=new Set();
  const 受注head=values[M.hr-1].map(v=>String(v||'').trim()), cステータス=受注head.indexOf('受注ステータス');
  for(let i=M.hr;i<values.length;i++){
    const row=values[i], ban=String(row[M.番号]||'').trim(), qty=Number(row[M.個数])||0;
    if(!ban || qty<=0 || 区分_(row[M.選択肢])!=='取り寄せ') continue;
    if(引当_別ルート判定_(row[M.選択肢], M.商品名>=0?row[M.商品名]:'')) continue; // 台湾・中国は台帳の対象外
    const 入荷日=M.入荷>=0?row[M.入荷]:'';
    const ステータス=cステータス>=0?String(row[cステータス]||''):'';
    if(String(入荷日==null?'':入荷日).trim()!=='') 着済スタンプ.add(ban); // 旧帳簿では届いているはず=棚を確認すべき注文
    if(/部分包装/.test(ステータス)) 部分包装.add(ban); // 梱包中=現物が必ずある注文(スタンプや一覧に関係なく拾う)
    orders.push({ban,氏名:M.氏名>=0?String(row[M.氏名]||''):'',code:String(row[M.コード]||''),sku:M.SKU>=0?String(row[M.SKU]||''):'',qty,
      入荷日,EMS:M.EMS>=0?String(row[M.EMS]||''):'',
      ステータス,
      予約:取り置き_予約判定_(row[M.選択肢], M.商品名>=0?row[M.商品名]:'')});
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
  let sheetRows=読む(TORIOKI_CFG.初期,['取置ID','現物取り置き数量','棚確認','メモ']);
  if(!sheetRows.length) sheetRows=読む(TORIOKI_CFG.初期,['取置ID','現物取り置き数量','メモ']);
  if(!sheetRows.length) sheetRows=読む('取り置き初期登録',['取置ID','現物取り置き数量','メモ']); // 旧名からの一度きりの引き継ぎ
  candidates=取り置き_登録シート引き継ぎ_(candidates,sheetRows,取り置き台帳_読む_());
  // 入荷日スタンプがある未入力行の「予約」は白紙へ戻す(過去の自動セットの引き継ぎ対策。届いた形跡が勝つ)
  candidates.forEach(c=>{
    if(String(c.旧入荷日||'').trim()!=='' && c.棚確認==='予約' &&
       String(c.現物取り置き数量==null?'':c.現物取り置き数量).trim()==='') c.棚確認='';
  });
  candidates=取り置き_登録絞り込み_(candidates); // 物がある可能性が高い行だけの最小リストへ
  // 出荷済み/未着と判断済みの行は記憶して以後表示しない(注文が消えれば記憶も自動掃除)
  let 棚記憶={};
  try{ 棚記憶=JSON.parse(PropertiesService.getDocumentProperties().getProperty('取り置き登録_棚確認済み')||'{}'); }catch(e){}
  const 記憶適用=取り置き_棚確認記憶を適用_(candidates,棚記憶);
  candidates=記憶適用.rows;
  try{ PropertiesService.getDocumentProperties().setProperty('取り置き登録_棚確認済み',JSON.stringify(記憶適用.store)); }catch(e){}
  candidates.forEach(c=>{ c.判定=取り置き_棚確認判定_(c); });
  // 要棚確認を上に(その中は状態グループ順のまま=元の並びを保つ安定ソート)
  candidates=candidates.map((c,i)=>({c,i}))
    .sort((a,b)=>((a.c.判定==='要棚確認')?0:1)-((b.c.判定==='要棚確認')?0:1) || a.i-b.i)
    .map(x=>x.c);
  取り置き_表を保存_(TORIOKI_CFG.初期,TORIOKI_CFG.初期HDR,candidates);
  const sh2=SpreadsheetApp.getActive().getSheetByName(TORIOKI_CFG.初期);
  // 列の並びが変わっても古い位置のプルダウンが残らないよう、シート全体の入力規則を消してから付け直す
  sh2.getRange(1,1,sh2.getMaxRows(),sh2.getMaxColumns()).clearDataValidations();
  if(candidates.length){
    // 棚確認列にプルダウン(空欄OK)を付け、要棚確認の行を黄色でマーク
    const col=TORIOKI_CFG.初期HDR.indexOf('棚確認')+1;
    sh2.getRange(2,col,candidates.length,1)
      .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(TORIOKI_棚確認.slice(),true).setAllowInvalid(true).build());
    const 幅=TORIOKI_CFG.初期HDR.length;
    sh2.getRange(2,1,candidates.length,幅).setBackgrounds(
      candidates.map(c=>new Array(幅).fill(c.判定==='要棚確認'? '#fff2cc' : null)));
  }
  const 入力済=candidates.filter(c=>String(c.現物取り置き数量)!=='').length;
  const 要確認数=candidates.filter(c=>c.判定==='要棚確認').length;
  ui.alert('取り置き登録を作成しました',
    '候補'+candidates.length+'行 ／ ⚠️要棚確認 '+要確認数+'行(黄色・上に集めました)\n\n'
    +'要棚確認=旧帳簿では着いているはずなのに、数量も棚確認も未入力の行です。\n'
    +'棚を見て、現物があれば「現物取り置き数量」を入力、無ければ「棚確認」で\n'
    +'出荷済み/未着を選んでください(選ぶと黄色の対象から外れます)。\n'
    +(入力済? '\n入力済みの数量'+入力済+'行は引き継いで表示しています。':''),ui.ButtonSet.OK);
}

function 取り置き初期登録を確定(){ 直列_(取り置き初期登録を確定本体_); }
function 取り置き初期登録を確定本体_(){
  const ui=SpreadsheetApp.getUi(), inputs=取り置き_表を読む_(TORIOKI_CFG.初期,TORIOKI_CFG.初期HDR);
  const plan=取り置き_初期確定計画_(inputs,取り置き台帳_読む_(),new Date());
  if(plan.errors.length){ ui.alert('取り置き登録を中止しました',plan.errors.join('\n'),ui.ButtonSet.OK); return; }
  const selected=inputs.filter(r=>(Number(r.現物取り置き数量)||0)>0), qty=selected.reduce((s,r)=>s+(Number(r.現物取り置き数量)||0),0);
  const answer=ui.alert('取り置き登録を確定します','対象'+selected.length+'行 / 合計'+qty+'個です。確定しますか？',ui.ButtonSet.OK_CANCEL);
  if(answer!==ui.Button.OK) return;
  取り置き台帳_保存_(plan.rows);
  SpreadsheetApp.getActive().toast('取り置き登録 '+selected.length+'行 / '+qty+'個を確定しました','取り置き台帳',7);
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

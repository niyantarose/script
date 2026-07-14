// ===== ダニエル便の余り（引当されなかった分）の自動計算 =====
// 背景: ダニエルEMSタブは取込のたびに上書きされるため、便の余りが帳簿から消えて
//       現物とズレていく。ここでは取込のたびに「ダニエル入荷記録」へ便を積み上げ、
//       受注明細・消込台帳から見える消費を差し引いた「余り推定」を「ダニエル余り」に出す。
// 設計: 消費は“誰がボタンを押したか”に依存せず、データ（消込台帳の出荷済み・
//       受注明細のダニエルEMS番号付き行）から毎回計算し直す。他の人がGoQの中だけで
//       出荷しても、次のGoQ取込で消込台帳に載った時点で自動的に差し引かれる。

const DANIEL_AMARI_CFG = {
  記録: 'ダニエル入荷記録',        // 便の入荷の積み上げ台帳(EMS番号単位で洗い替え)
  余り: 'ダニエル余り',            // 余り推定レポート(毎回作り直し)
  反映: 'ダニエル余りYahoo反映'    // 手入力: A=商品コード / B=Yahooへ足した数(任意)
};

// ---- 純粋ロジック(Nodeテスト対象) ----
// src:
//   記録:   [{ems, code, qty}]                ダニエル入荷記録の行
//   大邱ems:[{code, st, qty, arrival, ems}]   発注共有EMSリストの行(大邱の箱由来の出荷を除くため)
//   出荷済: [{ban, code, sku, qty, 入荷日}]    消込台帳_出荷済み行_()の戻り
//   受注:   [{code, sku, qty, ems}]           受注明細の現存注文(EMS番号付き)
//   反映:   {基底コード: 数}                   Yahooへ反映済みと手入力した数(null可)
// 戻り: [{code, 供給, 出荷済, 引当済, Yahoo反映, 余り}] 供給のあるコードだけ・余り降順
function ダニエル余り集計_(src){
  const 供給={}; const 便EMS=new Set();
  (src.記録||[]).forEach(r=>{
    const c=normCode_(r.code); if(!c) return;
    供給[c]=(供給[c]||0)+(Number(r.qty)||0);
    const e=String(r.ems||'').replace(/\s+/g,'').toUpperCase(); if(e) 便EMS.add(e);
  });

  // 大邱の箱の到着日集合(コード別)。この日付に紐づく出荷は大邱由来なので引かない
  const 大邱日={};
  (src.大邱ems||[]).forEach(r=>{
    if(r.ems!==undefined && !実EMS番号_(r.ems)) return;
    const c=normCode_(r.code); if(!c) return;
    const d=ymd_(r.arrival); if(!d) return;
    (大邱日[c]=大邱日[c]||new Set()).add(d);
  });

  // 出荷済み: 大邱の箱に紐づかない分(入荷日なし・大邱到着日と不一致)をダニエル消費とみなす
  const 出荷={};
  出荷済み重複排除_(src.出荷済||[]).forEach(r=>{
    const base=受注基底コード_(r.sku, r.code); if(!base || !(base in 供給)) return;
    const qty=Number(r.qty)||0; if(qty<=0) return;
    const d=ymd_(r.入荷日);
    if(d && 大邱日[base] && 大邱日[base].has(d)) return; // 大邱の箱から出た分
    出荷[base]=(出荷[base]||0)+qty;
  });

  // 引当済み(未出荷): 受注明細のEMS番号がダニエル便のEMS番号になっている行
  const 引当={};
  (src.受注||[]).forEach(r=>{
    const base=受注基底コード_(r.sku, r.code); if(!base || !(base in 供給)) return;
    const e=String(r.ems||'').replace(/\s+/g,'').toUpperCase();
    if(!e || !便EMS.has(e)) return;
    引当[base]=(引当[base]||0)+(Number(r.qty)||0);
  });

  const rows=Object.keys(供給).sort().map(c=>{
    const y=Number((src.反映||{})[c])||0;
    const 余り=供給[c]-(出荷[c]||0)-(引当[c]||0)-y;
    return { code:c, 供給:供給[c], 出荷済:出荷[c]||0, 引当済:引当[c]||0, Yahoo反映:y, 余り:余り };
  }).filter(r=>r.供給>0);
  rows.sort((a,b)=> b.余り-a.余り || (a.code<b.code? -1 : a.code>b.code? 1 : 0));
  return rows;
}

// ---- ダニエル入荷記録(積み上げ台帳) ----
// 取込_ダニエルEMSのrows([BOXNo,発送日,商品名,商品コード,数量,対象BOX,EMS番号,状態])を
// EMS番号単位で洗い替えして積み上げる(同じ便を再取込しても二重計上しない)
function ダニエル入荷記録へ追記_(rows){
  const ss=SpreadsheetApp.getActive();
  let sh=ss.getSheetByName(DANIEL_AMARI_CFG.記録);
  const HDR=['取込日','発送日','EMS番号','商品コード','数量'];
  if(!sh){ sh=ss.insertSheet(DANIEL_AMARI_CFG.記録);
    sh.getRange(1,1,1,HDR.length).setValues([HDR]).setFontWeight('bold'); sh.setFrozenRows(1); }
  const newEms=new Set(rows.map(r=>String(r[6]||'').replace(/\s+/g,'').toUpperCase()).filter(e=>e));
  if(!newEms.size) return 0;

  const last=sh.getLastRow();
  const keep=[];
  if(last>=2){
    sh.getRange(2,1,last-1,HDR.length).getValues().forEach(r=>{
      const e=String(r[2]||'').replace(/\s+/g,'').toUpperCase();
      if(e && !newEms.has(e)) keep.push(r); // 今回来たEMS番号は入れ替え(再取込・訂正対応)
    });
  }
  const today=Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy-MM-dd');
  rows.forEach(r=>{
    const e=String(r[6]||'').replace(/\s+/g,'').toUpperCase(); if(!e) return;
    keep.push([today, String(r[1]||''), e, String(r[3]||''), Number(r[4])||0]);
  });
  if(last>=2) sh.getRange(2,1,last-1,HDR.length).clearContent();
  if(keep.length) sh.getRange(2,1,keep.length,HDR.length).setValues(keep);
  return keep.length;
}

// 初回用: いまダニエルEMSタブに載っている便を記録台帳へ取り込む(エディタから実行可)
function ダニエル入荷記録_現タブから初期化(){
  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName(DANIEL_CFG.シート);
  if(!sh){ Logger.log('ダニエルEMSタブがありません'); return; }
  const hr=ヘッダー行検索_(sh,'BOXNo'), start=hr+1, last=sh.getLastRow();
  if(last<start){ Logger.log('ダニエルEMSタブにデータがありません'); return; }
  const vals=sh.getRange(start,1,last-start+1,8).getValues().filter(r=>String(r[6]||'').trim());
  const n=ダニエル入荷記録へ追記_(vals);
  Logger.log('ダニエル入荷記録を初期化: 台帳'+n+'行');
  ss.toast('ダニエル入荷記録へ '+vals.length+'行を取り込みました');
  ダニエル余りを計算();
}

// ---- 余り推定レポート ----
// UIを使わない(メニュー・エディタ・取込後フックのどこからでも実行できる)
function ダニエル余りを計算(){
  const ss=SpreadsheetApp.getActive();

  // 記録(供給)
  const 記録=[];
  { const sh=ss.getSheetByName(DANIEL_AMARI_CFG.記録);
    if(sh && sh.getLastRow()>=2){
      sh.getRange(2,1,sh.getLastRow()-1,5).getValues().forEach(r=>{
        if(String(r[2]||'').trim()) 記録.push({ems:r[2], code:r[3], qty:r[4]});
      });
    } }
  if(!記録.length){ ss.toast('ダニエル入荷記録が空です（先に📦取込するか、初期化を実行）','ダニエル余り',6); return; }

  // 大邱EMSリスト(読み取りのみ。読めなくても続行=大邱由来の除外が効かないだけ)
  let 大邱ems=[];
  try{
    const sh=発注共有を開く_().getSheetByName(P_KAKUTEI_CFG.シート);
    if(sh){
      const hr=P_KAKUTEI_CFG.ヘッダー行, last=sh.getLastRow();
      if(last>hr){
        const head=sh.getRange(hr,1,1,sh.getLastColumn()).getValues()[0].map(v=>String(v||'').trim());
        const f=(...names)=>{ for(const n of names){ const i=head.indexOf(n); if(i>=0) return i; } return -1; };
        const cSt=f('ステータス列'), cC=f('商品コード'), cQ=f('数量','個数'), cE=f('EMS番号');
        let cA=f('EMS到着日','到着日','到着'); if(cA<0) cA=4;
        if(cC>=0){
          大邱ems=sh.getRange(hr+1,1,last-hr,sh.getLastColumn()).getValues().map(r=>({
            code:r[cC], st:cSt>=0? r[cSt]:'到着済', qty:cQ>=0? r[cQ]:1, arrival:r[cA], ems:cE>=0? r[cE]:''}));
        }
      }
    }
  }catch(e){}

  // 消込台帳の出荷済み
  let 出荷済=[]; try{ 出荷済=消込台帳_出荷済み行_(); }catch(e){}

  // 受注明細(現存注文のEMS番号)
  const 受注=[];
  { const recv=ss.getSheetByName(HIKIATE_CFG.受注);
    if(recv){
      const M=列マップ_(recv); const R=recv.getDataRange().getValues();
      for(let i=M.hr;i<R.length;i++){
        const row=R[i]; if(!String(row[M.番号]||'').trim()) continue;
        受注.push({code:row[M.コード], sku:M.SKU>=0? row[M.SKU]:'', qty:row[M.個数],
          ems:M.EMS>=0? row[M.EMS]:''});
      }
    } }

  // Yahoo反映済み(手入力・任意)
  let 反映=null;
  { const sh=ss.getSheetByName(DANIEL_AMARI_CFG.反映);
    if(sh && sh.getLastRow()>=1){
      反映={};
      sh.getRange(1,1,sh.getLastRow(),2).getValues().forEach(r=>{
        const c=normCode_(r[0]); const n=Number(r[1])||0;
        if(c && n>0) 反映[c]=(反映[c]||0)+n;
      });
    } }

  const rows=ダニエル余り集計_({記録:記録, 大邱ems:大邱ems, 出荷済:出荷済, 受注:受注, 反映:反映});

  // 描画
  let rep=ss.getSheetByName(DANIEL_AMARI_CFG.余り); if(!rep) rep=ss.insertSheet(DANIEL_AMARI_CFG.余り);
  rep.clear();
  rep.getRange(1,1).setValue('ダニエル余り推定(読み取り専用): '
    +Utilities.formatDate(new Date(),'Asia/Tokyo','yyyy/MM/dd HH:mm')
    +' ／ 便の入荷記録 '+記録.length+'行');
  const HDR=['商品コード','便の入荷累計','出荷済み(推定)','引当済み(未出荷)','Yahoo反映済み(手入力)','余り推定'];
  rep.getRange(2,1,1,HDR.length).setValues([HDR])
    .setFontWeight('bold').setFontColor('#ffffff').setBackground('#4472c4').setFontSize(HIKIATE_CFG.字);
  rep.setFrozenRows(2);
  if(rows.length){
    const out=rows.map(r=>[r.code, r.供給, r.出荷済, r.引当済, r.Yahoo反映, r.余り]);
    rep.getRange(3,1,out.length,HDR.length).setValues(out).setFontSize(HIKIATE_CFG.字);
    const bg=rows.map(r=>{
      const c= r.余り>0? HIKIATE_CFG.色_黄 : r.余り<0? HIKIATE_CFG.色_赤 : null;
      return new Array(HDR.length).fill(c);
    });
    rep.getRange(3,1,out.length,HDR.length).setBackgrounds(bg);
  } else {
    rep.getRange(3,1).setValue('(対象なし)');
  }
  const 凡例行=3+Math.max(rows.length,1)+1;
  const 凡例=[
    '黄=余りあり: 便で届いた数のうち、出荷にも引当にもYahoo反映にも使われていない分(棚にあるはず)',
    '赤=マイナス: 便の数より消費が多い。大邱と同じコードを両ルートで仕入れている場合の推定誤差もある',
    '余りをYahooに足したら「'+DANIEL_AMARI_CFG.反映+'」シートに A=商品コード / B=足した数 を書くと差し引かれる',
    '※出荷済みは消込台帳ベースの推定(入荷日が大邱の箱と一致する出荷は大邱由来として除外)。何も書き換えない',
    '※ダニエル引当直後〜出荷までの間は余りが多めに見えることがある(出荷されて消込台帳に載ると自動で収束する)'
  ];
  rep.getRange(凡例行,1,凡例.length,1).setValues(凡例.map(s=>[s])).setFontSize(HIKIATE_CFG.字);
  const 余りあり=rows.filter(r=>r.余り>0).length;
  ss.toast('ダニエル余り: '+rows.length+'コード / 余りあり '+余りあり+'件','🧮ダニエル余り',6);
}

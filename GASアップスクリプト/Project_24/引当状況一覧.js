// ===== 引当状況一覧(読み取り専用)とGoQ差分 =====
// 全注文を注文単位で一覧化し、現在のGoQステータスと引当結果からの推奨を比較する。
// GoQへは一切書き戻さない(表示だけ)。仕様§12。

const 引当状況_CFG=Object.freeze({
  シート:'引当状況一覧',
  HDR:['受注番号','氏名','現在のGoQステータス','推奨GoQステータス','GoQ差分','差異理由','分類',
    'お届け希望日','支払い','商品コード','SKU','商品名','注文数量','現物確認済み','到着済引当','先行引当','不足','状態の理由','取り置きメモ']
});

const 引当状況_分類名={wait:'引当待ち',part:'部分在庫',hold:'希望日待ち',keep:'出荷GO未入金',ship:'出荷可能'};

// 分類・支払い・段階から推奨GoQ表示を返す(作業案内。APIへは書かない)
function 引当状況_推奨GoQ_(arr,paid,cod){
  const b=注文区分判定_(arr,paid,cod), s=注文充足集計_(arr);
  if(b==='wait') return '取り寄せ中';
  if(b==='part') return '一部確保('+s.確保総数+'/'+s.注文数量+')';
  const 先行含み=s.先行引当>0;
  if(b==='hold') return 先行含み?'希望日待ち(先行含み)':'希望日待ち';
  if(b==='keep') return 先行含み?'入金待ち(先行含み)':'入金待ち';
  if(cod&&!paid) return 先行含み?'出荷可能(代引き・先行待ち)':'出荷可能(代引き)';
  return 先行含み?'出荷可能(先行待ち)':'出荷可能';
}

// 差異理由: 段階・支払い・希望日から作業者が読める組合せ文言
function 引当状況_差異理由_(arr,paid,cod){
  const s=注文充足集計_(arr), out=[];
  if(s.確保総数<=0){ out.push('確保なし'); return out; }
  if(s.不足>0) out.push('不足'+s.不足+'個');
  if(s.先行引当>0&&s.現物確認済み+s.到着済引当<=0) out.push('未着先行のみ');
  else if(s.先行引当>0) out.push('先行'+s.先行引当+'個含み');
  if(s.不足===0&&s.先行引当===0) out.push('到着済み全数');
  if(s.現物確認済み>0) out.push('現物固定あり');
  if((arr||[]).some(l=>l&&!l.キャンセル&&希望日未来_(l.届))) out.push('希望日待ち');
  if(!paid) out.push(cod?'代引き':'未入金');
  return out;
}

// 一致/差異あり。GoQの現在値は受注明細の受注ステータス文字列から読む
function 引当状況_GoQ差分_(currentStatus, recommended, arr, paid, cod){
  const cur=String(currentStatus||'');
  const cur取寄せ=/取寄せ|取り寄せ|予約/.test(cur);
  const rec取寄せ=/^取り寄せ中|^一部確保/.test(String(recommended||''));
  const match=rec取寄せ?cur取寄せ:!cur取寄せ;
  return {判定:match?'一致':'差異あり',理由:match?[]:引当状況_差異理由_(arr,paid,cod)};
}

// 全注文の一覧行(1商品=1行・注文単位で連続)。書き込みAPIは呼ばない純粋関数
function 引当状況_一覧行_(linesByBan, paidByBan, codByBan, statusByBan){
  const rows=[];
  Object.keys(linesByBan||{}).forEach(ban=>{
    const arr=linesByBan[ban]||[], paid=!!(paidByBan&&paidByBan[ban]), cod=!!(codByBan&&codByBan[ban]);
    const cur=String(statusByBan&&statusByBan[ban]||'');
    const 分類=引当状況_分類名[注文区分判定_(arr,paid,cod)]||'';
    const 推奨=引当状況_推奨GoQ_(arr,paid,cod);
    const 差分=引当状況_GoQ差分_(cur,推奨,arr,paid,cod);
    arr.forEach(l=>{
      rows.push({受注番号:ban,氏名:String(l.氏名||''),現在のGoQステータス:cur,推奨GoQステータス:推奨,
        GoQ差分:差分.判定,差異理由:差分.理由.join('・'),分類,
        お届け希望日:String(l.届||''),支払い:paid?'済':(cod?'代引':'未'),
        商品コード:String(l.code||''),SKU:String(l.sku||''),商品名:String(l.商品名||''),
        注文数量:Number(l.qty)||0,現物確認済み:Number(l.現物確認済み数量)||0,
        到着済引当:Number(l.到着済引当数量)||0,先行引当:Number(l.先行引当数量)||0,
        不足:Number(l.未引当数量)||0,状態の理由:行状態理由_(l),取り置きメモ:String(l.メモ||'')});
    });
  });
  return rows;
}

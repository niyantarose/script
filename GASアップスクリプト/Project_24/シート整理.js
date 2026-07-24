// ===== バックアップ・監査シートの整理 =====
// 全件再計算の反映・現物確認移行のたびに退避シート(取り置き台帳_全件再計算前_◯◯等)が増えるため、
// 各種類の最新N世代だけ残して古いものを削除する。Driveのバックアップファイルは触らない(深い保険として残す)。

// Driveに反映ごとの丸ごとバックアップが残るため、シート側の退避は直近1世代で十分(2026-07-22ユーザー要望)
const シート整理_CFG=Object.freeze({保持世代:1});

// 削除対象の選定(純粋関数)。名前は「<種類>_(全件再計算前|移行前)_yyyyMMdd_HHmmss(_n)?」形式。
// シート名99文字制限で時刻が切れているもの(例: _12335)も拾う。protectedNamesは常に残す。
function シート整理_削除対象_(names, keep, protectedNames){
  const k=keep==null?シート整理_CFG.保持世代:keep;
  const guard=new Set((protectedNames||[]).map(String));
  const groups={};
  (names||[]).forEach(name=>{
    const text=String(name||'');
    if(guard.has(text)) return;
    const m=text.match(/^(.+?_(?:全件再計算前|移行前))_(\d{8}_\d{3,6})(?:_\d+)?$/);
    if(!m) return;
    (groups[m[1]]=groups[m[1]]||[]).push({name:text,ts:m[2]});
  });
  const out=[];
  Object.keys(groups).forEach(group=>{
    groups[group].sort((a,b)=>b.ts.localeCompare(a.ts));
    groups[group].slice(k).forEach(entry=>out.push(entry.name));
  });
  return out.sort();
}

function バックアップシートを整理(){ 直列_(バックアップシートを整理本体_); }
function バックアップシートを整理本体_(){
  const ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi();
  const names=ss.getSheets().map(sh=>sh.getName());
  // 現物確認移行の比較基準は移行運用が続く限り必要なので必ず残す
  const targets=シート整理_削除対象_(names,シート整理_CFG.保持世代,[現物移行_CFG.基準シート]);
  if(!targets.length){
    ui.alert('整理対象はありません','監査シートは各種類とも最新'+シート整理_CFG.保持世代+'世代以内です。',ui.ButtonSet.OK);
    return;
  }
  const ans=ui.alert('🧹 バックアップシートを整理',
    '古い監査シート'+targets.length+'枚を削除します。\n'
    +'(各種類の最新'+シート整理_CFG.保持世代+'世代、移行の基準シート、Driveのバックアップファイルは残ります)\n\n'
    +targets.slice(0,15).join('\n')+(targets.length>15?'\n…ほか'+(targets.length-15)+'枚':'')
    +'\n\n削除しますか？',ui.ButtonSet.OK_CANCEL);
  if(ans!==ui.Button.OK) return;
  let deleted=0;
  targets.forEach(name=>{ const sh=ss.getSheetByName(name); if(sh){ ss.deleteSheet(sh); deleted++; } });
  ss.toast(deleted+'枚を削除しました(Driveのバックアップは残っています)','🧹整理',6);
}

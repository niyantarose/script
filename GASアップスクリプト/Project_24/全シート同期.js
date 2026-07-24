// ===== 数量変更後の全派生シート同期 =====
// 数字を変更した最上位処理からだけ呼ぶ。台帳保存などの低レベル関数からは呼ばない。
// ②引当自身も台帳を保存するため、低レベル保存から呼ぶと再帰・二重実行になる。

const HIKIATE_SYNC_STATE_KEY='引当_全シート同期状態';

function 引当_同期状態_保存_(state){
  PropertiesService.getDocumentProperties().setProperty(HIKIATE_SYNC_STATE_KEY,JSON.stringify(state||{}));
}

function 引当_同期状態_読む_(){
  try{ return JSON.parse(PropertiesService.getDocumentProperties().getProperty(HIKIATE_SYNC_STATE_KEY)||'null'); }
  catch(e){ return null; }
}

// ②の主要出力後に、有効台帳から日常の管理画面を同じ時点へ揃える。
// ここは必須更新として扱い、例外を握り潰さない。
function 引当_同期管理シート_(){
  if(typeof 取り置き初期登録を作成本体_==='function') 取り置き初期登録を作成本体_({silent:true});
  if(typeof キャンセル戻し確認を更新本体_==='function') キャンセル戻し確認を更新本体_({silent:true});
  if(typeof Yahoo戻し候補を更新_==='function') Yahoo戻し候補を更新_();
}

// options:
//   理由: 完了/失敗表示に出す操作名
//   EMS更新: EMSリストの状態を変えた操作だけtrue
//   完了表示: falseなら中央同期自身のトースト/失敗ダイアログを出さない
function 引当_数値変更後全同期_(options){
  const opt=options||{}, reason=String(opt.理由||'数量変更'), started=Date.now();
  const props=PropertiesService.getDocumentProperties();
  props.deleteProperty('引当_整合状態');
  引当_同期状態_保存_({status:'running',reason,startedAt:started});
  try{
    if(opt.EMS更新) EMS在庫を更新_本体_();
    const result=引当実行_本体_({preview:false,silentSummary:true,skipManagementRefresh:true});
    if(!result || result.success!==true) throw new Error('②引当が完了結果を返しませんでした');

    引当_同期管理シート_();

    const integrity=JSON.parse(props.getProperty('引当_整合状態')||'null');
    const ledgerAt=Number(props.getProperty('取り置き台帳_最終更新')||0);
    if(!integrity || Number(integrity.ts||0)<ledgerAt) throw new Error('台帳更新後の整合時刻を確認できません');

    const finished=Date.now();
    引当_同期状態_保存_({status:'synced',reason,startedAt:started,finishedAt:finished});
    if(opt.完了表示!==false){
      try{ SpreadsheetApp.getActive().toast(reason+'後の全シート同期が完了しました（'+Math.round((finished-started)/1000)+'秒）','✅全シート同期',8); }catch(e){}
    }
    return {success:true,result,処理ms:finished-started};
  }catch(error){
    props.deleteProperty('引当_整合状態');
    const message=String(error&&error.message||error);
    引当_同期状態_保存_({status:'failed',reason,startedAt:started,failedAt:Date.now(),error:message});
    if(opt.完了表示!==false){
      try{
        SpreadsheetApp.getUi().alert('全シート同期に失敗しました',
          reason+'の元データは保存済みですが、派生シートは未同期です。\n'+message+'\n\nメニューから②を実行し直してください。',
          SpreadsheetApp.getUi().ButtonSet.OK);
      }catch(e){}
    }
    return {success:false,error};
  }
}

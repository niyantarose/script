/**
 * フォルダ内の新しいファイルをチェックして、
 * 見つかったら対応する処理関数を実行する
 */
function フォルダをチェック() {
  const props = PropertiesService.getScriptProperties();

  // ★ここを自分のフォルダID & 実行したい関数名に合わせて書き換え★
  const configList = [
    {
      key: 'QOO10_IMPORT',
      folderId: 'ここにQoo10インポートファイルのフォルダID',
      handlerName: 'Qoo10最新ファイル取込'       // 既にある関数名
    },
    {
      key: 'YAHOO_SRC',
      folderId: 'ここにyahoo変換元ファイルのフォルダID',
      handlerName: 'yahoo元ファイルを変換'       // 既にある関数名
    },
    {
      key: 'YAHOO_ALL',
      folderId: 'ここにyahoo全在庫のフォルダID',
      handlerName: 'Yahoo在庫上書き'             // 既にある関数名
    },
    {
      key: 'POSTAGE',
      folderId: 'ここに送料コード用フォルダID',
      handlerName: '送料コード反映'               // 既にある関数名
    }
  ];

  configList.forEach(cfg => {
    const folder = DriveApp.getFolderById(cfg.folderId);
    const files = folder.getFiles();

    // フォルダ内で「一番新しいファイル」を探す
    let latestFile = null;
    while (files.hasNext()) {
      const f = files.next();
      if (!latestFile || f.getLastUpdated() > latestFile.getLastUpdated()) {
        latestFile = f;
      }
    }
    if (!latestFile) return;  // フォルダが空なら何もしない

    const propKey = 'LAST_UPDATED_' + cfg.key;
    const lastTime = props.getProperty(propKey);
    const latestTime = String(latestFile.getLastUpdated().getTime());

    // 前回処理した時間と違えば「新しい or 更新された」とみなして実行
    if (lastTime !== latestTime) {
      // ログだけ出しておく
      console.log('新しいファイル検出: ' + cfg.key + ' → ' + latestFile.getName());

      // 名前から関数を呼び出す
      if (typeof this[cfg.handlerName] === 'function') {
        this[cfg.handlerName]();   // 既存の処理関数を引数なしで実行
      } else {
        console.warn('関数が見つかりません: ' + cfg.handlerName);
      }

      // このファイルを「処理済み」として記録
      props.setProperty(propKey, latestTime);
    }
  });
}

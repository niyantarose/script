/**
 * 確定発行ボタン設置.gs（韓国版）
 *
 * アクティブなシートに「確定発行」ボタン（画像+スクリプト割り当て）を設置する。
 * メニュー「🔘 このシートに確定発行ボタンを設置」から1回実行するだけ。
 * ボタンのスクリプト割り当ては同一スプレッドシートに紐づく全プロジェクト
 * （P05/P06/P07）から関数名で解決されるため、この1関数で全シート対応できる。
 *
 * やること（台湾CN側と同じ）:
 * 1) 2〜4行目にデータが詰まっていれば3行挿入してボタン置き場を空ける
 *    （ヘッダーは1行目のまま。ライブラリ・Webアプリともヘッダー=1行目前提のため動かさない）
 * 2) 1〜4行目を固定表示（スクロールしてもボタンが見える）
 * 3) ボタン画像を挿入し、シートに応じた確定発行関数を割り当てる
 *    （再実行時は古いボタンを消して作り直し＝安全）
 */

var 韓国確定発行ボタン_対象シート = {
  '韓国マンガ': '韓国マンガ_確定発行',
  '韓国書籍': '韓国書籍_確定発行',
  '韓国音楽映像': '韓国音楽映像_確定発行',
  '韓国雑誌': '韓国雑誌_確定発行',
  '韓国グッズ': '韓国グッズ_確定発行',
};

var 韓国確定発行ボタン_ALTタイトル = '確定発行ボタン';

var 韓国確定発行ボタン_PNG_BASE64 = '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';

function 韓国_確定発行ボタンを設置() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  const fn = 韓国確定発行ボタン_対象シート[sh.getName()];
  if (!fn) {
    ui.alert(
      'このシートには確定発行ボタンを設置できません。\n' +
      '対象シート: ' + Object.keys(韓国確定発行ボタン_対象シート).join(' / ') + '\n\n' +
      '設置したいシートを開いた状態で実行してください。'
    );
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    // 1) ボタン置き場の確保: 2〜4行目に何か入っていれば3行挿入してデータを下げる
    let 挿入した = false;
    if (sh.getLastRow() >= 2) {
      const 行数 = Math.min(3, sh.getLastRow() - 1);
      const zone = sh.getRange(2, 1, 行数, Math.max(sh.getLastColumn(), 1)).getDisplayValues();
      const 詰まっている = zone.some(row => row.some(v => String(v || '').trim() !== ''));
      if (詰まっている) {
        sh.insertRowsBefore(2, 3);
        挿入した = true;
        // 挿入行が引き継いだ書式・チェックボックス等をまっさらに戻す
        const 帯 = sh.getRange(2, 1, 3, sh.getMaxColumns());
        帯.clearContent();
        帯.clearDataValidations();
        帯.clearFormat();
      }
    }

    // 2) ヘッダー+ボタン帯を固定表示
    if (sh.getFrozenRows() < 4) sh.setFrozenRows(4);

    // 3) 既存の同名ボタンを消して作り直し（重複設置防止）
    sh.getImages().forEach(img => {
      try {
        if (img.getAltTextTitle() === 韓国確定発行ボタン_ALTタイトル) img.remove();
      } catch (e) {}
    });

    const blob = Utilities.newBlob(
      Utilities.base64Decode(韓国確定発行ボタン_PNG_BASE64), 'image/png', 'kakutei_button.png'
    );
    const img = sh.insertImage(blob, 2, 2); // B2 あたりに設置
    img.setAltTextTitle(韓国確定発行ボタン_ALTタイトル);
    img.assignScript(fn);
    img.setWidth(132);
    img.setHeight(36);

    SpreadsheetApp.flush();
    ui.alert(
      '✅ 「確定発行」ボタンを設置しました（' + sh.getName() + '）\n\n' +
      (挿入した ? '・データを3行下げてボタン置き場を作りました\n' : '・2〜4行目が空いていたのでそのまま使いました\n') +
      '・1〜4行目を固定表示にしました\n' +
      '・ボタンをクリックすると「① 確定発行」（コード重複プリフライトつき）が実行されます'
    );
  } finally {
    lock.releaseLock();
  }
}
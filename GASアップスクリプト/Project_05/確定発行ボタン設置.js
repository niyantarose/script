/**
 * 確定発行ボタン設置.gs（韓国版）
 *
 * アクティブなシートに操作ボタン（画像+スクリプト割り当て）を設置する。
 * メニュー「🔘 このシートに確定発行ボタンを設置」から1回実行するだけ。
 * ボタンのスクリプト割り当ては同一スプレッドシートに紐づく全プロジェクト
 * （P05/P06/P07）から関数名で解決されるため、この1関数で全シート対応できる。
 *
 * 設置されるボタン（そのシートに機能があるものだけ）:
 *   🔵 確定発行     … コード重複プリフライトつき確定発行（全5シート）
 *   🟠 重複チェック … 韓国グッズのみ（専用チェック）
 *
 * やること（台湾CN側と同じ）:
 * 1) 2〜4行目にデータが詰まっていれば3行挿入してボタン置き場を空ける
 * 2) 1〜4行目を固定表示
 * 3) ボタン画像を挿入してスクリプトを割り当て（再実行時は作り直し＝安全）
 */

var 韓国ボタンPNG_確定発行 = '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';
var 韓国ボタンPNG_重複チェック = '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';

var 韓国ボタン定義_ = [
  {
    alt: '確定発行ボタン', 列: 2, png: () => 韓国ボタンPNG_確定発行,
    対象: {
      '韓国マンガ': '韓国マンガ_確定発行',
      '韓国書籍': '韓国書籍_確定発行',
      '韓国音楽映像': '韓国音楽映像_確定発行',
      '韓国雑誌': '韓国雑誌_確定発行',
      '韓国グッズ': '韓国グッズ_確定発行',
    },
  },
  {
    alt: '重複チェックボタン', 列: 4, png: () => 韓国ボタンPNG_重複チェック,
    対象: {
      '韓国グッズ': '韓国グッズ_重複チェック',
    },
  },
];

function 韓国_確定発行ボタンを設置() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  const 設置予定 = 韓国ボタン定義_.filter(def => def.対象[sh.getName()]);
  if (!設置予定.length) {
    ui.alert(
      'このシートには操作ボタンを設置できません。\n' +
      '対象シート: ' + Object.keys(韓国ボタン定義_[0].対象).join(' / ') + '\n\n' +
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
        const 帯 = sh.getRange(2, 1, 3, sh.getMaxColumns());
        帯.clearContent();
        帯.clearDataValidations();
        帯.clearFormat();
      }
    }

    // 2) ヘッダー+ボタン帯を固定表示
    if (sh.getFrozenRows() < 4) sh.setFrozenRows(4);

    // 3) 既存の設置対象ボタンを消して作り直し（重複設置防止）
    const 全ALT = {};
    韓国ボタン定義_.forEach(def => { 全ALT[def.alt] = true; });
    sh.getImages().forEach(img => {
      try {
        if (全ALT[img.getAltTextTitle()]) img.remove();
      } catch (e) {}
    });

    const 設置名 = [];
    設置予定.forEach(def => {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(def.png()), 'image/png', def.alt + '.png'
      );
      const img = sh.insertImage(blob, def.列, 2);
      img.setAltTextTitle(def.alt);
      img.assignScript(def.対象[sh.getName()]);
      img.setWidth(132);
      img.setHeight(36);
      設置名.push(def.alt.replace('ボタン', ''));
    });

    SpreadsheetApp.flush();
    ui.alert(
      '✅ 操作ボタンを設置しました（' + sh.getName() + '）\n\n' +
      '・設置: ' + 設置名.join(' / ') + '\n' +
      (挿入した ? '・データを3行下げてボタン置き場を作りました\n' : '') +
      '・1〜4行目を固定表示にしました\n' +
      '・ボタンはドラッグで好きな位置に移動できます'
    );
  } finally {
    lock.releaseLock();
  }
}
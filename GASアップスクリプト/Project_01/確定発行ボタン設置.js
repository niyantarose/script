/**
 * 確定発行ボタン設置.gs
 *
 * アクティブなシートに操作ボタン一式（画像+スクリプト割り当て）を設置する。
 * メニュー「🔘 このシートに操作ボタンを設置」から1回実行するだけ。
 *
 * 設置されるボタン（そのシートに機能があるものだけ）:
 *   🔵 確定発行       … プリフライト検査つき確定発行
 *   🟠 重複チェック   … 書籍系=重複作品の検出ドライラン(色付け・非破壊) / グッズ=専用チェック
 *   🟢 画像名→SKU    … 画像フォルダをSKUにリネーム
 *   🟣 Yahoo送信      … チェック行をYahoo商品テンプレへ送信
 *
 * やること:
 * 1) 2〜4行目にデータが詰まっていれば3行挿入してボタン置き場を空ける
 *    （ヘッダーは1行目のまま。全スクリプトがヘッダー=1行目前提のため動かさない）
 * 2) 1〜4行目を固定表示
 * 3) ボタン画像を挿入してスクリプトを割り当て（再実行時は作り直し＝安全）
 */

var 台湾ボタンPNG_確定発行 = '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';
var 台湾ボタンPNG_重複チェック = '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';
var 台湾ボタンPNG_画像名変換 = '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';
var 台湾ボタンPNG_Yahoo送信 = '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';

var 台湾ボタン定義_ = [
  {
    alt: '確定発行ボタン', 列: 2, png: () => 台湾ボタンPNG_確定発行,
    対象: {
      '台湾まんが': '台湾まんが_確定発行',
      '台湾書籍その他': '台湾書籍その他_確定発行',
      '台湾グッズ': '台湾グッズ_確定発行',
    },
  },
  {
    alt: '重複チェックボタン', 列: 4, png: () => 台湾ボタンPNG_重複チェック,
    対象: {
      '台湾まんが': '台湾書籍系_重複作品_検出ドライラン',
      '台湾書籍その他': '台湾書籍系_重複作品_検出ドライラン',
      '台湾グッズ': '台湾グッズ_重複チェック',
    },
  },
  {
    alt: '画像名変換ボタン', 列: 6, png: () => 台湾ボタンPNG_画像名変換,
    対象: {
      '台湾まんが': '全シート_フォルダをSKUにリネーム',
      '台湾書籍その他': '全シート_フォルダをSKUにリネーム',
      '台湾グッズ': '全シート_フォルダをSKUにリネーム',
    },
  },
  {
    alt: 'Yahoo送信ボタン', 列: 8, png: () => 台湾ボタンPNG_Yahoo送信,
    対象: {
      '台湾まんが': '台湾_Yahooテンプレへ送信',
      '台湾書籍その他': '台湾_Yahooテンプレへ送信',
      '台湾グッズ': '台湾_Yahooテンプレへ送信',
    },
  },
];

function 台湾CN_確定発行ボタンを設置() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSheet();
  const 設置予定 = 台湾ボタン定義_.filter(def => def.対象[sh.getName()]);
  if (!設置予定.length) {
    ui.alert(
      'このシートには操作ボタンを設置できません。\n' +
      '対象シート: 台湾まんが / 台湾書籍その他 / 台湾グッズ\n\n' +
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
    台湾ボタン定義_.forEach(def => { 全ALT[def.alt] = true; });
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
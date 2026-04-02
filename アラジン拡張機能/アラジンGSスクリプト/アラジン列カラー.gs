/**
 * アラジン列カラー設定
 * トリガー列（入力するとアラジンAPIが動く列）にオレンジ色のヘッダーを設定する
 * メニュー → 🔴 アラジントリガー列に色付け から実行
 */

// アラジンのブランドカラー（オレンジ系）
const ALADIN_TRIGGER_COLOR  = '#e8463a'; // 赤オレンジ（トリガー列）
const ALADIN_TRIGGER_TEXT   = '#ffffff'; // 白文字
const ALADIN_OUTPUT_COLOR   = '#fde8e7'; // 薄いピンク（APIで自動入力される列）
const ALADIN_OUTPUT_TEXT    = '#c0392b'; // 赤文字

// シートごとのトリガー列・出力列定義
const ALADIN_COLOR_MAP = {
  '韓国書籍': {
    trigger: ['ISBN', 'アラジンURL'],
    output:  ['商品名(原題)', '著者', '出版社', '発売日', '原価', 'メイン画像URL', '追加画像URL', '商品説明', 'カテゴリ', 'アラジン商品コード'],
  },
  '韓国マンガ': {
    trigger: ['ISBN', 'アラジンURL'],
    output:  ['商品名(原題)', '著者', '出版社', '発売日', '原価', 'メイン画像URL', '追加画像URL', '商品説明', 'カテゴリ', 'アラジン商品コード'],
  },
  '韓国音楽映像': {
    trigger: ['アラジンURL'],
    output:  ['商品名(原題)', 'アーティスト名', '事務所レーベル', '発売日', '原価', 'メイン画像URL', '追加画像URL', '商品説明', 'カテゴリ', 'アラジン商品コード', 'JANコード'],
  },
  '韓国グッズ': {
    trigger: ['購入URL'],
    output:  ['商品名（原題）', '発売日', '原価', 'メイン画像URL', '追加画像URL', '商品説明', 'アラジン商品コード'],
  },
};

function アラジントリガー列に色付け() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const 対象シート = Object.keys(ALADIN_COLOR_MAP).filter(shName => ss.getSheetByName(shName));

  if (!対象シート.length) {
    ui.alert('アラジン列カラー設定', 'このファイルにアラジン対象シートが見つかりませんでした。', ui.ButtonSet.OK);
    return;
  }

  const 結果 = [];

  for (const shName of 対象シート) {
    const colorDef = ALADIN_COLOR_MAP[shName];
    const sh = ss.getSheetByName(shName);
    const lastCol = sh.getLastColumn();
    if (lastCol < 1) {
      結果.push(`⚠ ${shName}：列がありません`);
      continue;
    }

    const ヘッダー = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    let トリガー数 = 0;
    let 出力数 = 0;

    ヘッダー.forEach((h, i) => {
      const name = String(h).trim();
      const cell = sh.getRange(1, i + 1);

      if (colorDef.trigger.includes(name)) {
        cell.setBackground(ALADIN_TRIGGER_COLOR)
          .setFontColor(ALADIN_TRIGGER_TEXT)
          .setFontWeight('bold');
        cell.setNote('🔴 アラジンAPIトリガー列\nこの列にURLまたはISBNを入力すると自動取得が始まります');
        トリガー数++;
      } else if (colorDef.output.includes(name)) {
        cell.setBackground(ALADIN_OUTPUT_COLOR)
          .setFontColor(ALADIN_OUTPUT_TEXT)
          .setFontWeight('bold');
        cell.setNote('🤖 アラジンAPI自動入力列\nトリガー列入力後に自動で値がセットされます');
        出力数++;
      }
    });

    結果.push(`✅ ${shName}：トリガー${トリガー数}列・出力${出力数}列に色付け完了`);
  }

  ui.alert('アラジン列カラー設定完了', 結果.join('\n'), ui.ButtonSet.OK);
}



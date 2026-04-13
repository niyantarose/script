/**************************************************
 * Qoo10シート用
 * X列：商品説明（Magazine Info を整形してリボン説明の前に表示）
 * AC列：検索ワード（最大10個 / 1ワード30文字以内）
 *
 * 対象シート  : 「Qoo10」
 * データ開始行: 5行目（4行目までヘッダー）
 * 対象列      : E列=タイトル, X列=商品説明, AC列=検索ワード
 **************************************************/
function 雑誌用_X列説明整形_検索ワード() {
  var sheetName = "Qoo10";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Browser.msgBox("シート '" + sheetName + "' が見つかりません。");
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 5) {
    Browser.msgBox("データ行がありません。");
    return;
  }

  var rowCount = lastRow - 4; // 5行目から
  var rangeE  = sheet.getRange(5, 5,  rowCount); // タイトル
  var rangeX  = sheet.getRange(5, 24, rowCount); // 商品説明
  var rangeAC = sheet.getRange(5, 29, rowCount); // 検索ワード

  var valuesE = rangeE.getValues();
  var valuesX = rangeX.getValues();

  var newX  = [];
  var newAC = [];

  for (var i = 0; i < rowCount; i++) {
    var titleText    = valuesE[i][0] ? String(valuesE[i][0]) : "";
    var originalDesc = valuesX[i][0] ? String(valuesX[i][0]) : "";

    // ▼ 商品説明(X列)整形 ＋ 人名抽出用テキスト
    var resultHTMLAndText = 整形_説明HTML_と_抽出用テキスト取得(originalDesc);
    var finalHtml   = resultHTMLAndText.html;
    var textForName = resultHTMLAndText.text; // Magazine Info などのプレーンテキスト

    // ▼ 検索ワード抽出（タイトル＋Magazine Info）
    var keywords = [];

    // 1) タイトルから
    keywords = keywords.concat(抽出_人名_fromTitle(titleText));

    // 2) Magazine Info 等から
    keywords = keywords.concat(抽出_人名_fromMagazineText(textForName));

    // ▼ クリーニング＆重複排除＆制限（最大10個 / 1ワード30文字）
    var cleaned = [];
    for (var k = 0; k < keywords.length; k++) {
      var w = クリーン_人名候補(keywords[k]);
      if (!w) continue;

      // 30文字制限
      if (w.length > 30) {
        w = w.substring(0, 30);
      }

      // 重複排除（完全一致）
      if (cleaned.indexOf(w) === -1) {
        cleaned.push(w);
      }

      // 10個まで
      if (cleaned.length >= 10) break;
    }

    var acValue = cleaned.join("$$");

    newX.push([finalHtml]);
    newAC.push([acValue]);
  }

  rangeX.setValues(newX);
  rangeAC.setValues(newAC);

  Browser.msgBox("X列の商品説明整形とAC列の検索ワード作成が完了しました。");
}

/**************************************************
 * 商品説明 HTML 整形
 *  - 「■ Magazine Info」以降を整形して先頭に
 *  - それ以前（リボン説明など）はそのまま後ろに付ける
 * 戻り値: { html: 整形後HTML, text: 人名抽出用プレーンテキスト }
 **************************************************/
function 整形_説明HTML_と_抽出用テキスト取得(originalDesc) {
  if (!originalDesc) {
    return { html: "", text: "" };
  }

  var splitKey = "■ Magazine Info";
  var idx = originalDesc.indexOf(splitKey);
  var htmlResult = originalDesc;
  var textForName = "";

  if (idx !== -1) {
    // Magazine Info あり
    var ribbonPart = originalDesc.substring(0, idx).trim(); // リボンなど
    var infoPart   = originalDesc.substring(idx).trim();    // Magazine Info 以降

    // Magazine Info 部分をHTML整形
    var formattedInfoHtml = formatMagazineInfo(infoPart);

    htmlResult = formattedInfoHtml + "<br><br>" + ribbonPart;

    // 人名抽出用テキスト（タグ除去＋改行化）
    textForName = formattedInfoHtml
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<[^>]*>/g, "")
      .trim();
  } else {
    // Magazine Info がないパターン（★中国からの輸入雑誌です。COVER ... など）
    htmlResult = originalDesc;
    textForName = originalDesc
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<[^>]*>/g, "")
      .trim();
  }

  return { html: htmlResult, text: textForName };
}

/**************************************************
 * タイトルから人名候補を抽出
 **************************************************/
function 抽出_人名_fromTitle(title) {
  var results = [];
  if (!title) return results;

  var t = String(title);

  // 「：」があれば後ろを優先
  if (t.indexOf("：") !== -1) {
    t = t.split("：")[1];
  }

  // 1) 丸カッコ内
  var m = t.match(/\((.*?)\)/);
  if (m && m[1]) {
    results = results.concat(処理_1行_人名(m[1]));
  }

  // 2) 角カッコ内 [TREASURE別冊付録] など
  var b = t.match(/\[(.*?)\]/);
  if (b && b[1]) {
    results = results.concat(処理_1行_人名(b[1]));
  }

  // 3) 「フォトカード」「ミニポスター」などの前まで
  var cut = t.split(/フォトカード|ポストカード|ミニポスター|トレカ|グッズ|セット|1枚|2枚|3枚/)[0];
  results = results.concat(処理_1行_人名(cut));

  return results;
}

/**************************************************
 * Magazine Info から人名候補を抽出
 *  - ■ Magazine Info〜「外」までを対象
 *  - Magazine Info が無い場合は COVER / PEOPLE などのブロックを対象
 **************************************************/
function 抽出_人名_fromMagazineText(text) {
  var results = [];
  if (!text) return results;

  var sectionKeywords = ["COVER STORY", "COVER", "PEOPLE", "INTERVIEW", "SPECIAL", "FEATURE"];

  if (text.indexOf("■ Magazine Info") !== -1) {
    // ▼ パターンA：Magazine Info あり
    var lines = text.split(/\r\n|\r|\n/);
    var inBlock   = false; // ■ Magazine Info〜外
    var inSection = false; // COVER / PEOPLE などの中

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line) continue;

      // ブロック開始
      if (line.indexOf("■ Magazine Info") !== -1) {
        inBlock   = true;
        inSection = false;
        continue;
      }
      if (!inBlock) continue;

      // ブロック終了条件
      if (line === "外") {
        break;
      }
      if (line.indexOf("■ ") === 0 ||
          line.indexOf("※") === 0 ||
          line.indexOf("★") === 0 ||
          line.indexOf("********************************") !== -1) {
        break;
      }

      // セクション見出しか？
      var isHeader = false;
      for (var k = 0; k < sectionKeywords.length; k++) {
        var key = sectionKeywords[k];
        if (line.toUpperCase().indexOf(key) !== -1) {
          inSection = true;
          isHeader  = true;
          var content = line.replace(key, "").trim();
          if (content) {
            results = results.concat(処理_1行_人名(content));
          }
          break;
        }
      }
      if (isHeader) continue;

      // セクション内の行を処理
      if (inSection) {
        if (line === "外") continue;
        results = results.concat(処理_1行_人名(line));
      }
    }

    return results;
  }

  // ▼ パターンB：Magazine Info が無い（COVER だけの説明など）
  var lines2 = text.split(/\r\n|\r|\n/);
  var inSection2 = false;

  for (var j = 0; j < lines2.length; j++) {
    var line2 = lines2[j].trim();
    if (!line2) continue;

    // セクション見出しか？
    var isHeader2 = false;
    for (var kk = 0; kk < sectionKeywords.length; kk++) {
      var key2 = sectionKeywords[kk];
      if (line2.toUpperCase().indexOf(key2) !== -1) {
        inSection2 = true;
        isHeader2  = true;
        var content2 = line2.replace(key2, "").trim();
        if (content2) {
          results = results.concat(処理_1行_人名(content2));
        }
        break;
      }
    }
    if (isHeader2) continue;

    // 終了条件
    if (inSection2) {
      if (line2.indexOf("■ ") === 0 ||
          line2.indexOf("※") === 0 ||
          line2.indexOf("★") === 0 ||
          line2.indexOf("********************************") !== -1) {
        inSection2 = false;
        continue;
      }
      if (line2 === "外") continue;

      results = results.concat(処理_1行_人名(line2));
    }
  }

  return results;
}

/**************************************************
 * 1行のテキストから人名候補を抽出
 **************************************************/
function 処理_1行_人名(line) {
  var results = [];
  if (!line) return results;

  var t = String(line).trim();
  if (!t) return results;

  // 行頭の★などを除去
  t = t.replace(/^★+/, "").trim();
  if (!t || t === "外") return results;

  // 「表紙Aタイプ」などだけの行は無視
  if (/表紙.?タイプ/.test(t)) {
    return results;
  }

  // (14P) などページ表記をカット
  var pm = t.match(/(.+?)\s*\(\d+P\)/);
  if (pm && pm[1]) {
    t = pm[1].trim();
  }

  // 分割（＆ / & / × / X / , / ／ / /）
  var parts = t.split(/＆|&|×|x|X|,|／|\//);
  for (var i = 0; i < parts.length; i++) {
    var p = parts[i].trim();
    if (!p) continue;
    results.push(p);
  }

  return results;
}

/**************************************************
 * キーワード共通クリーニング
 **************************************************/
function クリーン_人名候補(str) {
  if (!str) return "";
  var s = String(str).trim();
  if (!s) return "";

  // 完全に「外」だけなら削除
  if (s === "外") return "";

  // 不要語削除
  s = s.replace(/両面表紙/g, "")
       .replace(/表紙選択/g, "")
       .replace(/表紙/g, "")
       .replace(/別冊付録/g, "")
       .replace(/特集/g, "")
       .replace(/記事/g, "")
       .replace(/special edition/gi, "")
       .replace(/ver\./gi, "")
       .replace(/[A-ZＡ-Ｚ]\s*ver\./gi, "")
       .replace(/[A-ZＡ-Ｚ]\s*タイプ/g, "")
       .replace(/タイプ/g, "")
       .replace(/外/g, "")
       .replace(/\*+/g, "")
       .replace(/Vol\.\s*\d+/gi, "")
       .replace(/vol\.\s*\d+/gi, "")
       .replace(/[0-9]{4}年[0-9]{1,2}月号?/g, "")
       .replace(/韓国\s*雑誌/g, "")
       .replace(/中国\s*雑誌/g, "")
       .replace(/台湾\s*雑誌/g, "")
       .replace(/【.*?】/g, "")
       .replace(/\[.*?\]/g, "")
       .replace(/「|」/g, "")
       .trim();

  // 前後の区切り文字
  s = s.replace(/^[\/・\s]+/, "").replace(/[\/・\s]+$/, "");

  return s.trim();
}

/**************************************************
 * Magazine Info テキストをHTML整形
 **************************************************/
function formatMagazineInfo(text) {
  var t = String(text);

  t = t.replace(/■ Magazine Info/g, '■ Magazine Info<br><br>')
       .replace(/COVER STORY/g, '<br>COVER STORY<br>')
       .replace(/COVER/g, '<br>COVER<br>')
       .replace(/INTERVIEW/g, '<br>INTERVIEW<br>')
       .replace(/PEOPLE/g, '<br>PEOPLE<br>')
       .replace(/\*{10,}/g, '<br>********************************<br>')
       .replace(/※この雑誌は/g, '<br>※この雑誌は')
       .replace(/※通関の/g, '<br>※通関の')
       .replace(/■ 発売日/g, '<br>■ 発売日')
       .replace(/■ ページ/g, '<br>■ ページ')
       .replace(/■ 出版社/g, '<br>■ 出版社')
       .replace(/■ 選　択/g, '<br>■ 選　択');

  return t;
}

/**************************************************
 * Qoo10：メニュー追加（スプレッドシートを開いた時に表示）
 **************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('🛒 Qoo10')
    .addItem('✅ 一括実行（雑誌：X列整形＋AC検索ワード）', 'Qoo10_一括実行_雑誌説明と検索ワード')
    .addSeparator()
    .addItem('個別実行：雑誌用_X列説明整形_検索ワード', '雑誌用_X列説明整形_検索ワード')
    .addToUi();
}

/**************************************************
 * Qoo10：一括実行（入口）
 * - 同時実行ロック
 * - 実行時間計測
 * - 失敗時も原因を出す
 **************************************************/
function Qoo10_一括実行_雑誌説明と検索ワード() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    SpreadsheetApp.getUi().alert('別の処理が実行中みたいや。少し待ってからもう一回やってな。');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const start = new Date();

  try {
    ss.toast('Qoo10 一括実行：開始', 'RUN', 5);

    // ★ここに「一括で走らせたい処理」を並べていくだけでOK
    雑誌用_X列説明整形_検索ワード();

    const sec = Math.round((new Date() - start) / 1000);
    ss.toast(`Qoo10 一括実行：完了（${sec}秒）`, 'DONE', 5);

  } catch (e) {
    const msg =
      'Qoo10 一括実行でエラーが出たで。\n\n' +
      '【エラー内容】\n' + (e && e.message ? e.message : e) + '\n\n' +
      '（詳細は「実行ログ」も見てな）';

    console.error(e);
    SpreadsheetApp.getUi().alert(msg);
  } finally {
    lock.releaseLock();
  }
}

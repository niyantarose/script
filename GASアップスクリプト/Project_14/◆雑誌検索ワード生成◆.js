// ==============================
//  設定
// ==============================
const TITLE_COL = 5;  // 商品タイトル列（E列）
const AC_COL    = 29; // AC列

// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
//  辞書１：除外ワード（雑誌名・国名・号情報・付録・広告など）
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
const EXCLUDE_WORDS = [
  // --- 国名 / カテゴリ ---
  "韓国", "韓国版", "中国", "中国版", "台湾", "台湾版", "海外",
  "香港", "Hong Kong", "Taiwan", "TAIWAN",
  "女性", "男性", "芸能", "教養", "映画",
  "雑誌", "マガジン",
  "Korea", "KOREA", "Japan", "Philippines", "Singapore",

  // “主要” 関連は完全にノイズ扱い
  "主要", "主要記事",

  // --- 雑誌名 ---
  "allure", "Allure",
  "ELLE", "ELLE MEN", "ELLE SINGAPORE",
  "VOGUE", "Vogue",
  "GQ", "GQ Korea", "GQ KOREA", "GQ中文版",
  "DAZED", "DAZED KOREA", "DAZED CONFUSED KOREA", "DAZED CONFUSED UK",
  "Harper's BAZAAR", "Harper's BAZAAR Man", "BAZAAR", "BAZAAR Man",
  "COSMOPOLITAN", "Cosmopolitan",
  "W Korea", "W KOREA", "W",
  "Esquire", "ESQUIRE",
  "ARENA HOMME+", "HOMME+", "HOMMES",
  "marie claire", "Marie Claire",
  "Singles", "SINGLES",
  "1st LOOK", "LOOK",
  "@Star1", "STAR1", "＠Star1",
  "HIGH CUT", "HIGHCUT",
  "THE BIG ISSUE", "BIG ISSUE",
  "CINE21", "Cine21",
  "CRAZY GIANT",
  "L'Officiel", "LOFFICIEL",
  "L'OFFICIEL HOMMES", "L'OFFICIEL HOMMES CHINA",
  "Luxury", "LUXURY",
  "Noblesse", "Noblesse MEN", "MEN Noblesse",
  "THE STAR", "10ASIA", "10＋Star", "10+Star",
  "OhBoy!", "OhBoy! Magazine",
  "HIM",
  "K-BEAUTY", "K BEAUTY",
  "Pilates S", "ピラティスS", "ピラティス S",
  "Rolling Stone Korea",
  "STAR FOCUS", "Star Focus",
  "MAPS",
  "MAXIM KOREA", "MAXIM",
  "Living sense", "Living Sense",
  "THE MUSICAL",
  "HUMAN AID",
  "PMK ISSUE",
  "MAG AND JINA",
  "NYLON", "Nylon",
  "MEN'S UNO", "MEN'S UNO HK",
  "#legend",

  // --- 号情報 ---
  "号", "No.", "No", "Vol.", "Vol", "Volume",
  "月号",
  "2025", "2024", "2023", "2022", "2021", "2020", "2019", "2018",
  "2025年", "2024年", "2023年", "2022年", "2021年", "2020年",
  "2019年", "2018年",

  // --- タイプ / バージョン ---
  "Aタイプ", "Bタイプ", "Cタイプ", "Dタイプ", "Eタイプ", "Fタイプ",
  "Gタイプ", "Hタイプ", "Kタイプ", "Lタイプ",
  "A ver.", "B ver.", "C ver.", "D ver.", "Ver.", "ver.",
  "type", "Type", "Editio", "Edition",

  // --- 付録 / 特典 ---
  "付録", "別冊付録", "特典", "特装版", "附録",
  "ポスター", "折り畳みポスター", "折りたたみポスター",
  "フォトカード", "フォトカード5種", "+フォトカード5種",
  "ミニカード", "はがき4種", "カレンダー", "カレンダーつき",
  "ランダム発送", "表紙ランダム発送", "両面表紙", "両面", "裏面", "裏面広告",
  "広告収録", "広告", "記事", "ほか記事", "他記事",
  "種類", "ランダム", "枚",

  // --- サイズ情報など ---
  "二つ折りページ", "ページ約", "mm",
  "255×335", "255×335mm",

  // 語関連
  "語 まんが",
  "語版",

  // --- 折りたたみ系 ---
  "折りたたみ", "折り畳み", "折りたたみ 1種",

  // --- 表紙関連 ---
  "表紙選択", "表紙",

  // --- その他 ---
  "送料無料", "予約", "在庫放出",
  "[il]", "[ポスター]",
  "夏号", "冬号", "春号", "秋号",
  "中国語版", "中国語", "英語", "英文版",
  "月刊中国版", "月刊", "季刊", "年4回発行",
];

// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
//  辞書２：K-POPグループ名
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
const KPOP_GROUPS = [
  "東方神起", "TVXQ",
  "少女時代", "Girls' Generation",
  "SUPER JUNIOR",
  "SHINee",
  "EXO",
  "NCT", "NCT 127", "NCT DREAM", "WayV",
  "Stray Kids",
  "SEVENTEEN",
  "BTS", "防弾少年団",
  "TOMORROW X TOGETHER", "TXT",
  "ENHYPEN",
  "NewJeans",
  "LE SSERAFIM",
  "IVE",
  "ITZY",
  "Kep1er",
  "ZEROBASEONE", "ZB1",
  "TREASURE",
  "THE BOYZ",
  "ATEEZ",
  "STAYC",
  "(G)I-DLE", "G I-DLE",
  "ILLIT",
  "RIIZE",
  "BOYNEXTDOOR",
  "DAY6",
  "Red Velvet",
  "aespa",
  "TWICE",
  "KARA",
  "Highlight", "HIGHLIGHT",
  "BTOB",
  "MONSTA X",
  "ASTRO",
  "SF9",
  "NMIXX",
  "QWER",
  "TWS",
  "NATURE",
  "LABOUM",
  "APRIL",
  "Rocket Punch",
  "EVNNE",
  "VANNER",
  "n.SSign",
];

// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
//  辞書３：よく出る個人名
// ＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
const PERSON_NAMES = [
  "チョン・ユミ", "ユンホ", "チャンミン",
  "ソルヒョン", "ジミン", "ジョングク", "IU",
  // 必要に応じてここに追記
];

// カタカナ韓国名っぽいパターン
const KOREAN_NAME_REGEX = /^(?:[ァ-ヶー]{2,4}・[ァ-ヶー]{2,6}|[ァ-ヶー]{3,8})$/;

// ----------------------------------------------------
//  ノイズ削除（normalize_token）
// ----------------------------------------------------
function normalizeToken(text) {
  if (!text) return "";
  let t = String(text).trim();

  // 「表紙」以降を全部落とす
  t = t.replace(/表紙.*$/g, "");

  // 冒頭の「12月 タイプB」「12月タイプB」などを削除
  t = t.replace(/^[0-9]{1,2}月[\s　]*タイプ[ A-Za-zＡ-Ｚａ-ｚ]*/g, "");

  // タイプ / ver / type（末尾）
  t = t.replace(/[A-Z]\s*タイプ$/g, "");
  t = t.replace(/[A-D]\s*ver\.?$/gi, "");
  t = t.replace(/[A-Z]\s*type$/gi, "");

  // 「1月」「2月号」「10月号」など（末尾）
  t = t.replace(/[0-9]{1,2}月(号)?$/g, "");
  // 行頭の「4月」「5月」など
  t = t.replace(/^[0-9]{1,2}月[\s　]*/g, "");

  // 「…のV」「…のRM」
  t = t.replace(/の[0-9A-Za-z]+$/g, "");

  // 年号削除（2024年 / 1999年 など）
  t = t.replace(/[0-9]{2,4}年/g, "");

  // サイズ表記削除（255×335mm / 210x297mm など）
  t = t.replace(/[0-9]{2,4}[×xX][0-9]{2,4}mm/g, "");

  // 「年」「語」単独
  t = t.replace(/^年[\s　]*/g, "");
  if (t.trim() === "年" || t.trim() === "語") return "";

  // コロン
  t = t.replace(/[：:]/g, "");

  // 「ほか記事」「他記事」「表紙選択」
  t = t.replace(/ほか記事$/g, "");
  t = t.replace(/他記事$/g, "");
  t = t.replace(/表紙選択$/g, "");

  // 前後の空白・中黒など
  t = t.replace(/^[ ･・　]+|[ ･・　]+$/g, "");

  return t;
}

// ----------------------------------------------------
//  人名っぽいか？
// ----------------------------------------------------
function isPersonLike(text) {
  if (!text) return false;
  if (PERSON_NAMES.indexOf(text) !== -1) return true;
  if (KOREAN_NAME_REGEX.test(text)) return true;
  return false;
}

// ----------------------------------------------------
//  タイトルからACキーワード抽出
// ----------------------------------------------------
function cleanTitle(title) {
  if (typeof title !== "string") return "";

  const kanaKanjiRegex = /[ぁ-んァ-ン一-龥가-힣]/;

  let text = title.replace(/　/g, " ");
  const original = text;

  const keywords = [];
  const groups   = [];
  const persons  = [];

  // ---------- 0) カッコ内 ----------
  const parenRegex = /（([^）]+)）|\(([^)]+)\)/g;
  let m;
  while ((m = parenRegex.exec(original)) !== null) {
    const inner = m[1] || m[2];
    if (!inner) continue;

    let tmp = inner.replace(/、/g, "$$").replace(/\//g, "$$");
    const tokens = tmp.split("$$");

    tokens.forEach(tokenRaw => {
      let token = tokenRaw.trim();
      if (!token) return;

      // 除外ワード削除
      EXCLUDE_WORDS.forEach(w => {
        if (w) token = token.split(w).join(" ");
      });
      token = token.replace(/\s+/g, " ").trim();
      token = normalizeToken(token);
      if (!token) return;

      if (!kanaKanjiRegex.test(token)) return;

      if (isPersonLike(token)) {
        if (persons.indexOf(token) === -1) persons.push(token);
      } else {
        if (keywords.indexOf(token) === -1) keywords.push(token);
      }
    });
  }

  // ---------- カッコを本体から削る ----------
  let originalNoParen = original.replace(/（.*?）/g, " ");
  originalNoParen = originalNoParen.replace(/\(.*?\)/g, " ");

  // ---------- 1) グループ名 ----------
  KPOP_GROUPS.forEach(grp => {
    if (grp && originalNoParen.indexOf(grp) !== -1 && groups.indexOf(grp) === -1) {
      groups.push(grp);
    }
  });

  // ---------- 2) PERSON_NAMES ----------
  PERSON_NAMES.forEach(name => {
    if (name && originalNoParen.indexOf(name) !== -1 && persons.indexOf(name) === -1) {
      persons.push(name);
    }
  });

  // ---------- 3) 「〇〇の△△表紙」パターン ----------
  const coverRegex = /(.+?)表紙/g;
  let cm;
  while ((cm = coverRegex.exec(originalNoParen)) !== null) {
    let frag = cm[1].trim();
    let cand = frag.indexOf("の") !== -1 ? frag.split("の").pop() : frag;
    cand = normalizeToken(cand);
    if (!cand) continue;

    if (EXCLUDE_WORDS.some(w => w && cand.indexOf(w) !== -1)) continue;
    if (/^[0-9]+$/.test(cand)) continue;

    if (isPersonLike(cand)) {
      if (persons.indexOf(cand) === -1) persons.push(cand);
    } else if (keywords.indexOf(cand) === -1) {
      keywords.push(cand);
    }
  }

  // ---------- 4) $$ 区切りの各パート ----------
  originalNoParen.split(/\$\$+/).forEach(partRaw => {
    let part = partRaw.trim();
    if (!part) return;

    let tmp = part;
    EXCLUDE_WORDS.forEach(w => {
      if (w) tmp = tmp.split(w).join(" ");
    });
    tmp = tmp.replace(/\s+/g, " ").trim();
    if (!tmp) return;

    tmp = normalizeToken(tmp);
    if (!tmp) return;
    if (!kanaKanjiRegex.test(tmp)) return;

    if (isPersonLike(tmp)) {
      if (persons.indexOf(tmp) === -1) persons.push(tmp);
    } else {
      if (keywords.indexOf(tmp) === -1) keywords.push(tmp);
    }
  });

  // ---------- 5) 結果まとめ ----------
  const result = [].concat(groups, persons, keywords);
  const seen = new Set();
  const final = [];

  result.forEach(itemRaw => {
    let item = String(itemRaw).trim();
    if (!item) return;
    if (seen.has(item)) return;
    if (EXCLUDE_WORDS.some(w => w && w === item)) return;
    if (item === "年" || item === "語") return;

    seen.add(item);
    final.push(item);
  });

  if (final.length === 0) return "";

  // Qoo10制限：最大10個・各30文字
  const limited = [];
  for (let i = 0; i < final.length && i < 10; i++) {
    limited.push(final[i].slice(0, 30));
  }

  return limited.join("$$");
}

// ====================================================
//  メイン：ACキーワード一括更新  （ボタンから呼ぶ）
// ====================================================
function runAcUpdate() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("Qoo10");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("シート『Qoo10』が見つかりません。");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("データ行がありません。");
    return;
  }

  const lastCol = sheet.getLastColumn();
  const numRows = lastRow - 1;

  const values = sheet.getRange(2, 1, numRows, lastCol).getValues();

  const newAcList = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const title = row[TITLE_COL - 1] || "";
    const cleaned = cleanTitle(title).trim();

    const acValue = cleaned ? cleaned : "";
    newAcList.push([acValue]);

    if ((i + 2) % 300 === 0) {
      Logger.log("... " + (i + 2) + " 行目まで処理完了");
    }
  }

  sheet.getRange(2, AC_COL, numRows, 1).setValues(newAcList);

  SpreadsheetApp.getUi().alert("ACキーワード一括更新が完了しました。（" + numRows + " 行）");
}

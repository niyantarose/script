/******************************************************
 * 画像翻訳 & 整形ツール（翻訳エンジン強化版 V2.4）
 * - 修正: 人名保護ロジックを「全文保護」から「候補抽出式」に変更
 * (翻訳精度が劇的に向上し、文法が壊れなくなります)
 ******************************************************/

// ====== 設定 ======
const APP_CONFIG = {
  defaultOcrLang: 'ko',
  defaultTargetLang: 'ja',

  // ★★★ 韓国語→翻訳前の前処理（誤訳防止のヒント埋め込み：タグ方式） ★★★
  preKoHints: [
    // --- 版・エディション系 ---
    [/특장판/g, '특장판〔H:特装版〕'],
    [/한정판/g, '한정판〔H:限定版〕'],
    [/초판/g, '초판〔H:初回限定〕'],
    [/초회/g, '초회〔H:初回〕'],
    [/통상판/g, '통상판〔H:通常版〕'],
    [/일반판/g, '일반판〔H:通常版〕'],
    [/디럭스/g, '디럭스〔H:デラックス〕'],
    [/스페셜/g, '스페셜〔H:スペシャル〕'],
    [/리미티드/g, '리미티드〔H:限定〕'],
    [/에디션/g, '에디션〔H:エディション〕'],
    [/컬렉터즈/g, '컬렉터즈〔H:コレクターズ〕'],
    [/프리미엄/g, '프리미엄〔H:プレミアム〕'],

    // --- 特典・付属品系 ---
    [/특전/g, '특전〔H:特典〕'],
    [/구성품/g, '구성품〔H:セット内容〕'],
    [/구성/g, '구성〔H:セット内容〕'],
    [/증정/g, '증정〔H:プレゼント〕'],
    [/사은품/g, '사은품〔H:購入特典〕'],
    [/굿즈/g, '굿즈〔H:グッズ〕'],
    [/포토카드/g, '포토카드〔H:フォトカード〕'],
    [/포카/g, '포카〔H:フォトカード〕'],
    [/엽서/g, '엽서〔H:ポストカード〕'],
    [/스티커/g, '스티커〔H:ステッカー〕'],
    [/포스터/g, '포스터〔H:ポスター〕'],
    [/브로마이드/g, '브로마이드〔H:ブロマイド〕'],
    [/아크릴/g, '아크릴〔H:アクリル〕'],
    [/키링/g, '키링〔H:キーリング〕'],
    [/북마크/g, '북마크〔H:しおり〕'],
    [/책갈피/g, '책갈피〔H:しおり〕'],
    [/미니북/g, '미니북〔H:ミニブック〕'],
    [/화보집/g, '화보집〔H:写真集〕'],
    [/화보/g, '화보〔H:グラビア〕'],
    [/셀피/g, '셀피〔H:セルフィー〕'],

    // --- 発送・配送系 ---
    [/랜덤/g, '랜덤〔H:ランダム〕'],
    [/랜덤발송/g, '랜덤발송〔H:ランダム発送〕'],
    [/선택불가/g, '선택불가〔H:選択不可〕'],
    [/품절/g, '품절〔H:品切れ〕'],
    [/재입고/g, '재입고〔H:再入荷〕'],
    [/예약/g, '예약〔H:予約〕'],
    [/출시/g, '출시〔H:発売〕'],
    [/발매/g, '발매〔H:発売〕'],
    [/입고/g, '입고〔H:入荷〕'],
    [/배송/g, '배송〔H:配送〕'],

    // --- 数量・種類系 ---
    [/종/g, '종〔H:種〕'],
    [/세트/g, '세트〔H:セット〕'],
    [/버전/g, '버전〔H:バージョン〕'],
    [/타입/g, '타입〔H:タイプ〕'],
    [/중(\d+)/g, '중$1〔H:全$1種のうち〕'],

    // --- 状態・注意系 ---
    [/미개봉/g, '미개봉〔H:未開封〕'],
    [/새상품/g, '새상품〔H:新品〕'],
    [/소진시/g, '소진시〔H:なくなり次第〕'],
    [/한정수량/g, '한정수량〔H:数量限定〕'],
    [/단독/g, '단독〔H:単独/独占〕'],

    // --- 人物・グループ系 ---
    [/멤버/g, '멤버〔H:メンバー〕'],
    [/솔로/g, '솔로〔H:ソロ〕'],
    [/유닛/g, '유닛〔H:ユニット〕'],
    [/단체/g, '단체〔H:団体/グループ〕'],
  ],

  // ★★★ 日本語の後処理（誤訳・不自然な表現の修正） ★★★
  glossaryJaFix: [
    // === 版・エディション系の確定修正 ===
    [/特長版/g, '特装版'],
    [/特裝版/g, '特装版'],
    [/特別版/g, '特装版'],
    [/初回版/g, '初回限定版'],
    [/初版限定/g, '初回限定版'],
    [/一般版/g, '通常版'],
    [/一般盤/g, '通常盤'],

    // === 特典・付属品系 ===
    [/贈呈品/g, '特典'],
    [/贈呈/g, 'プレゼント'],
    [/謝恩品/g, '購入特典'],
    [/構成品/g, 'セット内容'],
    [/構成物/g, 'セット内容'],
    [/同梱内容/g, 'セット内容'],
    [/フォトカード/g, 'トレカ'],
    [/写真カード/g, 'トレカ'],
    [/はがき/g, 'ポストカード'],
    [/葉書/g, 'ポストカード'],

    // === 発送・在庫系 ===
    [/ランダム\s*発送/g, 'ランダム発送'],
    [/ランダム\s*封入/g, 'ランダム封入'],
    [/選択\s*不可/g, '選択不可'],
    [/消尽時/g, 'なくなり次第終了'],
    [/消尽したら/g, 'なくなり次第'],
    [/在庫切れ/g, '品切れ'],
    [/売り切れ次第/g, 'なくなり次第'],
    [/数量限定/g, '限定数量'],

    // === 動詞・表現の自然化 ===
    [/提供します/g, '付属します'],
    [/提供されます/g, '付属します'],
    [/提供いたします/g, '付属します'],
    [/含まれます/g, '含まれています'],
    [/含まれております/g, '含まれています'],
    [/構成されます/g, 'セットになっています'],
    [/構成されています/g, 'セットになっています'],
    [/進行します/g, '行います'],
    [/進行されます/g, '行われます'],
    [/発送されます/g, 'お届けします'],
    [/配送されます/g, 'お届けします'],

    // === 機械翻訳っぽい表現の修正 ===
    [/ご参考ください/g, 'ご参照ください'],
    [/参考してください/g, 'ご参照ください'],
    [/ご了承お願いします/g, 'ご了承ください'],
    [/了解お願いします/g, 'ご了承ください'],
    [/確認お願いします/g, 'ご確認ください'],
    [/注意お願いします/g, 'ご注意ください'],
    [/お問い合わせお願いします/g, 'お問い合わせください'],
    [/よろしくお願いいたします/g, 'お願いいたします'],
    [/の場合があります/g, 'することがあります'],
    [/になる場合があります/g, 'になることがあります'],
    [/異なる場合があります/g, '異なることがあります'],

    // === 不自然な敬語の修正 ===
    [/させていただきます/g, 'いたします'],
    [/してくださいますよう/g, 'していただきますよう'],
    [/ございますので/g, 'ありますので'],
    [/でございます/g, 'です'],

    // === 冗長表現の簡潔化 ===
    [/することができます/g, 'できます'],
    [/することが可能です/g, 'できます'],
    [/していくことになります/g, 'します'],
    [/となっております/g, 'です'],
    [/となります/g, 'です'],
    [/になります/g, 'です'],
    [/の方/g, ''],

    // === 記号・表記の統一 ===
    [/！！+/g, '！'],
    [/？？+/g, '？'],
    [/\.\.\.+/g, '…'],
    [/~~~+/g, '〜'],
    
    // ★重要変更: 改行(\n)は消さず、スペース・タブだけを正規化する
    [/[ \t　]+/g, ' '],
  ],

  // ★★★ 文末・文体の調整 ★★★
  sentencePatterns: [
    // 「〜ます。」で終わる文を自然に
    [/ますます\./g, 'ます。'],
    [/ですです\./g, 'です。'],

    // 体言止めの後の処理
    [/です。です。/g, 'です。'],

    // 箇条書き風の整理
    [/^[・•]\s*/gm, '・'],
    [/^[-－]\s*/gm, '・'],
    [/^[★☆]\s*/gm, '★ '],

    // 括弧の統一
    [/（/g, '('],
    [/）/g, ')'],
    [/【/g, '['],
    [/】/g, ']'],
  ],

  // 日本語の仕上げ設定
  jpPolish: {
    squeezeBlankLines: true,
    unifyPunct: true,
    removeOverExclaim: true,
    naturalizeEndings: true
  }
};


// --- Webアプリの入り口 ---
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('画像翻訳 & 整形ツール')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// --- メイン処理：OCR & 翻訳（Drive v2 OCR + convert=true方式 + 新・人名保護） ---
function processImageTranslationEx(base64Data, targetLang) {
  let fileId = null;
  try {
    const tgt = targetLang || APP_CONFIG.defaultTargetLang;

    const mimeType = base64Data.substring(5, base64Data.indexOf(';'));
    const title = 'temp_ocr_' + Utilities.getUuid();
    const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
    const blob = Utilities.newBlob(bytes, mimeType, title);

    // ★OCR (Googleドキュメント化して安定させる)
    const resource = { title: title, mimeType: 'application/vnd.google-apps.document' };
    const options = {
      ocr: true,
      ocrLanguage: APP_CONFIG.defaultOcrLang,
      convert: true // 念のため明示
    };

    const file = Drive.Files.insert(resource, blob, options);
    fileId = file.id;

    // OCR完了待ち（空文字回避）
    let originalText = '';
    for (let i = 0; i < 10; i++) {
      Utilities.sleep(500);
      try {
        const doc = DocumentApp.openById(fileId);
        originalText = (doc.getBody().getText() || '').trim();
        if (originalText) break;
      } catch(e){}
    }

    try { Drive.Files.remove(fileId); } catch (e) {}

    if (!originalText) {
      return { success: false, error: "文字が検出されませんでした（OCR結果が空）" };
    }

    // ★★★ 新・人名保護機能（候補抽出式） ★★★
    // 文脈から人名っぽいものだけを保護するため、翻訳が崩壊しない
    const protectedPack = protectNames_(originalText);
    const protectedText = protectedPack.text;
    const nameMap = protectedPack.map;

    // 翻訳
    const translatedTextRaw = translateNatural(protectedText, APP_CONFIG.defaultOcrLang, tgt);

    // 人名を復元
    const translatedText = restoreNames_(translatedTextRaw, nameMap);

    // 整形
    const formattedText = formatRuleBased(translatedText);

    return { success: true, original: originalText, translated: translatedText, formatted: formattedText };

  } catch (e) {
    if (fileId) { try { Drive.Files.remove(fileId); } catch(err){} }
    return { success: false, error: e.toString() };
  }
}


// --- AI整形機能（ボタン押下時用） ---
function formatForProductPageEx(original, translated) {
  try {
    const formatted = formatWithGemini(original, translated);
    return { success: true, formatted: formatted };
  } catch (e) {
    console.warn("Gemini整形失敗:", e.toString());
    const fallback = formatRuleBased(translated);
    return { success: true, formatted: fallback, isFallback: true, errorDetail: e.toString() };
  }
}


// ★★★ 翻訳エンジン本体 ★★★
function translateNatural(originalText, sourceLang, targetLang) {
  let pre = String(originalText).trim();
  pre = pre.replace(/\r\n/g, '\n');
  pre = pre.replace(/[ \t　]+/g, ' ');

  pre = applyReplacements(pre, APP_CONFIG.preKoHints);

  let ja = LanguageApp.translate(pre, sourceLang || 'ko', targetLang || 'ja');

  ja = ja.replace(/〔H:[^〕]+〕/g, '');
  ja = applyReplacements(ja, APP_CONFIG.glossaryJaFix);
  ja = applyReplacements(ja, APP_CONFIG.sentencePatterns);
  ja = polishJapanese(ja);

  return ja.trim();
}


// ★★★ 内部ロジック: AI整形 (Gemini 1.5 Flash) ★★★
function formatWithGemini(original, translated) {
  const API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!API_KEY) throw new Error('API Key未設定');

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${API_KEY}`;

  const prompt = `
あなたは日本のECサイト（Amazon、Yahoo!ショッピング、Qoo10）向けの商品説明文ライターです。
以下の韓国語原文と機械翻訳を参考に、自然で読みやすい日本語の商品説明文を作成してください。

【重要なルール】
1. 文体は「です・ます調」
2. 特装版/限定版/初回限定/トレカ/セット内容/特典/ランダム 等の用語は正しく使う
3. 簡潔に書く
4. 「〜させていただきます」→「〜いたします」
5. 「〜することができます」→「〜できます」
6. 誇張表現は避ける

【原文】
${original}

【機械翻訳】
${translated}

【出力形式】
商品説明文のみを出力
`;

  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { 
      temperature: 0.2,
      topP: 0.8,
      topK: 40
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const json = JSON.parse(response.getContentText());
  if (json.error) throw new Error(json.error.message);

  const candidate = json?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!candidate) throw new Error('Gemini応答空');

  let result = applyReplacements(candidate.trim(), APP_CONFIG.glossaryJaFix);
  result = applyReplacements(result, APP_CONFIG.sentencePatterns);

  return result;
}


// --- ルールベース整形 ---
function formatRuleBased(text) {
  if (!text) return "";
  let t = String(text);
  t = applyReplacements(t, APP_CONFIG.glossaryJaFix);
  t = applyReplacements(t, APP_CONFIG.sentencePatterns);
  t = polishJapanese(t);
  return t.trim();
}


// ★★★ 日本語の磨き上げ ★★★
function polishJapanese(text) {
  let t = String(text);

  if (APP_CONFIG.jpPolish.removeOverExclaim) {
    t = t.replace(/[！!]{2,}/g, '！');
    t = t.replace(/！+/g, '。');
  }

  if (APP_CONFIG.jpPolish.unifyPunct) {
    t = t.replace(/[？?]{2,}/g, '？');
  }

  if (APP_CONFIG.jpPolish.squeezeBlankLines) {
    t = t.replace(/(\n\s*){3,}/g, '\n\n');
    t = t.replace(/^\s+/gm, '');
  }

  t = t.replace(/ですです/g, 'です');
  t = t.replace(/ますます([^。])/g, 'ます$1');
  t = t.replace(/。。+/g, '。');
  t = t.replace(/、、+/g, '、');

  if (APP_CONFIG.jpPolish.naturalizeEndings) {
    t = t.replace(/^この商品は/gm, '本商品は');
    t = t.replace(/この商品には/g, '本商品には');
    t = t.replace(/になっています。([^。]+)になっています。/g, 'です。$1になっています。');
    t = t.replace(/購入の方/g, '購入');
    t = t.replace(/注文の方/g, '注文');
    t = t.replace(/となっております/g, 'です');
    t = t.replace(/となっています/g, 'です');
    t = t.replace(/ご確認してください/g, 'ご確認ください');
    t = t.replace(/ご注意してください/g, 'ご注意ください');
    t = t.replace(/ご了承してください/g, 'ご了承ください');
  }

  t = t.replace(/[ \t　]+/g, ' '); 
  t = t.replace(/ ([。、！？）」』】])/g, '$1');
  t = t.replace(/([（「『【]) /g, '$1');

  return t;
}


// ===== ユーティリティ =====
function applyReplacements(text, pairs) {
  let t = String(text);
  (pairs || []).forEach(([from, to]) => {
    t = t.replace(from, to);
  });
  return t;
}


// ★★★ 改良版：人名保護機能（ピンポイント抽出式） ★★★

/**
 * 文脈から「人名候補」だけを特定してトークン化する
 * （全文保護による翻訳崩壊を防ぐ）
 */
function protectNames_(text) {
  const original = String(text);
  // 候補抽出
  const candidates = extractNameCandidatesKo_(original);
  
  if (!candidates.length) return { text: original, map: {} };

  let out = original;
  const map = {};
  let idx = 0;

  function tokenFor_(value) {
    const token = `__PN_${idx}__`;
    map[token] = value;
    idx++;
    return token;
  }

  // 長い順に処理して、短い名前の部分一致誤爆を防ぐ
  candidates.sort((a, b) => b.length - a.length);

  for (const name of candidates) {
    // 特殊文字エスケープ
    const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    // グローバル置換
    out = out.replace(new RegExp(escaped, "g"), () => tokenFor_(name));
  }

  return { text: out, map };
}

/**
 * 人名抽出ロジック（文脈依存）
 * - 「○○の(의)」の前
 * - 「男 ○○」「○○ 登場」
 * - 敬称付き
 */
function extractNameCandidatesKo_(text) {
  const t = String(text);
  const set = new Set();

  // 1. 〜の (미스즈의)
  t.replace(/(^|[\n\r\s,!.?…“”"()［\]【】])([가-힣]{2,6})(?=의[\s\n\r])/g, (_, __, name) => {
    set.add(name);
    return _;
  });

  // 2. 男 〜 (남자 라이카)
  t.replace(/남자\s+([가-힣]{2,6})/g, (_, name) => {
    set.add(name);
    return _;
  });

  // 3. 〜登場 (라이카 등장)
  t.replace(/([가-힣]{2,6})\s+등장/g, (_, name) => {
    set.add(name);
    return _;
  });

  // 4. 敬称 (씨, 님, 양, 군)
  t.replace(/([가-힣]{2,6})(씨|님|양|군)\b/g, (_, name) => {
    set.add(name);
    return _;
  });

  // 誤爆しやすい一般語を除外
  const NG = new Set([
    '과거','알고','있는','수수께끼','같은','남자','등장','그가','그녀가','그것','이것','저것',
    '오늘','내일','어제','정말','진짜','매우','항상'
  ]);

  return [...set].filter(x => !NG.has(x));
}

/**
 * 復元処理（変更なし）
 */
function restoreNames_(text, map) {
  let out = String(text);
  const keys = Object.keys(map).sort((a,b)=>b.length-a.length);
  for (const k of keys) out = out.split(k).join(map[k]);
  return out;
}


// --- 承認用テスト関数 ---
function 承認テスト_Drive() {
  const name = DriveApp.getRootFolder().getName();
  Logger.log("root=" + name);
}

function auth_AdvancedDrive_Test() {
  const list = Drive.Files.list({ maxResults: 1 });
  Logger.log("Drive.Files.list OK: " + (list.items ? list.items.length : 0));
  const blob = Utilities.newBlob("test", "text/plain", "auth_test.txt");
  const created = Drive.Files.insert(
    { title: "auth_test.txt", mimeType: "text/plain" },
    blob
  );
  Logger.log("Drive.Files.insert OK: fileId=" + created.id);
}
function resolveTodayString() {
  if (typeof formatToday === 'function') return formatToday();
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}
// popup.goods.js - 台湾 books.com.tw 拡張 グッズ処理

function stripGoodsMediaTag(value) {
  return trimValue(value)
    .replace(/\s*[（(][^）)]*(?:漫畫|漫画|コミック|小說|小说|小説|動畫|动画|アニメ|影集|ドラマ|電影|映画|劇場版)[^）)]*[)）]\s*/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function extractGoodsWorkTitleText(value) {
  const fallback = trimValue(value).replace(/^[台臺][湾灣]版\s*/u, '');
  if (!fallback) return '';

  const normalizeCandidate = candidate => {
    const normalized = stripGoodsMediaTag(extractOriginalTitleText(candidate) || candidate);
    return normalized || fallback;
  };

  const mediaTagged = fallback.match(/^(.+?)\s*[（(][^）)]*(?:漫畫|漫画|コミック|小說|小说|小説|動畫|动画|アニメ|影集|ドラマ|電影|映画|劇場版)[^）)]*[)）](?:\s+.*)?$/u);
  if (mediaTagged?.[1]) {
    return normalizeCandidate(mediaTagged[1]);
  }

  const merchTagged = fallback.match(/^(.+?)\s+(?:全棉托特袋|托特袋|帆布袋|手提袋|透明立方|壓克力(?:牌|立牌|磚|砖)?|压克力(?:牌|立牌|磚|砖)?|立牌|海報|海报|掛軸|挂轴|徽章|胸章|卡片|貼紙|贴纸|明信片|抱枕|滑鼠墊|鼠標墊|鼠标垫|桌墊|桌垫|吊飾|吊饰|鑰匙圈|钥匙圈|杯墊|杯垫|色紙|色纸|套組|套装|福袋|拼圖|拼图|公仔|玩偶|資料夾|资料夹|文件夾|文件夹|T恤|毛巾|桌曆|桌历|年曆|年历)(?:\b|\s|$)/u);
  if (merchTagged?.[1]) {
    return normalizeCandidate(merchTagged[1]);
  }

  return normalizeCandidate(fallback);
}

function normalizeGoodsSheetDescription(value) {
  const lines = String(value || '')
    .replace(/\r/g, '')
    .split('\n')
    .map(line => line.trim())
    .filter(Boolean)
    .filter(line => !/^(?:產品說明|产品说明|商品簡介|商品介绍|商品介紹|商品說明|商品说明)$/.test(line));

  const uniqueLines = [];
  for (const line of lines) {
    if (!uniqueLines.includes(line)) uniqueLines.push(line);
  }
  return uniqueLines.join('\n');
}

function buildGoodsSheetRow(product) {
  const code = extractProductCode(product);
  const description = normalizeGoodsSheetDescription(product?.商品説明 || '');
  const additional = getAdditionalImagesValue(product);
  const rawTitle = trimValue(product?.商品名 || '');
  const goodsWorkTitle = trimValue(product?.作品名原題 || extractGoodsWorkTitleText(rawTitle));
  const bonusMemo = extractBonusMemo(product);
  const japaneseTitle = trimValue(product?.日本語タイトル || product?.作品名日本語 || product?.['作品名（日本語）'] || '');
  const japaneseProductName = trimValue(product?.['商品名（日本語）'] || product?.商品名日本語 || '');
  const listingName = japaneseProductName || rawTitle;
  const pageUrl = trimValue(product?.URL || '');
  const mainImageUrl = trimValue(product?.画像URL || '');

  return {
    '発番発行': '',
    '登録状況': '',
    '言語': getLanguageValue(product),
    '親コード': trimValue(product?.親コード || ''),
    '商品名（出品用）': '',
    '商品名（日本語）': japaneseProductName,
    '日本語タイトル': japaneseTitle,
    '補足項目': bonusMemo,
    '特典メモ': '',
    '売価': trimValue(product?.売価 || ''),
    '原価': trimValue(product?.価格 || ''),
    '粗利益率': trimValue(product?.粗利益率 || ''),
    '配送パターン': trimValue(product?.配送パターン || ''),
    '登録者': trimValue(product?.登録者 || ''),
    '商品説明': description,
    '備考': description,
    '作品名（原題）': goodsWorkTitle,
    '作品名（日本語）': japaneseTitle,
    '原題タイトル': goodsWorkTitle,
    '商品名（原題）': rawTitle,
    '原題商品タイトル': rawTitle,
    '博客來商品コード': code,
    'サイト商品コード': code,
    '博客來URL': pageUrl,
    'リンク': pageUrl,
    'メイン画像URL': mainImageUrl,
    'メイン画像': mainImageUrl,
    '追加画像URL': additional,
    '追加画像': additional,
    '発売日': trimValue(product?.発売日 || ''),
    '重複チェックキー': code,
    '登録日': resolveTodayString(),
  };
}


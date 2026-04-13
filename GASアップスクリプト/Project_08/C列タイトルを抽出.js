function 作品タイトル候補(text) {
  if (Array.isArray(text)) {
    return text.map(row => [作品タイトル候補_1件(row[0])]);
  }
  return 作品タイトル候補_1件(text);
}

function 作品タイトルマスター(text) {
  let list = [];

  if (Array.isArray(text)) {
    list = text.map(row => row[0]);
  } else {
    list = [text];
  }

  const cleaned = list
    .map(v => 作品タイトル候補_1件(v))
    .filter(v => v && v !== '');

  const uniq = [];
  const seen = new Set();

  for (const v of cleaned) {
    if (!seen.has(v)) {
      seen.add(v);
      uniq.push([v]);
    }
  }

  return uniq;
}

function 作品タイトル候補_1件(text) {
  if (!text) return '';

  let s = String(text);

  // HTMLエンティティ
  s = s
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&amp;/gi, '&')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'");

  // スペース整理
  s = s.replace(/\u3000/g, ' ');
  s = s.replace(/\s+/g, ' ').trim();

  // かっこ以降を削除
  s = s.replace(/\s*[\(（【\[].*$/, '');

  // part / vol / volume / episode
  s = s.replace(/\s+part\.?\s*\d+.*$/i, '');
  s = s.replace(/\s+vol\.?\s*\d+.*$/i, '');
  s = s.replace(/\s+volume\s*\d+.*$/i, '');
  s = s.replace(/\s+episode\s*\d+.*$/i, '');

  // セット・巻数・部数・話数・数字終わり
  s = s.replace(/\s+\d+\s*[~〜～\-‐–—]\s*\d+\s*巻.*$/i, '');
  s = s.replace(/\s+\d+\+\d+.*$/i, '');
  s = s.replace(/\s+\d+\s*巻.*$/i, '');
  s = s.replace(/\s+\d+\s*部.*$/i, '');
  s = s.replace(/\s+\d+\s*話.*$/i, '');
  s = s.replace(/\s+\d+\s*$/, '');

  // 商品語が出たらそこで切る
  const productWords = [
    'OFFICIAL ARTBOOK',
    'ARTBOOK',
    '公式設定集',
    '公式ガイドブック',
    'イラストはがきセット',
    'はがきセット',
    'ブックケースセット',
    'ボックスセット',
    '単行本セット',
    'オールインワンセット',
    '推しセット',
    'セット',
    'コレクティングカード',
    'フォトカード',
    'トレカ',
    'ポラロイド',
    'ポストカード',
    'クリアカード',
    'イラストカード',
    '缶バッジ',
    'スクエアバッジ',
    'グリッターバッジ',
    'バッジ',
    'アクリルスタンド',
    'アクスタ',
    'スタンド',
    'キーホルダー',
    'キーリング',
    'ミニドール',
    'テーマ人形',
    'ヒトデ人形',
    '人形',
    'ぬいぐるみ',
    'ぬい',
    'クッション',
    'タペストリー',
    'ポスター',
    '色紙',
    'カレンダー',
    '卓上カレンダー',
    '壁掛けカレンダー',
    'コラボカフェ',
    'カフェ',
    '写真集',
    'OST',
    '演奏曲集'
  ];

  for (const word of productWords) {
    const re = new RegExp('\\s*' + escapeRegExp_(word) + '.*$', 'i');
    s = s.replace(re, '');
  }

  // 英語タイトル + 日本語作品名 のとき、日本語側を優先
  // 例: ALIEN STAGE エイリアンステージ → エイリアンステージ
  const japChunks = s.match(/[ぁ-んァ-ン一-龥ー・]+/g);
  if (japChunks && japChunks.length > 0) {
    const lastJap = japChunks[japChunks.length - 1];
    if (/[A-Za-z]/.test(s) && lastJap.length >= 3) {
      // 末尾に近い日本語塊があるなら、それを優先
      const idx = s.lastIndexOf(lastJap);
      if (idx >= 0 && idx > 0) {
        const after = s.slice(idx).trim();
        if (after === lastJap || after.startsWith(lastJap + ' ')) {
          s = lastJap;
        }
      }
    }
  }

  // 記号整理
  s = s.replace(/\s+/g, ' ').trim();
  s = s.replace(/[-‐–—―ー・:：/／\s]+$/g, '').trim();

  return s;
}

function escapeRegExp_(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
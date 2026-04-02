const fs = require('fs');
const path = require('path');
const root = 'Z:/script';
const extDir = fs.readdirSync(root, { withFileTypes: true })
  .find(entry => entry.isDirectory() && entry.name.includes('商品登録拡張機能'));
if (!extDir) throw new Error('extension directory not found');
const file = path.join(root, extDir.name, 'content', 'taiwan', 'content.js');
let text = fs.readFileSync(file, 'utf8');

const newBlock = String.raw`  const stripTrailingVolumeSuffix = rawText => toSingleLine(rawText)
    .replace(/\s*[（(［\[]\s*(?:第\s*)?\d+\s*(?:巻|卷|冊|册|集|話|话|部)?\s*[）)］\]]\s*$/u, '')
    .replace(/\s*第?\s*\d+\s*(?:巻|卷|冊|册|集|話|话|部)\s*$/u, '')
    .replace(/\s+\d+\s*$/u, '')
    .trim();

  const extractOriginalTitle = rawTitle => {
    const fallback = toSingleLine(rawTitle);
    if (!fallback) return '';

    const normalized = stripTrailingVolumeSuffix(
      fallback
        .replace(/^[台臺][湾灣]版\s*/u, '')
        .replace(/\s*【[^】]*(?:首刷|初回|限定|特裝|特装|通常|普通)[^】]*】\s*/gu, ' ')
        .replace(/\s*\([^)]*(?:首刷|初回|限定|特裝|特装|通常|普通)[^)]*\)\s*/gu, ' ')
        .replace(/\s*\[[^\]]*(?:首刷|初回|限定|特裝|特装|通常|普通)[^\]]*\]\s*/gu, ' ')
        .replace(/\s*（[^）]*(?:首刷|初回|限定|特裝|特装|通常|普通)[^）]*）\s*/gu, ' ')
        .replace(/\s*[-/／]\s*(?:首刷|初回|限定|特裝|特装|通常|普通)[^ ]*\s*/gu, ' ')
        .replace(/\s+/g, ' ')
    );

    return normalized || fallback;
  };

  const normalizeJapaneseTitle = rawText => stripTrailingVolumeSuffix(
    toSingleLine(rawText)
      .replace(/^(?:日文(?:書名)?|日語(?:書名)?|日语(?:书名)?|日本語(?:タイトル)?|原文(?:書名|书名)?|原題(?:標題|标题|タイトル)?|書名原文)\s*[：:]\s*/u, '')
      .replace(/\s*【[^】]*(?:首刷|初回|限定|特裝|特装|通常|普通|特典|附錄|附录|豪華|豪华)[^】]*】\s*/gu, ' ')
      .replace(/\s*\([^)]*(?:首刷|初回|限定|特裝|特装|通常|普通|特典|附錄|附录|豪華|豪华)[^)]*\)\s*/gu, ' ')
      .replace(/\s*\[[^\]]*(?:首刷|初回|限定|特裝|特装|通常|普通|特典|附錄|附录|豪華|豪华)[^\]]*\]\s*/gu, ' ')
      .replace(/\s*（[^）]*(?:首刷|初回|限定|特裝|特装|通常|普通|特典|附錄|附录|豪華|豪华)[^）]*）\s*/gu, ' ')
      .replace(/\s+/g, ' ')
  );

  const isJapaneseTitleNoise = text => /(?:作者|譯者|译者|出版社|出版日期|ISBN|規格|规格|語言|语言|定價|定价|優惠價|优惠价|價格|价格|庫存|库存|可購買版本|可购买版本|本系列共|分享|放入購物車|加入下次再買|立即購買|立即购买|博客來|books\.com\.tw|NT\$|元)/u.test(String(text || ''));

  const looksLikeJapaneseTitle = (text, requireKana = true) => {
    const normalized = toSingleLine(text);
    if (!normalized) return false;
    if (normalized.length > 120) return false;
    if (isJapaneseTitleNoise(normalized)) return false;
    if (requireKana && !/[ぁ-ゖァ-ヺー]/u.test(normalized)) return false;
    return true;
  };

  const getJapaneseTitle = (titleElement, rawTitle = '') => {
    const labeled = normalizeJapaneseTitle(getLabeledValue(['日文書名', '日語書名', '日语书名', '日本語タイトル', '原文書名', '原文书名', '原題タイトル']));
    if (looksLikeJapaneseTitle(labeled, false)) return labeled;

    const seen = new Set();
    const candidates = [];
    const pushCandidate = (rawText, score = 0, requireKana = true) => {
      const normalized = normalizeJapaneseTitle(rawText);
      if (!normalized || seen.has(normalized)) return;
      seen.add(normalized);
      if (!looksLikeJapaneseTitle(normalized, requireKana)) return;
      if (rawTitle && normalized === normalizeJapaneseTitle(rawTitle)) return;
      candidates.push({ text: normalized, score });
    };

    let sibling = titleElement?.nextElementSibling || null;
    let siblingDepth = 0;
    while (sibling && siblingDepth < 4) {
      pushCandidate(sibling.innerText, 100 - siblingDepth, false);
      sibling = sibling.nextElementSibling;
      siblingDepth += 1;
    }

    const titleParent = titleElement?.parentElement || null;
    if (titleParent) {
      Array.from(titleParent.children).forEach((node, index) => {
        if (node === titleElement) return;
        pushCandidate(node.innerText, 70 - index, index < 2);
      });
    }

    const titleContainer = titleElement?.closest('.type02_p003, .book_title, .mod, .type02_p003.clearfix') || null;
    if (titleContainer && titleContainer !== titleParent) {
      Array.from(titleContainer.children).forEach((node, index) => {
        if (node === titleElement || node.contains?.(titleElement) || titleElement?.contains?.(node)) return;
        pushCandidate(node.innerText, 40 - index, false);
      });
    }

    candidates.sort((a, b) => (b.score - a.score) || (a.text.length - b.text.length));
    return candidates[0]?.text || '';
  };`;

const pattern1 = /  const extractOriginalTitle = rawTitle => \{[\s\S]*?    return normalized \|\| fallback;\r?\n  \};/;
if (!pattern1.test(text)) throw new Error('original title block not found');
text = text.replace(pattern1, newBlock);

const pattern2 = /  \/\/ 商品名\r?\n  const name = document\.querySelector\('h1\.BD_NAME, h1\[itemprop="name"\], \.book_title h1, h1'\)\?\.innerText\?\.trim\(\) \|\| '';?/;
if (!pattern2.test(text)) throw new Error('name block not found');
text = text.replace(pattern2, String.raw`  // 商品名
  const nameEl = document.querySelector('h1.BD_NAME, h1[itemprop="name"], .book_title h1, h1');
  const name = nameEl?.innerText?.trim() || '';
  const japaneseTitle = getJapaneseTitle(nameEl, name);`);

const pattern3 = /    商品名: name,\r?\n    ページタイトル: pageTitle,\r?\n    原題タイトル: originalTitle,/;
if (!pattern3.test(text)) throw new Error('result block not found');
text = text.replace(pattern3, String.raw`    商品名: name,
    日本語タイトル: japaneseTitle,
    ページタイトル: pageTitle,
    原題タイトル: originalTitle,`);

fs.writeFileSync(file, text, 'utf8');
console.log(file);

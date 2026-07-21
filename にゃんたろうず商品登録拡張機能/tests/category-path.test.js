const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');

// ============================================================
// 1) content.js getCategoryPath: パンくず/本書分類だけを拾い、
//    作者フォロー窓（修改/確定/取消）等のUIリンクを混ぜない
// ============================================================
function makeAnchor(text) {
  return { innerText: text };
}
function makeContainer(anchorTexts) {
  return { querySelectorAll: () => anchorTexts.map(makeAnchor) };
}
function makeDocument({ containersBySelector = {}, metaDescription = '' } = {}) {
  return {
    querySelectorAll: selector => containersBySelector[selector] || [],
    querySelector: selector => {
      if (selector === 'meta[name="description"]' && metaDescription) {
        return { getAttribute: name => (name === 'content' ? metaDescription : null) };
      }
      return null;
    },
  };
}

function loadContentJs(doc) {
  const ctx = {
    console,
    document: doc,
    chrome: { runtime: { onMessage: { addListener: () => {} } } },
  };
  ctx.globalThis = ctx;
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(root, 'content/taiwan/content.js'), 'utf8'), ctx, {
    filename: 'content/taiwan/content.js',
  });
  return ctx;
}

let failed = 0;
function eq(name, actual, expected) {
  const a = JSON.stringify(actual);
  const e = JSON.stringify(expected);
  if (a !== e) { failed++; console.error(`[NG] ${name}: expected=${e} actual=${a}`); }
  else { console.log(`[OK] ${name}`); }
}

// 二宮和也『獨斷與偏見』ページの実構造を模したケース
const ninomiyaDoc = makeDocument({
  containersBySelector: {
    '#breadcrumb-trail': [makeContainer(['博客來', '中文書', '心理勵志', '個人成長', '人生規劃/自我改變'])],
    'ul.sort': [makeContainer(['心理勵志', '個人成長', '人生規劃/自我改變', '影視偶像', '偶像書', '海外偶像'])],
    '.type02_p003 a': [], // コンテナ単位に変えたので旧の全体マッチは使わない
  },
});
{
  const ctx = loadContentJs(ninomiyaDoc);
  eq('getCategoryPath がトップレベルに定義されている', typeof ctx.getCategoryPath, 'function');
  eq(
    'パンくず（博客來は除去）を採用する',
    ctx.getCategoryPath(),
    '中文書 > 心理勵志 > 個人成長 > 人生規劃/自我改變'
  );
}

// パンくずが無いページ → 詳細資料「本書分類」(ul.sort) を使う
{
  const ctx = loadContentJs(makeDocument({
    containersBySelector: {
      'ul.sort': [makeContainer(['漫畫/圖文書', 'BL'])],
    },
  }));
  eq('パンくず無しは本書分類(ul.sort)を使う', ctx.getCategoryPath(), '漫畫/圖文書 > BL');
}

// UIリンクだけのコンテナはスキップされ、次のコンテナに進む
{
  const ctx = loadContentJs(makeDocument({
    containersBySelector: {
      '#breadcrumb-trail': [makeContainer(['修改', '確定', '取消', '追蹤作者', '新功能介紹', '訂閱出版社新書快訊'])],
      'ul.sort': [makeContainer(['心理勵志', '個人成長'])],
    },
  }));
  eq('UI文言のみのコンテナはスキップ', ctx.getCategoryPath(), '心理勵志 > 個人成長');
}

// どのコンテナも無い → meta description の「類別：X」から取る
{
  const ctx = loadContentJs(makeDocument({
    metaDescription: '書名：獨斷與偏見，原文名稱：独断と偏見，語言：繁體中文，類別：心理勵志',
  }));
  eq('コンテナ無しはmeta類別から取得', ctx.getCategoryPath(), '心理勵志');
}

// 何も無ければ空
{
  const ctx = loadContentJs(makeDocument({}));
  eq('取得源が無ければ空文字', ctx.getCategoryPath(), '');
}

// ============================================================
// 2) popup: カテゴリ分類とプルダウン汚染ガード
// ============================================================
const popupCtx = {
  console,
  chrome: { storage: { local: { get: () => {}, set: () => {} } } },
};
popupCtx.globalThis = popupCtx;
vm.createContext(popupCtx);
vm.runInContext(fs.readFileSync(path.join(root, 'popup/taiwan/popup.shared.js'), 'utf8'), popupCtx, {
  filename: 'popup/taiwan/popup.shared.js',
});
vm.runInContext(fs.readFileSync(path.join(root, 'popup/taiwan/popup.books.js'), 'utf8'), popupCtx, {
  filename: 'popup/taiwan/popup.books.js',
});

// 心理勵志（自己啓発）系はエッセイに分類する
eq(
  '心理勵志系カテゴリはエッセイ',
  popupCtx.getBookGenreLabel({ カテゴリ: '中文書 > 心理勵志 > 個人成長 > 人生規劃/自我改變' }),
  'エッセイ'
);

// 既存の汚染済みカテゴリ値（分類キーワード入り）でも正しく分類される
eq(
  '汚染済みでも分類キーワードがあればエッセイ',
  popupCtx.getBookGenreLabel({ カテゴリ: '修改 > 確定 > 取消 > 二宮和也 > 心理勵志 > 個人成長' }),
  'エッセイ'
);

// 分類できない生パンくず/長文はプルダウン列へ漏らさずデフォルトに落とす
eq(
  '分類不能な生パスは書籍に落ちる',
  popupCtx.getBookGenreLabel({ カテゴリ: '修改 > 確定 > 取消 > 王筱玲 > 大塊文化' }),
  '書籍'
);

// 短い単独カテゴリ名はそのまま通る（従来挙動）
eq('短い単独名は素通し', popupCtx.getBookGenreLabel({ カテゴリ: '料理' }), '料理');

// まんが分類の回帰確認
eq('漫畫はまんが', popupCtx.getBookGenreLabel({ カテゴリ: '中文書 > 漫畫/圖文書 > BL' }), 'まんが');

// 空はデフォルト書籍
eq('空は書籍', popupCtx.getBookGenreLabel({}), '書籍');

if (failed) {
  console.error(`${failed} 件失敗`);
  process.exit(1);
}
console.log('category-path.test.js: ok');

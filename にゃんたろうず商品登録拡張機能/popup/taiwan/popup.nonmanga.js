// popup.nonmanga.js - 台湾 books.com.tw 拡張 非まんが書籍処理

function getNonMangaBookGenreLabel(product) {
  const label = getBookGenreLabel(product);
  return label && label !== 'まんが' ? label : '書籍';
}

function buildNonMangaBookSheetRow(product) {
  return buildCommonBookSheetRow(product, {
    category: getNonMangaBookGenreLabel(product),
  });
}

function buildNonMangaBookCsvRow(product) {
  return buildCommonBookCsvRow(product);
}

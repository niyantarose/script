// popup.comics.js - 台湾 books.com.tw 拡張 まんが書籍処理

function buildComicBookSheetRow(product) {
  return buildCommonBookSheetRow(product, {
    category: 'まんが',
  });
}

function buildComicBookCsvRow(product) {
  return buildCommonBookCsvRow(product);
}

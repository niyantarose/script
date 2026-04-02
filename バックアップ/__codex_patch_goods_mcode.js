const fs = require('fs');
const path = require('path');
const root = 'Z:/script';
const extDir = fs.readdirSync(root, { withFileTypes: true })
  .find(entry => entry.isDirectory() && entry.name.includes('商品登録拡張機能'));
if (!extDir) throw new Error('extension directory not found');

const contentFile = path.join(root, extDir.name, 'content', 'taiwan', 'content.js');
let content = fs.readFileSync(contentFile, 'utf8');
content = content.replace("  // 書籍: B, E, F など / グッズ: N\n  const isBook = !productCode.startsWith('N');", "  // 書籍: 数字始まりや B/E/F など / グッズ: N, M\n  const isBook = !!productCode && !/^[NM]/i.test(productCode);");
fs.writeFileSync(contentFile, content, 'utf8');

const booksFile = path.join(root, extDir.name, 'popup', 'taiwan', 'popup.books.js');
let books = fs.readFileSync(booksFile, 'utf8');
books = books.replace("  return !!code && !code.startsWith('N');", "  return !!code && !/^[NM]/i.test(code);");
fs.writeFileSync(booksFile, books, 'utf8');

console.log(contentFile);
console.log(booksFile);

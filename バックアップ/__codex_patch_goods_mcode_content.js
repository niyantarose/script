const fs = require('fs');
const path = require('path');
const root = 'Z:/script';
const extDir = fs.readdirSync(root, { withFileTypes: true })
  .find(entry => entry.isDirectory() && entry.name.includes('商品登録拡張機能'));
if (!extDir) throw new Error('extension directory not found');
const contentFile = path.join(root, extDir.name, 'content', 'taiwan', 'content.js');
let content = fs.readFileSync(contentFile, 'utf8');
const pattern = /  \/\/ 商品種別判定（書籍かグッズか）\r?\n  \/\/ 書籍: .*?\r?\n  const isBook = .*?;/;
if (!pattern.test(content)) throw new Error('content isBook block not found');
content = content.replace(pattern, "  // 商品種別判定（書籍かグッズか）\n  // 書籍: 数字始まりや B/E/F など / グッズ: N, M\n  const isBook = !!productCode && !/^[NM]/i.test(productCode);");
fs.writeFileSync(contentFile, content, 'utf8');
console.log(contentFile);

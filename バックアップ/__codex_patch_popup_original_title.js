const fs = require('fs');
const path = require('path');
const root = 'Z:/script';
const extDir = fs.readdirSync(root, { withFileTypes: true })
  .find(entry => entry.isDirectory() && entry.name.includes('商品登録拡張機能'));
if (!extDir) throw new Error('extension directory not found');
const sharedFile = path.join(root, extDir.name, 'popup', 'taiwan', 'popup.shared.js');
let shared = fs.readFileSync(sharedFile, 'utf8');

const pattern = /function extractOriginalTitleText\(value\) \{[\s\S]*?\n\}/;
const replacement = String.raw`function extractOriginalTitleText(value) {
  const fallback = trimValue(value);
  if (!fallback) return '';

  const stripTrailingEditionSuffix = text => trimValue(text)
    .replace(/\s*(?:首刷限定|初回限定|限定|特裝|特装|特仕|豪華|豪华|通常|普通)(?:版|版本)?\s*$/u, '')
    .trim();

  const normalized = stripTrailingEditionSuffix(
    fallback
      .replace(/^[台臺][湾灣]版\s*/u, '')
      .replace(/\s*【[^】]*(?:首刷|初回|限定|特裝|特装|通常|普通)[^】]*】\s*/gu, ' ')
      .replace(/\s*\([^)]*(?:首刷|初回|限定|特裝|特装|通常|普通)[^)]*\)\s*/gu, ' ')
      .replace(/\s*\[[^\]]*(?:首刷|初回|限定|特裝|特装|通常|普通)[^\]]*\]\s*/gu, ' ')
      .replace(/\s*（[^）]*(?:首刷|初回|限定|特裝|特装|通常|普通)[^）]*）\s*/gu, ' ')
      .replace(/\s*[-/／]\s*(?:首刷|初回|限定|特裝|特装|通常|普通)[^ ]*\s*/gu, ' ')
      .replace(/\s*[（(［\[]\s*(?:第\s*)?\d+\s*(?:巻|卷|冊|册|集|話|话|部)?\s*[）)］\]]\s*$/u, '')
      .replace(/\s*第?\s*\d+\s*(?:巻|卷|冊|册|集|話|话|部)\s*$/u, '')
      .replace(/\s+\d+\s*$/u, '')
      .replace(/\s+/g, ' ')
  );

  return normalized || fallback;
}`;
if (!pattern.test(shared)) throw new Error('extractOriginalTitleText block not found');
shared = shared.replace(pattern, replacement);
fs.writeFileSync(sharedFile, shared, 'utf8');
console.log(sharedFile);

const fs = require('fs');
const path = require('path');
const root = 'Z:/script';
const extDir = fs.readdirSync(root, { withFileTypes: true })
  .find(entry => entry.isDirectory() && entry.name.includes('商品登録拡張機能'));
if (!extDir) throw new Error('extension directory not found');

const contentFile = path.join(root, extDir.name, 'content', 'taiwan', 'content.js');
let content = fs.readFileSync(contentFile, 'utf8');
const oldContent = String.raw`  const stripTrailingVolumeSuffix = rawText => toSingleLine(rawText)
    .replace(/\s*[（(［\[]\s*(?:第\s*)?\d+\s*(?:巻|卷|冊|册|集|話|话|部)?\s*[）)］\]]\s*$/u, '')
    .replace(/\s*第?\s*\d+\s*(?:巻|卷|冊|册|集|話|话|部)\s*$/u, '')
    .replace(/\s+\d+\s*$/u, '')
    .trim();`;
const newContent = String.raw`  const stripTrailingEditionSuffix = rawText => toSingleLine(rawText)
    .replace(/\s*(?:首刷限定|初回限定|限定|特裝|特装|特仕|豪華|豪华|通常|普通)(?:版|版本)?\s*$/u, '')
    .trim();

  const stripTrailingVolumeSuffix = rawText => stripTrailingEditionSuffix(toSingleLine(rawText))
    .replace(/\s*[（(［\[]\s*(?:第\s*)?\d+\s*(?:巻|卷|冊|册|集|話|话|部)?\s*[）)］\]]\s*$/u, '')
    .replace(/\s*第?\s*\d+\s*(?:巻|卷|冊|册|集|話|话|部)\s*$/u, '')
    .replace(/\s+\d+\s*$/u, '')
    .trim();`;
if (!content.includes(oldContent)) throw new Error('content strip block not found');
content = content.replace(oldContent, newContent);
fs.writeFileSync(contentFile, content, 'utf8');

const sharedFile = path.join(root, extDir.name, 'popup', 'taiwan', 'popup.shared.js');
let shared = fs.readFileSync(sharedFile, 'utf8');
const oldShared = String.raw`function extractOriginalTitleText(value) {
  const fallback = trimValue(value);
  if (!fallback) return '';

  const normalized = fallback
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
    .trim();

  return normalized || fallback;
}`;
const newShared = String.raw`function extractOriginalTitleText(value) {
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
if (!shared.includes(oldShared)) throw new Error('shared original title block not found');
shared = shared.replace(oldShared, newShared);
fs.writeFileSync(sharedFile, shared, 'utf8');

console.log(contentFile);
console.log(sharedFile);

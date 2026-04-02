const fs = require('fs');
const path = require('path');
const root = 'Z:/script';
const extDir = fs.readdirSync(root, { withFileTypes: true })
  .find(entry => entry.isDirectory() && entry.name.includes('商品登録拡張機能'));
if (!extDir) throw new Error('extension directory not found');
const file = path.join(root, extDir.name, 'content', 'taiwan', 'content.js');
let text = fs.readFileSync(file, 'utf8');

const oldBlock = String.raw`    description = baseDescription;
    bonusInfoText = firstEditionBonus.text;
    formatSize = formatInfo.size;
    formatText = formatInfo.format;
    originalTitle = extractOriginalTitle(name);
    mainImageUrl = bookImages.main || fallbackImageUrl;
    additionalImageUrls = dedupeProductImageUrls([
      ...bookImages.gallery.slice(1),
      ...bookImages.detail,
      ...extendedContent.images,
      ...firstEditionBonus.images,
    ], mainImageUrl);`;

const newBlock = String.raw`    description = mergeTextBlocks([
      baseDescription,
      extendedContent.text,
      firstEditionBonus.text,
    ]);
    bonusInfoText = firstEditionBonus.text;
    formatSize = formatInfo.size;
    formatText = formatInfo.format;
    originalTitle = extractOriginalTitle(name);
    mainImageUrl = bookImages.main || fallbackImageUrl;
    additionalImageUrls = dedupeProductImageUrls([
      ...bookImages.gallery.slice(1),
      ...bookImages.detail,
      ...extendedContent.images,
      ...firstEditionBonus.images,
    ], mainImageUrl);`;

if (!text.includes(oldBlock)) throw new Error('book description block not found');
text = text.replace(oldBlock, newBlock);
fs.writeFileSync(file, text, 'utf8');
console.log(file);

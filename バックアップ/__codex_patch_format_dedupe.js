const fs = require('fs');
const path = require('path');
const root = 'Z:/script';
const extDir = fs.readdirSync(root, { withFileTypes: true })
  .find(entry => entry.isDirectory() && entry.name.includes('商品登録拡張機能'));
if (!extDir) throw new Error('extension directory not found');
const file = path.join(root, extDir.name, 'content', 'taiwan', 'content.js');
let text = fs.readFileSync(file, 'utf8');

const oldBlock = String.raw`  const mergeFormatParts = parts => {
    const merged = [];
    for (const part of (parts || []).map(toSingleLine).filter(Boolean)) {
      const normalizedPart = part.replace(/\s+/g, '').toLowerCase();
      if (!normalizedPart) continue;
      const duplicated = merged.some(existing => existing.replace(/\s+/g, '').toLowerCase() === normalizedPart);
      if (!duplicated) merged.push(part);
    }
    return merged.join(' / ');
  };`;

const newBlock = String.raw`  const mergeFormatParts = parts => {
    const merged = [];
    const seen = new Set();
    const pushSegment = rawSegment => {
      const segment = toSingleLine(rawSegment)
        .replace(/^規格\s*[：:]\s*/u, '')
        .trim();
      if (!segment) return;

      const normalizedSegment = segment.replace(/\s+/g, '').toLowerCase();
      if (!normalizedSegment || seen.has(normalizedSegment)) return;
      seen.add(normalizedSegment);
      merged.push(segment);
    };

    for (const part of (parts || []).map(toSingleLine).filter(Boolean)) {
      part.split(/\s*\/\s*/).forEach(pushSegment);
    }

    return merged.join(' / ');
  };`;

if (!text.includes(oldBlock)) throw new Error('mergeFormatParts block not found');
text = text.replace(oldBlock, newBlock);
fs.writeFileSync(file, text, 'utf8');
console.log(file);

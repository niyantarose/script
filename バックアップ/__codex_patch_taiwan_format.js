const fs = require('fs');
const path = require('path');
const root = 'Z:/script';
const extDir = fs.readdirSync(root, { withFileTypes: true })
  .find(entry => entry.isDirectory() && entry.name.includes('商品登録拡張機能'));
if (!extDir) throw new Error('extension directory not found');
const file = path.join(root, extDir.name, 'content', 'taiwan', 'content.js');
let text = fs.readFileSync(file, 'utf8');

const pattern = /  const getFormatInfo = \(\) => \{[\s\S]*?  \};\r?\n\r?\n  const buildBookDescriptionSupplement =/;
if (!pattern.test(text)) throw new Error('getFormatInfo block not found');

const replacement = String.raw`  const extractDetailValueFromText = (rawText, labels) => {
    const text = toMultiline(rawText);
    if (!text) return '';

    for (const label of labels) {
      const matched = text.match(new RegExp(
        \\`${escapeRegExp(label)}\\s*[：:]\\s*([^\\n]+)\\`,
        'i'
      ));
      if (matched?.[1]) {
        return cleanupLabeledValue(matched[1]);
      }
    }

    return '';
  };

  const mergeFormatParts = parts => {
    const merged = [];
    for (const part of (parts || []).map(toSingleLine).filter(Boolean)) {
      const normalizedPart = part.replace(/\s+/g, '').toLowerCase();
      if (!normalizedPart) continue;
      const duplicated = merged.some(existing => existing.replace(/\s+/g, '').toLowerCase() === normalizedPart);
      if (!duplicated) merged.push(part);
    }
    return merged.join(' / ');
  };

  const getFormatInfo = () => {
    const detailBlock = Array.from(document.querySelectorAll('.mod_b, .mod, .prod_cont')).find(mod =>
      /詳細資料|详细资料/.test(getText(mod.querySelector('h3, h2, h1')))
    );
    const detailText = toMultiline(getText(detailBlock));
    const formatLabel = toSingleLine(
      getLabeledValue(['規格', '规格']) ||
      extractDetailValueFromText(detailText, ['規格', '规格'])
    );
    const binding = toSingleLine(
      getLabeledValue(['裝訂方式', '装订方式', '裝訂', '装订']) ||
      extractDetailValueFromText(detailText, ['裝訂方式', '装订方式', '裝訂', '装订'])
    );
    const pages = toSingleLine(
      getLabeledValue(['頁數', '页数']) ||
      extractDetailValueFromText(detailText, ['頁數', '页数'])
    );
    const size = parseSizeFromText(formatLabel || detailText);
    const format = mergeFormatParts([formatLabel, binding, pages, size]);

    return {
      format,
      size,
    };
  };

  const buildBookDescriptionSupplement =`;

text = text.replace(pattern, replacement);
fs.writeFileSync(file, text, 'utf8');
console.log(file);

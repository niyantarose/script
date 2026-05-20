function trimValue(value) {
  return String(value ?? '').trim();
}

function cleanupOriginalTitleTextOnce(value) {
  return value
    .replace(/^[台臺][湾灣]版\s*/u, '')
    // 1. Bracket noise with edition/bonus keywords (added 獨家/独家/博客來獨家)
    .replace(/\s*【[^】]*(?:首刷|初回|初版|限定|特裝|特装|特別|特别|通常|普通|獨家|独家|博客來獨家|博客来独家)[^】]*】\s*/gu, ' ')
    .replace(/\s*\([^)]*(?:首刷|初回|初版|限定|特裝|特装|特別|特别|通常|普通|獨家|独家|博客來獨家|博客来独家)[^)]*\)\s*/gu, ' ')
    .replace(/\s*\[[^\]]*(?:首刷|初回|初版|限定|特裝|特装|特別|特别|通常|普通|獨家|独家|博客來獨家|博客来独家)[^\]]*\]\s*/gu, ' ')
    .replace(/\s*（[^）]*(?:首刷|初回|初版|限定|特裝|特装|特別|特别|通常|普通|獨家|独家|博客來獨家|博客来独家)[^）]*）\s*/gu, ' ')
    // 2. Bracket with complete/exclusive markers at the very end
    .replace(/\s*[（(\[【]\s*(?:完|全|獨家|独家|博客來獨家|博客来独家|套書|套书|合集)\s*[）)\]】]\s*$/giu, ' ')
    // 3. Volume number followed by complete/exclusive bracket at the very end, e.g. " 3 (完)"
    .replace(/\s*(?:第?\s*)?\d+\s*(?:巻|卷|冊|册|集|話|话|部)?\s*[（(\[【]\s*(?:完|全|獨家|独家|博客來獨家|博客来独家|套書|套书|合集)\s*[）)\]】]\s*$/giu, ' ')
    // 4. Standard bracketed volume numbers at the end
    .replace(/\s*[\(（【\[]\s*(?:第?\s*)?\d+\s*(?:巻|卷|冊|册|集|話|话|部)?\s*[\)）】\]]\s*$/u, ' ')
    .replace(/\s*[-/／]\s*(?:首刷|初回|初版|限定|特裝|特装|特別|特别|通常|普通|獨家|独家)[^ ]*\s*/gu, ' ')
    .replace(/\s*(?:首刷(?:限定)?|初回(?:限定)?|初版(?:限定)?|限定|特裝|特装|特別|特别|通常|普通|獨家|独家)(?:版)?\s*$/u, '')
    .replace(/\s*第?\s*\d+\s*(?:巻|卷|冊|册|集|話|话|部)\s*$/u, '')
    .replace(/\s*\d+\s*$/u, '')
    .replace(/^《\s*([^》]+?)\s*》$/u, '$1')
    .replace(/^「\s*([^」]+?)\s*」$/u, '$1')
    .replace(/^『\s*([^』]+?)\s*』$/u, '$1')
    .replace(/\s+/g, ' ')
    .trim();
}

function extractOriginalTitleText(value) {
  const fallback = trimValue(value);
  if (!fallback) return '';

  let normalized = fallback;
  for (let i = 0; i < 4; i += 1) {
    const next = cleanupOriginalTitleTextOnce(normalized);
    if (next === normalized) break;
    normalized = next;
  }

  return normalized || fallback;
}

function stripVolumeNoiseForSheetOriginal(value) {
  let text = trimValue(value);
  if (!text) return '';
  text = text
    .replace(/\s*(?:vol\.?|v\.?|第)\s*[0-9０-９]{1,4}\s*(?:巻|卷|集|冊|話|回|期|号|號)?\s*$/iu, '')
    .replace(/\s*[#＃]?\s*[0-9０-９]{1,4}\s*$/u, '')
    .replace(/\s*(?:第)?\s*[0-9０-９]{1,4}\s*(?:漫畫|漫画|コミック|單行本|单行本)\s*$/iu, '')
    .replace(/\s*(?:漫畫|漫画|コミック|單行本|单行本)\s*(?:第)?\s*[0-9０-９]{1,4}\s*$/iu, '')
    .replace(/\s*(?:漫畫|漫画|コミック|單行本|单行本)\s*$/iu, '')
    .replace(/\s+/g, ' ')
    .trim();
  return text;
}

function finalizeSheetOriginalWorkTitleRow(value) {
  const extracted = extractOriginalTitleText(value);
  const stripped = stripVolumeNoiseForSheetOriginal(extracted);
  return trimValue(stripped || extracted || value);
}

const titles = [
  '全球高考 3 (完)',
  '全球高考 (博客來獨家)',
  '全球高考3(完)',
  '全球高考(博客來獨家)'
];

for (const t of titles) {
  console.log(`Original: "${t}"`);
  console.log(`  extracted: "${extractOriginalTitleText(t)}"`);
  console.log(`  finalize: "${finalizeSheetOriginalWorkTitleRow(t)}"`);
  console.log('------------------');
}

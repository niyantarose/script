// カテゴリ入力値を補正_ のガード検証（韓国アラジン取込の主経路）。
// シートのカテゴリ列はプルダウン（入力規則）なので、分類できなかったときに
// アラジンAPIの categoryName（"국내도서>만화>BL만화" のような '>' 区切りパス）を
// そのまま書き込むと入力規則エラーになる。台湾側と同じく「分類不能なら空欄」に揃える。
// 実行: node tests/category-guard.test.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ctx = { console };
ctx.globalThis = ctx;
ctx.SpreadsheetApp = { getActive: () => null, getUi: () => null };
ctx.LockService = {};
ctx.PropertiesService = {};
ctx.Logger = { log: () => {} };
ctx.Utilities = {};
ctx.CacheService = {};
ctx.UrlFetchApp = {};
ctx.ContentService = {};
ctx.XmlService = {};
vm.createContext(ctx);
vm.runInContext(
  fs.readFileSync(path.join(__dirname, '..', 'アラジンAPI取得(アラジン.gs).js'), 'utf8'),
  ctx,
  { filename: 'アラジンAPI取得(アラジン.gs).js' }
);

let failed = 0;
function eq(name, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a !== e) { failed++; console.error(`[NG] ${name}: expected=${e} actual=${a}`); }
  else { console.log(`[OK] ${name}`); }
}

const fix = (sheetName, source) => ctx.カテゴリ入力値を補正_(sheetName, source);

// === 正常系: 従来どおり分類できる場合はそのラベルを返す ===
eq('韓国書籍: 에세이 → エッセイ',
   fix('韓国書籍', { categoryName: '국내도서>에세이' }), 'エッセイ');
eq('韓国マンガ: 만화 → まんが',
   fix('韓国マンガ', { categoryName: '국내도서>만화>BL만화' }), 'まんが');
eq('韓国音楽映像: 음반 → CD',
   fix('韓国音楽映像', { categoryName: '음반>가요>댄스뮤직' }), 'CD');

// === 判定済みカテゴリ（拡張機能が渡す sheetCategory）は従来どおり優先 ===
eq('sheetCategory が有効値なら優先',
   fix('韓国書籍', { sheetCategory: '参考書', categoryName: '국내도서>수험서' }), '参考書');
eq('sheetCategory がシートの有効値外なら無視して再判定',
   fix('韓国音楽映像', { sheetCategory: 'まんが', categoryName: '음반>가요' }), 'CD');

// === 本題: 分類できないときに生のパス文字列を書き込まない ===
eq('韓国音楽映像: 分類不能な生パスは空欄',
   fix('韓国音楽映像', { categoryName: '국내도서>인문학>철학' }), '');
eq('韓国書籍: 分類不能な生パスは空欄',
   fix('韓国書籍', { categoryName: '국내도서>가정/요리/뷰티>반려동물' }), '');
eq('未知シート名でも生パスは書き込まない',
   fix('その他シート', { categoryName: '국내도서>만화>BL만화' }), '');

// === 空・未指定は空欄 ===
eq('categoryName 空なら空欄', fix('韓国書籍', { categoryName: '' }), '');
eq('source 未指定でも落ちない', fix('韓国書籍', null), '');

process.exitCode = failed ? 1 : 0;
console.log(failed ? `\n${failed} failed` : '\nall passed');

/**
 * 粗利益率列追加.gs
 * 実質粗利益率 = (売価 - 原価×レート×1.3 - 売価×0.22) / 売価
 *
 * コスト内訳:
 *   送料     = 原価×30%（原価×レート×0.3）
 *   Yahoo    = 売価×9%
 *   カード   = 売価×3%
 *   消費税   = 売価×10%
 *   合計販売コスト = 売価×22%
 *
 * 為替レートマスター:
 *   B2=TWD, B3=CN, B4=KR, B5=HK, B6=TH, B7=USD
 */

const 粗利益率_対象シート = [
  { シート名: '韓国音楽映像',                     売価列名: '売価', 原価列名: '原価', 通貨: 'KR',  言語列名: null   },
  { シート名: '韓国書籍',                         売価列名: '売価', 原価列名: '原価', 通貨: 'KR',  言語列名: null   },
  { シート名: '韓国マンガ',                       売価列名: '売価', 原価列名: '原価', 通貨: 'KR',  言語列名: null   },
  { シート名: '韓国グッズ',                       売価列名: '売価', 原価列名: '原価', 通貨: 'KR',  言語列名: null   },
  { シート名: '台湾グッズ',                       売価列名: '売価', 原価列名: '原価', 通貨: null,  言語列名: '言語' },
  { シート名: '台湾_書籍（コミック/小説/設定集）', 売価列名: '売価', 原価列名: '原価', 通貨: null,  言語列名: '言語' },
];

// 日本語・英語コード両対応
const レートセル = {
  'TW':     '為替レートマスター!$B$2',
  '台湾':   '為替レートマスター!$B$2',
  'CN':     '為替レートマスター!$B$3',
  '中国':   '為替レートマスター!$B$3',
  'KR':     '為替レートマスター!$B$4',
  '韓国':   '為替レートマスター!$B$4',
  'HK':     '為替レートマスター!$B$5',
  'TH':     '為替レートマスター!$B$6',
  'タイ':   '為替レートマスター!$B$6',
  'USD':    '為替レートマスター!$B$7',
  'アメリカ': '為替レートマスター!$B$7',
  '英語':   '為替レートマスター!$B$7',
};

const 販売コスト率 = 0.22;  // Yahoo9% + カード3% + 消費税10%
const 送料率      = 1.3;    // 原価 × 1.3（原価＋送料30%）

function 全シートに粗利益率列を追加() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const 結果 = [];
  for (const 設定 of 粗利益率_対象シート) {
    const sh = ss.getSheetByName(設定.シート名);
    if (!sh) { 結果.push(`${設定.シート名}: シートなし`); continue; }
    結果.push(`${設定.シート名}: ${粗利益率列を追加_(sh, 設定)}`);
  }
  ui.alert('粗利益率列追加結果\n\n' + 結果.join('\n'));
}

function アクティブシートに粗利益率列を追加() {
  const sh = SpreadsheetApp.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const 設定 = 粗利益率_対象シート.find(s => s.シート名 === sh.getName());
  if (!設定) { ui.alert(`「${sh.getName()}」は対象シートに含まれていません`); return; }
  ui.alert(`${sh.getName()}: ${粗利益率列を追加_(sh, 設定)}`);
}

function 粗利益率列を追加_(sh, 設定) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const 列 = {};
  headers.forEach((h, i) => { if (h) 列[String(h).trim()] = i + 1; });

  const 売価Col = 列[設定.売価列名];
  const 原価Col = 列[設定.原価列名];
  if (!売価Col) return `「${設定.売価列名}」列が見つかりません`;
  if (!原価Col) return `「${設定.原価列名}」列が見つかりません`;

  const 言語Col = 設定.言語列名 ? 列[設定.言語列名] : null;
  if (設定.言語列名 && !言語Col) return `「${設定.言語列名}」列が見つかりません`;

  if (列['粗利益率']) {
    粗利益率数式と書式を設定_(sh, 列['粗利益率'], 売価Col, 原価Col, 言語Col, 設定.通貨);
    return '既存列を更新しました';
  }

  sh.insertColumnAfter(原価Col);
  const 挿入列 = 原価Col + 1;
  const hCell = sh.getRange(1, 挿入列);
  hCell.setValue('粗利益率');
  hCell.setBackground('#cc0000').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  粗利益率数式と書式を設定_(sh, 挿入列, 売価Col, 原価Col, 言語Col, 設定.通貨);
  return '✅ 追加しました';
}

function 粗利益率数式と書式を設定_(sh, 粗利益率Col, 売価Col, 原価Col, 言語Col, 固定通貨) {
  const lastRow = Math.max(sh.getLastRow(), 100);
  if (lastRow < 2) return;

  const 売 = 列番号を文字に_(売価Col);
  const 原 = 列番号を文字に_(原価Col);
  const 言 = 言語Col ? 列番号を文字に_(言語Col) : null;

  const formulas = [];
  for (let r = 2; r <= lastRow; r++) {
    let レート式;
    if (言) {
      const IFS = Object.entries(レートセル)
        .map(([通貨, セル]) => `${言}${r}="${通貨}",${セル}`)
        .join(',');
      レート式 = `IFS(${IFS},TRUE,1)`;
    } else {
      レート式 = レートセル[固定通貨] || '1';
    }

    // 実質粗利益率 = (売価 - 原価×レート×1.3 - 売価×0.22) / 売価
    const 数式 = `=IFERROR(IF(${売}${r}="","",IF(${売}${r}=0,"エラー",(${売}${r}-${原}${r}*${レート式}*${送料率}-${売}${r}*${販売コスト率})/${売}${r})),"エラー")`;
    formulas.push([数式]);
  }

  sh.getRange(2, 粗利益率Col, lastRow - 1, 1).setValues(formulas);
  sh.getRange(2, 粗利益率Col, lastRow - 1, 1).setNumberFormat('0.0%');
  sh.setColumnWidth(粗利益率Col, 80);

  const range = sh.getRange(2, 粗利益率Col, lastRow - 1, 1);
  const 新ルール = sh.getConditionalFormatRules().filter(rule =>
    !rule.getRanges().some(r => r.getColumn() === 粗利益率Col && r.getNumColumns() === 1)
  );

  新ルール.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0).setBackground('#ea4335').setFontColor('#ffffff').setRanges([range]).build());
  新ルール.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0, 0.0999).setBackground('#ff6d00').setFontColor('#ffffff').setRanges([range]).build());
  新ルール.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0.10, 0.1999).setBackground('#ffd600').setFontColor('#000000').setRanges([range]).build());

  sh.setConditionalFormatRules(新ルール);
}

function 列番号を文字に_(col) {
  let s = '';
  while (col > 0) {
    const r = (col - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}
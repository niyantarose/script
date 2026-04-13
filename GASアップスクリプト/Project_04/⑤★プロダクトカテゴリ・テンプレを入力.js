/**
 * プロダクトカテゴリを入力
 * v2: getUi() 安全化
 */

// ====================================================================
// ★ 安全なUI通知（UIが無くても落ちない）
// ====================================================================

function CAT_uiSafeAlert_(msg) {
  try {
    SpreadsheetApp.getUi().alert(String(msg));
    return;
  } catch (e) {}

  // フォールバック: toast → log
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(String(msg), 'プロダクトカテゴリ', 10);
  } catch (e2) {}

  Logger.log('[CAT_UI_FALLBACK] ' + msg);
}

function CAT_uiSafeToast_(msg, title, sec) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(String(msg), title || 'INFO', sec || 5);
  } catch (e) {
    Logger.log('[CAT_TOAST_FALLBACK] ' + (title ? title + ': ' : '') + msg);
  }
}


function プロダクトカテゴリを入力() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('①商品入力シート');
  const categorySheet = ss.getSheetByName('プロダクトカテゴリ');
  
  if (!inputSheet) {
    CAT_uiSafeAlert_('エラー: ①商品入力シートが見つかりません');
    return;
  }
  if (!categorySheet) {
    CAT_uiSafeAlert_('エラー: プロダクトカテゴリシートが見つかりません');
    return;
  }

  // ーーーーーーーーーーーーーーーーーー
  // ✨ テンプレ名を半角英数字に整形する関数
  // ーーーーーーーーーーーーーーーーーー
  function normalizeTemplateName(str) {
    if (!str) return '';

    // 全角 → 半角
    const hankaku = str.replace(/[！-～]/g, s =>
      String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
    );

    // 半角英数字以外を削除し、trim
    return hankaku.replace(/[^0-9A-Za-z]/g, '').trim();
  }

  const countryNames = ['韓国語の','韓国の作家','韓国の','韓国','台湾','タイ','中国','海外','日本','アメリカ','フランス','イギリス'];

  const categoryLastRow = categorySheet.getLastRow();
  const categoryData = categorySheet.getRange('A2:C' + categoryLastRow).getValues();

  // ★ 2つのマップを作成：国名付き（原本）と国名なし
  const categoryMapOriginal = {};  // 国名付きのまま（台湾小説 など）
  const categoryMapNoCountry = {}; // 国名削除後（小説 など）

  categoryData.forEach(row => {
    const rawCategoryName = row[0];
    const rawTemplateName = row[1];
    const rawProductCategory = row[2];

    if (rawCategoryName) {
      // 原本（空白除去のみ）
      const originalName = rawCategoryName.toString().replace(/\s+/g, '').trim();
      
      if (originalName) {
        const entry = {
          template: normalizeTemplateName(rawTemplateName),
          code: (rawProductCategory || '').toString().trim()
        };
        
        // 国名付きマップに登録
        categoryMapOriginal[originalName] = entry;
        
        // 国名削除版も作成
        let nameNoCountry = originalName;
        countryNames.forEach(country => {
          nameNoCountry = nameNoCountry.replace(new RegExp(country, 'g'), '');
        });
        nameNoCountry = nameNoCountry.trim();
        
        // 国名削除後に残った名前があれば登録（重複時は長い方優先のため後で判定）
        if (nameNoCountry && nameNoCountry !== originalName) {
          // 既存エントリがなければ、または既存より元カテゴリ名が長ければ上書き
          if (!categoryMapNoCountry[nameNoCountry] || 
              originalName.length > (categoryMapNoCountry[nameNoCountry].originalLength || 0)) {
            categoryMapNoCountry[nameNoCountry] = {
              ...entry,
              originalLength: originalName.length
            };
          }
        }
      }
    }
  });

  const lastRow = inputSheet.getLastRow();
  const startRow = 3;
  const productNames = inputSheet.getRange('B' + startRow + ':B' + lastRow).getValues();

  const resultArray = [];

  productNames.forEach((row, index) => {
    const productName = row[0];
    let foundTemplate = '';
    let foundCategoryCode = '';
    let longestMatch = '';

    if (productName) {
      const productNameClean = productName.toString().replace(/\s+/g, '');
      
      // ★ 優先順位1: 国名付きのままでマッチング（台湾小説 → 台湾小説）
      for (let category in categoryMapOriginal) {
        if (productNameClean.includes(category)) {
          if (category.length > longestMatch.length) {
            longestMatch = category;
            foundTemplate = categoryMapOriginal[category].template || '';
            foundCategoryCode = categoryMapOriginal[category].code || '';
          }
        }
      }
      
      // ★ 優先順位2: 国名付きでマッチしなかった場合のみ、国名削除版でマッチング
      if (!longestMatch) {
        let productNameNoCountry = productNameClean;
        countryNames.forEach(country => {
          productNameNoCountry = productNameNoCountry.replace(new RegExp(country, 'g'), '');
        });

        for (let category in categoryMapNoCountry) {
          if (productNameNoCountry.includes(category)) {
            if (category.length > longestMatch.length) {
              longestMatch = category;
              foundTemplate = categoryMapNoCountry[category].template || '';
              foundCategoryCode = categoryMapNoCountry[category].code || '';
            }
          }
        }
      }
    }

    resultArray.push([
      normalizeTemplateName(foundTemplate),
      foundCategoryCode || ''
    ]);
  });

  inputSheet.getRange('M' + startRow + ':N' + lastRow).setValues(resultArray);

  CAT_uiSafeToast_(
    'テンプレ名（M列）とプロダクトカテゴリ（N列）を反映しました。',
    '完了',
    5
  );
}
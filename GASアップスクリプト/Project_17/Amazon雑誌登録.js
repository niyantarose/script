/**
 * Yahoo→Amazon完全自動化スクリプト（高速版 + 入力規則設定）
 * メニューから実行してください
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Yahoo→Amazon変換')
    .addItem('CSVファイルを変換', 'CSV変換開始')
    .addItem('テンプレートシートをコピー', 'テンプレートシートをコピー')
    .addItem('入力規則を設定', 'テンプレートシートに入力規則を設定')
    .addToUi();
}

function CSV変換開始() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          * { margin: 0; padding: 0; box-sizing: border-box; }
          body { font-family: 'Google Sans', Arial, sans-serif; padding: 32px; background: #f8f9fa; }
          .container { max-width: 600px; margin: 0 auto; background: white; border-radius: 12px; padding: 32px; box-shadow: 0 1px 3px rgba(0,0,0,0.12); }
          h1 { font-size: 24px; color: #202124; margin-bottom: 8px; }
          .subtitle { color: #5f6368; font-size: 14px; margin-bottom: 32px; }
          .upload-area { border: 2px dashed #dadce0; border-radius: 8px; padding: 48px 24px; text-align: center; background: #f8f9fa; cursor: pointer; transition: all 0.3s; margin-bottom: 24px; }
          .upload-area:hover { border-color: #1a73e8; background: #e8f0fe; }
          .upload-icon { font-size: 48px; margin-bottom: 16px; }
          .upload-text { font-size: 16px; color: #202124; margin-bottom: 8px; }
          .upload-hint { font-size: 12px; color: #5f6368; }
          input[type="file"] { display: none; }
          .file-info { background: #e8f0fe; border-radius: 8px; padding: 16px; margin-bottom: 24px; display: none; }
          .file-info.show { display: block; }
          .file-name { font-size: 14px; color: #1967d2; font-weight: 500; margin-bottom: 4px; }
          .file-size { font-size: 12px; color: #5f6368; }
          .button { width: 100%; background: #1a73e8; color: white; border: none; border-radius: 4px; padding: 12px 24px; font-size: 14px; font-weight: 500; cursor: pointer; }
          .button:hover:not(:disabled) { background: #1557b0; }
          .button:disabled { background: #dadce0; cursor: not-allowed; }
          .status { margin-top: 24px; padding: 16px; border-radius: 8px; display: none; font-size: 14px; line-height: 1.6; }
          .status.show { display: block; }
          .status.processing { background: #fef7e0; color: #7c4d00; border-left: 4px solid #f9ab00; }
          .status.success { background: #e6f4ea; color: #137333; border-left: 4px solid #34a853; }
          .status.error { background: #fce8e6; color: #c5221f; border-left: 4px solid #ea4335; }
          .progress { margin-top: 16px; height: 4px; background: #dadce0; border-radius: 2px; overflow: hidden; display: none; }
          .progress.show { display: block; }
          .progress-bar { height: 100%; background: #1a73e8; width: 0%; animation: progress 2s ease-in-out infinite; }
          @keyframes progress { 0% { width: 0%; } 50% { width: 70%; } 100% { width: 100%; } }
        </style>
      </head>
      <body>
        <div class="container">
          <h1>📊 Yahoo→Amazon CSV変換（高速版）</h1>
          <p class="subtitle">1000行でも高速処理！</p>
          
          <div class="upload-area" id="uploadArea" onclick="document.getElementById('fileInput').click()">
            <div class="upload-icon">📁</div>
            <div class="upload-text">CSVファイルを選択</div>
            <div class="upload-hint">1000行以上でもOK</div>
          </div>
          
          <input type="file" id="fileInput" accept=".csv" onchange="ファイル選択(event)">
          
          <div class="file-info" id="fileInfo">
            <div class="file-name" id="fileName"></div>
            <div class="file-size" id="fileSize"></div>
          </div>
          
          <button class="button" id="convertBtn" onclick="変換開始()" disabled>変換開始</button>
          
          <div class="progress" id="progress"><div class="progress-bar"></div></div>
          <div class="status" id="status"></div>
        </div>
        
        <script>
          let selectedFile = null;
          
          function ファイル選択(event) {
            const file = event.target.files[0];
            if (!file || !file.name.endsWith('.csv')) { alert('CSVファイルを選択してください'); return; }
            selectedFile = file;
            document.getElementById('fileName').textContent = '📄 ' + file.name;
            document.getElementById('fileSize').textContent = 'サイズ: ' + (file.size / 1024).toFixed(2) + ' KB';
            document.getElementById('fileInfo').classList.add('show');
            document.getElementById('convertBtn').disabled = false;
          }
          
          const uploadArea = document.getElementById('uploadArea');
          uploadArea.addEventListener('dragover', (e) => { e.preventDefault(); uploadArea.style.background = '#e8f0fe'; });
          uploadArea.addEventListener('dragleave', () => { uploadArea.style.background = '#f8f9fa'; });
          uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.style.background = '#f8f9fa';
            const file = e.dataTransfer.files[0];
            if (file && file.name.endsWith('.csv')) {
              document.getElementById('fileInput').files = e.dataTransfer.files;
              ファイル選択({ target: { files: [file] } });
            }
          });
          
          function 変換開始() {
            if (!selectedFile) { alert('ファイルを選択してください'); return; }
            
            const statusDiv = document.getElementById('status');
            const convertBtn = document.getElementById('convertBtn');
            const progress = document.getElementById('progress');
            
            statusDiv.className = 'status processing show';
            statusDiv.innerHTML = '⏳ ファイルを読み込み中...';
            convertBtn.disabled = true;
            progress.classList.add('show');
            
            const reader = new FileReader();
            reader.onload = function(e) {
              const csvContent = e.target.result;
              if (!csvContent || csvContent.trim().length === 0) {
                statusDiv.className = 'status error show';
                statusDiv.innerHTML = '❌ ファイルが空です';
                convertBtn.disabled = false;
                progress.classList.remove('show');
                return;
              }
              
              statusDiv.innerHTML = '⏳ データ変換中...（高速処理モード）';
              
              google.script.run
                .withSuccessHandler(function(result) {
                  progress.classList.remove('show');
                  statusDiv.className = 'status success show';
                  statusDiv.innerHTML = 
                    '<strong>✅ 変換完了！</strong><br><br>' +
                    '📊 処理行数：<strong>' + result.rowsProcessed + '行</strong><br>' +
                    '⚡ 処理時間：<strong>' + result.processingTime + '秒</strong><br><br>' +
                    '✓ サフィックス判断完了<br>' +
                    '✓ ページ数・発売日整形完了<br>' +
                    '✓ S列→AW列コピー完了<br><br>' +
                    '<small>このウィンドウを閉じてシートを確認してください</small>';
                })
                .withFailureHandler(function(error) {
                  progress.classList.remove('show');
                  statusDiv.className = 'status error show';
                  statusDiv.innerHTML = '<strong>❌ エラー</strong><br><br>' + error.message;
                  convertBtn.disabled = false;
                })
                .YahooからAmazonへ完全自動変換_高速版(csvContent);
            };
            reader.readAsText(selectedFile, 'Shift-JIS');
          }
        </script>
      </body>
    </html>
  `).setWidth(700).setHeight(600);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Yahoo CSV変換');
}

/**
 * 高速版メイン処理（配列で一括処理）
 */
function YahooからAmazonへ完全自動変換_高速版(csvContent) {
  const startTime = new Date();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // アクティブシートが存在するか確認
    if (!ss) {
      throw new Error('スプレッドシートが見つかりません');
    }
    
    const currentSheet = ss.getActiveSheet();
    
    if (!currentSheet) {
      throw new Error('アクティブシートが見つかりません');
    }
    
    Logger.log('処理開始: シート名 = ' + currentSheet.getName());
    Logger.log('ステップ1: CSV読み込み');
    
    const dataRowCount = csvData.length - 1;
    const startRow = 8;
    
    // ヘッダー設定
    ヘッダー設定(currentSheet);
    
    // 最大列数を計算（JP列 = 296列）
    const maxCol = 296;
    
    // 全データを格納する2次元配列を作成
    const outputData = [];
    
    Logger.log('ステップ2-7: データ処理（配列）');
    
    for (let i = 1; i < csvData.length; i++) {
      const row = csvData[i];
      if (!row || row.length === 0) continue;
      
      // 296列の空配列を作成
      const outputRow = new Array(maxCol).fill('');
      
      // C列 → A列 (SKU)
      if (row[2]) outputRow[0] = row[2];
      
      // B列 → D列 (商品名)
      if (row[1]) outputRow[3] = row[1];
      
      // F列 → S列 (価格)
      if (row[5]) {
        const price = 価格データクリーン(row[5]);
        outputRow[18] = price;
        outputRow[48] = price; // AW列にもコピー
      }
      
      // M列 → BW列 (商品説明)
      let description = row[12] || '';
      if (description) {
        const originalText = description;
        let processedText = description.replace(/<br>/gi, '\n').replace(/<br\/>/gi, '\n').replace(/<br \/>/gi, '\n');
        
        // 発売日抽出 → DC列
        const releaseDate = 発売日抽出(processedText);
        if (releaseDate) outputRow[106] = releaseDate;
        
        // ページ数抽出 → CF列
        let pageNum = ページ数抽出(processedText);
        outputRow[83] = pageNum || '200';
        
        // サイズ抽出
        const size = サイズ抽出(processedText);
        if (size) {
          const {width, depth} = サイズ分割(size);
          if (width) {
            outputRow[288] = width; // JI列
            outputRow[289] = 'ミリメートル'; // JJ列
          }
          if (depth) {
            outputRow[290] = depth; // JK列
            outputRow[291] = 'ミリメートル'; // JL列
          }
        }
        
        // 固定サイズ・重量
        outputRow[292] = '2.5'; // JM列
        outputRow[293] = 'センチメートル'; // JN列
        outputRow[294] = '0.95'; // JO列
        outputRow[295] = 'グラム'; // JP列
        
        // 説明文整形
        if (originalText.indexOf('<br>') === -1 && originalText.indexOf('\n') === -1) {
          description = originalText.replace(/\r\n|\n|\r/g, ' ');
        } else {
          description = originalText;
        }
        outputRow[74] = description; // BW列
      }
      
      // CL列 → IY列 (画像URL)
      if (row[89]) {
        const urlText = row[89];
        if (urlText.indexOf('https://') > -1) {
          const urls = URL抽出(urlText);
          const cleanedUrls = urls.map(url => url.replace(/\.(jpg|JPG|jpeg|JPEG)$/g, '').trim());
          if (cleanedUrls[0]) outputRow[260] = cleanedUrls[0]; // IY列
          if (cleanedUrls[1]) outputRow[261] = cleanedUrls[1]; // IZ列
          if (cleanedUrls[2]) outputRow[262] = cleanedUrls[2]; // JA列
        }
      }
      
      // BG列 → BV列 (配送グループ)
      if (row[58]) outputRow[73] = row[58];
      
      // サフィックス判断
      const skuValue = String(outputRow[0]).trim();
      if (skuValue) {
        const suffixValue = サフィックス値取得(skuValue);
        if (suffixValue > 0) {
          outputRow[84] = suffixValue;  // CG列
          outputRow[122] = suffixValue; // DS列
          outputRow[123] = suffixValue; // DT列
          outputRow[157] = suffixValue; // FB列
        }
      }
      
      // タイトル抽出
      const titleText = String(outputRow[3]).trim();
      if (titleText) {
        const extractedTitle = タイトル抽出(titleText);
        if (extractedTitle) outputRow[121] = extractedTitle; // DR列
        
        const bookTitle = 書籍タイトル抽出(titleText);
        if (bookTitle) {
          outputRow[158] = bookTitle; // FC列
          outputRow[77] = bookTitle;  // BZ列
        }
      }
      
      // 固定値設定
      outputRow[2] = 'ABIS_BOOK';        // C列
      outputRow[4] = 'ABIS_BOOK';        // E列
      outputRow[12] = 'WWD Korea';       // M列
      outputRow[5] = 'GTIN免除';         // F列
      outputRow[16] = '新品';            // Q列
      outputRow[43] = 'DEFAULT';         // AR列
      outputRow[44] = 0;                 // AS列
      outputRow[73] = '200円でゆうパケット'; // BV列
      
      outputData.push(outputRow);
    }
    
    Logger.log('ステップ8: シートに一括書き込み');
    
    // 一括書き込み（超高速）
    if (outputData.length > 0) {
      const writeRange = currentSheet.getRange(startRow, 1, outputData.length, maxCol);
      writeRange.setValues(outputData);
    }
    
    const endTime = new Date();
    const processingTime = ((endTime - startTime) / 1000).toFixed(1);
    
    Logger.log('✅ 完了！処理時間: ' + processingTime + '秒');
    
    return {
      success: true,
      rowsProcessed: outputData.length,
      processingTime: processingTime
    };
    
  } catch (error) {
    Logger.log('エラー: ' + error.toString());
    throw new Error(error.toString());
  }
}

/**
 * テンプレートシートをコピー（UI不要版）
 */
function テンプレートシートをコピー() {
  try {
    // コピー元のスプレッドシート
    const sourceSpreadsheetId = '1jD74RlK8g8F--IWggzRnjrcMJXhoNwzO';
    const sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
    const sourceSheet = sourceSpreadsheet.getSheetByName('テンプレート');
    
    if (!sourceSheet) {
      Logger.log('❌ 「テンプレート」シートが見つかりません');
      SpreadsheetApp.getUi().alert('エラー', '「テンプレート」シートが見つかりません', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // コピー先のスプレッドシート（現在のスプレッドシート）
    const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 既存の「テンプレート」シートを削除（存在する場合）
    const existingSheet = targetSpreadsheet.getSheetByName('テンプレート');
    if (existingSheet) {
      Logger.log('既存のテンプレートシートを削除中...');
      targetSpreadsheet.deleteSheet(existingSheet);
    }
    
    // シートをコピー
    Logger.log('シートをコピー中...');
    const copiedSheet = sourceSheet.copyTo(targetSpreadsheet);
    copiedSheet.setName('テンプレート');
    
    // 最初のシートとして移動
    targetSpreadsheet.setActiveSheet(copiedSheet);
    targetSpreadsheet.moveActiveSheet(1);
    
    Logger.log('✅ テンプレートシートのコピーが完了しました');
    
    SpreadsheetApp.getUi().alert(
      '完了',
      'テンプレートシートをコピーしました！',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert(
      'エラー',
      'シートのコピーに失敗しました:\n' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * テンプレートシートに入力規則を設定
 */
function テンプレートシートに入力規則を設定() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('テンプレート');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('「テンプレート」シートが見つかりません');
    return;
  }
  
  // データ開始行（8行目から）
  const startRow = 8;
  const lastRow = 1000; // 最大1000行分設定
  
  try {
    // C列: ABIS_BOOK（固定値）
    const ruleC = SpreadsheetApp.newDataValidation()
      .requireValueInList(['ABIS_BOOK'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow, 3, lastRow - startRow + 1, 1).setDataValidation(ruleC);
    Logger.log('✓ C列に入力規則を設定');
    
    // E列: 商品IDタイプ（固定値）
    const ruleE = SpreadsheetApp.newDataValidation()
      .requireValueInList(['ABIS_BOOK'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow, 5, lastRow - startRow + 1, 1).setDataValidation(ruleE);
    Logger.log('✓ E列に入力規則を設定');
    
    // F列: GTIN免除（固定値）
    const ruleF = SpreadsheetApp.newDataValidation()
      .requireValueInList(['GTIN免除'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow, 6, lastRow - startRow + 1, 1).setDataValidation(ruleF);
    Logger.log('✓ F列に入力規則を設定');
    
    // M列: メーカー名（WWD Korea固定）
    const ruleM = SpreadsheetApp.newDataValidation()
      .requireValueInList(['WWD Korea'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow, 13, lastRow - startRow + 1, 1).setDataValidation(ruleM);
    Logger.log('✓ M列に入力規則を設定');
    
    // Q列: 商品状態
    const ruleQ = SpreadsheetApp.newDataValidation()
      .requireValueInList(['新品', '中古品 - 良い', '中古品 - 非常に良い', '中古品 - ほぼ新品', 'コレクター商品 - 良い', 'コレクター商品 - 非常に良い', 'コレクター商品 - ほぼ新品', '再生品'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow, 17, lastRow - startRow + 1, 1).setDataValidation(ruleQ);
    Logger.log('✓ Q列に入力規則を設定');
    
    // AR列: 最大注文個数
    const ruleAR = SpreadsheetApp.newDataValidation()
      .requireValueInList(['DEFAULT'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow, 44, lastRow - startRow + 1, 1).setDataValidation(ruleAR);
    Logger.log('✓ AR列に入力規則を設定');
    
    // BV列: 配送方法
    const ruleBV = SpreadsheetApp.newDataValidation()
      .requireValueInList(['200円でゆうパケット', 'SAGAWA', 'その他'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow, 74, lastRow - startRow + 1, 1).setDataValidation(ruleBV);
    Logger.log('✓ BV列に入力規則を設定');
    
    // JJ列: 幅単位
    const ruleJJ = SpreadsheetApp.newDataValidation()
      .requireValueInList(['ミリメートル', 'センチメートル', 'メートル'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow, 290, lastRow - startRow + 1, 1).setDataValidation(ruleJJ);
    Logger.log('✓ JJ列に入力規則を設定');
    
    // JL列: 奥行き単位
    const ruleJL = SpreadsheetApp.newDataValidation()
      .requireValueInList(['ミリメートル', 'センチメートル', 'メートル'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow, 292, lastRow - startRow + 1, 1).setDataValidation(ruleJL);
    Logger.log('✓ JL列に入力規則を設定');
    
    // JN列: 高さ単位
    const ruleJN = SpreadsheetApp.newDataValidation()
      .requireValueInList(['センチメートル', 'ミリメートル', 'メートル'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow, 294, lastRow - startRow + 1, 1).setDataValidation(ruleJN);
    Logger.log('✓ JN列に入力規則を設定');
    
    // JP列: 重量単位
    const ruleJP = SpreadsheetApp.newDataValidation()
      .requireValueInList(['グラム', 'キログラム'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow, 296, lastRow - startRow + 1, 1).setDataValidation(ruleJP);
    Logger.log('✓ JP列に入力規則を設定');
    
    SpreadsheetApp.getUi().alert('✅ 入力規則の設定が完了しました！');
    
  } catch (error) {
    Logger.log('❌ エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert('エラー: ' + error.message);
  }
}

// ========================================
// ヘルパー関数
// ========================================

function ヘッダー設定(sheet) {
  // シートが存在するか確認
  if (!sheet) {
    throw new Error('シートが見つかりません');
  }
  
  const headers = {
    'A7': 'SKU', 'C7': '商品ID', 'D7': '商品名', 'E7': '商品ID タイプ', 'F7': 'GTIN',
    'M7': 'メーカー名', 'Q7': '商品状態', 'S7': '税込みの参考価格',
    'AR7': '最大注文個数', 'AS7': '販売開始日', 'AW7': '販売価格',
    'BV7': '配送方法', 'BW7': '商品説明', 'BZ7': '書籍タイトル2',
    'CG7': 'サフィックス値1', 'CL7': '商品タイプ1', 'CS7': '商品タイプ2', 'CT7': '商品タイプ3',
    'DC7': '出版日', 'CF7': 'ページ数', 'DS7': 'サフィックス値2', 'DT7': 'サフィックス値3',
    'DV7': '商品タイプ4', 'DW7': '商品タイプ5', 'DR7': '抽出タイトル',
    'EH7': '商品タイプ9', 'EI7': '商品タイプ6', 'ER7': '商品タイプ7', 'ES7': '商品タイプ8',
    'FB7': 'サフィックス値4', 'FC7': '書籍タイトル', 'FM7': '商品タイプ10', 'GP7': '商品タイプ11',
    'IY7': 'メイン画像', 'IZ7': '画像2', 'JA7': '画像3',
    'JI7': '幅', 'JJ7': '幅単位', 'JK7': '奥行き', 'JL7': '奥行き単位',
    'JM7': '高さ', 'JN7': '高さ単位', 'JO7': '重量', 'JP7': '重量単位'
  };
  
  try {
    for (const [cell, value] of Object.entries(headers)) {
      sheet.getRange(cell).setValue(value);
    }
  } catch (error) {
    Logger.log('ヘッダー設定エラー: ' + error.toString());
    throw error;
  }
}
function 価格データクリーン(priceText) {
  let result = String(priceText).trim().replace(/,/g, '').replace(/円/g, '').replace(/\\/g, '').replace(/\s/g, '');
  if (result && !isNaN(result)) return result.indexOf('.') > -1 ? "'" + result : result;
  return priceText;
}

function サフィックス値取得(skuValue) {
  if (!skuValue || skuValue.length === 0) return 1;
  const suffixMap = {'a': 1, 'b': 2, 'c': 3, 'd': 4, 'e': 5, 'f': 6, 'g': 7, 'h': 8, 'i': 9, 'j': 10};
  return suffixMap[skuValue.slice(-1).toLowerCase()] || 1;
}

function URL抽出(urlText) {
  const urls = [], regex = /https:\/\/[^\s;]+/g;
  let match;
  while ((match = regex.exec(urlText)) !== null) {
    urls.push(match[0].replace(/;/g, ''));
    if (urls.length >= 3) break;
  }
  return urls;
}

function 発売日抽出(text) {
  const pos = text.indexOf('発売日');
  if (pos === -1) return '';
  let dateStr = text.substring(pos);
  const lineBreak = dateStr.search(/[\r\n]/);
  if (lineBreak > -1) dateStr = dateStr.substring(0, lineBreak);
  const colonPos = Math.max(dateStr.indexOf('：'), dateStr.indexOf(':'));
  if (colonPos > -1) dateStr = dateStr.substring(colonPos + 1);
  dateStr = dateStr.replace(/以後|以降|頃|予定/g, '').trim();
  let result = '', dotCount = 0;
  for (let i = 0; i < dateStr.length; i++) {
    const c = dateStr[i];
    if (c >= '0' && c <= '9') result += c;
    else if (/[.\-/／]/.test(c) && dotCount < 2) { result += c; dotCount++; }
    else if (dotCount === 2 && result.length >= 8) break;
    else if (result.length > 0 && /[\s　]/.test(c)) break;
  }
  return 日付文字列整形(result.trim());
}

function 日付文字列整形(dateStr) {
  if (!dateStr) return '';
  let result = dateStr.replace(/[.\/／]/g, '-');
  if (result.indexOf('-') === -1) return dateStr;
  const parts = result.split('-');
  if (parts.length < 3) return dateStr;
  let [year, month, day] = parts.map(p => p.replace(/\D/g, ''));
  if (year.length === 2) year = (parseInt(year) > 50 ? '19' : '20') + year;
  if (year.length !== 4) return dateStr;
  return `${year}-${month.padStart(2, '0').substring(0, 2)}-${day.padStart(2, '0').substring(0, 2)}`;
}

function ページ数抽出(text) {
  const pos = text.search(/ページ[：:]/);
  if (pos === -1) return '';
  let pageStr = text.substring(pos + 4);
  const spacePos = pageStr.search(/[\s　]/);
  if (spacePos > -1) pageStr = pageStr.substring(0, spacePos);
  const lineBreak = pageStr.search(/[\r\n]/);
  if (lineBreak > -1) pageStr = pageStr.substring(0, lineBreak);
  pageStr = pageStr.trim();
  if (pageStr[0] === '-' || pageStr[0] === '－') return '';
  let result = '', digitCount = 0;
  for (let i = 0; i < pageStr.length; i++) {
    const c = pageStr[i];
    if (c >= '0' && c <= '9') { result += c; digitCount++; if (digitCount >= 3) break; }
    else if (result.length > 0) break;
  }
  return (result && result.length <= 3) ? result : '';
}

function サイズ抽出(text) {
  const pos = text.indexOf('ページ');
  if (pos === -1) return '';
  let sizeStr = text.substring(pos);
  const lineBreak = sizeStr.search(/[\r\n]/);
  if (lineBreak > -1) sizeStr = sizeStr.substring(0, lineBreak);
  const slashPos = Math.max(sizeStr.indexOf('／'), sizeStr.indexOf('/'));
  if (slashPos === -1) return '';
  sizeStr = sizeStr.substring(slashPos + 1).trim();
  let result = '';
  for (let i = 0; i < sizeStr.length; i++) {
    const c = sizeStr[i];
    if (c >= '0' && c <= '9') result += c;
    else if (/[xX*＊]/.test(c)) result += '*';
    else if (c === 'm') break;
  }
  return result;
}

function サイズ分割(sizeStr) {
  let width = '', depth = '';
  if (sizeStr.indexOf('*') > -1) {
    const parts = sizeStr.split('*');
    width = parts[0].replace(/\D/g, '');
    depth = parts[1] ? parts[1].replace(/\D/g, '') : '';
  }
  return {width, depth};
}

function タイトル抽出(titleText) {
  return titleText.replace(/\[.*?\]/g, '').replace(/\(.*?\)/g, '').replace(/（.*?）/g, '')
    .replace(/韓国\s*雑誌|韓国雑誌|雑誌|韓国/g, '').replace(/\s+/g, ' ').trim();
}

function 書籍タイトル抽出(titleText) {
  let result = titleText.replace(/韓国\s*雑誌\s*|韓国雑誌\s*|韓国\s*雑誌|韓国雑誌/g, '')
    .replace(/\[韓国雑誌\]|【韓国雑誌】/g, '').replace(/雑誌\s*|雑誌|韓国\s*|韓国/g, '');
  const volPos = result.search(/vol[.\s]/i);
  if (volPos > -1) result = result.substring(0, volPos);
  return result.replace(/\(.*?\)/g, '').replace(/（.*?）/g, '').replace(/\[.*?\]/g, '')
    .replace(/［.*?］/g, '').replace(/【.*?】/g, '').replace(/\s+/g, ' ').trim();
}
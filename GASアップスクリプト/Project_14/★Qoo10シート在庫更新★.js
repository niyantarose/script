/**********************
 * Qoo10在庫更新スクリプト（完全版 - 全パターン対応）
 **********************/

function 通知_安全_(message, title) {
  title = title || '通知';
  message = (message == null) ? '' : String(message);

  try {
    SpreadsheetApp.getUi().alert(title + '\n' + message);
    return;
  } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) ss.toast(message, title, 5);
  } catch (e2) {}

  try {
    Logger.log(title + ': ' + message);
  } catch (e3) {}
}

/**
 * コードから在庫タイプを判定（即納 or お取り寄せ）
 */
function 在庫タイプ判定(コード) {
  if (!コード) return null;
  
  // 末尾の文字を取得
  const 末尾 = コード.slice(-1);
  
  // 末尾が小文字の a で判定（即納）
  if (末尾 === 'a') {
    return '即納';
  }
  
  // 末尾が小文字の b で判定（お取り寄せ）
  if (末尾 === 'b') {
    return 'お取り寄せ';
  }
  
  // 末尾が数字の場合
  if (/\d$/.test(末尾)) {
    // まず a/b があるかチェック（優先）
    const match = コード.match(/([ab])(?:[^ab]*?)$/);
    if (match) {
      const 判定文字 = match[1];
      if (判定文字 === 'a') {
        return '即納';
      } else if (判定文字 === 'b') {
        return 'お取り寄せ';
      }
    }
    
    // a/b がない場合、末尾の数字で判定
    // 1 → 即納
    if (末尾 === '1') {
      return '即納';
    }
    // 0, 2 → お取り寄せ
    if (末尾 === '0' || 末尾 === '2') {
      return 'お取り寄せ';
    }
  }
  
  return null;
}

/**
 * N列のオプション文字列から全てのsub-codeを抽出
 */
function N列からコード抽出(N列値) {
  const コード一覧 = [];
  
  if (!N列値) return コード一覧;
  
  const ブロック配列 = N列値.split('$$');
  
  for (let i = 0; i < ブロック配列.length; i++) {
    const ブロック = ブロック配列[i];
    if (!ブロック) continue;
    
    const パーツ = ブロック.split('||');
    
    for (let j = パーツ.length - 1; j >= 0; j--) {
      const パーツ値 = パーツ[j] || '';
      if (パーツ値.startsWith('*')) {
        const コード = パーツ値.substring(1).trim();
        if (コード) {
          コード一覧.push({
            コード: コード,
            ブロックIndex: i,
            パーツIndex: j
          });
          break;
        }
      }
    }
  }
  
  return コード一覧;
}

/**
 * Qoo10シートの在庫を更新
 */
function Qoo10在庫を更新() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const Yahoo全在庫シート = ss.getSheetByName('Yahoo全在庫');
  const Qoo10シート = ss.getSheetByName('Qoo10');

  if (!Yahoo全在庫シート) {
    通知_安全_('「Yahoo全在庫」シートが見つかりません');
    return;
  }

  if (!Qoo10シート) {
    通知_安全_('「Qoo10」シートが見つかりません');
    return;
  }

  // Yahoo全在庫シートのデータを取得
  const Yahoo全在庫LastRow = Yahoo全在庫シート.getLastRow();
  if (Yahoo全在庫LastRow < 2) {
    通知_安全_('Yahoo全在庫シートにデータがありません');
    return;
  }

  const Yahoo全在庫データ = Yahoo全在庫シート
    .getRange(2, 1, Yahoo全在庫LastRow - 1, 4)
    .getValues();

  Logger.log('===== Yahoo全在庫データ読み込み開始 =====');

  // 在庫データをMapに変換
  const 在庫Map = {};
  
  Yahoo全在庫データ.forEach(function (row) {
    const code = String(row[0] || '').trim();
    const subCode = String(row[2] || '').trim();
    const quantity = Number(row[3]) || 0;

    if (!code) return;

    if (subCode) {
      // sub-codeをキーとして在庫数を保存
      if (在庫Map[subCode] === undefined) {
        在庫Map[subCode] = quantity;
      } else {
        在庫Map[subCode] += quantity;
      }
      
      // 親コードに集約
      const タイプ = 在庫タイプ判定(subCode);
      
      if (!在庫Map[code]) {
        在庫Map[code] = { 即納: 0, お取り寄せ: 0, sub_code有無: true };
      }
      
      if (タイプ === '即納') {
        在庫Map[code].即納 += quantity;
      } else if (タイプ === 'お取り寄せ') {
        在庫Map[code].お取り寄せ += quantity;
      }
    } else {
      // sub-codeがない場合は親コードに即納として保存
      if (!在庫Map[code]) {
        在庫Map[code] = { 即納: 0, お取り寄せ: 0, sub_code有無: false };
      }
      // 空白の場合は即納として扱う
      在庫Map[code].即納 += quantity;
      在庫Map[code].sub_code有無 = false;
    }
  });

  Logger.log('在庫Map構築完了');

  // Qoo10シートのデータを取得（1-4行目ヘッダー、5行目からデータ）
  const Qoo10LastRow = Qoo10シート.getLastRow();
  if (Qoo10LastRow < 5) {
    通知_安全_('Qoo10シートにデータがありません');
    return;
  }

  const データ行数 = Qoo10LastRow - 4;

  const Qoo10データB = Qoo10シート.getRange(5, 2, データ行数, 1).getValues();  // B列
  const Qoo10データM = Qoo10シート.getRange(5, 13, データ行数, 1).getValues(); // M列
  const Qoo10データN = Qoo10シート.getRange(5, 14, データ行数, 1).getValues(); // N列

  // 更新用配列
  const M列更新 = [];
  const N列更新 = [];
  const AA列更新 = [];
  let 更新件数 = 0;

  // 各行を処理
  for (let i = 0; i < データ行数; i++) {
    const seller_unique_item_id = String(Qoo10データB[i][0] || '').trim();
    const 現在のN列 = String(Qoo10データN[i][0] || '');

    // 在庫情報を取得
    const 在庫情報 = 在庫Map[seller_unique_item_id];

    let M列値 = Qoo10データM[i][0] || 0;
    let N列値 = 現在のN列;
    let AA列値 = '';

    if (在庫情報) {
      const 即納 = 在庫情報.即納 || 0;
      const お取り寄せ = 在庫情報.お取り寄せ || 0;
      const sub_code有無 = 在庫情報.sub_code有無;

      // sub-codeがない（空白）の場合
      if (sub_code有無 === false) {
        // M列は更新しない（元の値を保持）
        M列値 = Qoo10データM[i][0] || 0;
        // AA列は3
        AA列値 = '3';
      } else {
        // 通常の処理
        // M列：即納優先、なければお取り寄せ10、どちらもなければ0
        if (即納 > 0) {
          M列値 = 即納;
          AA列値 = '3';
        } else if (お取り寄せ > 0) {
          M列値 = 10;
          AA列値 = '14';
        } else {
          M列値 = 0;
          AA列値 = '14';
        }
      }

      // N列：オプション在庫を更新
      if (現在のN列) {
        const コード情報一覧 = N列からコード抽出(現在のN列);
        
        N列値 = Qoo10オプション在庫を更新する(
          現在のN列,
          コード情報一覧,
          在庫Map
        );
      }

      更新件数++;
    }

    M列更新.push([M列値]);
    N列更新.push([N列値]);
    AA列更新.push([AA列値]);
  }

  // 一括書き込み
  if (M列更新.length > 0) {
    Qoo10シート.getRange(5, 13, M列更新.length, 1).setValues(M列更新);
    Qoo10シート.getRange(5, 14, N列更新.length, 1).setValues(N列更新);
    Qoo10シート.getRange(5, 27, AA列更新.length, 1).setValues(AA列更新);
  }

  Logger.log('処理完了: ' + 更新件数 + '件');
  通知_安全_('更新完了\n' + 更新件数 + '件の商品を更新しました');
}

/**
 * オプション在庫を更新する
 */
function Qoo10オプション在庫を更新する(現在の値, コード情報一覧, 在庫Map) {
  if (!現在の値 || !コード情報一覧 || コード情報一覧.length === 0) {
    return 現在の値;
  }

  // 即納ブロックがあるかチェック
  let 即納ブロックあり = false;
  
  for (let i = 0; i < コード情報一覧.length; i++) {
    const コード = コード情報一覧[i].コード;
    const タイプ = 在庫タイプ判定(コード);
    const 在庫数 = 在庫Map[コード];
    
    if (タイプ === '即納' && 在庫数 > 0) {
      即納ブロックあり = true;
      break;
    }
  }

  // $$で分割
  const ブロック配列 = 現在の値.split('$$');
  const 更新済みブロック = [];

  for (let i = 0; i < ブロック配列.length; i++) {
    const ブロック = ブロック配列[i];
    
    if (!ブロック) {
      更新済みブロック.push(ブロック);
      continue;
    }

    // このブロックのコード情報を取得
    const このブロックのコード情報 = コード情報一覧.find(function(info) {
      return info.ブロックIndex === i;
    });

    if (!このブロックのコード情報) {
      更新済みブロック.push(ブロック);
      continue;
    }

    const コード = このブロックのコード情報.コード;
    const パーツIndex = このブロックのコード情報.パーツIndex;
    
    // ||で分割
    const パーツ = ブロック.split('||');
    
    // 在庫数更新位置（コードの1つ前）
    const 在庫数更新位置 = パーツIndex - 1;
    
    if (在庫数更新位置 < 0 || 在庫数更新位置 >= パーツ.length) {
      更新済みブロック.push(ブロック);
      continue;
    }

    // 新しい在庫数を決定
    let 新しい在庫数 = 0;
    const タイプ = 在庫タイプ判定(コード);
    
    if (在庫Map[コード] !== undefined) {
      const 在庫数 = Number(在庫Map[コード]) || 0;
      
      if (タイプ === '即納') {
        新しい在庫数 = 在庫数;
      } else if (タイプ === 'お取り寄せ') {
        // 即納ブロックがある場合はお取り寄せを0にする
        if (即納ブロックあり) {
          新しい在庫数 = 0;
        } else {
          新しい在庫数 = 在庫数;
        }
      }
    }

    // 在庫数を更新
    パーツ[在庫数更新位置] = '*' + 新しい在庫数;
    
    // ブロックを再構築
    更新済みブロック.push(パーツ.join('||'));
  }

  return 更新済みブロック.join('$$');
}

/**
 * 一括更新（ボタン用）
 */
function Qoo10一括更新_ボタン用() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30 * 1000)) {
    通知_安全_('別の処理が実行中です。少し待ってから再実行してください。', '一括更新');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    ss.toast('一括更新：開始', 'Qoo10', 3);

    // Qoo10在庫更新（M/N/AA）
    Qoo10在庫を更新();
    ss.toast('Qoo10在庫更新：完了', 'Qoo10', 3);

    通知_安全_('一括更新が完了しました。', '一括更新');

  } catch (err) {
    Logger.log(err);
    通知_安全_('一括更新でエラー：\n' + (err && err.message ? err.message : err), '一括更新エラー');
  } finally {
    lock.releaseLock();
  }
}
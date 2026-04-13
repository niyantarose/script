/******************************************************
 * 商品コード管理ツール v10.10（台湾まんが）
 *
 * 【設計方針】
 * ・入力するだけでタイトル・SKUが段階的に自動生成される
 * ・Edit時にWorksの確定列(F/G/H)には書かない
 * ・Edit時にWorksの予約列(I/J)には書く
 * ・一括更新 / 重複削除はセーフティネットとして用意
 *
 * ✅ v10.10の修正点（v10.9からの変更）
 * 
 * 【Works照合キー】
 * - 原題あり → 「原題||正規化原題」のみをWorksKeyとして使用
 *   → 作者の表記ゆれ・日本語タイトル修正でIDが動かない
 * - 原題なし → 従来通り「日本語タイトル||作者」（フォールバック）
 * - 原題が既存にあるが作者が明らかに違う → 「【同名疑い・要確認】」をステータスに出す
 *   IDは増やさず手動判断に委ねる
 *
 * 【商品行重複チェック】
 * - ISBN あり → ISBNで最優先照合（1巻単位で完全一致）
 * - ISBN なし → 原題+巻数+形態+言語 の複合キーで照合
 *               セット商品は 原題+開始巻-終了巻+形態+言語
 *
 * ✅ v10.9からの継続
 * - Works登録タイミング: タイトル+作者+言語+カテゴリの4条件が揃ってから
 * - 4条件揃うまでは作品IDを'????'で仮表示
 * - 🧹孤立Worksエントリー削除メニュー
 * - ③一括更新に孤立エントリー自動削除を組み込み
 * - セット商品の予約巻数書き込み対応
 ******************************************************/

/* ============================
 * 設定
 * ============================ */
const 設定 = {
  マスターシート名: '台湾まんがマスターシート',
  作品シート名: 'Works',
  言語マスター名: '言語マスター',
  カテゴリマスター名: 'カテゴリマスター',

  作品ヘッダー: ['WorksKey', '作品ID', '日本語タイトル', '作者', '原題タイトル', '登録済み巻', '最新巻', '更新日時', '最新巻(予約込み)', '予約更新日時'],
  作品列数: 10,

  言語ヘッダー: ['言語', 'コード', '色'],
  カテゴリヘッダー: ['カテゴリ', 'コード', '色'],

  言語初期値: [
    ['台湾', 'TW', '#b7e1cd'],
    ['韓国', 'KR', '#fce8b2'],
    ['日本', 'JP', '#f4c7c3'],
    ['中国', 'CN', '#c9daf8'],
    ['タイ', 'TH', '#d9d2e9'],
    ['英語', 'US', '#fce5cd']
  ],

  カテゴリ初期値: [
    ['コミック', 'CM', '#fff2cc'],
    ['まんが', 'CM', '#fff2cc'],
    ['小説', 'NV', '#d9ead3'],
    ['グッズ', 'GD', '#cfe2f3'],
    ['設定集', 'ART', '#f4cccc'],
    ['アートブック', 'ART', '#f4cccc'],
    ['雑誌', 'MZ', '#d9d2e9']
  ],

  色パレット: ['#b7e1cd', '#fce8b2', '#f4c7c3', '#c9daf8', '#d9d2e9', '#fce5cd', '#fff2cc', '#d9ead3', '#cfe2f3', '#f4cccc'],

  形態プレフィックス: {
    '特装版': 'S',
    '初版限定版': 'F',
    '初回限定版': 'F'
  },

  特典自動付与: true,
  形態別特典: {
    '初版限定版': '※初版限定版特典付き',
    '初回限定版': '※初回限定版特典付き',
    '特装版': '※特装版限定特典付き'
  }
};

/* ============================
 * 列名（半角カッコ・半角スラッシュで統一）
 * ============================ */
const 列名 = {
  発行チェック: '発番発行',
  商品コード: '商品コード(SKU)',
  タイトル: 'タイトル',
  作者: '作者',
  日本語タイトル: '日本語タイトル',
  原題: '原題タイトル',
  形態: '形態(通常/初回限定/特装)',
  言語: '言語',
  カテゴリ: 'カテゴリ',
  単巻数: '単巻数',
  セット開始: 'セット巻数開始番号',
  セット終了: 'セット巻数終了番号',
  特典メモ: '特典メモ',
  ISBN: 'ISBN',
  作品ID: '作品ID(W)(自動)',
  SKU: 'SKU(自動)',
  コードステータス: '商品コードステータス',
  登録状況: '登録状況'
};

const 監視列 = [
  列名.作者, 列名.日本語タイトル, 列名.原題,
  列名.言語, 列名.カテゴリ, 列名.形態,
  列名.単巻数, 列名.セット開始, 列名.セット終了, 列名.特典メモ,
  列名.ISBN
];

/* ============================
 * ヘッダー正規化
 * ============================ */
function ヘッダー正規化_(s) {
  return String(s || '')
    .replace(/（/g, '(').replace(/）/g, ')')
    .replace(/／/g, '/').replace(/　/g, ' ')
    .replace(/\s+/g, ' ').trim();
}

function 列番号を取得(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const name = ヘッダー正規化_(headers[i]);
    if (name) map[name] = i + 1;
  }
  return map;
}

/* ============================
 * メニュー
 * ============================ */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('商品コード管理')
      .addItem('① 商品コード確定発行', 'メニュー_確定発行')
      .addItem('② 商品コード削除', 'メニュー_削除')
      .addSeparator()
      .addItem('③ 既存データ一括更新(Works整理含む)', 'メニュー_一括更新')
      .addSeparator()
      .addItem('④ Works初期化', 'メニュー_Works初期化')
      .addItem('⑤ マスター作成', 'メニュー_マスター作成')
      .addItem('⑥ プルダウン更新', 'メニュー_プルダウン更新')
      .addSeparator()                                                         // ← 追加
      .addItem('📋 韓国マンガシート作成', 'メニュー_韓国マンガシート作成')    // ← 追加
      .addSeparator()
      .addItem('🔍 Works重複チェック・統合', 'メニュー_重複統合')
      .addItem('🔄 WorksKey再正規化', 'メニュー_WorksKey再正規化')
      .addItem('🔢 Works ID振り直し', 'メニュー_ID振り直し')
      .addItem('🧹 Works孤立エントリー削除', 'メニュー_孤立削除')
      .addToUi();
  } catch (e) {
    Logger.log('onOpenはスプレッドシートから実行してください: ' + e.message);
  }
}

/* ============================================================
 * ✅ v10.10: WorksKey生成（原題優先）
 *
 * 原題あり → 「原題||正規化原題」のみ
 * 原題なし → 「正規化タイトル||正規化作者」（フォールバック）
 * ============================================================ */
function WorksKeyを作る(日本語タイトル, 作者, 原題 = '') {
  const 正規化原題 = キー用正規化_(原題);
  if (正規化原題) {
    return '原題||' + 正規化原題;
  }
  return キー用正規化_(日本語タイトル) + '||' + キー用正規化_(作者);
}

/* ============================================================
 * ✅ v10.10: Works照合と同名疑い警告
 *
 * - 原題キーで既存が見つかり、作者が明らかに違う → 警告文字列を返す
 * - IDは増やさず、手動確認を促す
 * ============================================================ */
function Works照合と警告(worksKey, 作者, 作品データ) {
  const result = { 作品ID: null, 警告: '' };
  const 既存ID = 作品データ.keyToId[worksKey];
  if (!既存ID) return result;

  result.作品ID = String(既存ID).padStart(4, '0');

  // 原題キーの場合のみ作者チェック（作者フォールバックキーは作者込みなのでスキップ）
  if (worksKey.startsWith('原題||') && 作者) {
    const 既存作者 = 作品データ.keyToData[worksKey]?.作者 || '';
    if (既存作者 && キー用正規化_(作者) !== キー用正規化_(既存作者)) {
      result.警告 = '【同名疑い・要確認】';
    }
  }

  return result;
}

/* ============================================================
 * ✅ v10.10: 商品行重複チェックキー生成
 *
 * ISBN があれば呼び出し元で先に処理（ISBNトラッカーで判定）
 * ISBNなしのとき → 原題+巻数+形態+言語 の複合キー
 * ============================================================ */
function 商品行キーを作る(原題, 単巻数, セット開始, セット終了, 形態, 言語) {
  const タイトル部 = キー用正規化_(原題) || 'NOTITLE';
  const 形態部 = キー用正規化_(形態) || '通常';
  const 言語部 = String(言語 || '').trim();

  const sf = String(セット開始 || '').trim();
  const st = String(セット終了 || '').trim();
  let 巻数部;

  if (sf && st) {
    巻数部 = `${sf}-${st}`;
  } else {
    const v = 数値変換(単巻数);
    巻数部 = v != null ? String(v).padStart(2, '0') : 'NOVOL';
  }

  return `${タイトル部}||${巻数部}||${形態部}||${言語部}`;
}

/* ============================================================
 * ✅ v10.10: Edit
 * ============================================================ */
function Edit(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== 設定.マスターシート名) return;

  const 開始行 = e.range.getRow();
  const 行数 = e.range.getNumRows();
  if (開始行 + 行数 - 1 < 2) return;

  const lock = LockService.getDocumentLock();
  try { if (!lock.tryLock(5000)) return; } catch (_) { return; }

  try {
    if (自己更新中か_()) return;

    const 列マップ = 列番号を取得(sh);
    const 編集開始列 = e.range.getColumn();
    const 編集終了列 = e.range.getLastColumn();
    const 監視列番号 = 監視列.map(h => 列マップ[h]).filter(Boolean);
    const 対象列が含まれる = 監視列番号.some(c => c >= 編集開始列 && c <= 編集終了列);
    if (!対象列が含まれる) return;

    自己更新を開始_();
    try {
      const 最終列 = sh.getLastColumn();
      const 行データ一覧 = sh.getRange(開始行, 1, 行数, 最終列).getValues();

      const ss = SpreadsheetApp.getActive();
      const 作品シート = 作品シートを確保(ss);
      const 作品データ = 全作品データを読み込み(作品シート);
      const 言語マップ = 言語マップを取得(ss);
      const カテゴリマップ = カテゴリマップを取得(ss);

      const 出力SKU = [];
      const 出力タイトル = [];
      const 出力作品ID = [];
      const 出力SKU自動 = [];
      const 出力ステータス = [];

      // ✅ v10.10: ISBN重複チェックトラッカー（1巻単位）
      const ISBNトラッカー = {};
      // ✅ v10.10: 商品行複合キートラッカー（原題+巻数+形態+言語）
      const 商品行キートラッカー = {};
      // Works I/J列書き込み用
      const 予約巻数マップ = {};

      for (let i = 0; i < 行数; i++) {
        const 行番号 = 開始行 + i;
        const 行データ = 行データ一覧[i];
        if (行番号 < 2) {
          出力SKU.push(['']); 出力タイトル.push(['']); 出力作品ID.push(['']);
          出力SKU自動.push(['']); 出力ステータス.push(['']);
          continue;
        }

        const 取得 = (名前) => 正規化(行データ[(列マップ[名前] || 1) - 1]);
        const 取得生 = (名前) => 行データ[(列マップ[名前] || 1) - 1];

        const 日本語タイトル = 取得(列名.日本語タイトル);
        let 作者 = 取得(列名.作者);
        let 原題 = 取得(列名.原題);
        const 言語 = 取得(列名.言語);
        const カテゴリ = 取得(列名.カテゴリ);
        const 形態 = 取得(列名.形態);
        const 単巻数 = 取得生(列名.単巻数);
        const セット開始 = 取得生(列名.セット開始);
        const セット終了 = 取得生(列名.セット終了);
        const 特典メモ = 取得(列名.特典メモ);
        const ISBN = 取得(列名.ISBN);

        if (!日本語タイトル || !作者) {
          出力SKU.push([取得生(列名.商品コード) || '']);
          出力タイトル.push([取得生(列名.タイトル) || '']);
          出力作品ID.push([取得生(列名.作品ID) || '']);
          出力SKU自動.push([取得生(列名.SKU) || '']);
          出力ステータス.push([取得生(列名.コードステータス) || '']);
          continue;
        }

        // ====================================================
        // ✅ v10.10: Works照合（原題優先キー）
        // ====================================================
        const worksKey = WorksKeyを作る(日本語タイトル, 作者, 原題);
        const 全条件揃い = !!(日本語タイトル && 作者 && 言語 && カテゴリ);

        let 作品ID = '';
        let 同名警告 = '';

        // 既存照合（原題キー or タイトル+作者キー）
        const 照合結果 = Works照合と警告(worksKey, 作者, 作品データ);

        if (照合結果.作品ID) {
          // 既存作品: IDを引き継ぐ（F/G/H列は書かない）
          作品ID = 照合結果.作品ID;
          同名警告 = 照合結果.警告;
          const 既存 = 作品データ.keyToData[worksKey];
          if (既存) {
            if (!作者 && 既存.作者) 作者 = 既存.作者;
            if (!原題 && 既存.原題) 原題 = 既存.原題;
          }
        } else {
          // 新規作品
          if (全条件揃い) {
            // ✅ 原題キーのとき: 既存titleToKeyからも探す（日本語タイトル表記ゆれ保険）
            if (!worksKey.startsWith('原題||')) {
              const titleKey = キー用正規化_(日本語タイトル);
              const 既存worksKey = 作品データ.titleToKey[titleKey];
              if (既存worksKey && 作品データ.keyToId[既存worksKey]) {
                // 日本語タイトル一致 → 既存IDを流用
                作品ID = String(作品データ.keyToId[既存worksKey]).padStart(4, '0');
                作品データ.keyToId[worksKey] = 作品ID;
                作品データ.keyToRow[worksKey] = 作品データ.keyToRow[既存worksKey];
                作品データ.keyToData[worksKey] = 作品データ.keyToData[既存worksKey];
              }
            }

            if (!作品ID) {
              // 本当の新規
              作品データ.maxId++;
              作品ID = String(作品データ.maxId).padStart(4, '0');
              作品データ.keyToId[worksKey] = 作品ID;
              作品データ.keyToRow[worksKey] = null;
              作品データ.keyToData[worksKey] = { 作者, 原題 };
              作品データ.newRows.push([worksKey, 作品ID, 日本語タイトル, 作者, 原題, '', '', '', '', '']);
              // titleToKeyに登録（フォールバックキー用）
              if (!worksKey.startsWith('原題||')) {
                const titleKey = キー用正規化_(日本語タイトル);
                if (!作品データ.titleToKey[titleKey]) 作品データ.titleToKey[titleKey] = worksKey;
              }
            }
          } else {
            作品ID = '????';
          }
        }

        // ====================================================
        // ✅ v10.10: 商品行重複チェック
        // ① ISBN優先 → ② 原題+巻数+形態+言語
        // ====================================================
        let 巻数警告 = '';

        if (ISBN) {
          // ① ISBNで判定（1巻単位で最強）
          if (ISBNトラッカー[ISBN]) {
            巻数警告 = '【重複注意】';
          }
          ISBNトラッカー[ISBN] = true;
        } else if (原題) {
          // ② 原題+巻数+形態+言語の複合キー
          const 商品行キー = 商品行キーを作る(原題, 単巻数, セット開始, セット終了, 形態, 言語);
          if (商品行キートラッカー[商品行キー]) {
            巻数警告 = '【重複注意】';
          }
          商品行キートラッカー[商品行キー] = true;
        }

        // 予約巻数マップへ追加（I/J列用）
        const 巻数 = 数値変換(単巻数);
        const セット終了巻 = 数値変換(セット終了);
        const 最新巻候補 = 巻数 != null ? 巻数 : セット終了巻;
        if (最新巻候補 != null && 作品ID !== '????') {
          if (!予約巻数マップ[worksKey]) 予約巻数マップ[worksKey] = [];
          予約巻数マップ[worksKey].push(最新巻候補);
        }

        // ====================================================
        // 出力生成
        // ====================================================
        const 言語コード = 言語 ? (言語マップ[言語] || 'XX') : '';
        const カテゴリコード = カテゴリ ? (カテゴリマップ[カテゴリ] || 'XX') : '';
        const SKU = SKUを段階生成(言語コード, 形態, 作品ID, カテゴリコード, 単巻数, セット開始, セット終了);
        const タイトル = 同名警告 + 巻数警告 + タイトルを段階生成(言語, カテゴリ, 形態, 日本語タイトル, 単巻数, セット開始, セット終了, 作者, 原題, 特典メモ);

        let ステータス = '';
        if (同名警告) {
          ステータス = 同名警告;
        } else if (巻数警告) {
          ステータス = '【重複注意】';
        } else {
          ステータス = 全条件揃い ? '商品コード(予約)' : '入力中...';
        }

        出力SKU.push([SKU]);
        出力タイトル.push([タイトル]);
        出力作品ID.push([作品ID]);
        出力SKU自動.push([SKU]);
        出力ステータス.push([ステータス]);
      }

      // マスターシートへ一括書き出し
      if (列マップ[列名.商品コード]) sh.getRange(開始行, 列マップ[列名.商品コード], 行数, 1).setValues(出力SKU);
      if (列マップ[列名.タイトル]) sh.getRange(開始行, 列マップ[列名.タイトル], 行数, 1).setValues(出力タイトル);
      if (列マップ[列名.作品ID]) sh.getRange(開始行, 列マップ[列名.作品ID], 行数, 1).setValues(出力作品ID);
      if (列マップ[列名.SKU]) sh.getRange(開始行, 列マップ[列名.SKU], 行数, 1).setValues(出力SKU自動);
      if (列マップ[列名.コードステータス]) sh.getRange(開始行, 列マップ[列名.コードステータス], 行数, 1).setValues(出力ステータス);

      // WorksへI/J列書き込み（新規登録 + 予約更新）
      作品データを更新_予約込み(作品シート, 作品データ, 予約巻数マップ);
    } finally {
      自己更新を終了_();
    }
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================
 * Works更新（Edit専用: 新規登録 + 予約I/J列更新）
 * ============================================================ */
function 作品データを更新_予約込み(作品シート, 作品データ, 予約巻数マップ) {
  if (作品データ.keyUpdates && 作品データ.keyUpdates.length > 0) {
    for (const upd of 作品データ.keyUpdates) 作品シート.getRange(upd.行, 1).setValue(upd.key);
    作品データ.keyUpdates = [];
  }

  if (作品データ.newRows.length > 0) {
    const 開始行 = Math.max(2, 作品シート.getLastRow() + 1);
    作品シート.getRange(開始行, 1, 作品データ.newRows.length, 設定.作品列数).setValues(作品データ.newRows);
    for (let i = 0; i < 作品データ.newRows.length; i++) {
      const key = String(作品データ.newRows[i][0] || '').trim();
      if (key) 作品データ.keyToRow[key] = 開始行 + i;
    }
  }

  const now = new Date();
  for (const [key, 巻数リスト] of Object.entries(予約巻数マップ)) {
    const 行番号 = 作品データ.keyToRow[key];
    if (!行番号) continue;
    const 今回最大 = Math.max(...巻数リスト);
    const 既存予約最新巻 = 作品データ.keyTo予約最新巻[key] || 0;
    const 新最新巻 = Math.max(今回最大, 既存予約最新巻);
    作品シート.getRange(行番号, 9).setValue(新最新巻);
    作品シート.getRange(行番号, 10).setValue(now);
  }

  作品データ.newRows = [];
}

/* ============================================================
 * コア処理
 * ============================================================ */
function WorksKey再正規化を実行(作品シート) {
  if (!作品シート || 作品シート.getLastRow() < 2) return { キー更新数: 0, 統合数: 0, 削除行数: 0 };

  const データ = 作品シート.getRange(2, 1, 作品シート.getLastRow() - 1, 設定.作品列数).getValues();
  const keyMap = new Map();

  for (let i = 0; i < データ.length; i++) {
    const r = データ[i];
    const t = 正規化(r[2] || ''), a = 正規化(r[3] || ''), o = 正規化(r[4] || '');
    if (!t && !a) continue;
    const key = WorksKeyを作る(t, a, o);
    const 元key = String(r[0] || '').trim();
    if (!keyMap.has(key)) keyMap.set(key, []);
    keyMap.get(key).push({ 行: i + 2, データ: r, 元key, キー変更: 元key !== key });
  }

  let キー更新数 = 0, 統合数 = 0;
  const 削除行 = [];

  for (const [key, 行配列] of keyMap.entries()) {
    if (行配列.length === 1) {
      if (行配列[0].キー変更) { 作品シート.getRange(行配列[0].行, 1).setValue(key); キー更新数++; }
    } else {
      行配列.sort((a, b) => parseInt(String(a.データ[1] || '9999'), 10) - parseInt(String(b.データ[1] || '9999'), 10));
      const 残す = 行配列[0];
      作品シート.getRange(残す.行, 1).setValue(key);

      const 全巻 = new Set();
      for (const item of 行配列) {
        String(item.データ[5] || '').split(',').forEach(v => { const n = parseInt(String(v).trim(), 10); if (!isNaN(n)) 全巻.add(n); });
      }
      if (全巻.size > 0) {
        const arr = Array.from(全巻).sort((a, b) => a - b);
        作品シート.getRange(残す.行, 6).setValue(arr.join(','));
        作品シート.getRange(残す.行, 7).setValue(Math.max(...arr));
        作品シート.getRange(残す.行, 8).setValue(new Date());
      }

      let 最大予約巻 = 0;
      for (const item of 行配列) {
        const v = parseInt(String(item.データ[8] || '0'), 10);
        if (!isNaN(v) && v > 最大予約巻) 最大予約巻 = v;
      }
      if (最大予約巻 > 0) {
        作品シート.getRange(残す.行, 9).setValue(最大予約巻);
        作品シート.getRange(残す.行, 10).setValue(new Date());
      }

      if (!正規化(残す.データ[4] || '')) {
        for (let j = 1; j < 行配列.length; j++) {
          const t = 正規化(行配列[j].データ[4] || '');
          if (t) { 作品シート.getRange(残す.行, 5).setValue(t); break; }
        }
      }
      for (let j = 1; j < 行配列.length; j++) 削除行.push(行配列[j].行);
      統合数++;
    }
  }
  削除行.sort((a, b) => b - a);
  for (const 行 of 削除行) 作品シート.deleteRow(行);
  return { キー更新数, 統合数, 削除行数: 削除行.length };
}

function WorksID振り直しを実行(作品シート) {
  const result = { 変更数: 0, 旧新マップ: {} };
  if (!作品シート || 作品シート.getLastRow() < 2) return result;

  const データ = 作品シート.getRange(2, 1, 作品シート.getLastRow() - 1, 設定.作品列数).getValues();
  const 行リスト = データ.map((r, i) => ({ 旧ID: parseInt(String(r[1] || '0'), 10), データ: r }));
  行リスト.sort((a, b) => a.旧ID - b.旧ID);

  const 新データ = 行リスト.map((item, i) => {
    const r = item.データ.slice();
    const 新ID = String(i + 1).padStart(4, '0');
    if (item.旧ID !== i + 1) { result.旧新マップ[String(item.旧ID).padStart(4, '0')] = 新ID; result.変更数++; }
    r[1] = 新ID;
    return r;
  });

  作品シート.getRange(2, 1, 新データ.length, 設定.作品列数).setValues(新データ);
  return result;
}

function Works重複を検出(作品シート) {
  if (!作品シート || 作品シート.getLastRow() < 2) return [];
  const データ = 作品シート.getRange(2, 1, 作品シート.getLastRow() - 1, 設定.作品列数).getValues();
  const keyMap = new Map();
  for (let i = 0; i < データ.length; i++) {
    const t = 正規化(データ[i][2]), a = 正規化(データ[i][3]), o = 正規化(データ[i][4]);
    if (!t || !a) continue;
    const key = WorksKeyを作る(t, a, o);
    if (!keyMap.has(key)) keyMap.set(key, []);
    keyMap.get(key).push({ 行: i + 2, データ: データ[i] });
  }
  return Array.from(keyMap.entries()).filter(([_, v]) => v.length > 1).map(([key, 行配列]) => ({ key, 行配列 }));
}

/* ============================================================
 * 孤立Worksエントリー削除
 * ============================================================ */
function Works孤立エントリーを削除(作品シート, マスターシート, 列マップ) {
  const result = { 削除数: 0, 削除リスト: [] };
  if (!作品シート || 作品シート.getLastRow() < 2) return result;

  const 最終行 = データがある最終行を取得(マスターシート, 列マップ);
  const 使用中Keys = new Set();
  if (最終行 >= 2) {
    const 全データ = マスターシート.getRange(2, 1, 最終行 - 1, マスターシート.getLastColumn()).getValues();
    for (const r of 全データ) {
      const t = 正規化(r[(列マップ[列名.日本語タイトル] || 1) - 1]);
      const a = 正規化(r[(列マップ[列名.作者] || 1) - 1]);
      const o = 正規化(r[(列マップ[列名.原題] || 1) - 1]);
      if (t && a) 使用中Keys.add(WorksKeyを作る(t, a, o));
    }
  }

  const worksData = 作品シート.getRange(2, 1, 作品シート.getLastRow() - 1, 設定.作品列数).getValues();
  const 削除行 = [];
  for (let i = 0; i < worksData.length; i++) {
    const t = 正規化(worksData[i][2] || ''), a = 正規化(worksData[i][3] || ''), o = 正規化(worksData[i][4] || '');
    if (!t && !a) { 削除行.push(i + 2); continue; }
    const key = WorksKeyを作る(t, a, o);
    if (!使用中Keys.has(key)) {
      削除行.push(i + 2);
      result.削除リスト.push(`ID:${worksData[i][1]} ${worksData[i][2]} / ${worksData[i][3]}`);
    }
  }

  削除行.sort((a, b) => b - a);
  for (const 行 of 削除行) 作品シート.deleteRow(行);
  result.削除数 = 削除行.length;
  return result;
}

/* ============================================================
 * メニューラッパー
 * ============================================================ */
function メニュー_WorksKey再正規化() {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(設定.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  if (ui.alert('確認', '全WorksKeyを再計算し重複を統合します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const r = WorksKey再正規化を実行(作品シート);
    ui.alert(`✅ 完了\nキー更新: ${r.キー更新数}\n重複統合: ${r.統合数}\n削除: ${r.削除行数}` + (r.統合数 > 0 ? '\n\n③一括更新を実行してください。' : ''));
  } finally { lock.releaseLock(); }
}

function メニュー_重複統合() {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(設定.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const 重複リスト = Works重複を検出(作品シート);
    if (重複リスト.length === 0) { ui.alert('✅ 重複はありません！'); return; }
    let レポート = `🔍 重複検出: ${重複リスト.length}件\n\n`;
    for (const dup of 重複リスト) {
      レポート += `【${dup.行配列[0].データ[2]}】 著：${dup.行配列[0].データ[3]}\n`;
      for (const item of dup.行配列) レポート += `  ID:${item.データ[1]} (行${item.行})\n`;
      レポート += '\n';
    }
    if (ui.alert('重複チェック結果', レポート + '\n自動統合しますか？', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    const r = WorksKey再正規化を実行(作品シート);
    ui.alert(`✅ 統合完了\n統合: ${r.統合数}件, 削除: ${r.削除行数}行\n\n③一括更新を実行してください。`);
  } finally { lock.releaseLock(); }
}

function メニュー_ID振り直し() {
  const ui = SpreadsheetApp.getUi();
  const 作品シート = SpreadsheetApp.getActive().getSheetByName(設定.作品シート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  if (ui.alert('確認', 'Works IDを1から連番に振り直します。\n③一括更新とセットで実行してください。\n続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const r = WorksID振り直しを実行(作品シート);
    ui.alert(`✅ ID振り直し完了\n変更: ${r.変更数}件\n\n③一括更新を実行してください。`);
  } finally { lock.releaseLock(); }
}

function メニュー_孤立削除() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const 作品シート = ss.getSheetByName(設定.作品シート名);
  const マスターシート = ss.getSheetByName(設定.マスターシート名);
  if (!作品シート || 作品シート.getLastRow() < 2) { ui.alert('Worksにデータがありません'); return; }
  if (!マスターシート) { ui.alert('マスターシートがありません'); return; }

  const 列マップ = 列番号を取得(マスターシート);
  if (ui.alert('確認', 'マスターシートのどの行からも参照されていないWorksエントリーを削除します。\n（作者名間違い等で生まれた不要なIDを掃除します）\n\n続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  const lock = LockService.getDocumentLock(); lock.waitLock(30000);
  try {
    const r = Works孤立エントリーを削除(作品シート, マスターシート, 列マップ);
    if (r.削除数 === 0) {
      ui.alert('✅ 孤立エントリーはありません！');
    } else {
      let msg = `✅ 孤立エントリー削除完了: ${r.削除数}件\n\n【削除されたエントリー】\n`;
      msg += r.削除リスト.slice(0, 20).join('\n');
      if (r.削除リスト.length > 20) msg += `\n...他 ${r.削除リスト.length - 20}件`;
      msg += '\n\n③一括更新を実行してIDを振り直すことを推奨します。';
      ui.alert(msg);
    }
  } finally { lock.releaseLock(); }
}

/* ============================================================
 * ① 商品コード確定発行（F/G/H列に書き込む）
 * ============================================================ */
function メニュー_確定発行() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定.マスターシート名);
  if (!sh) return;
  const 列マップ = 列番号を取得(sh);
  const 最終行 = データがある最終行を取得(sh, 列マップ);
  if (最終行 < 2) return;
  const 言語マップ = 言語マップを取得(ss), カテゴリマップ = カテゴリマップを取得(ss);
  if (Object.keys(言語マップ).length === 0 || Object.keys(カテゴリマップ).length === 0) {
    ui.alert('言語/カテゴリマスターがありません。⑤を実行してください。'); return;
  }

  const 全データ = sh.getRange(2, 1, 最終行 - 1, sh.getLastColumn()).getValues();
  const 作品シート = 作品シートを確保(ss);
  const lock = LockService.getDocumentLock(); lock.waitLock(60000);
  try {
    WorksKey再正規化を実行(作品シート);
    const 作品データ = 全作品データを読み込み(作品シート);
    const out作品ID = [], outSKU = [], out商品コード = [], outタイトル = [], outステータス = [];
    let 発行数 = 0;

    for (let i = 0; i < 全データ.length; i++) {
      const r = 全データ[i];
      const 取得 = (名前) => 正規化(r[(列マップ[名前] || 1) - 1]);
      const 取得生 = (名前) => r[(列マップ[名前] || 1) - 1];

      if (取得生(列名.発行チェック) !== true) {
        out作品ID.push([取得生(列名.作品ID) || '']); outSKU.push([取得生(列名.SKU) || '']);
        out商品コード.push([取得生(列名.商品コード) || '']); outタイトル.push([取得生(列名.タイトル) || '']);
        outステータス.push([取得生(列名.コードステータス) || '']);
        continue;
      }

      const 日本語タイトル = 取得(列名.日本語タイトル), 作者 = 取得(列名.作者);
      let 原題 = 取得(列名.原題);
      const 言語 = 取得(列名.言語), カテゴリ = 取得(列名.カテゴリ), 形態 = 取得(列名.形態);

      if (!日本語タイトル || !作者 || !言語 || !カテゴリ) {
        out作品ID.push([取得生(列名.作品ID) || '']); outSKU.push([取得生(列名.SKU) || '']);
        out商品コード.push([取得生(列名.商品コード) || '']); outタイトル.push([取得生(列名.タイトル) || '']);
        outステータス.push(['入力中...']);
        continue;
      }

      const 言語コード = 言語マップ[言語] || 'XX', カテゴリコード = カテゴリマップ[カテゴリ] || 'XX';
      const worksKey = WorksKeyを作る(日本語タイトル, 作者, 原題);
      let 作品ID = 作品データ.keyToId[worksKey];
      const 既存 = 作品データ.keyToData[worksKey];
      if (既存 && !原題 && 既存.原題) 原題 = 正規化(既存.原題);

      if (!作品ID) {
        // titleToKeyフォールバック
        if (!worksKey.startsWith('原題||')) {
          const titleKey = キー用正規化_(日本語タイトル);
          const 既存worksKey = 作品データ.titleToKey[titleKey];
          if (既存worksKey && 作品データ.keyToId[既存worksKey]) {
            作品ID = 作品データ.keyToId[既存worksKey];
            作品データ.keyToId[worksKey] = 作品ID;
          }
        }
        if (!作品ID) {
          作品データ.maxId++;
          作品ID = String(作品データ.maxId).padStart(4, '0');
          作品データ.keyToId[worksKey] = 作品ID;
          作品データ.newRows.push([worksKey, 作品ID, 日本語タイトル, 作者, 原題, '', '', '', '', '']);
          if (!worksKey.startsWith('原題||')) {
            const titleKey = キー用正規化_(日本語タイトル);
            if (!作品データ.titleToKey[titleKey]) 作品データ.titleToKey[titleKey] = worksKey;
          }
        }
      } else {
        作品ID = String(作品ID).padStart(4, '0');
      }

      const 巻数 = 数値変換(取得生(列名.単巻数));
      const セット終了巻 = 数値変換(取得生(列名.セット終了));
      if (巻数 != null) {
        if (!作品データ.keyToVols[worksKey]) 作品データ.keyToVols[worksKey] = new Set();
        作品データ.keyToVols[worksKey].add(巻数);
      }
      if (セット終了巻 != null) {
        const セット開始巻 = 数値変換(取得生(列名.セット開始)) || 1;
        if (!作品データ.keyToVols[worksKey]) 作品データ.keyToVols[worksKey] = new Set();
        for (let v = セット開始巻; v <= セット終了巻; v++) 作品データ.keyToVols[worksKey].add(v);
      }

      const SKU = SKUを生成(言語コード, 形態, 作品ID, カテゴリコード, 取得生(列名.単巻数), 取得生(列名.セット開始), 取得生(列名.セット終了));
      const タイトル = タイトルを段階生成(言語, カテゴリ, 形態, 日本語タイトル, 取得生(列名.単巻数), 取得生(列名.セット開始), 取得生(列名.セット終了), 作者, 原題, 取得(列名.特典メモ));

      out作品ID.push([作品ID]); outSKU.push([SKU]); out商品コード.push([SKU]);
      outタイトル.push([タイトル]); outステータス.push(['商品コード(発行済み確定)']);
      発行数++;
    }

    if (列マップ[列名.作品ID]) sh.getRange(2, 列マップ[列名.作品ID], out作品ID.length, 1).setValues(out作品ID);
    if (列マップ[列名.SKU]) sh.getRange(2, 列マップ[列名.SKU], outSKU.length, 1).setValues(outSKU);
    if (列マップ[列名.商品コード]) sh.getRange(2, 列マップ[列名.商品コード], out商品コード.length, 1).setValues(out商品コード);
    if (列マップ[列名.タイトル]) sh.getRange(2, 列マップ[列名.タイトル], outタイトル.length, 1).setValues(outタイトル);
    if (列マップ[列名.コードステータス]) sh.getRange(2, 列マップ[列名.コードステータス], outステータス.length, 1).setValues(outステータス);
    作品データを更新_確定(作品シート, 作品データ);
    ui.alert(`✅ 確定発行完了: ${発行数}件`);
  } finally { lock.releaseLock(); }
}

/* ============================================================
 * ② 商品コード削除
 * ============================================================ */
function メニュー_削除() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName(設定.マスターシート名);
  if (!sh) return;
  const 列マップ = 列番号を取得(sh);
  const 最終行 = データがある最終行を取得(sh, 列マップ);
  if (最終行 < 2) return;
  const 全データ = sh.getRange(2, 1, 最終行 - 1, sh.getLastColumn()).getValues();
  const 削除行 = [];
  for (let i = 0; i < 全データ.length; i++) {
    if (全データ[i][(列マップ[列名.発行チェック] || 1) - 1] === true) 削除行.push(i + 2);
  }
  if (削除行.length === 0) { ui.alert('チェックが入った行がありません'); return; }
  if (ui.alert('確認', `${削除行.length}件を削除します。続行？`, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  削除行.sort((a, b) => b - a);
  for (const 行 of 削除行) sh.deleteRow(行);
  ui.alert(`✅ 削除完了: ${削除行.length}件`);
}

/* ============================================================
 * ③ 既存データ一括更新（Works整理含む）
 * ============================================================ */
function メニュー_一括更新() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('確認', '全行を再生成します（Works重複整理+ID振り直し含む）。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(設定.マスターシート名);
  if (!sh) return;
  const 列マップ = 列番号を取得(sh);
  const 最終行 = データがある最終行を取得(sh, 列マップ);
  if (最終行 < 2) { ui.alert('データがありません'); return; }
  const 言語マップ = 言語マップを取得(ss), カテゴリマップ = カテゴリマップを取得(ss);
  if (Object.keys(言語マップ).length === 0 || Object.keys(カテゴリマップ).length === 0) {
    ui.alert('言語/カテゴリマスターがありません。⑤を実行してください。'); return;
  }

  const 全データ = sh.getRange(2, 1, 最終行 - 1, sh.getLastColumn()).getValues();
  const 作品シート = 作品シートを確保(ss);
  const lock = LockService.getDocumentLock(); lock.waitLock(60000);
  try {
    const 正規化結果 = WorksKey再正規化を実行(作品シート);
    const 孤立結果 = Works孤立エントリーを削除(作品シート, sh, 列マップ);
    const ID結果 = WorksID振り直しを実行(作品シート);
    const 作品データ = 全作品データを読み込み(作品シート);

    const out作品ID = [], outSKU = [], out商品コード = [], outタイトル = [];
    const outステータス = [], out作者 = [], out原題 = [];

    for (let i = 0; i < 全データ.length; i++) {
      const r = 全データ[i];
      const 取得 = (名前) => 正規化(r[(列マップ[名前] || 1) - 1]);
      const 取得生 = (名前) => r[(列マップ[名前] || 1) - 1];

      const 日本語タイトル = 取得(列名.日本語タイトル);
      const 言語 = 取得(列名.言語), カテゴリ = 取得(列名.カテゴリ);
      let 作者 = 取得(列名.作者), 原題 = 取得(列名.原題);
      const 形態 = 取得(列名.形態);

      if (!日本語タイトル || !言語 || !カテゴリ) {
        out作品ID.push([取得生(列名.作品ID) || '']); outSKU.push([取得生(列名.SKU) || '']);
        out商品コード.push([取得生(列名.商品コード) || '']); outタイトル.push([取得生(列名.タイトル) || '']);
        outステータス.push([取得生(列名.コードステータス) || '']);
        out作者.push([作者 || '']); out原題.push([原題 || '']);
        continue;
      }

      const 言語コード = 言語マップ[言語] || 'XX', カテゴリコード = カテゴリマップ[カテゴリ] || 'XX';
      let 作品ID = '';

      if (日本語タイトル && 作者) {
        const worksKey = WorksKeyを作る(日本語タイトル, 作者, 原題);
        const 既存 = 作品データ.keyToData[worksKey];
        if (既存) {
          if (!作者 && 既存.作者) 作者 = 正規化(既存.作者);
          if (!原題 && 既存.原題) 原題 = 正規化(既存.原題);
        }
        作品ID = 作品データ.keyToId[worksKey];
        if (!作品ID) {
          // titleToKeyフォールバック
          if (!worksKey.startsWith('原題||')) {
            const titleKey = キー用正規化_(日本語タイトル);
            const 既存worksKey = 作品データ.titleToKey[titleKey];
            if (既存worksKey && 作品データ.keyToId[既存worksKey]) {
              作品ID = 作品データ.keyToId[既存worksKey];
              作品データ.keyToId[worksKey] = 作品ID;
            }
          }
          if (!作品ID) {
            作品データ.maxId++;
            作品ID = String(作品データ.maxId).padStart(4, '0');
            作品データ.keyToId[worksKey] = 作品ID;
            作品データ.newRows.push([worksKey, 作品ID, 日本語タイトル, 作者, 原題, '', '', '', '', '']);
            if (!worksKey.startsWith('原題||')) {
              const titleKey = キー用正規化_(日本語タイトル);
              if (!作品データ.titleToKey[titleKey]) 作品データ.titleToKey[titleKey] = worksKey;
            }
          }
        } else {
          作品ID = String(作品ID).padStart(4, '0');
        }
      }

      const SKU = SKUを段階生成(言語コード, 形態, 作品ID, カテゴリコード, 取得生(列名.単巻数), 取得生(列名.セット開始), 取得生(列名.セット終了));
      const タイトル = タイトルを段階生成(言語, カテゴリ, 形態, 日本語タイトル, 取得生(列名.単巻数), 取得生(列名.セット開始), 取得生(列名.セット終了), 作者, 原題, 取得(列名.特典メモ));

      out作品ID.push([作品ID]); outSKU.push([SKU]); out商品コード.push([SKU]); outタイトル.push([タイトル]);
      outステータス.push([(日本語タイトル && 作者 && 言語 && カテゴリ) ? '商品コード(予約)' : '入力中...']);
      out作者.push([作者 || '']); out原題.push([原題 || '']);
    }

    if (列マップ[列名.作品ID]) sh.getRange(2, 列マップ[列名.作品ID], out作品ID.length, 1).setValues(out作品ID);
    if (列マップ[列名.SKU]) sh.getRange(2, 列マップ[列名.SKU], outSKU.length, 1).setValues(outSKU);
    if (列マップ[列名.商品コード]) sh.getRange(2, 列マップ[列名.商品コード], out商品コード.length, 1).setValues(out商品コード);
    if (列マップ[列名.タイトル]) sh.getRange(2, 列マップ[列名.タイトル], outタイトル.length, 1).setValues(outタイトル);
    if (列マップ[列名.コードステータス]) sh.getRange(2, 列マップ[列名.コードステータス], outステータス.length, 1).setValues(outステータス);
    if (列マップ[列名.作者]) sh.getRange(2, 列マップ[列名.作者], out作者.length, 1).setValues(out作者);
    if (列マップ[列名.原題]) sh.getRange(2, 列マップ[列名.原題], out原題.length, 1).setValues(out原題);
    作品データを更新_確定(作品シート, 作品データ);

    let msg = `✅ 一括更新完了: ${out作品ID.length}件`;
    if (正規化結果.統合数 > 0) msg += `\nWorks重複統合: ${正規化結果.統合数}件`;
    if (孤立結果.削除数 > 0) msg += `\nWorks孤立削除: ${孤立結果.削除数}件`;
    if (ID結果.変更数 > 0) msg += `\nID振り直し: ${ID結果.変更数}件`;
    ui.alert(msg);
  } finally { lock.releaseLock(); }
}

/* ============================================================
 * 作品シート関連
 * ============================================================ */
function 作品シートを確保(ss) {
  let sh = ss.getSheetByName(設定.作品シート名);
  if (!sh) {
    sh = ss.insertSheet(設定.作品シート名);
    sh.getRange(1, 1, 1, 設定.作品列数).setValues([設定.作品ヘッダー]);
  } else {
    const 現在列数 = sh.getLastColumn();
    if (現在列数 < 設定.作品列数) {
      for (let i = 現在列数; i < 設定.作品列数; i++) sh.getRange(1, i + 1).setValue(設定.作品ヘッダー[i]);
    }
  }
  return sh;
}

/* ✅ v10.10: titleToKey / ISBNToKey も読み込む */
function 全作品データを読み込み(作品シート) {
  const result = {
    keyToId: {}, keyToData: {}, keyToRow: {}, keyToVols: {},
    keyTo予約最新巻: {}, maxId: 0, newRows: [], keyUpdates: [],
    titleToKey: {}   // 日本語タイトル正規化キー → worksKey（フォールバック用）
  };
  const 最終行 = 作品シート.getLastRow();
  if (最終行 < 2) return result;

  const 列数 = Math.max(作品シート.getLastColumn(), 設定.作品列数);
  const データ = 作品シート.getRange(2, 1, 最終行 - 1, Math.min(列数, 設定.作品列数)).getValues();

  for (let i = 0; i < データ.length; i++) {
    const r = データ[i];
    const idStr = String(r[1] == null ? '' : r[1]).trim();
    if (!idStr) continue;
    const t = 正規化(r[2] || ''), a = 正規化(r[3] || ''), o = 正規化(r[4] || '');
    let key = (t && a) ? WorksKeyを作る(t, a, o) : String(r[0] || '').trim();
    if (!key) continue;

    const 保存済みKey = String(r[0] || '').trim();
    if (保存済みKey !== key) result.keyUpdates.push({ 行: i + 2, key });

    if (result.keyToId[key]) {
      const 既存ID = parseInt(result.keyToId[key], 10), 新ID = parseInt(idStr, 10);
      if (!isNaN(新ID) && !isNaN(既存ID) && 新ID >= 既存ID) {
        String(r[5] || '').split(',').forEach(v => { const n = parseInt(String(v).trim(), 10); if (!isNaN(n)) { if (!result.keyToVols[key]) result.keyToVols[key] = new Set(); result.keyToVols[key].add(n); } });
        const num = parseInt(idStr, 10); if (!isNaN(num) && num > result.maxId) result.maxId = num;
        continue;
      }
    }

    result.keyToId[key] = idStr.padStart(4, '0');
    result.keyToRow[key] = i + 2;
    result.keyToData[key] = { 作者: 正規化(r[3] || ''), 原題: 正規化(r[4] || '') };
    if (!result.keyToVols[key]) result.keyToVols[key] = new Set();
    String(r[5] || '').split(',').forEach(v => { const n = parseInt(String(v).trim(), 10); if (!isNaN(n)) result.keyToVols[key].add(n); });

    const 予約巻 = parseInt(String(r[8] || '0'), 10);
    if (!isNaN(予約巻) && 予約巻 > 0) result.keyTo予約最新巻[key] = 予約巻;

    const num = parseInt(idStr, 10); if (!isNaN(num) && num > result.maxId) result.maxId = num;

    // titleToKey登録（フォールバックキーのみ: 原題キーはスキップ）
    if (t && !key.startsWith('原題||')) {
      const titleKey = キー用正規化_(t);
      if (!result.titleToKey[titleKey]) result.titleToKey[titleKey] = key;
    }
  }
  return result;
}

/* ============================================================
 * 作品データ更新（確定版: ①確定発行・③一括更新で使用）
 * F/G/H列（確定データ）に書き込む
 * ============================================================ */
function 作品データを更新_確定(作品シート, 作品データ) {
  if (作品データ.keyUpdates && 作品データ.keyUpdates.length > 0) {
    for (const upd of 作品データ.keyUpdates) 作品シート.getRange(upd.行, 1).setValue(upd.key);
    作品データ.keyUpdates = [];
  }
  if (作品データ.newRows.length > 0) {
    const 開始行 = Math.max(2, 作品シート.getLastRow() + 1);
    作品シート.getRange(開始行, 1, 作品データ.newRows.length, 設定.作品列数).setValues(作品データ.newRows);
    for (let i = 0; i < 作品データ.newRows.length; i++) {
      const key = String(作品データ.newRows[i][0] || '').trim();
      if (key) 作品データ.keyToRow[key] = 開始行 + i;
    }
  }
  for (const [key, vols] of Object.entries(作品データ.keyToVols)) {
    const 行番号 = 作品データ.keyToRow[key];
    if (!行番号 || !vols || vols.size === 0) continue;
    const arr = Array.from(vols).sort((a, b) => a - b);
    作品シート.getRange(行番号, 6).setValue(arr.join(','));
    作品シート.getRange(行番号, 7).setValue(Math.max(...arr));
    作品シート.getRange(行番号, 8).setValue(new Date());
  }
  作品データ.newRows = [];
}

/* ============================================================
 * SKU / タイトル生成
 * ============================================================ */
function SKUを段階生成(言語コード, 形態, 作品ID, カテゴリコード, 単巻数, セット開始, セット終了) {
  const parts = [];
  if (言語コード) parts.push(言語コード);
  const 形態プレ = 設定.形態プレフィックス[String(形態 || '').trim()] || '';
  if (形態プレ) parts.push(形態プレ);
  if (作品ID) { parts.push(String(作品ID).padStart(4, '0')); } else if (言語コード) { parts.push('????'); }
  if (カテゴリコード) parts.push('-' + カテゴリコード);
  const sf = String(セット開始 || '').trim(), st = String(セット終了 || '').trim();
  if (sf && st) { parts.push('-' + sf.padStart(2, '0') + st.padStart(2, '0')); }
  else { const v = String(単巻数 || '').trim(); if (v) { const n = parseInt(v, 10); parts.push('-' + (!isNaN(n) ? String(n).padStart(2, '0') : v)); } }
  return parts.join('');
}

function SKUを生成(言語コード, 形態, 作品ID, カテゴリコード, 単巻数, セット開始, セット終了) {
  const 形態プレ = 設定.形態プレフィックス[String(形態 || '').trim()] || '';
  const base = String(言語コード || '') + 形態プレ + String(作品ID || '').padStart(4, '0') + '-' + String(カテゴリコード || '');
  const sf = String(セット開始 || '').trim(), st = String(セット終了 || '').trim();
  if (sf && st) return base + '-' + sf.padStart(2, '0') + st.padStart(2, '0');
  const v = String(単巻数 || '').trim();
  if (v) { const n = parseInt(v, 10); return base + '-' + (!isNaN(n) ? String(n).padStart(2, '0') : v); }
  return base;
}

function タイトルを段階生成(言語, カテゴリ, 形態, 日本語タイトル, 単巻数, セット開始, セット終了, 作者, 原題, 特典メモ) {
  const parts = [];
  const head = [言語 ? `${言語}版` : '', カテゴリ].filter(Boolean).join(' ');
  if (head) parts.push(head);
  if (形態) parts.push(`(${形態})`);
  if (日本語タイトル) parts.push(`『${日本語タイトル}』`);

  const sf = String(セット開始 || '').trim(), st = String(セット終了 || '').trim();
  if (sf && st) { parts.push(`${sf.padStart(2, '0')}-${st.padStart(2, '0')}巻`); }
  else { const v = String(単巻数 || '').trim(); if (v) { const n = parseInt(v, 10); parts.push(`第${!isNaN(n) ? String(n).padStart(2, '0') : v}巻`); } }

  if (作者) parts.push(`著：${作者}`);

  if (原題) {
    let 原題巻数 = '';
    if (sf && st) { 原題巻数 = `${sf}-${st}`; }
    else { const v = String(単巻数 || '').trim(); if (v) { const n = parseInt(v, 10); 原題巻数 = !isNaN(n) ? String(n) : v; } }

    let 台湾版形態 = '';
    if (形態 === '初版限定版' || 形態 === '初回限定版') { 台湾版形態 = '首刷限定版'; }
    else if (形態 === '特装版') { 台湾版形態 = '特裝版'; }
    else if (形態 && 形態 !== '通常' && 形態 !== '通常版') { 台湾版形態 = 形態; }

    let 構築原題 = 原題;
    if (原題巻数) 構築原題 += 原題巻数;
    if (台湾版形態) 構築原題 += ` (${台湾版形態})`;
    parts.push(構築原題);
  }

  if (特典メモ) { parts.push(特典メモ.startsWith('※') ? 特典メモ : `※${特典メモ}`); }
  else if (設定.特典自動付与 && 形態 && 設定.形態別特典[形態]) { parts.push(設定.形態別特典[形態]); }

  return parts.join(' ');
}

/* ============================================================
 * ヘルパー
 * ============================================================ */
function 言語マップを取得(ss) {
  const map = {}, sh = ss.getSheetByName(設定.言語マスター名);
  if (!sh || sh.getLastRow() < 2) return map;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (const [n, c] of data) { const name = String(n || '').trim(), code = String(c || '').trim(); if (name && code) map[name] = code; }
  return map;
}

function カテゴリマップを取得(ss) {
  const map = {}, sh = ss.getSheetByName(設定.カテゴリマスター名);
  if (!sh || sh.getLastRow() < 2) return map;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  for (const [n, c] of data) { const name = String(n || '').trim(), code = String(c || '').trim(); if (name && code) map[name] = code; }
  return map;
}

function 正規化(v) { return String(v || '').replace(/\u3000/g, ' ').replace(/\s+/g, ' ').trim(); }
function 数値変換(v) { const n = parseInt(String(v || '').trim(), 10); return Number.isFinite(n) ? n : null; }

function データがある最終行を取得(sh, 列マップ) {
  const 基準列 = 列マップ[列名.日本語タイトル] || 1;
  const 最大行 = sh.getLastRow();
  if (最大行 <= 1) return 1;
  const data = sh.getRange(2, 基準列, 最大行 - 1, 1).getValues();
  for (let i = data.length - 1; i >= 0; i--) { if (data[i][0] !== '' && data[i][0] !== null && data[i][0] !== undefined) return i + 2; }
  return 1;
}

/* ============================================================
 * WorksKey 正規化
 * ============================================================ */
function キー用正規化_(v) {
  let s = 正規化(v).toLowerCase();
  s = s.replace(/^著[:：]\s*/g, '').replace(/^作[:：]\s*/g, '')
    .replace(/[［\[][^］\]]*[］\]]/g, '').replace(/[（\(][^）\)]*[）\)]/g, '').replace(/[｛\{][^｝\}]*[｝\}]/g, '')
    .replace(/[・･]/g, ' ').replace(/[～〜~]/g, '').replace(/[：:]/g, '').replace(/[、,]/g, '')
    .replace(/[。\.]/g, '').replace(/[！!]/g, '').replace(/[？?]/g, '').replace(/[『』「」]/g, '').replace(/["'"]/g, '')
    .replace(/[‐−–—―]/g, '-')
    .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/[Ａ-Ｚａ-ｚ]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/\s*[\/／]\s*/g, ' / ').replace(/\s+/g, ' ').replace(/-+/g, '-').trim();
  return s;
}

/* ============================================================
 * ④⑤⑥ マスター管理
 * ============================================================ */
function メニュー_Works初期化() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('警告', 'Worksを全削除します。続行？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(設定.作品シート名);
  if (sh) { const last = sh.getLastRow(); if (last > 1) sh.deleteRows(2, last - 1); }
  else { sh = ss.insertSheet(設定.作品シート名); }
  sh.getRange(1, 1, 1, 設定.作品列数).setValues([設定.作品ヘッダー]);
  ui.alert('✅ Works初期化完了');
}

function メニュー_マスター作成() {
  const ss = SpreadsheetApp.getActive(); const created = [];
  if (!ss.getSheetByName(設定.言語マスター名)) {
    const sh = ss.insertSheet(設定.言語マスター名);
    sh.getRange(1, 1, 1, 設定.言語ヘッダー.length).setValues([設定.言語ヘッダー]);
    sh.getRange(2, 1, 設定.言語初期値.length, 設定.言語初期値[0].length).setValues(設定.言語初期値);
    created.push('言語');
  }
  if (!ss.getSheetByName(設定.カテゴリマスター名)) {
    const sh = ss.insertSheet(設定.カテゴリマスター名);
    sh.getRange(1, 1, 1, 設定.カテゴリヘッダー.length).setValues([設定.カテゴリヘッダー]);
    sh.getRange(2, 1, 設定.カテゴリ初期値.length, 設定.カテゴリ初期値[0].length).setValues(設定.カテゴリ初期値);
    created.push('カテゴリ');
  }
  if (!ss.getSheetByName(設定.作品シート名)) {
    const sh = ss.insertSheet(設定.作品シート名);
    sh.getRange(1, 1, 1, 設定.作品列数).setValues([設定.作品ヘッダー]);
    created.push('Works');
  }
  if (created.length > 0) { メニュー_プルダウン更新(); SpreadsheetApp.getUi().alert(`✅ 作成: ${created.join(', ')}`); }
  else { SpreadsheetApp.getUi().alert('既に全て存在します'); }
}

function メニュー_プルダウン更新() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  if (!sh) return; // ← 最初にチェック

  const 設定 = シート設定を取得(sh.getName());
  if (!設定) {
    SpreadsheetApp.getUi().alert('このシートはプルダウン更新の対象外です');
    return;
  }

  const 列マップ = 列番号を取得(sh);
  const 最終行 = Math.max(データがある最終行を取得(sh, 列マップ) + 100, 200);

  const 言語データ = 言語データを色付きで取得(ss);
  if (言語データ.values.length > 0 && 列マップ[列名.言語]) {
    sh.getRange(2, 列マップ[列名.言語], 最終行 - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(言語データ.values, true).build());
    条件付き書式をクリア(sh, 列マップ[列名.言語]);
    条件付き書式を設定(sh, 列マップ[列名.言語], 言語データ.values, 言語データ.colors, 最終行);
  }

  const カテゴリデータ = カテゴリデータを色付きで取得(ss);
  if (カテゴリデータ.values.length > 0 && 列マップ[列名.カテゴリ]) {
    sh.getRange(2, 列マップ[列名.カテゴリ], 最終行 - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(カテゴリデータ.values, true).build());
    条件付き書式をクリア(sh, 列マップ[列名.カテゴリ]);
    条件付き書式を設定(sh, 列マップ[列名.カテゴリ], カテゴリデータ.values, カテゴリデータ.colors, 最終行);
  }

  // ← 形態も追加
  const 形態データ = 形態データを色付きで取得(ss);
  if (形態データ.values.length > 0 && 列マップ[列名.形態]) {
    sh.getRange(2, 列マップ[列名.形態], 最終行 - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(形態データ.values, true).build());
    条件付き書式をクリア(sh, 列マップ[列名.形態]);
    条件付き書式を設定(sh, 列マップ[列名.形態], 形態データ.values, 形態データ.colors, 最終行);
  }

  ss.toast('プルダウン更新完了', '商品コード管理', 3);
}
function 言語データを色付きで取得(ss) {
  const result = { values: [], colors: [] }, sh = ss.getSheetByName(設定.言語マスター名);
  if (!sh || sh.getLastRow() < 2) return result;
  let ci = 0;
  for (const row of sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()) {
    const n = String(row[0] || '').trim(); if (!n) continue;
    result.values.push(n); result.colors.push(String(row[2] || '').trim() || 設定.色パレット[ci++ % 設定.色パレット.length]);
  }
  return result;
}

function カテゴリデータを色付きで取得(ss) {
  const result = { values: [], colors: [] }, sh = ss.getSheetByName(設定.カテゴリマスター名);
  if (!sh || sh.getLastRow() < 2) return result;
  let ci = 0;
  for (const row of sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()) {
    const n = String(row[0] || '').trim(); if (!n) continue;
    result.values.push(n); result.colors.push(String(row[2] || '').trim() || 設定.色パレット[ci++ % 設定.色パレット.length]);
  }
  return result;
}

function 条件付き書式をクリア(sh, colNum) {
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().filter(rule =>
    !rule.getRanges().some(rng => {
      const c1 = rng.getColumn();
      const c2 = c1 + rng.getNumColumns() - 1;
      // 開始列と終了列が両方ともcolNumと一致する単一列ルールのみ削除
      return c1 === colNum && c2 === colNum;
    })
  ));
}

/* ============================================================
 * 自己書き込みループ防止（10秒タイムアウト付き）
 * ============================================================ */
function 自己更新中か_() {
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty('__SELF_EDIT_LOCK__') !== '1') return false;
  const ts = Number(props.getProperty('__SELF_EDIT_TS__') || '0');
  if (ts > 0 && (Date.now() - ts) > 10000) { props.deleteProperty('__SELF_EDIT_LOCK__'); props.deleteProperty('__SELF_EDIT_TS__'); return false; }
  return true;
}

function 自己更新を開始_() {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('__SELF_EDIT_LOCK__', '1');
  props.setProperty('__SELF_EDIT_TS__', String(Date.now()));
}

function 自己更新を終了_() {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty('__SELF_EDIT_LOCK__');
  props.deleteProperty('__SELF_EDIT_TS__');
}

function シート設定を取得(シート名) {
  const マップ = {
    '台湾_書籍（コミック/小説/設定集）': 設定,
    '韓国マンガ': 設定,  // ← 追加（韓国マンガ用設定オブジェクトができるまで暫定）
  };
  return マップ[シート名] || null;
}

function 形態データを色付きで取得(ss) {
  const result = { values: [], colors: [] };
  const sh = ss.getSheetByName('形態マスターシート');
  if (!sh || sh.getLastRow() < 2) return result;
  let ci = 0;
  for (const row of sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()) {
    const n = String(row[0] || '').trim();
    if (!n) continue;
    result.values.push(n);
    result.colors.push(String(row[3] || '').trim() || 設定.色パレット[ci++ % 設定.色パレット.length]);
  }
  return result;
}

function 条件付き書式を設定(sh, colNum, values, colors, lastRow) {
  const rules = sh.getConditionalFormatRules();
  const range = sh.getRange(2, colNum, lastRow - 1, 1);
  for (let i = 0; i < values.length; i++) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(values[i])
        .setBackground(colors[i] || '#fff')
        .setRanges([range])
        .build()
    );
  }
  sh.setConditionalFormatRules(rules);
}
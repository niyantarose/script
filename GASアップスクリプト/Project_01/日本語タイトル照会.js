/**
 * 選択行の「日本語タイトル」をシートから直接サイト照会して埋める（aniList → MangaUpdates）。
 *
 * 背景: 照会は従来、拡張機能のスクレイプ時にしか走らず、シート上で原題タイトルを
 * 修正しても照会をやり直す手段が無かった（例: 「今生我來當家主3+」→「今生我來當家主」に
 * 直しても日本語タイトルが空のまま）。このメニューで任意の行をいつでも再照会できる。
 *
 * 誤マッチ防止: クエリ（原題）が作品の登録名（タイトル/別名）と正規化一致した作品だけを
 * 採用し、その作品の「かな入り」の名前を日本語タイトルとして返す。
 * 副題付きで一致しない場合は、末尾の語を落として再試行する（例: 后宮的Ω王子 雪花之章 →
 * 后宮的Ω王子）。この場合も登録名との完全一致が必要なので誤マッチはしない。
 */

// 図形ボタン割当可（末尾アンダースコアなし）
function 台湾_日本語タイトルを照会() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  const sheetName = sh.getName();

  let 設定 = null;
  if (sheetName === '台湾まんが' && typeof 設定_台湾まんが !== 'undefined') 設定 = 設定_台湾まんが;
  if (sheetName === '台湾書籍その他' && typeof 設定_台湾書籍その他 !== 'undefined') 設定 = 設定_台湾書籍その他;
  if (!設定) {
    ui.alert('台湾まんが / 台湾書籍その他 のシートで、照会したい行を選択して実行してください');
    return;
  }

  const 列 = 台湾書籍系_列マップを取得_(sh);
  const cn = 設定.列名 || {};
  const n日本語 = 台湾書籍系_実列名を取得_(列, [cn.日本語タイトル, '日本語タイトル']);
  const n原題 = 台湾書籍系_実列名を取得_(列, [cn.原題, '原題タイトル']);
  const n原題商品 = 台湾書籍系_実列名を取得_(列, [cn.原題商品タイトル, '原題商品タイトル']);
  const col日本語 = n日本語 ? 列[n日本語] : 0;
  const col原題 = n原題 ? 列[n原題] : 0;
  const col原題商品 = n原題商品 ? 列[n原題商品] : 0;
  if (!col日本語 || !col原題) {
    ui.alert('日本語タイトル / 原題タイトル 列が見つかりません');
    return;
  }

  // 選択範囲から対象行を集める（複数範囲・複数行対応）
  const rangeList = sh.getActiveRangeList();
  const ranges = rangeList ? rangeList.getRanges() : [sh.getActiveRange()].filter(Boolean);
  const rowSet = {};
  ranges.forEach(function (r) {
    for (let i = r.getRow(); i <= r.getLastRow(); i += 1) {
      if (i >= 2) rowSet[i] = true;
    }
  });
  const rows = Object.keys(rowSet).map(Number).sort(function (a, b) { return a - b; });
  if (!rows.length) {
    ui.alert('2行目以降の行を選択してから実行してください');
    return;
  }
  if (rows.length > 30) {
    ui.alert('一度に照会できるのは30行までです（外部API負荷対策）。選択を減らしてください');
    return;
  }

  const 結果 = [];
  let 取得数 = 0;

  rows.forEach(function (row) {
    const 現況 = String(sh.getRange(row, col日本語).getDisplayValue() || '').trim();
    if (現況 && !/^(登録なし|照会失敗|未照会|MU登録あり)/.test(現況)) {
      結果.push(row + '行: 入力済みのためスキップ（' + 現況 + '）');
      return;
    }
    const 原題 = String(sh.getRange(row, col原題).getDisplayValue() || '').trim()
      || (col原題商品 ? String(sh.getRange(row, col原題商品).getDisplayValue() || '').trim() : '');
    if (!原題) {
      結果.push(row + '行: 原題タイトルが空のためスキップ');
      return;
    }

    let jp = '';
    let 供給元 = '';
    const クエリ候補 = 台湾照会_クエリ候補_(原題);
    for (let q = 0; q < クエリ候補.length && !jp; q += 1) {
      try {
        jp = 台湾照会_aniListで日本語題_(クエリ候補[q]);
        if (jp) 供給元 = 'aniList';
      } catch (e) { /* 続行してMUへ */ }
      if (!jp) {
        try {
          jp = 台湾照会_MUで日本語題_(クエリ候補[q]);
          if (jp) 供給元 = 'MangaUpdates';
        } catch (e) { /* 続行してBangumiへ */ }
      }
      if (!jp) {
        try {
          jp = 台湾照会_Bangumiで日本語題_(クエリ候補[q]);
          if (jp) 供給元 = 'Bangumi';
        } catch (e) { /* 次の候補へ */ }
      }
    }

    if (jp) {
      sh.getRange(row, col日本語).setValue(jp);
      // タイトル生成・Works反映まで一気に（onEditと同じ安全エンジン）
      try {
        台湾書籍系_1行補完_共通_(sh, row, 設定, { Works新規作成: true });
      } catch (e) { /* 補完失敗しても照会結果は残す */ }
      取得数 += 1;
      結果.push(row + '行: ✅ ' + jp + '（' + 供給元 + '）');
    } else {
      結果.push(row + '行: ❌ 見つからず（' + 原題 + '）→ DB未収録の可能性。手入力してください');
    }
    Utilities.sleep(250);
  });

  ui.alert(
    '日本語タイトル照会 結果（取得 ' + 取得数 + '/' + rows.length + '件）',
    結果.join('\n'),
    ui.ButtonSet.OK
  );
}

/** クエリ候補: 原題そのまま → 末尾の空白区切り語を1つ落とす → 2つ落とす（副題対策） */
function 台湾照会_クエリ候補_(原題) {
  const base = String(原題 || '').trim();
  const out = [base];
  const parts = base.split(/[\s　]+/).filter(Boolean);
  if (parts.length >= 2) out.push(parts.slice(0, -1).join(' '));
  if (parts.length >= 3) out.push(parts.slice(0, -2).join(' '));
  return out;
}

/** 照合用正規化キー（空白・記号・約物を除去して小文字化） */
function 台湾照会_キー_(v) {
  return String(v || '')
    .toLowerCase()
    .replace(/[\s　]+/g, '')
    .replace(/[!！?？.,，。、・·:：;；'"’‘“”（）()【】\[\]{}~～\-–—_+＋*＊/／\\|｜]/g, '');
}

function 台湾照会_かな有_(v) {
  return /[ぁ-ゖァ-ヺ]/.test(String(v || ''));
}

/**
 * aniList: クエリが作品のタイトル/別名と正規化一致した作品から「かな入り」の名前を返す。
 * （例: 今生我來當家主 → 別名 今世は当主になります）
 */
function 台湾照会_aniListで日本語題_(query) {
  const payload = JSON.stringify({
    query: 'query($s:String){Page(perPage:6){media(search:$s,type:MANGA){countryOfOrigin title{native romaji english}synonyms}}}',
    variables: { s: query },
  });
  const res = UrlFetchApp.fetch('https://graphql.anilist.co', {
    method: 'post',
    contentType: 'application/json',
    payload: payload,
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) return '';
  let data;
  try { data = JSON.parse(res.getContentText()); } catch (e) { return ''; }
  const media = data && data.data && data.data.Page && Array.isArray(data.data.Page.media)
    ? data.data.Page.media
    : [];
  const qk = 台湾照会_キー_(query);
  if (!qk) return '';
  for (let i = 0; i < media.length; i += 1) {
    const m = media[i] || {};
    const t = m.title || {};
    const names = [t.native, t.romaji, t.english].concat(Array.isArray(m.synonyms) ? m.synonyms : []).filter(Boolean);
    const 一致 = names.some(function (n) { return 台湾照会_キー_(n) === qk; });
    if (!一致) continue; // クエリが登録名と一致した作品だけ採用（誤マッチ防止）
    // 日本原作（countryOfOrigin=JP）なら native は日本語題そのもの。
    // 漢字のみの題（例: 呪術廻戦、怪獣８号）もかな縛りなしで安全に採用できる。
    // ただしクエリのエコー（クエリ自身がnativeに一致しただけ）は除外。
    const nat = String(t.native || '').trim();
    if (nat && String(m.countryOfOrigin || '') === 'JP' && 台湾照会_キー_(nat) !== qk) {
      return nat;
    }
    // 韓国・中国原作は native が原語（ハングル/中文）なので、synonyms から
    // かな入りの日本語ライセンス題を探す（例: 이번 생은…→今世は当主になります）。
    const 候補 = [t.native].concat(Array.isArray(m.synonyms) ? m.synonyms : []).filter(Boolean);
    for (let c = 0; c < 候補.length; c += 1) {
      const cand = String(候補[c]).trim();
      if (台湾照会_キー_(cand) === qk) continue; // クエリのエコーは除外
      if (台湾照会_かな有_(cand)) return cand;
    }
  }
  return '';
}

/** 漢字（CJK統合漢字）だけを重複なしで集める */
function 台湾照会_漢字集合_(v) {
  const out = {};
  const s = String(v || '');
  for (let i = 0; i < s.length; i += 1) {
    const c = s[i];
    if (/[㐀-䶿一-鿿]/.test(c)) out[c] = true;
  }
  return out;
}

/**
 * Bangumi (bgm.tv): 中華圏ACGデータベース。「中文名⇔日本語原名」ペアを人手登録しており、
 * aniList/MUに無い日本原作（BL・マイナー作品）に強い（例: 伏魔師祓清 → 悪祓士のキヨシくん、
 * 男孩子氣的女友超級可愛 → ボーイッシュ彼女が可愛すぎる）。
 * 検証: 中文名とクエリの漢字重なり（簡繁字体差があるため共通漢字数で判定）＋原名かな入り必須。
 */
function 台湾照会_Bangumiで日本語題_(query) {
  const q = String(query || '').trim();
  if (!q) return '';
  const qHan = 台湾照会_漢字集合_(q);
  const qHanCount = Object.keys(qHan).length;
  if (!qHanCount) return '';

  const res = UrlFetchApp.fetch(
    'https://api.bgm.tv/search/subject/' + encodeURIComponent(q) + '?type=1&responseGroup=small&max_results=6',
    { muteHttpExceptions: true }
  );
  if (res.getResponseCode() !== 200) return '';
  let data;
  try { data = JSON.parse(res.getContentText()); } catch (e) { return ''; }
  const list = data && Array.isArray(data.list) ? data.list : [];

  const 必要重なり = Math.max(2, Math.ceil(qHanCount * 0.5));
  for (let i = 0; i < list.length; i += 1) {
    const item = list[i] || {};
    const 原名 = String(item.name || '').trim();
    const 中文名 = String(item.name_cn || '').trim();
    if (!原名 || !中文名) continue; // 中文名が無い項目（巻分エントリ等）は検証不能なのでスキップ
    if (!台湾照会_かな有_(原名)) continue; // 日本語原名（かな入り）のみ。ハングル・中文名は不採用
    // 簡体字/繁体字の字体差があっても共通の漢字は多く残る（伏魔師祓清⇔伏魔师祓清→4字共通）
    const cnHan = 台湾照会_漢字集合_(中文名);
    let 共通 = 0;
    for (const c in qHan) {
      if (Object.prototype.hasOwnProperty.call(cnHan, c)) 共通 += 1;
    }
    if (共通 < 必要重なり) continue;
    // 末尾の巻数・上下表記を除去して作品名として返す（例: 后宮のオメガ (上) → 后宮のオメガ）
    return 原名
      .replace(/\s*[（(]\s*(?:\d{1,3}|上|下|前編|後編)\s*[)）]\s*$/u, '')
      .trim();
  }
  return '';
}

/**
 * MangaUpdates: 検索ヒットの詳細（関連名）にクエリと正規化一致する名前がある作品から
 * 「かな入り」の名前を返す。（例: 后宮的Ω王子 → 后宮のオメガ）
 */
function 台湾照会_MUで日本語題_(query) {
  const qk = 台湾照会_キー_(query);
  if (!qk) return '';
  const res = UrlFetchApp.fetch('https://api.mangaupdates.com/v1/series/search', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ search: query, perpage: 5 }),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) return '';
  let data;
  try { data = JSON.parse(res.getContentText()); } catch (e) { return ''; }
  const results = data && Array.isArray(data.results) ? data.results : [];
  for (let i = 0; i < results.length; i += 1) {
    const rec = results[i] && results[i].record;
    const id = rec && rec.series_id;
    if (!id) continue;
    const dRes = UrlFetchApp.fetch('https://api.mangaupdates.com/v1/series/' + encodeURIComponent(id), {
      muteHttpExceptions: true,
    });
    if (dRes.getResponseCode() !== 200) continue;
    let d;
    try { d = JSON.parse(dRes.getContentText()); } catch (e) { continue; }
    const assoc = Array.isArray(d.associated)
      ? d.associated.map(function (a) { return a && a.title; }).filter(Boolean)
      : [];
    const names = [d.title].concat(assoc).filter(Boolean);
    const 一致 = names.some(function (n) { return 台湾照会_キー_(n) === qk; });
    if (!一致) { Utilities.sleep(150); continue; }
    for (let j = 0; j < names.length; j += 1) {
      if (台湾照会_かな有_(names[j])) return String(names[j]).trim();
    }
    Utilities.sleep(150);
  }
  return '';
}

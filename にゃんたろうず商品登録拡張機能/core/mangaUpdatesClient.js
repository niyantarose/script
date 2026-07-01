/**
 * MangaUpdates 公式 API を拡張コンテキストから直接呼び出す（GAS からの 403 回避用）。
 * OpenAPI: POST /v1/series/search , GET /v1/series/{id}
 */
(function registerMangaUpdatesClient(global) {
  'use strict';

  var MU_API = 'https://api.mangaupdates.com/v1';
  var ANILIST_URL = 'https://graphql.anilist.co';
  // サイトHTMLは API が別名ヒットしないときの検索フォールバックのみ（拡張UIへは注入しない）。
  // それでも取れない場合は AniList 経由で補う。
  var MU_SITE_SEARCH = 'https://www.mangaupdates.com/site/search/result?search=';
  var MU_SEARCH_TOP = 8;
  var _searchCache = new Map();
  var _detailCache = new Map();
  var _CACHE_MAX = 200;

  function cacheTake(map, key, value) {
    if (map.size >= _CACHE_MAX) {
      var first = map.keys().next().value;
      map.delete(first);
    }
    map.set(key, value);
  }

  function hasKanaJapaneseSignal(value) {
    return /[\u3041-\u3096\u30A1-\u30FA\u30FC\u309D\u309E\u3005\u3006\u3007]/.test(String(value || ''));
  }

  function hasJapaneseTitleSignal(value) {
    var s = String(value || '');
    if (hasKanaJapaneseSignal(s)) return true;
    // カナ無しの日本語漢字題（例: K-9 警視庁公安部公安第9課異能対策係）
    if (/[\u3400-\u4DBF\u4E00-\u9FFF]/.test(s)) return true;
    return false;
  }

  function pickPreferredJapaneseCandidate(candidates, echoQueries) {
    var list = candidates || [];
    if (!list.length) return '';
    // 原語題（中文原題など）がクエリと一致してエコーされている場合は除外する。
    // 例: 中華漫画の Associated に「中文原題 + 日本語ライセンス題」が並ぶケースで
    // 中文原題ではなく日本語題を採用するため。字体差のある日本語漢字題は除外しない。
    var skip = {};
    var qs = Array.isArray(echoQueries) ? echoQueries : (echoQueries ? [echoQueries] : []);
    var i;
    for (i = 0; i < qs.length; i += 1) {
      var qk = normalizeTitleKey(qs[i]);
      if (qk) skip[qk] = true;
    }
    var pool = [];
    for (i = 0; i < list.length; i += 1) {
      var key = normalizeTitleKey(list[i]);
      if (key && skip[key]) continue;
      pool.push(list[i]);
    }
    if (!pool.length) pool = list;
    // MangaUpdates の Associated Names 表示順（先頭）を尊重して採用する。
    return String(pool[0] || '').trim();
  }

  function uniqAniListQueries(values) {
    var seen = {};
    var out = [];
    var i;
    var s;
    var k;
    for (i = 0; i < (values || []).length; i += 1) {
      s = String(values[i] || '').trim();
      if (!s || s.length < 2) continue;
      k = s.toLowerCase();
      if (seen[k]) continue;
      seen[k] = true;
      out.push(s);
    }
    return out;
  }

  async function aniListFetchNativeJapanese(searchStrings, validationQueries) {
    var queries = uniqAniListQueries(searchStrings).slice(0, 5);
    // 元クエリ（中華題など）に漢字があれば、AniList が返す native タイトルが
    // その作品と無関係な誤答（例: 別作品「バトルファッカーB子」）でないか漢字重なりで検証する。
    // 検索は広めのクエリ群で行うが、採用判定は必ず元クエリと突き合わせる。
    var valQueries = Array.isArray(validationQueries)
      ? validationQueries
      : (validationQueries ? [validationQueries] : []);
    var requireHan = false;
    var vci;
    for (vci = 0; vci < valQueries.length; vci += 1) {
      if (collectHanChars(valQueries[vci]).length) { requireHan = true; break; }
    }
    // 検証クエリ（中華題など）の正規化キー集合。
    var valKeySet = {};
    for (vci = 0; vci < valQueries.length; vci += 1) {
      var vk = normalizeTitleKey(valQueries[vci]);
      if (vk) valKeySet[vk] = true;
    }
    // AniList が返す作品のタイトル/別名(synonyms)にクエリが含まれる＝作品一致が確定している。
    // 例: 青春之箱 は AniList の「アオのハコ」の synonyms に入っている。
    function seriesMatchesQuery(m) {
      if (!m) return false;
      var titles = [];
      if (m.title) titles.push(m.title.native, m.title.romaji, m.title.english);
      if (Array.isArray(m.synonyms)) titles = titles.concat(m.synonyms);
      for (var i = 0; i < titles.length; i += 1) {
        var k = normalizeTitleKey(titles[i]);
        if (k && valKeySet[k]) return true;
      }
      return false;
    }
    function nativeIsValid(nat, m) {
      if (!requireHan) return true; // 検証クエリに漢字が無ければ検証不能 → 許容
      // クエリが作品のタイトル/別名に一致＝作品確定なら、意訳・カナ題(漢字重なり0)でも採用。
      if (seriesMatchesQuery(m)) return true;
      var best = 0;
      var k;
      for (k = 0; k < valQueries.length; k += 1) {
        best = Math.max(best, hanOverlapScore(nat, valQueries[k]));
      }
      return best >= 2; // 最低2文字の漢字重なりを要求（誤マッチ防止）
    }
    var qi;
    for (qi = 0; qi < queries.length; qi += 1) {
      var q = queries[qi];
      var res = await fetch(ANILIST_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Accept: 'application/json',
        },
        body: JSON.stringify({
          query:
            'query ($q: String) { Page(page: 1, perPage: 6) { media(search: $q) { title { native romaji english } synonyms } } }',
          variables: { q: q },
        }),
      });
      if (!res.ok) continue;
      var body = await res.json().catch(function() {
        return null;
      });
      var media =
        body && body.data && body.data.Page && Array.isArray(body.data.Page.media)
          ? body.data.Page.media
          : [];
      var mi;
      for (mi = 0; mi < media.length; mi += 1) {
        var m = media[mi];
        var nat = m && m.title ? String(m.title.native || '').trim() : '';
        if (nat && hasJapaneseTitleSignal(nat) && nativeIsValid(nat, m)) return nat;
      }
    }
    return '';
  }

  function jsonHeaders() {
    return {
      Accept: 'application/json',
      'Content-Type': 'application/json',
    };
  }

  function detailHeaders() {
    return {
      Accept: 'application/json',
    };
  }

  function base36SlugToDecimalString(slug) {
    var s = String(slug || '').trim().toLowerCase();
    if (!s || !/^[a-z0-9]+$/.test(s)) return '';
    try {
      var n = 0n;
      var i;
      for (i = 0; i < s.length; i += 1) {
        var v = parseInt(s[i], 36);
        if (!Number.isFinite(v)) return '';
        n = n * 36n + BigInt(v);
      }
      return n.toString();
    } catch (_) {
      return '';
    }
  }

  async function searchSeriesSlugsOnMuSite(query) {
    var q = String(query || '').trim();
    if (!q) return [];
    var res = await fetch(MU_SITE_SEARCH + encodeURIComponent(q), {
      method: 'GET',
      headers: {
        Accept: 'text/html,application/xhtml+xml;q=0.9,*/*;q=0.8',
      },
    });
    if (!res.ok) return [];
    var html = await res.text().catch(function() {
      return '';
    });
    if (!html) return [];
    var slugs = [];
    var seen = {};
    var re = /href="(?:https?:\/\/www\.mangaupdates\.com)?\/series\/([a-z0-9]+)(?:\/|["?#])/gi;
    var m;
    while ((m = re.exec(html)) !== null) {
      var slug = String(m[1] || '').trim().toLowerCase();
      if (!slug || seen[slug]) continue;
      seen[slug] = true;
      slugs.push(slug);
      if (slugs.length >= 6) break;
    }
    return slugs;
  }

  async function mangaUpdatesSeriesDetailBySlugOrId(seriesRef) {
    var ref = String(seriesRef || '').trim();
    if (!ref) return null;
    // MU API は URL スラッグ (xvytm0b) を受け付けず数値 series_id のみ有効
    var numericId = /^\d+$/.test(ref)
      ? ref
      : (/^[a-z0-9]+$/i.test(ref) ? base36SlugToDecimalString(ref) : '');
    if (numericId) {
      var byNumeric = await mangaUpdatesSeriesDetail(numericId).catch(function() {
        return null;
      });
      if (byNumeric) return byNumeric;
    }
    if (/^\d+$/.test(ref)) return null;
    return mangaUpdatesSeriesDetail(ref).catch(function() {
      return null;
    });
  }

  async function mangaUpdatesSearch(search) {
    var q = String(search || '').trim();
    var sk = 's:' + q.toLowerCase();
    if (_searchCache.has(sk)) return _searchCache.get(sk);
    var res = await fetch(MU_API + '/series/search', {
      method: 'POST',
      headers: jsonHeaders(),
      body: JSON.stringify({
        search: q,
        stype: 'title',
        page: 1,
        perpage: 8,
      }),
    });
    if (!res.ok) throw new Error('MangaUpdates API search error: ' + res.status);
    var data = await res.json().catch(function() {
      return null;
    });
    var results = data && Array.isArray(data.results) ? data.results : [];
    cacheTake(_searchCache, sk, results);
    return results;
  }

  async function mangaUpdatesSeriesDetail(seriesId) {
    var id = String(seriesId || '').trim();
    if (!id) return null;
    if (_detailCache.has(id)) return _detailCache.get(id);
    var res = await fetch(MU_API + '/series/' + encodeURIComponent(id), {
      method: 'GET',
      headers: detailHeaders(),
    });
    if (!res.ok) throw new Error('MangaUpdates API detail error: ' + res.status);
    var parsed = await res.json().catch(function() {
      return null;
    });
    cacheTake(_detailCache, id, parsed);
    return parsed;
  }

  function asTitleText(value) {
    if (typeof value === 'string') return value.trim();
    if (value && typeof value === 'object') {
      return String(value.title || value.name || value.label || '').trim();
    }
    return '';
  }

  function pickJapaneseTitle(values) {
    var i;
    for (i = 0; i < (values || []).length; i += 1) {
      var t = asTitleText(values[i]);
      if (hasJapaneseTitleSignal(t)) return t;
    }
    return '';
  }

  function normalizeTitleKey(value) {
    return String(value || '')
      .toLowerCase()
      .replace(/[\s　\-_:：!！?？.,，。、'"’‘“”（）()【】\[\]{}~～]/g, '')
      .trim();
  }

  function collectDetailTitles(detail) {
    if (!detail || typeof detail !== 'object') return [];
    var associated = Array.isArray(detail.associated) ? detail.associated : [];
    return [detail.title].concat(associated)
      .map(asTitleText)
      .filter(Boolean);
  }

  function collectTitlesForMatch(detail, searchRow) {
    return uniqText(
      collectDetailTitles(detail).concat(getSearchResultTitles(searchRow || {})),
      24
    );
  }

  function anyQueryHasHan(queryOrQueries) {
    var qs = Array.isArray(queryOrQueries) ? queryOrQueries : [queryOrQueries];
    var i;
    for (i = 0; i < qs.length; i += 1) {
      if (collectHanChars(qs[i]).length) return true;
    }
    return false;
  }

  function maxHanOverlapBetweenTitlesAndQueries(titles, queries) {
    var best = 0;
    var i;
    var j;
    for (i = 0; i < (titles || []).length; i += 1) {
      for (j = 0; j < (queries || []).length; j += 1) {
        best = Math.max(best, hanOverlapScore(titles[i], queries[j]));
      }
    }
    return best;
  }

  function detailMatchesQueries(detail, queryKeys, rawQueries, searchRow) {
    if (!queryKeys || !queryKeys.length) return false;
    var titles = collectTitlesForMatch(detail, searchRow);
    // 連想一致で別シリーズを拾うのを防ぐ（CJKクエリでは漢字重なりを最低限要求）
    var raw = Array.isArray(rawQueries) ? rawQueries : [];
    var requireHan = anyQueryHasHan(raw);
    if (requireHan) {
      var titlesHaveHan = false;
      var th;
      for (th = 0; th < titles.length; th += 1) {
        if (collectHanChars(titles[th]).length) {
          titlesHaveHan = true;
          break;
        }
      }
      // MU側が英題/韓題中心で漢字タイトルを返さない作品はここで弾かない
      if (titlesHaveHan) {
        var overlap = maxHanOverlapBetweenTitlesAndQueries(titles, raw);
        if (overlap < 2) return false;
        // 強い漢字重なりはシリーズ一致の十分条件とみなす。
        // 繁体字クエリ（博客來原題）↔ 日本語題（字体差: 廳/庁, 對/対, 係/組 等）で
        // 正規化キーが一致しない正解シリーズを API 検索段で採用するための救済。
        var distinctQueryHan = {};
        var dq;
        for (dq = 0; dq < raw.length; dq += 1) {
          var hs = collectHanChars(raw[dq]);
          var hi;
          for (hi = 0; hi < hs.length; hi += 1) distinctQueryHan[hs[hi]] = true;
        }
        var distinctCount = 0;
        var dk;
        for (dk in distinctQueryHan) {
          if (Object.prototype.hasOwnProperty.call(distinctQueryHan, dk)) distinctCount += 1;
        }
        if (overlap >= 4 && overlap >= Math.ceil(distinctCount * 0.5)) return true;
      }
    }
    var i;
    for (i = 0; i < titles.length; i += 1) {
      var key = normalizeTitleKey(titles[i]);
      if (!key) continue;
      var qj;
      for (qj = 0; qj < queryKeys.length; qj += 1) {
        var qk = queryKeys[qj];
        if (!qk) continue;
        if (key === qk) return true;
        if (key.length >= 4 && qk.length >= 4 && (key.indexOf(qk) >= 0 || qk.indexOf(key) >= 0)) return true;
      }
    }
    return false;
  }

  function collectHanChars(text) {
    var s = String(text || '');
    var m = s.match(/[\u3400-\u4DBF\u4E00-\u9FFF]/g);
    return m ? m : [];
  }

  function hasAnyHan(text) {
    return collectHanChars(text).length > 0;
  }

  function hanOverlapScore(a, b) {
    var aa = collectHanChars(a);
    var bb = collectHanChars(b);
    if (!aa.length || !bb.length) return 0;
    var setA = {};
    var i;
    for (i = 0; i < aa.length; i += 1) setA[aa[i]] = true;
    var hit = 0;
    var uniqB = {};
    for (i = 0; i < bb.length; i += 1) {
      var ch = bb[i];
      if (uniqB[ch]) continue;
      uniqB[ch] = true;
      if (setA[ch]) hit += 1;
    }
    // ざっくり: 重なり文字数を重視（「守」「護」などが効く）
    return hit;
  }

  function scoreTitleMatchWithKeys(candidate, preferKeys) {
    var cKey = normalizeTitleKey(candidate);
    if (!cKey) return 0;
    var best = 0;
    var preferHasHan = false;
    var bestHan = 0;
    var i;
    for (i = 0; i < (preferKeys || []).length; i += 1) {
      var raw = preferKeys[i];
      if (!preferHasHan && hasAnyHan(raw)) preferHasHan = true;
      var pk = normalizeTitleKey(raw);
      if (cKey === pk) return 1000;
      if (cKey.length >= 4 && pk.length >= 4 && (cKey.indexOf(pk) >= 0 || pk.indexOf(cKey) >= 0)) {
        best = Math.max(best, 700 + Math.min(cKey.length, pk.length));
      }
      // CJK（漢字）重なりで誤採用を防ぐ：中華題名→日本語翻訳タイトルの対応が強い
      var hanScore = hanOverlapScore(candidate, raw);
      if (hanScore > bestHan) bestHan = hanScore;
      if (hanScore > 0) best = Math.max(best, 300 + hanScore * 30);
    }
    // preferKeys に漢字が含まれる（繁体字/簡体字/日本語漢字クエリ）場合、
    // 漢字重なりゼロの日本語候補（無関係な別名）を強く落とす
    if (preferHasHan && bestHan === 0) best -= 250;
    return best;
  }

  function pickJapaneseFromDetail(detail, preferKeys, options) {
    options = options || {};
    if (!detail || typeof detail !== 'object') return '';
    var associated = Array.isArray(detail.associated) ? detail.associated : [];
    var jpCandidates = [];
    var i;
    for (i = 0; i < associated.length; i += 1) {
      var assocTitle = asTitleText(associated[i]);
      if (assocTitle && hasJapaneseTitleSignal(assocTitle)) jpCandidates.push(assocTitle);
    }
    var main = asTitleText(detail.title);
    if (main && hasJapaneseTitleSignal(main)) jpCandidates.push(main);
    jpCandidates = uniqText(jpCandidates, 12);
    if (!jpCandidates.length) return '';

    // 漢字クエリ（繁体/簡体/漢字日本語）がある場合、
    // 漢字重なりゼロの日本語候補は誤マッチの可能性が高いため採用対象から除外する。
    // ただしシリーズ一致が確認済み（associated に原題がある等）なら除外しない。
    var hasHanQuery = false;
    for (i = 0; i < (preferKeys || []).length; i += 1) {
      if (collectHanChars(preferKeys[i]).length) {
        hasHanQuery = true;
        break;
      }
    }
    if (hasHanQuery && !options.seriesVerified) {
      var filtered = [];
      for (i = 0; i < jpCandidates.length; i += 1) {
        var c0 = jpCandidates[i];
        var maxHan = 0;
        var k;
        for (k = 0; k < (preferKeys || []).length; k += 1) {
          maxHan = Math.max(maxHan, hanOverlapScore(c0, preferKeys[k]));
        }
        if (maxHan > 0) filtered.push(c0);
      }
      if (filtered.length) jpCandidates = filtered;
      else jpCandidates = [];
    }

    if (!jpCandidates.length) return '';

    // シリーズ一致済みなら Associated Names の先頭日本語を優先（GAS と同様）。
    // ただし中文原題などクエリと一致する原語題エコーは除外してから先頭を採る
    // （字体差のある日本語漢字題は除外しないので K-9 等は先頭の日本語題を採用）。
    if (options.seriesVerified) {
      return pickPreferredJapaneseCandidate(jpCandidates, options.echoQueries);
    }

    // クエリや英題（Saving My Sweetheart など）に最も近い日本語候補を優先する
    var best = jpCandidates[0];
    var bestScore = scoreTitleMatchWithKeys(best, preferKeys || []);
    for (i = 1; i < jpCandidates.length; i += 1) {
      var c = jpCandidates[i];
      var s = scoreTitleMatchWithKeys(c, preferKeys || []);
      if (s > bestScore) {
        best = c;
        bestScore = s;
      }
    }
    return String(best || '').trim();
  }

  function getSearchResultSeriesId(row) {
    if (!row || typeof row !== 'object') return '';
    var record = row.record && typeof row.record === 'object' ? row.record : {};
    return record.series_id || record.id || row.series_id || row.id || '';
  }

  function getSearchResultTitles(row) {
    if (!row || typeof row !== 'object') return [];
    var record = row.record && typeof row.record === 'object' ? row.record : {};
    var titles = [
      row.hit_title,
      row.title,
      record.title,
    ];
    if (Array.isArray(record.associated)) {
      titles = titles.concat(record.associated.map(asTitleText));
    }
    return titles.map(asTitleText).filter(Boolean);
  }

  function tryResolveMatchedDetail_(detail, searchRow, queryKeys, queries, matchedTitles) {
    if (!detail) return { japaneseTitle: '', matchedTitles: matchedTitles || [] };
    if (!detailMatchesQueries(detail, queryKeys, queries, searchRow)) {
      return { japaneseTitle: '', matchedTitles: matchedTitles || [] };
    }
    var titles = uniqText(
      (matchedTitles || []).concat(collectTitlesForMatch(detail, searchRow)),
      12
    );
    var detailPreferKeys = queries
      .concat(getSearchResultTitles(searchRow || {}))
      .filter(Boolean);
    var jp = pickJapaneseFromDetail(detail, detailPreferKeys, { seriesVerified: true, echoQueries: queries });
    return { japaneseTitle: jp, matchedTitles: titles };
  }

  function stripVolumeSuffix(value) {
    return String(value || '')
      .replace(/[\s　]*(?:第)?\s*[0-9０-９]{1,4}\s*(?:漫畫|漫画|コミック|單行本|单行本)\s*$/i, '')
      .replace(/[\s　]*(?:漫畫|漫画|コミック|單行本|单行本)\s*(?:第)?\s*[0-9０-９]{1,4}\s*$/i, '')
      .replace(/[\s　]*[#＃]?\s*[0-9０-９]{1,4}\s*$/i, '')
      .replace(/[\s　]*(?:vol\.?|v\.?|第)?\s*[0-9０-９]{1,4}\s*(?:巻|卷|集|冊|話|回|話$|$)?\s*$/i, '')
      .replace(/[\s　]+$/g, '')
      .trim();
  }

  function expandQueriesWithStripped(values) {
    var out = [];
    var i;
    for (i = 0; i < (values || []).length; i += 1) {
      var s = String(values[i] || '').trim();
      if (!s) continue;
      out.push(s);
      var stripped = stripVolumeSuffix(s);
      if (stripped && stripped !== s && stripped.length >= 2) out.push(stripped);
    }
    return out;
  }

  function uniqQueries(values) {
    var seen = {};
    var out = [];
    var i;
    for (i = 0; i < (values || []).length; i += 1) {
      var s = String(values[i] || '').trim();
      if (!s || s.length < 2) continue;
      var k = s.toLowerCase();
      if (seen[k]) continue;
      seen[k] = true;
      out.push(s);
    }
    return out;
  }

  function uniqText(values, limit) {
    var seen = {};
    var out = [];
    var i;
    for (i = 0; i < (values || []).length; i += 1) {
      var s = String(values[i] || '').trim();
      var key = s.toLowerCase();
      if (!s || seen[key]) continue;
      seen[key] = true;
      out.push(s);
      if (limit && out.length >= limit) break;
    }
    return out;
  }

  function toCandidateList(titles, provider, score) {
    var values = uniqText(titles || [], 12);
    var out = [];
    var i;
    for (i = 0; i < values.length; i += 1) {
      out.push({
        title: values[i],
        provider: provider || 'mangaUpdates',
        score: score || 0,
      });
    }
    return out;
  }

  function makeTraceQuery(queries) {
    var list = queries || [];
    if (!list.length) return '';
    return list.length > 1 ? list[0] + '+' + (list.length - 1) + 'cand' : list[0];
  }

  async function searchAndCollectDetails(q, errors) {
    var results = [];
    try {
      results = await mangaUpdatesSearch(q);
    } catch (err) {
      if (errors) errors.push('search(' + q + '): ' + (err.message || err));
      results = [];
    }
    var slice = results.slice(0, MU_SEARCH_TOP);
    var details = await Promise.all(
      slice.map(function(row) {
        var seriesId = getSearchResultSeriesId(row);
        return seriesId
          ? mangaUpdatesSeriesDetailBySlugOrId(seriesId).catch(function(err) {
              if (errors) errors.push('detail(' + seriesId + '): ' + (err.message || err));
              return null;
            })
          : Promise.resolve(null);
      })
    );
    return { results: results, details: details };
  }

  /**
   * @param {string|string[]} queryOrQueries
   * @returns {Promise<object>}
   */
  async function lookupJapaneseTitleDetailed(queryOrQueries) {
    var inputQueries = Array.isArray(queryOrQueries) ? queryOrQueries : [queryOrQueries];
    var queries = uniqQueries(expandQueriesWithStripped(inputQueries)).slice(0, 3);
    var errors = [];
    if (!queries.length) {
      return {
        status: 'not_found',
        japaneseTitle: '',
        provider: '',
        queries: [],
        matchedTitles: [],
        candidates: [],
        trace: 'mangaUpdates:no_query',
        errors: errors,
      };
    }
    var queryKeys = uniqQueries(
      queries.concat(queries.map(stripVolumeSuffix)).filter(Boolean)
    ).map(normalizeTitleKey).filter(Boolean);
    var aniQueries = queries.slice();
    var verifiedJp = '';
    var matchedTitles = [];
    var traceQuery = makeTraceQuery(queries);
    var qi;
    for (qi = 0; qi < queries.length; qi += 1) {
      var pack = await searchAndCollectDetails(queries[qi], errors);
      var dj;
      for (dj = 0; dj < pack.details.length; dj += 1) {
        var searchRow = (pack.results || [])[dj] || {};
        var resolved = tryResolveMatchedDetail_(
          pack.details[dj],
          searchRow,
          queryKeys,
          queries,
          matchedTitles
        );
        if (!resolved.matchedTitles.length) continue;
        matchedTitles = resolved.matchedTitles;
        if (resolved.japaneseTitle) {
          verifiedJp = resolved.japaneseTitle;
          break;
        }
      }
      if (verifiedJp) break;
      var ri;
      for (ri = 0; ri < pack.results.length; ri += 1) {
        var rec = pack.results[ri] && pack.results[ri].record;
        var rt = rec && rec.title ? String(rec.title).trim() : '';
        var hit = pack.results[ri] && pack.results[ri].hit_title
          ? String(pack.results[ri].hit_title).trim()
          : '';
        if (rt) aniQueries.push(rt);
        if (hit && hit !== rt) aniQueries.push(hit);
        var assoc = rec && Array.isArray(rec.associated) ? rec.associated : [];
        var ai;
        for (ai = 0; ai < assoc.length; ai += 1) {
          var at = asTitleText(assoc[ai]);
          if (at && /[A-Za-z]/.test(at) && !collectHanChars(at).length) aniQueries.push(at);
        }
      }
    }
    if (!verifiedJp) {
      var si;
      var sq;
      for (qi = 0; qi < queries.length; qi += 1) {
        sq = queries[qi];
        var slugs = await searchSeriesSlugsOnMuSite(sq).catch(function(err) {
          errors.push('siteSearch(' + sq + '): ' + (err.message || err));
          return [];
        });
        // スラッグ詳細は並列取得する（逐次 await だと最大数十リクエストが直列化し、
        // MUに日本語タイトルが無い作品で照会が15〜30秒級に膨らむ主因だった）。
        var siteRefs = slugs.slice(0, 4);
        var siteDetails = await Promise.all(
          siteRefs.map(function(slug) {
            return mangaUpdatesSeriesDetailBySlugOrId(slug).catch(function(err) {
              errors.push('siteDetail(' + slug + '): ' + (err.message || err));
              return null;
            });
          })
        );
        for (si = 0; si < siteDetails.length; si += 1) {
          // クエリ文字列を hit_title として注入しない。注入するとそれ自体が候補タイトルに
          // 混ざり、detailMatchesQueries がどのシリーズでも自明に一致して検証を素通りし、
          // 無関係なシリーズ（例: 別作「Police Tribe K-9」）を採用してしまう。
          // 検証はシリーズ自身の Associated Names に対して行う。
          var siteRow = {};
          var siteResolved = tryResolveMatchedDetail_(
            siteDetails[si],
            siteRow,
            queryKeys,
            queries,
            matchedTitles
          );
          if (!siteResolved.matchedTitles.length) continue;
          matchedTitles = siteResolved.matchedTitles;
          if (siteResolved.japaneseTitle) {
            verifiedJp = siteResolved.japaneseTitle;
            traceQuery = makeTraceQuery(queries) + ',site';
            break;
          }
        }
        if (verifiedJp) break;
      }
    }
    if (verifiedJp) {
      return {
        status: 'resolved',
        japaneseTitle: verifiedJp,
        provider: 'mangaUpdates',
        queries: queries,
        matchedTitles: matchedTitles,
        candidates: toCandidateList([verifiedJp], 'mangaUpdates', 900),
        trace: 'mangaUpdates:hit(q=' + traceQuery + ')',
        errors: errors,
      };
    }
    // 検索は aniQueries（MU検索で拾った英題等を含む広めの集合）で行うが、
    // 採用判定は元クエリ（商品名由来 queries）と漢字重なりで検証する。
    var fromAni = await aniListFetchNativeJapanese(aniQueries, queries).catch(function(err) {
      errors.push('aniList: ' + (err.message || err));
      return '';
    });
    if (fromAni) {
      return {
        status: 'resolved',
        japaneseTitle: fromAni,
        provider: 'aniList(via_mangaupdatesClient)',
        queries: queries,
        matchedTitles: matchedTitles,
        candidates: toCandidateList([fromAni], 'aniList(via_mangaupdatesClient)', 850),
        trace: 'aniListNative:hit(q=' + traceQuery + ')',
        errors: errors,
      };
    }
    if (matchedTitles.length) {
      return {
        status: 'series_found_no_japanese',
        japaneseTitle: '',
        provider: 'mangaUpdates',
        queries: queries,
        matchedTitles: matchedTitles,
        candidates: toCandidateList(matchedTitles, 'mangaUpdates(series)', 100),
        trace: 'mangaUpdates:series_hit_no_japanese(q=' + traceQuery + ' titles=' + matchedTitles.slice(0, 4).join('/') + ')',
        errors: errors,
      };
    }
    return {
      status: 'not_found',
      japaneseTitle: '',
      provider: '',
      queries: queries,
      matchedTitles: [],
      candidates: [],
      trace: 'mangaUpdates:not_found(q=' + traceQuery + ')',
      errors: errors,
    };
  }

  /**
   * @param {string|string[]} queryOrQueries
   * @returns {Promise<string>}
   */
  async function lookupJapaneseTitle(queryOrQueries) {
    var result = await lookupJapaneseTitleDetailed(queryOrQueries);
    return result && result.japaneseTitle ? result.japaneseTitle : '';
  }

  global.titleLookupMangaUpdates = {
    lookupJapaneseTitle: lookupJapaneseTitle,
    lookupJapaneseTitleDetailed: lookupJapaneseTitleDetailed,
    mangaUpdatesSearch: mangaUpdatesSearch,
    mangaUpdatesSeriesDetail: mangaUpdatesSeriesDetail,
    __test: {
      detailMatchesQueries: detailMatchesQueries,
      pickJapaneseFromDetail: pickJapaneseFromDetail,
      collectTitlesForMatch: collectTitlesForMatch,
      normalizeTitleKey: normalizeTitleKey,
      base36SlugToDecimalString: base36SlugToDecimalString,
      tryResolveMatchedDetail_: tryResolveMatchedDetail_,
    },
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);

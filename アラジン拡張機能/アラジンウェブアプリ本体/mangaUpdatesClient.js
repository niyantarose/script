/**
 * MangaUpdates 公式 API を拡張コンテキストから直接呼び出す（GAS からの 403 回避用）。
 * OpenAPI: POST /v1/series/search , GET /v1/series/{id}
 */
(function registerMangaUpdatesClient(global) {
  'use strict';

  var MU_API = 'https://api.mangaupdates.com/v1';
  var ANILIST_URL = 'https://graphql.anilist.co';
  var ENABLE_MU_SITE_HTML_FALLBACK = false;
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

  function hasJapaneseTitleSignal(value) {
    return /[\u3041-\u3096\u30A1-\u30FA\u30FC\u309D\u309E\u3005\u3006\u3007]/.test(String(value || ''));
  }

  function normalizeTitleKey(value) {
    return String(value || '')
      .trim()
      .toLowerCase()
      .replace(/[\s\u3000"'“”‘’・･:：!！?？,，、.．\-ー‐―–—~〜/／\\|()\[\]{}【】「」『』<>]/g, '');
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
    return hit;
  }

  function collectDetailTitles(detail) {
    if (!detail || typeof detail !== 'object') return [];
    var associated = Array.isArray(detail.associated) ? detail.associated : [];
    var titles = [detail.title];
    var i;
    for (i = 0; i < associated.length; i += 1) {
      titles.push(associated[i] && associated[i].title ? associated[i].title : '');
    }
    return titles.map(function(v) { return String(v || '').trim(); }).filter(Boolean);
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

  function detailMatchesQuery(detail, rawQuery) {
    var q = String(rawQuery || '').trim();
    if (!q) return false;
    var qKey = normalizeTitleKey(q);
    var titles = collectDetailTitles(detail);
    // CJKクエリでは漢字重なりが最低限ないものは除外
    if (collectHanChars(q).length) {
      var titlesHaveHan = false;
      var th;
      for (th = 0; th < titles.length; th += 1) {
        if (collectHanChars(titles[th]).length) {
          titlesHaveHan = true;
          break;
        }
      }
      if (titlesHaveHan && maxHanOverlapBetweenTitlesAndQueries(titles, [q]) < 2) return false;
    }
    var i;
    for (i = 0; i < titles.length; i += 1) {
      var key = normalizeTitleKey(titles[i]);
      if (!key) continue;
      if (key === qKey) return true;
      if (key.length >= 4 && qKey.length >= 4 && (key.indexOf(qKey) >= 0 || qKey.indexOf(key) >= 0)) return true;
    }
    return false;
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
      var hanScore = hanOverlapScore(candidate, raw);
      if (hanScore > bestHan) bestHan = hanScore;
      if (hanScore > 0) best = Math.max(best, 300 + hanScore * 30);
    }
    if (preferHasHan && bestHan === 0) best -= 250;
    return best;
  }

  function stripVolumeSuffix(value) {
    return String(value || '')
      .replace(/[\s　]*(?:第)?\s*[0-9０-９]{1,4}\s*(?:漫畫|漫画|コミック|單行本|单行本)\s*$/i, '')
      .replace(/[\s　]*(?:漫畫|漫画|コミック|單行本|单行本)\s*(?:第)?\s*[0-9０-９]{1,4}\s*$/i, '')
      .replace(/[\s　]*[#＃]?\s*[0-9０-９]{1,4}\s*$/i, '')
      .replace(/[\s　]*(?:vol\.?|v\.?|第)?\s*[0-9０-９]{1,4}\s*(?:巻|卷|集|冊|話|回)?\s*$/i, '')
      .replace(/[\s　]+$/g, '')
      .trim();
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

  function base36SlugToDecimalString(slug) {
    var s = String(slug || '').trim().toLowerCase();
    if (!s || !/^[a-z0-9]+$/.test(s)) return '';
    try {
      var n = 0n;
      var i;
      for (i = 0; i < s.length; i += 1) {
        var ch = s[i];
        var v = parseInt(ch, 36);
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
    var url = 'https://www.mangaupdates.com/site/search/result?search=' + encodeURIComponent(q);
    var res = await fetch(url, {
      method: 'GET',
      headers: {
        Accept: 'text/html,application/xhtml+xml;q=0.9,*/*;q=0.8',
      },
    });
    if (!res.ok) return [];
    var html = await res.text().catch(function() { return ''; });
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

  async function aniListFetchNativeJapanese(searchStrings) {
    var queries = uniqAniListQueries(searchStrings).slice(0, 5);
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
            'query ($q: String) { Page(page: 1, perPage: 6) { media(search: $q) { title { native romaji english } } } }',
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
        if (nat && hasJapaneseTitleSignal(nat)) return nat;
      }
    }
    return '';
  }

  function browserJsonHeaders() {
    return {
      Accept: 'application/json',
      'Content-Type': 'application/json',
    };
  }

  async function mangaUpdatesSearch(search) {
    var q = String(search || '').trim();
    var sk = 's:' + q.toLowerCase();
    if (_searchCache.has(sk)) return _searchCache.get(sk);
    var res = await fetch(MU_API + '/series/search', {
      method: 'POST',
      headers: browserJsonHeaders(),
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
      headers: browserJsonHeaders(),
    });
    if (!res.ok) throw new Error('MangaUpdates API detail error: ' + res.status);
    var parsed = await res.json().catch(function() {
      return null;
    });
    cacheTake(_detailCache, id, parsed);
    return parsed;
  }

  function pickJapaneseFromDetail(detail, preferKeys) {
    if (!detail || typeof detail !== 'object') return '';
    var associated = Array.isArray(detail.associated) ? detail.associated : [];
    var jpCandidates = [];
    var i;
    for (i = 0; i < associated.length; i += 1) {
      var item = associated[i];
      var t = String(item && item.title ? item.title : '').trim();
      if (t && hasJapaneseTitleSignal(t)) jpCandidates.push(t);
    }
    var main = String(detail.title || '').trim();
    if (main && hasJapaneseTitleSignal(main)) jpCandidates.push(main);
    if (!jpCandidates.length) return '';

    var hasHanQuery = false;
    for (i = 0; i < (preferKeys || []).length; i += 1) {
      if (collectHanChars(preferKeys[i]).length) {
        hasHanQuery = true;
        break;
      }
    }
    if (hasHanQuery) {
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
    }

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

  /**
   * @param {string} query
   * @returns {Promise<string>}
   */
  async function lookupJapaneseTitle(query) {
    var q = String(query || '').trim();
    if (!q) return '';
    var q2 = stripVolumeSuffix(q);
    var searchQueries = [];
    if (q) searchQueries.push(q);
    if (q2 && q2 !== q) searchQueries.push(q2);
    var aniQueries = searchQueries.slice();
    var sqi;
    for (sqi = 0; sqi < searchQueries.length; sqi += 1) {
      var sq = searchQueries[sqi];
      var results = [];
      try {
        results = await mangaUpdatesSearch(sq);
      } catch (_) {
        results = [];
      }
      var slice = results.slice(0, 4);
      var details = await Promise.all(
        slice.map(function(row) {
          var seriesId = row && row.record && row.record.series_id;
          return seriesId
            ? mangaUpdatesSeriesDetail(seriesId).catch(function() {
                return null;
              })
            : Promise.resolve(null);
        })
      );
      var j;
      for (j = 0; j < details.length; j += 1) {
        if (!details[j]) continue;
        if (!detailMatchesQuery(details[j], sq)) continue;
        var row = results[j] || {};
        var rec = row && row.record ? row.record : {};
        var detailPreferKeys = [sq];
        var rt = rec && rec.title ? String(rec.title).trim() : '';
        var hit = row && row.hit_title ? String(row.hit_title).trim() : '';
        if (rt) detailPreferKeys.push(rt);
        if (hit && hit !== rt) detailPreferKeys.push(hit);
        var jp = pickJapaneseFromDetail(details[j], detailPreferKeys);
        if (jp) return jp;
      }
      var ri;
      for (ri = 0; ri < results.length; ri += 1) {
        var r = results[ri] || {};
        var rr = r.record || {};
        var art = rr && rr.title ? String(rr.title).trim() : '';
        var ahit = r && r.hit_title ? String(r.hit_title).trim() : '';
        if (art) aniQueries.push(art);
        if (ahit && ahit !== art) aniQueries.push(ahit);
      }
      if (ENABLE_MU_SITE_HTML_FALLBACK && (!details.length || !results.length)) {
        var slugs = await searchSeriesSlugsOnMuSite(sq).catch(function() { return []; });
        var si;
        for (si = 0; si < slugs.length; si += 1) {
          var numericId = base36SlugToDecimalString(slugs[si]);
          if (!numericId) continue;
          var d = await mangaUpdatesSeriesDetail(numericId).catch(function() { return null; });
          if (!d) continue;
          if (!detailMatchesQuery(d, sq)) continue;
          var sjp = pickJapaneseFromDetail(d, [sq]);
          if (sjp) return sjp;
        }
      }
    }
    var fromAni = await aniListFetchNativeJapanese(aniQueries).catch(function() {
      return '';
    });
    return fromAni || '';
  }

  global.titleLookupMangaUpdates = {
    lookupJapaneseTitle: lookupJapaneseTitle,
    mangaUpdatesSearch: mangaUpdatesSearch,
    mangaUpdatesSeriesDetail: mangaUpdatesSeriesDetail,
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);

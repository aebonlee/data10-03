/* 도면 한 장을 읽어 목표 1·2 결과를 만드는 흐름 — 브라우저(app.js)와 node(테스트·측정 스크립트)가 같이 씁니다.
   pdf.js 의 page 객체에서 글자·선을 꺼내는 부분(extractPage)만 pdf.js 에 기대고, 나머지는 DRLogic 의 순수 함수입니다. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./logic.js'));
  else root.DRFlow = factory(root.DRLogic);
})(typeof self !== 'undefined' ? self : this, function (L) {
  'use strict';

  /* pdf.js page → { page, w, h, items, tokens, segs, circles, tris } (좌표 pt · 왼쪽 위 원점) */
  function extractPage(page, OPS) {
    var vp = page.getViewport({ scale: 1 }), n = page.pageNumber;
    return page.getTextContent().then(function (tc) {
      return page.getOperatorList().then(function (ol) {
        var items = L.pdfTextItems(tc.items, vp.transform, n);
        var paths = L.pathsFromOpList(ol.fnArray, ol.argsArray, OPS, vp.transform);
        return { page: n, w: vp.width, h: vp.height, items: items, tokens: L.groupTokens(items, n), segs: paths.segs,
          circles: L.findCircles(paths.subpaths), tris: L.findWeldTriangles(paths.subpaths) };
      });
    });
  }
  function extractAll(pdf, OPS) {
    var out = [], p = Promise.resolve();
    for (var i = 1; i <= pdf.numPages; i++) (function (k) {
      p = p.then(function () { return pdf.getPage(k); }).then(function (pg) { return extractPage(pg, OPS); }).then(function (d) { out.push(d); });
    })(i);
    return p.then(function () { return out; });
  }

  function docOf(doc) {
    doc = doc || {};
    ['excl', 'removed', 'changed', 'added', 'nums', 'bpos', 'linkOv'].forEach(function (k) { if (!doc[k]) doc[k] = {}; });
    return doc;
  }
  /* 한 쪽의 후보 = 자동 후보 − 지운 것 + 고친 값 + 직접 더한 것 */
  function pageCandidates(pg, s, doc, opt) {
    doc = docOf(doc); opt = opt || {};
    var skip = {};
    (opt.balloons || []).forEach(function (b) { if (b.tokenId) skip[b.tokenId] = 1; });
    var det = L.detectCandidates(pg, s, { skip: skip, userBoxes: doc.excl[pg.page] || [] });
    var removed = [];
    var cands = det.cands.filter(function (c) { if (doc.removed[c.id]) { removed.push(c); return false; } return true; }).map(function (c) {
      var ch = doc.changed[c.id];
      return ch ? Object.assign({}, c, ch, { edited: true }) : c;
    });
    (doc.added[pg.page] || []).forEach(function (a) { cands.push(Object.assign({ src: 'manual', tokens: [], check: false, note: '' }, a)); });
    return { cands: cands, excluded: det.excluded, boxes: det.boxes, removed: removed };
  }

  /* 목표 1 — 풍선 없는 도면: 후보에 번호를 매기고 풍선 자리를 잡습니다 */
  function goal1(pages, s, doc) {
    doc = docOf(doc);
    var ordered = [];
    pages.forEach(function (pg) {
      // 이미 있는 풍선 숫자(원 안의 정수)는 검사 대상이 아니므로 뺍니다
      var pc = pageCandidates(pg, s, doc, { balloons: L.detectBalloons(pg, s) });
      L.orderCandidates(pc.cands, pg, s).forEach(function (c) { ordered.push({ c: c, pg: pg }); });
    });
    var used = {};
    ordered.forEach(function (x) { var n = doc.nums[x.c.id]; if (n) used[n] = 1; });
    var next = s.startNo || 1;
    var items = ordered.map(function (x) {
      var no = doc.nums[x.c.id];
      if (!no) { while (used[next]) next++; no = next; used[next] = 1; }
      return { no: no, page: x.pg.page, zone: L.zoneOf(x.c, x.pg, s).label, cand: Object.assign({}, x.c, { no: no }), pg: x.pg };
    });
    pages.forEach(function (pg) {
      var mine = items.filter(function (x) { return x.page === pg.page; }).sort(function (a, b) { return a.no - b.no; });
      var ss = Object.assign({}, s, { userBoxes: doc.excl[pg.page] || [] });
      var placed = L.placeBalloons(mine.map(function (x) { return x.cand; }), pg, ss, doc.bpos);
      placed.forEach(function (b, i) { mine[i].balloon = b; if (b.overlap) mine[i].note = '풍선이 다른 글자와 겹침 — 옮겨 주세요'; });
    });
    items.sort(function (a, b) { return a.no - b.no; });
    items.forEach(function (x) { x.method = '자동 배치'; delete x.pg; });
    return items;
  }

  /* 목표 2 — 풍선 있는 도면: 풍선 번호를 읽고 가리키는 치수·공차·용접을 연결합니다 */
  function goal2(pages, s, doc) {
    doc = docOf(doc);
    var out = [], allBalloons = [];
    pages.forEach(function (pg) {
      var balloons = L.detectBalloons(pg, s);
      var pc = pageCandidates(pg, s, doc, { balloons: balloons });
      var links = L.linkBalloons(balloons, pc.cands, pg, s);
      var byId = {};
      pc.cands.forEach(function (c) { byId[c.id] = c; });
      links.forEach(function (lk) {
        var key = lk.page + ':' + lk.no, ov = doc.linkOv[key];
        if (ov === 'none') { lk.candId = null; lk.method = '사람이 비움'; }
        else if (ov && byId[ov]) { lk.candId = ov; lk.method = '사람이 연결'; lk.check = false; lk.crop = null; }
        var c = lk.candId ? byId[lk.candId] : null;
        out.push({ no: lk.no, page: lk.page, zone: c ? L.zoneOf(c, pg, s).label : '', cand: c, balloon: { x: lk.balloon.cx, y: lk.balloon.cy, r: lk.balloon.r, src: lk.balloon.src },
          leader: lk.leader, target: lk.target, method: lk.method === 'leader' ? '지시선' : lk.method === 'near' ? '가까운 글자' : lk.method === 'none' ? '연결 못 함' : lk.method,
          check: lk.check || (lk.balloon.check) || (c && c.check), crop: lk.crop || (c && c.crop) || null,
          note: lk.method === 'none' ? (lk.target ? '지시선 끝 근처에 글자가 없음 — 형상만 있는 표기일 수 있음' : '가까운 글자가 없음') : '' });
        allBalloons.push(lk);
      });
      // 같은 번호가 두 번 나오면 표시
    });
    var seen = {};
    out.forEach(function (x) { if (seen[x.no]) { x.check = true; x.note = (x.note ? x.note + ' · ' : '') + '같은 번호 풍선이 또 있음'; } seen[x.no] = 1; });
    out.sort(function (a, b) { return a.no - b.no || a.page - b.page; });
    return out;
  }
  /* 도면에 풍선이 있는지 대략 판정(원 안 정수 3개 이상) */
  function looksBallooned(pages, s) {
    var n = 0;
    pages.forEach(function (pg) { n += L.detectBalloons(pg, Object.assign({}, s, { isolatedBalloons: false })).length; });
    return n >= 3;
  }

  return { extractPage: extractPage, extractAll: extractAll, pageCandidates: pageCandidates, goal1: goal1, goal2: goal2, looksBallooned: looksBallooned };
});

/* 검사 기준서 작성 자동화 — 순수 로직 (화면·저장소와 무관, node 에서도 불러 씀)

   좌표는 모두 「PDF 쪽 단위(pt, 1/72 인치) · 왼쪽 위 원점 · 아래로 갈수록 y 증가」입니다. mm = pt × 25.4 / 72.

   1. 도우미·기하            6. 제외 영역(테두리·표제란·개정란·주기)   11. 일치율(정답 대조)
   2. pdf.js 결과 정리        7. 검사 대상 후보                         12. 풍선 PDF 쓰기
   3. 원·용접 삼각형 찾기     8. 번호 순서(줄·열·구역)                  13. AI 프롬프트(애매한 후보)
   4. 글자 조각 → 토큰        9. 풍선 자리 잡기(겹침 피하기)            14. 엑셀 행 · 저장 형식
   5. 토큰 분류(치수·공차·기하공차·용접)   10. 풍선 읽기 · 지시선 연결
*/
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.DRLogic = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ── 1. 도우미·기하 ─────────────────────────────────────────── */
  var PT2MM = 25.4 / 72;
  function str(v) { return v == null ? '' : String(v).trim(); }
  function round(n, d) { if (n == null || !isFinite(n)) return null; var p = Math.pow(10, d == null ? 3 : d); return Math.round(n * p) / p; }
  function mm(pt) { return round(pt * PT2MM, 1); }
  function toPt(v) { return v / PT2MM; }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }
  function mul(m1, m2) { // pdf.js Util.transform 과 같음: m2 를 먼저, m1 을 나중에
    return [m1[0] * m2[0] + m1[2] * m2[1], m1[1] * m2[0] + m1[3] * m2[1], m1[0] * m2[2] + m1[2] * m2[3],
      m1[1] * m2[2] + m1[3] * m2[3], m1[0] * m2[4] + m1[2] * m2[5] + m1[4], m1[1] * m2[4] + m1[3] * m2[5] + m1[5]];
  }
  function apply(m, x, y) { return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]]; }
  function box(x0, y0, x1, y1) { return { x0: Math.min(x0, x1), y0: Math.min(y0, y1), x1: Math.max(x0, x1), y1: Math.max(y0, y1) }; }
  function boxOfPts(pts) {
    var b = { x0: Infinity, y0: Infinity, x1: -Infinity, y1: -Infinity };
    pts.forEach(function (p) { b.x0 = Math.min(b.x0, p[0]); b.y0 = Math.min(b.y0, p[1]); b.x1 = Math.max(b.x1, p[0]); b.y1 = Math.max(b.y1, p[1]); });
    return b;
  }
  function union(a, b) { if (!a) return b && Object.assign({}, b); if (!b) return Object.assign({}, a); return { x0: Math.min(a.x0, b.x0), y0: Math.min(a.y0, b.y0), x1: Math.max(a.x1, b.x1), y1: Math.max(a.y1, b.y1) }; }
  function pad(b, d) { return { x0: b.x0 - d, y0: b.y0 - d, x1: b.x1 + d, y1: b.y1 + d }; }
  function center(b) { return [(b.x0 + b.x1) / 2, (b.y0 + b.y1) / 2]; }
  function inBox(b, x, y) { return x >= b.x0 && x <= b.x1 && y >= b.y0 && y <= b.y1; }
  function ptBoxDist(b, x, y) { var dx = Math.max(b.x0 - x, 0, x - b.x1), dy = Math.max(b.y0 - y, 0, y - b.y1); return Math.sqrt(dx * dx + dy * dy); }
  function boxDist(a, b) { var dx = Math.max(a.x0 - b.x1, 0, b.x0 - a.x1), dy = Math.max(a.y0 - b.y1, 0, b.y0 - a.y1); return Math.sqrt(dx * dx + dy * dy); }
  function boxOverlap(a, b) { return a.x0 < b.x1 && b.x0 < a.x1 && a.y0 < b.y1 && b.y0 < a.y1; }
  function dist(x1, y1, x2, y2) { return Math.sqrt((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2)); }
  function ptSegDist(px, py, s) {
    var dx = s.x2 - s.x1, dy = s.y2 - s.y1, L = dx * dx + dy * dy;
    var t = L ? Math.max(0, Math.min(1, ((px - s.x1) * dx + (py - s.y1) * dy) / L)) : 0;
    return dist(px, py, s.x1 + t * dx, s.y1 + t * dy);
  }
  function nearestOnBox(b, x, y) { return [Math.max(b.x0, Math.min(x, b.x1)), Math.max(b.y0, Math.min(y, b.y1))]; }

  /* ── 2. pdf.js 결과 정리 ───────────────────────────────────── */
  /* getTextContent().items + 보기 변환(viewport.transform, 배율 1) → 글자 조각
     x, y = 기준선 시작점, w = 글자 진행 길이, h = 글자 크기, a = 진행 방향 각도(라디안, 화면 좌표) */
  function itemBox(x, y, w, h, a) {
    var c = Math.cos(a), s = Math.sin(a), ux = s, uy = -c; // ux,uy = 글자 「위」 방향
    var ex = x + c * w, ey = y + s * w, top = 0.8 * h, bot = 0.22 * h;
    return boxOfPts([[x - ux * bot, y - uy * bot], [x + ux * top, y + uy * top], [ex - ux * bot, ey - uy * bot], [ex + ux * top, ey + uy * top]]);
  }
  function pdfTextItems(items, vt, page) {
    var out = [];
    (items || []).forEach(function (it, i) {
      var s = it.str == null ? '' : String(it.str);
      if (!s.trim()) return;
      var m = mul(vt, it.transform);
      var h = Math.sqrt(m[2] * m[2] + m[3] * m[3]) || Math.abs(it.height) || 1;
      var a = Math.atan2(m[1], m[0]);
      var w = Math.abs(it.width) || s.length * h * 0.5;
      var x = m[4], y = m[5];
      out.push({ id: 'i' + (page || 1) + '_' + i, page: page || 1, str: s, x: x, y: y, w: w, h: h, a: a, box: itemBox(x, y, w, h, a) });
    });
    return out;
  }
  /* getOperatorList() → 선분·부분 경로(곡선은 꺾은선으로 근사). OPS 는 pdfjsLib.OPS */
  function pathsFromOpList(fnArray, argsArray, OPS, vt) {
    var ctm = [1, 0, 0, 1, 0, 0], stack = [], subpaths = [], cur = null;
    function dev(x, y) { return apply(mul(vt, ctm), x, y); }
    function start(p) { cur = { pts: [p], closed: false, curves: 0 }; subpaths.push(cur); }
    for (var i = 0; i < fnArray.length; i++) {
      var fn = fnArray[i], args = argsArray[i];
      if (fn === OPS.save) stack.push(ctm.slice());
      else if (fn === OPS.restore) { if (stack.length) ctm = stack.pop(); }
      else if (fn === OPS.transform) ctm = mul(ctm, args);
      else if (fn === OPS.paintFormXObjectBegin) { stack.push(ctm.slice()); if (args && Array.isArray(args[0]) && args[0].length === 6) ctm = mul(ctm, args[0]); }
      else if (fn === OPS.paintFormXObjectEnd) { if (stack.length) ctm = stack.pop(); }
      else if (fn === OPS.constructPath) {
        var ops = args[0], co = args[1], k = 0, last = null, first = null;
        for (var j = 0; j < ops.length; j++) {
          var op = ops[j];
          if (op === OPS.moveTo) { last = [co[k], co[k + 1]]; first = last; k += 2; start(dev(last[0], last[1])); }
          else if (op === OPS.lineTo) { last = [co[k], co[k + 1]]; k += 2; if (!cur) start(dev(last[0], last[1])); else cur.pts.push(dev(last[0], last[1])); }
          else if (op === OPS.curveTo || op === OPS.curveTo2 || op === OPS.curveTo3) {
            var p0 = last || [0, 0], c1, c2, p3;
            if (op === OPS.curveTo) { c1 = [co[k], co[k + 1]]; c2 = [co[k + 2], co[k + 3]]; p3 = [co[k + 4], co[k + 5]]; k += 6; }
            else if (op === OPS.curveTo2) { c1 = p0; c2 = [co[k], co[k + 1]]; p3 = [co[k + 2], co[k + 3]]; k += 4; }
            else { c1 = [co[k], co[k + 1]]; c2 = [co[k + 2], co[k + 3]]; p3 = c2; k += 4; }
            if (!cur) start(dev(p0[0], p0[1]));
            for (var t = 1; t <= 6; t++) {
              var u = t / 6, a1 = (1 - u) * (1 - u) * (1 - u), b1 = 3 * u * (1 - u) * (1 - u), c3 = 3 * u * u * (1 - u), d1 = u * u * u;
              cur.pts.push(dev(a1 * p0[0] + b1 * c1[0] + c3 * c2[0] + d1 * p3[0], a1 * p0[1] + b1 * c1[1] + c3 * c2[1] + d1 * p3[1]));
            }
            cur.curves++; last = p3;
          } else if (op === OPS.closePath) { if (cur) cur.closed = true; if (first) { last = first; cur = null; } }
          else if (op === OPS.rectangle) {
            var rx = co[k], ry = co[k + 1], rw = co[k + 2], rh = co[k + 3]; k += 4;
            start(dev(rx, ry)); cur.pts.push(dev(rx + rw, ry), dev(rx + rw, ry + rh), dev(rx, ry + rh)); cur.closed = true; cur.rect = true;
            last = [rx, ry]; first = last; cur = null;
          }
        }
      }
    }
    var segs = [];
    subpaths.forEach(function (sp) {
      if (sp.curves) return; // 지시선·치수선은 곧은 선
      var p = sp.pts;
      for (var q = 1; q < p.length; q++) segs.push({ x1: p[q - 1][0], y1: p[q - 1][1], x2: p[q][0], y2: p[q][1] });
      if (sp.closed && p.length > 2) segs.push({ x1: p[p.length - 1][0], y1: p[p.length - 1][1], x2: p[0][0], y2: p[0][1] });
    });
    segs = segs.filter(function (s) { return dist(s.x1, s.y1, s.x2, s.y2) > 0.3; });
    return { subpaths: subpaths, segs: segs };
  }

  /* ── 3. 원(풍선)·용접 삼각형 찾기 ───────────────────────────── */
  function findCircles(subpaths) {
    var out = [];
    subpaths.forEach(function (sp) {
      if (sp.curves < 3 || sp.pts.length < 12) return;
      var p = sp.pts, n = p.length, cx = 0, cy = 0;
      var closedish = sp.closed || dist(p[0][0], p[0][1], p[n - 1][0], p[n - 1][1]) < 1;
      if (!closedish) return;
      p.forEach(function (q) { cx += q[0]; cy += q[1]; }); cx /= n; cy /= n;
      var rs = p.map(function (q) { return dist(q[0], q[1], cx, cy); });
      var r = rs.reduce(function (a, b) { return a + b; }, 0) / n;
      var sd = Math.sqrt(rs.reduce(function (a, b) { return a + (b - r) * (b - r); }, 0) / n);
      if (r >= 2.5 && r <= 70 && sd / r < 0.08) out.push({ cx: cx, cy: cy, r: r });
    });
    return out;
  }
  function dedupePts(pts) {
    var o = [];
    pts.forEach(function (p) { if (!o.length || dist(p[0], p[1], o[o.length - 1][0], o[o.length - 1][1]) > 0.4) o.push(p); });
    if (o.length > 1 && dist(o[0][0], o[0][1], o[o.length - 1][0], o[o.length - 1][1]) <= 0.4) o.pop();
    return o;
  }
  /* 필릿 용접 기호 = 다리 길이가 비슷한 직각 삼각형. 치수선 화살촉(가늘고 긴 삼각형)은 빼냅니다 */
  function findWeldTriangles(subpaths) {
    var out = [];
    subpaths.forEach(function (sp) {
      if (sp.curves) return;
      var p = dedupePts(sp.pts);
      if (p.length !== 3) return;
      var L = [dist(p[0][0], p[0][1], p[1][0], p[1][1]), dist(p[1][0], p[1][1], p[2][0], p[2][1]), dist(p[2][0], p[2][1], p[0][0], p[0][1])].sort(function (a, b) { return a - b; });
      var a = L[0], b = L[1], c = L[2];
      if (c < 5 || c > 34) return;
      var right = Math.abs(a * a + b * b - c * c) < 0.15 * c * c;
      if (!right || a / b < 0.6) return;
      var bx = boxOfPts(p);
      if (out.some(function (t) { return boxDist(t.box, bx) < 0.5 && Math.abs(t.box.x0 - bx.x0) < 1; })) return;
      out.push({ box: bx, cx: (bx.x0 + bx.x1) / 2, cy: (bx.y0 + bx.y1) / 2, pts: p.map(function (q) { return [round(q[0], 2), round(q[1], 2)]; }) });
    });
    return out;
  }

  /* ── 4. 글자 조각 → 토큰 ──────────────────────────────────── */
  var GDT = { '⏥': '평면도', '▱': '평면도', '⌭': '원통도', '○': '진원도', '⌒': '선의 윤곽도', '⌓': '면의 윤곽도', '∥': '평행도', '⫽': '평행도',
    '⊥': '직각도', '⟂': '직각도', '∠': '경사도', '⌖': '위치도', '⊕': '위치도', '◎': '동심도', '⌯': '대칭도', '↗': '원주 흔들림', '⌰': '온 흔들림' };
  function gdtSym(s) { var c = str(s).charAt(0); return GDT[c] ? c : (/^\/\//.test(str(s)) ? '∥' : ''); }
  function local(x, y, a) { var c = Math.cos(a), s = Math.sin(a); return [x * c + y * s, -x * s + y * c]; }
  function fromLocal(u, v, a) { var c = Math.cos(a), s = Math.sin(a); return [u * c - v * s, u * s + v * c]; }
  function sameAngle(a, b) { var d = Math.abs(a - b) % (2 * Math.PI); return Math.min(d, 2 * Math.PI - d) < 0.05; }
  function makeToken(members, page) {
    members.sort(function (p, q) { return p.u0 - q.u0; });
    var f = members[0], h = Math.max.apply(null, members.map(function (m) { return m.h; })), s = '';
    members.forEach(function (m, i) {
      if (i) { var g = m.u0 - members[i - 1].u1; if (g > 0.18 * h && !/\s$/.test(s) && !/^\s/.test(m.str)) s += ' '; }
      s += m.str;
    });
    var u0 = f.u0, u1 = Math.max.apply(null, members.map(function (m) { return m.u1; })), v = f.v, a = f.a;
    var pts = [fromLocal(u0, v - 0.8 * h, a), fromLocal(u1, v - 0.8 * h, a), fromLocal(u1, v + 0.22 * h, a), fromLocal(u0, v + 0.22 * h, a)];
    var b = boxOfPts(pts);
    var parts = [];
    members.forEach(function (m) { parts = parts.concat(m.parts || [m.id]); });
    return { id: 't' + page + '_' + Math.round((b.x0 + b.x1) / 2) + '_' + Math.round((b.y0 + b.y1) / 2), page: page, str: s.replace(/\s+/g, ' ').trim(),
      a: a, h: h, u0: u0, u1: u1, v: v, box: b, cx: (b.x0 + b.x1) / 2, cy: (b.y0 + b.y1) / 2, parts: parts };
  }
  var CELL = /^(Ø?\d*\.?\d+[MLPⓂⓁⓅ]?|[A-Z](-[A-Z])?|[MLPⓂⓁⓅ])$/;
  var SMALL_NUM = /^([+\-±]\s*\d*\.?\d+|0)$/;
  function groupTokens(items, page) {
    page = page || (items[0] && items[0].page) || 1;
    var L = items.map(function (it) {
      var p = local(it.x, it.y, it.a);
      return { id: it.id, str: normText(it.str), h: it.h, a: it.a, u0: p[0], u1: p[0] + it.w, v: p[1] };
    }).filter(function (x) { return x.str; });
    // ① 같은 줄·같은 크기·가까운 조각 붙이기
    var parent = L.map(function (_, i) { return i; });
    function find(i) { while (parent[i] !== i) { parent[i] = parent[parent[i]]; i = parent[i]; } return i; }
    for (var i = 0; i < L.length; i++) for (var j = 0; j < L.length; j++) {
      if (i === j) continue;
      var A = L[i], B = L[j];
      if (!sameAngle(A.a, B.a)) continue;
      var h = Math.max(A.h, B.h), ratio = Math.min(A.h, B.h) / h;
      if (ratio < 0.8 || Math.abs(A.v - B.v) > 0.3 * h) continue;
      var g = B.u0 - A.u1;
      if (g < -0.5 * h || g > 0.55 * h) continue;
      if (gdtSym(A.str) || gdtSym(B.str)) continue; // 기하공차 기호는 ②에서
      parent[find(j)] = find(i);
    }
    var groups = {};
    L.forEach(function (x, k) { var r = find(k); (groups[r] = groups[r] || []).push(x); });
    var toks = Object.keys(groups).map(function (k) { return makeToken(groups[k], page); });
    // ② 기하공차 틀: 기호 뒤의 칸 글자(공차값·데이텀)를 넓은 간격까지 붙임
    toks.sort(function (p, q) { return p.u0 - q.u0; });
    toks.forEach(function (t) {
      if (t.dead || !gdtSym(t.str)) return;
      var grew = true, cells = [t];
      while (grew) {
        grew = false;
        var lastC = cells[cells.length - 1];
        for (var q = 0; q < toks.length; q++) {
          var o = toks[q];
          if (o === t || o.dead || cells.indexOf(o) >= 0 || !sameAngle(o.a, t.a)) continue;
          var g2 = o.u0 - lastC.u1;
          if (Math.abs(o.v - t.v) <= 0.35 * t.h && g2 > -0.5 * t.h && g2 < 3 * t.h && CELL.test(o.str.replace(/\s+/g, ''))) { cells.push(o); o.dead = true; grew = true; break; }
        }
      }
      if (cells.length > 1) {
        var nt = makeToken(cells.map(function (c) { return { id: c.id, parts: c.parts, str: c.str, h: c.h, a: c.a, u0: c.u0, u1: c.u1, v: c.v }; }), page);
        Object.assign(t, nt);
      }
    });
    toks = toks.filter(function (t) { return !t.dead; });
    // ③ 위·아래로 쌓인 작은 공차(+0.05 / -0.02)를 앞 치수에 붙임
    toks.forEach(function (t) {
      if (t.dead || !/\d$/.test(t.str) || SMALL_NUM.test(t.str)) return;
      var smalls = toks.filter(function (o) {
        if (o === t || o.dead || !sameAngle(o.a, t.a) || !SMALL_NUM.test(o.str)) return false;
        if (o.h > 0.85 * t.h || o.h < 0.35 * t.h) return false;
        return o.u0 >= t.u1 - 0.3 * t.h && o.u0 <= t.u1 + 1.2 * t.h && o.v >= t.v - 1.3 * t.h && o.v <= t.v + 0.5 * t.h;
      }).sort(function (p, q) { return p.v - q.v; });
      if (!smalls.length) return;
      var use = smalls.slice(0, 2);
      use.forEach(function (o) { o.dead = true; });
      var tail = use.length === 2 ? use[0].str + '/' + use[1].str : use[0].str;
      var b = t.box; use.forEach(function (o) { b = union(b, o.box); });
      t.str = t.str + ' ' + tail.replace(/\s+/g, '');
      t.box = b; t.u1 = Math.max(t.u1, Math.max.apply(null, use.map(function (o) { return o.u1; })));
      t.parts = t.parts.concat.apply(t.parts, use.map(function (o) { return o.parts; }));
      t.stacked = use.length === 2;
      t.cx = (b.x0 + b.x1) / 2; t.cy = (b.y0 + b.y1) / 2;
    });
    return toks.filter(function (t) { return !t.dead; }).map(function (t) { delete t.dead; return t; });
  }

  /* ── 5. 토큰 분류 ─────────────────────────────────────────── */
  function normText(s) {
    return str(s).replace(/[⌀øφΦ∅]/g, 'Ø').replace(/[−–—‐]/g, '-').replace(/[×✕]/g, 'x').replace(/º/g, '°')
      .replace(/＋/g, '+').replace(/\+\s*\/\s*-|\+-(?=\s*\d)/g, '±').replace(/(\d),(\d)/g, '$1.$2').replace(/\s+/g, ' ').trim();
  }
  var NUM = '(\\d+(?:\\.\\d+)?|\\.\\d+)';
  var KIND_BY_PREFIX = { 'Ø': '지름', 'R': '반지름', 'SØ': '구면 지름', 'SR': '구면 반지름', 'M': '나사', 'C': '모따기' };
  /* ISO 286 끼워맞춤 — 자주 쓰는 등급만(μm). 지름 구간: ≤3, 3~6, 6~10, 10~18, 18~30, 30~50, 50~80, 80~120, 120~180 */
  var IT = { 5: [4, 5, 6, 8, 9, 11, 13, 15, 18], 6: [6, 8, 9, 11, 13, 16, 19, 22, 25], 7: [10, 12, 15, 18, 21, 25, 30, 35, 40], 8: [14, 18, 22, 27, 33, 39, 46, 54, 63], 9: [25, 30, 36, 43, 52, 62, 74, 87, 100] };
  var SHAFT_ES = { f: [-6, -10, -13, -16, -20, -25, -30, -36, -43], g: [-2, -4, -5, -6, -7, -9, -10, -12, -14], h: [0, 0, 0, 0, 0, 0, 0, 0, 0] };
  var SHAFT_EI = { k: [0, 1, 1, 1, 2, 2, 2, 3, 3], m: [2, 4, 6, 7, 8, 9, 11, 13, 15], n: [4, 8, 10, 12, 15, 17, 20, 23, 27], p: [6, 12, 15, 18, 22, 26, 32, 37, 43] };
  var HOLE_EI = { F: [6, 10, 13, 16, 20, 25, 30, 36, 43], G: [2, 4, 5, 6, 7, 9, 10, 12, 14], H: [0, 0, 0, 0, 0, 0, 0, 0, 0] };
  function sizeIdx(d) { var lim = [3, 6, 10, 18, 30, 50, 80, 120, 180]; for (var i = 0; i < lim.length; i++) if (d <= lim[i]) return i; return -1; }
  function isoFit(d, cls) {
    var m = /^([A-Za-z]{1,2})(\d{1,2})$/.exec(str(cls)); if (!m || !(d > 0)) return null;
    var letter = m[1], g = +m[2], i = sizeIdx(d), it = IT[g];
    if (i < 0 || !it) return null;
    var T = it[i], up, lo;
    if (letter === 'js' || letter === 'JS') { up = T / 2; lo = -T / 2; }
    else if (SHAFT_ES[letter]) { up = SHAFT_ES[letter][i]; lo = up - T; }
    else if (SHAFT_EI[letter] && g <= 7) { lo = SHAFT_EI[letter][i]; up = lo + T; }
    else if (HOLE_EI[letter]) { lo = HOLE_EI[letter][i]; up = lo + T; }
    else return null;
    return { upper: round(up / 1000, 4), lower: round(lo / 1000, 4) };
  }
  var OK_SUFFIX = /^(TYP\.?|THRU|EQS|EQ\.?\s?SP\.?|DP\.?\s*\d+(\.\d+)?|관통|등간격|\(?\d+\s*PLACES?\)?|\d+\s*개소)$/i;
  function parseTail(t) {
    t = str(t);
    var r = { upper: null, lower: null, cls: '', note: '', check: false, bad: false };
    if (!t) return r;
    var m = /^-?\s*(\d[A-Za-z]|[A-Za-z]{1,2}\d{1,2})(?![A-Za-z0-9])\s*(.*)$/.exec(t);
    if (m && !/^(DP|RA|RZ)\d/i.test(m[1])) { r.cls = m[1]; t = m[2].replace(/^\(|\)$/g, '').trim(); }
    var n;
    if ((n = new RegExp('^±\\s*' + NUM + '\\s*°?$').exec(t))) { r.upper = +n[1]; r.lower = -n[1]; t = ''; }
    else if ((n = new RegExp('^([+\\-]\\s*' + NUM + '|0)\\s*/?\\s*([+\\-]\\s*' + NUM + '|0)\\s*°?$').exec(t))) {
      var p = parseFloat(n[1].replace(/\s/g, '')), q = parseFloat(n[3].replace(/\s/g, ''));
      r.upper = Math.max(p, q); r.lower = Math.min(p, q); t = '';
    } else if ((n = new RegExp('^([+\\-])\\s*' + NUM + '\\s*°?$').exec(t))) {
      var v = parseFloat(n[1] + n[2]); r.upper = v > 0 ? v : 0; r.lower = v < 0 ? v : 0; r.check = true; r.note = '한쪽 공차만 있음'; t = '';
    }
    if (!t) return r;
    if (OK_SUFFIX.test(t)) { r.note = t; return r; }
    if (/^(MAX|MIN)\.?$/i.test(t)) { r.note = t.toUpperCase(); r.check = true; return r; }
    if (/[A-Za-z가-힣]{3,}/.test(t) || /[A-Za-z]{2,}\s+[A-Za-z]/.test(t)) { r.bad = true; return r; }
    r.note = t; r.check = true; return r;
  }
  /* 토큰 글 → { cat: 치수|공차|기하공차|용접|기타|null, kind, nominal, upper, lower, unit, count, cls, datum, check, note, reason } */
  function classifyToken(raw) {
    var s = normText(raw), m;
    var r = { cat: null, kind: '', nominal: null, upper: null, lower: null, unit: 'mm', count: null, cls: '', datum: '', check: false, note: '', reason: '' };
    if (!s) { r.reason = '빈 글자'; return r; }
    if (/^\(\s*[^()]*\d[^()]*\)$/.test(s)) { r.reason = '참고 치수(괄호)'; return r; }
    if (/\d{4}\s*[-./]\s*\d{1,2}\s*[-./]\s*\d{1,2}/.test(s)) { r.reason = '날짜'; return r; }
    if (/^\d+\s*:\s*\d+$/.test(s)) { r.reason = '척도'; return r; }
    if (/^\d+\s*\/\s*\d+$/.test(s)) { r.reason = '쪽·매수'; return r; }
    var g = gdtSym(s);
    if (g) {
      var rest = s.replace(/^\/\/|^./, '').trim(), tm = /Ø?\s*(\d*\.?\d+)/.exec(rest);
      r.cat = '기하공차'; r.kind = GDT[g]; r.unit = 'mm';
      if (tm) { r.upper = +tm[1]; r.lower = 0; } else { r.check = true; r.note = '공차값 없음'; }
      var after = tm ? rest.slice(tm.index + tm[0].length) : rest;
      r.datum = (after.match(/\b[A-Z](?:-[A-Z])?\b/g) || []).filter(function (x) { return !/^[MLP]$/.test(x); }).join(' ');
      if (/Ø/.test(rest.slice(0, tm ? tm.index + 1 : 0))) r.note = '지름 공차역';
      return r;
    }
    if (/^C\.?\s?W\.?$/i.test(s) || /^(FULL\s*PEN|CJP|PJP)$/i.test(s)) { r.cat = '용접'; r.kind = '연속 용접'; r.unit = ''; return r; }
    if ((m = new RegExp('^([az])\\s?' + NUM + '$').exec(s))) { r.cat = '용접'; r.kind = m[1] === 'a' ? '목두께' : '각장'; r.nominal = +m[2]; return r; }
    if ((m = /^(\d+)\s*-\s*(\d+)\s*\(\s*(\d+)\s*\)$/.exec(s))) { r.cat = '용접'; r.kind = '단속 용접'; r.note = '개수 ' + m[1] + ' · 길이 ' + m[2] + ' · 피치 ' + m[3]; r.unit = ''; return r; }
    if (/^[▽△◺⊿▷◁▼▲◸◹◿]+$/.test(s)) { r.cat = '용접'; r.kind = '용접 기호 글자'; r.unit = ''; r.check = true; r.note = '표면 기호일 수도 있음'; return r; }
    if ((m = new RegExp('^(Ra|Rz|Ry|Rmax)\\s*' + NUM + '$', 'i').exec(s))) { r.cat = '기타'; r.kind = '표면 거칠기'; r.nominal = +m[2]; r.unit = 'μm'; return r; }
    if ((m = new RegExp('^' + NUM + '\\s*[xX]\\s*45\\s*°$').exec(s))) { r.cat = '치수'; r.kind = '모따기'; r.nominal = +m[1]; r.note = 'x 45°'; return r; }
    m = new RegExp('^(?:(\\d+)\\s*[-xX]\\s*(?=SØ|SR|Ø|R|M\\d|C\\d))?(SØ|SR|Ø|R|M(?=\\d)|C(?=\\d))?\\s*' + NUM + '(?:\\s*[xX]\\s*' + NUM + ')?\\s*(°)?\\s*(?:[xX]\\s*45\\s*°)?(.*)$').exec(s);
    if (!m) { r.reason = '치수 형식 아님'; return r; }
    var prefix = m[2] || '', tail = parseTail(m[6]);
    if (tail.bad) { r.reason = '문장(주기)'; return r; }
    if (m[4] && prefix !== 'M') { r.reason = '치수 형식 아님'; return r; }
    r.nominal = +m[3];
    r.count = m[1] ? +m[1] : null;
    r.kind = KIND_BY_PREFIX[prefix] || (m[5] ? '각도' : '길이');
    r.unit = m[5] ? '°' : 'mm';
    if (prefix === 'M' && m[4]) r.note = '피치 ' + m[4];
    if (/x\s*45\s*°/i.test(s)) { r.kind = '모따기'; }
    r.upper = tail.upper; r.lower = tail.lower; r.cls = tail.cls; r.check = tail.check;
    if (tail.note) r.note = (r.note ? r.note + ' · ' : '') + tail.note;
    if (tail.cls) {
      var fit = prefix === 'M' ? null : isoFit(r.nominal, tail.cls);
      if (fit && r.upper == null) { r.upper = fit.upper; r.lower = fit.lower; r.note = (r.note ? r.note + ' · ' : '') + tail.cls + ' → ISO 286 환산'; }
      else if (!fit) { r.check = r.check || prefix !== 'M'; r.note = (r.note ? r.note + ' · ' : '') + (prefix === 'M' ? '나사 등급 ' + tail.cls : tail.cls + ' 표 확인'); }
    }
    r.cat = (r.upper != null || r.lower != null || r.cls) ? '공차' : '치수';
    if (r.cat === '치수' && r.kind === '길이' && r.nominal < 10 && !/\./.test(m[3])) { r.check = true; r.note = (r.note ? r.note + ' · ' : '') + '한 자리 숫자'; }
    return r;
  }

  /* ── 6. 제외 영역 ─────────────────────────────────────────── */
  var DEFAULT_TITLE_KEYWORDS = 'TITLE, DWG NO, DRAWING NO, DWG, SCALE, MATERIAL, MATL, DATE, REV, SHEET, DRAWN, CHECKED, APPROVED, DESIGNED, UNIT, PART NO, TOLERANCE, DESCRIPTION, 도면번호, 도번, 도명, 품명, 품번, 척도, 재질, 단위, 일자, 작성, 검토, 승인, 설계, 개정';
  function keywordRe(list) {
    var ks = str(list).split(/[,，\n]/).map(function (k) { return k.trim(); }).filter(Boolean)
      .map(function (k) { return k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '[\\s.]*'); });
    if (!ks.length) return null;
    return new RegExp('(^|[^A-Za-z])(' + ks.join('|') + ')(?![A-Za-z])', 'i');
  }
  function autoExcludeBoxes(pg, s) {
    s = s || {};
    var out = [], w = pg.w, h = pg.h, toks = pg.tokens || [];
    if (s.autoBorder !== false) {
      var bm = (s.borderPct == null ? 4 : s.borderPct) / 100 * Math.min(w, h);
      if (bm > 0) out.push({ x0: 0, y0: 0, x1: w, y1: bm, reason: '테두리 구역', auto: true }, { x0: 0, y0: h - bm, x1: w, y1: h, reason: '테두리 구역', auto: true },
        { x0: 0, y0: 0, x1: bm, y1: h, reason: '테두리 구역', auto: true }, { x0: w - bm, y0: 0, x1: w, y1: h, reason: '테두리 구역', auto: true });
    }
    if (s.autoTitle !== false) {
      var re = keywordRe(s.titleKeywords == null ? DEFAULT_TITLE_KEYWORDS : s.titleKeywords);
      var hits = re ? toks.filter(function (t) { return re.test(t.str) && !classifyToken(t.str).cat; }) : [];
      [['bottom', '표제란'], ['top', '개정란']].forEach(function (q) {
        var grp = hits.filter(function (t) { return t.cx > 0.45 * w && (q[0] === 'bottom' ? t.cy > 0.55 * h : t.cy < 0.3 * h); });
        if (grp.length < 2) return;
        var b = null; grp.forEach(function (t) { b = union(b, t.box); });
        var hh = Math.max.apply(null, grp.map(function (t) { return t.h; }));
        b = pad(b, hh * 1.2);
        if (q[0] === 'bottom') { b.y1 = h; } else { b.y0 = 0; }
        b.x1 = w;
        b.reason = q[1]; b.auto = true;
        out.push(b);
      });
    }
    if (s.autoNotes !== false) {
      toks.forEach(function (hd) {
        if (!/^(NOTES?|주\s?기|주\s?서|비\s?고)\s*[:：]?$/i.test(hd.str)) return;
        var lines = toks.filter(function (t) { return t !== hd && sameAngle(t.a, hd.a) && t.cy > hd.cy && t.box.x0 >= hd.box.x0 - 2 * hd.h && t.box.x0 <= hd.box.x0 + 3 * hd.h; })
          .sort(function (p, q) { return p.cy - q.cy; });
        var b = hd.box, lastY = hd.cy;
        for (var i = 0; i < lines.length; i++) {
          if (lines[i].cy - lastY > 2.6 * hd.h) break;
          var row = lines[i];
          toks.forEach(function (t) { if (Math.abs(t.cy - row.cy) < 0.4 * row.h && t.box.x0 >= row.box.x0 - row.h && t.box.x0 - row.box.x1 < 40 * row.h) b = union(b, t.box); });
          lastY = row.cy;
        }
        if (b !== hd.box) { var nb = pad(b, 2); nb.reason = '주기란'; nb.auto = true; out.push(nb); }
      });
    }
    return out;
  }
  function insideAny(t, boxes) {
    for (var i = 0; i < boxes.length; i++) if (inBox(boxes[i], t.cx, t.cy)) return boxes[i];
    return null;
  }
  /* 같은 줄에 낱말이 둘 이상 있으면 주기 문장으로 봅니다 */
  function isNoteLine(t, toks) {
    var words = 0;
    for (var i = 0; i < toks.length; i++) {
      var o = toks[i];
      if (o === t || !sameAngle(o.a, t.a) || Math.abs(o.v - t.v) > 0.4 * t.h) continue;
      var gap = Math.max(o.u0 - t.u1, t.u0 - o.u1);
      if (gap > 12 * t.h) continue;
      if (/[A-Za-z가-힣]{3,}/.test(o.str) && !classifyToken(o.str).cat) words++;
    }
    return words >= 2;
  }

  /* ── 7. 검사 대상 후보 ────────────────────────────────────── */
  var CATS = ['치수', '공차', '기하공차', '용접', '기타'];
  function candId(page, b) { return 'c' + page + '_' + Math.round((b.x0 + b.x1) / 2) + '_' + Math.round((b.y0 + b.y1) / 2); }
  function makeCand(page, t, c, extra) {
    var o = { id: candId(page, t.box), page: page, raw: t.str, box: t.box, cx: (t.box.x0 + t.box.x1) / 2, cy: (t.box.y0 + t.box.y1) / 2, h: t.h || 10,
      cat: c.cat, kind: c.kind, nominal: c.nominal, upper: c.upper, lower: c.lower, unit: c.unit, count: c.count, cls: c.cls, datum: c.datum,
      check: !!c.check, note: c.note || '', src: 'auto', tokens: t.id ? [t.id] : [] };
    return Object.assign(o, extra || {});
  }
  /* pg = { page, w, h, items, tokens, segs, subpaths, tris }, opt = { skip: {토큰id:1}, userBoxes: [], includeEtc } */
  function detectCandidates(pg, s, opt) {
    s = s || {}; opt = opt || {};
    var boxes = autoExcludeBoxes(pg, s).concat((opt.userBoxes || []).map(function (b) { return Object.assign({ reason: '사용자 제외 영역' }, b); }));
    var skip = opt.skip || {}, cands = [], excluded = [];
    (pg.tokens || []).forEach(function (t) {
      if (skip[t.id]) return;
      var c = classifyToken(t.str);
      if (!c.cat) return;
      if (c.cat === '기타' && s.includeEtc === false) return;
      var zb = insideAny(t, boxes);
      if (zb) { excluded.push({ token: t, raw: t.str, box: t.box, reason: zb.reason }); return; }
      if (s.autoNotes !== false && c.cat !== '기하공차' && isNoteLine(t, pg.tokens)) { excluded.push({ token: t, raw: t.str, box: t.box, reason: '주기 문장' }); return; }
      cands.push(makeCand(pg.page, t, c));
    });
    // 용접 삼각형 옆 글자를 한 항목으로 묶고, 글자가 없으면 「용접 형상」
    (pg.tris || []).forEach(function (tri) {
      if (insideAny({ cx: tri.cx, cy: tri.cy }, boxes)) return;
      var near = cands.filter(function (c) {
        if (c.used) return false;
        var weldish = c.cat === '용접' || (c.cat === '치수' && c.kind === '길이' && c.upper == null);
        return weldish && boxDist(c.box, tri.box) <= 0.85 * c.h && Math.abs(c.cy - tri.cy) < 1.5 * c.h;
      }).sort(function (p, q) { return p.cx - q.cx; });
      var b = tri.box, raws = [], nominal = null, kind = '필릿', note = [], toks = [];
      near.forEach(function (c) {
        c.used = true; b = union(b, c.box); raws.push(c.raw); toks = toks.concat(c.tokens);
        if (c.cat === '치수' && nominal == null) nominal = c.nominal;
        else if (c.kind === '목두께' || c.kind === '각장') { kind = c.kind; if (nominal == null) nominal = c.nominal; }
        else if (c.kind === '단속 용접') { kind = '단속 필릿'; note.push(c.note); }
        else if (c.kind === '연속 용접') note.push('연속 용접(' + c.raw + ')');
      });
      var shape = !near.length;
      var bb = shape ? pad(tri.box, 8) : b;
      cands.push({ id: candId(pg.page, bb), page: pg.page, raw: shape ? '(용접 형상)' : raws.join(' '), box: bb, cx: (bb.x0 + bb.x1) / 2, cy: (bb.y0 + bb.y1) / 2, h: near[0] ? near[0].h : 10,
        cat: '용접', kind: shape ? '용접 형상' : kind, nominal: nominal, upper: null, lower: null, unit: shape ? '' : 'mm', count: null, cls: '', datum: '',
        check: shape, note: shape ? '글자 없는 용접 기호 — 형상 그림으로 확인' : note.join(' · '), src: 'auto', shape: shape, crop: shape ? pad(tri.box, 22) : null, tokens: toks });
    });
    cands = cands.filter(function (c) { return !c.used; });
    return { cands: cands, excluded: excluded, boxes: boxes };
  }

  /* ── 8. 번호 순서 ─────────────────────────────────────────── */
  var ROWS = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
  function zoneOf(c, pg, s) {
    var cols = Math.max(1, s.zoneCols || 8), rows = Math.max(1, s.zoneRows || 6);
    var ci = Math.min(cols - 1, Math.max(0, Math.floor(c.cx / (pg.w / cols)))), ri = Math.min(rows - 1, Math.max(0, Math.floor(c.cy / (pg.h / rows))));
    return { col: ci, row: ri, label: ROWS.charAt(ri) + (ci + 1) };
  }
  /* mode: rows(위→아래 띠, 띠 안에서 왼→오) · cols(왼→오 띠, 띠 안에서 위→아래) · zone(구역 A1, A2 … 순서) */
  function orderCandidates(cands, pg, s) {
    s = s || {};
    var mode = s.order || 'rows', n = Math.max(1, s.bands || 6);
    var list = cands.map(function (c) { return { c: c, z: zoneOf(c, pg, s) }; });
    list.sort(function (p, q) {
      var a = p.c, b = q.c;
      if (mode === 'cols') { var bw = pg.w / n, ka = Math.floor(a.cx / bw), kb = Math.floor(b.cx / bw); return ka - kb || a.cy - b.cy || a.cx - b.cx; }
      if (mode === 'zone') return p.z.row - q.z.row || p.z.col - q.z.col || a.cy - b.cy || a.cx - b.cx;
      var bh = pg.h / n, ra = Math.floor(a.cy / bh), rb = Math.floor(b.cy / bh);
      return ra - rb || a.cx - b.cx || a.cy - b.cy;
    });
    return list.map(function (x) { return x.c; });
  }

  /* ── 9. 풍선 자리 잡기 ────────────────────────────────────── */
  function circleHitsBox(cx, cy, r, b) { return ptBoxDist(b, cx, cy) < r; }
  /* cands(번호 순서대로, c.no 있음) → [{ id, no, x, y, r, leader, overlap }]. fixed = { 후보id: {x,y} } 사람이 옮긴 자리 */
  function placeBalloons(cands, pg, s, fixed) {
    s = s || {}; fixed = fixed || {};
    var r = toPt(s.balloonR || 4), margin = Math.min(pg.w, pg.h) * 0.01;
    var texts = (pg.tokens || []).map(function (t) { return pad(t.box, 1); });
    var userBoxes = (s.userBoxes || []);
    var segs = pg.segs || [], placed = [], out = [];
    cands.forEach(function (c) { var f = fixed[c.id]; if (f) placed.push({ x: f.x, y: f.y, r: r }); });
    var dirs = [];
    [-45, -90, 0, -135, 180, 45, 90, 135, -22.5, -67.5, 22.5, 157.5].forEach(function (d) { var t = d * Math.PI / 180; dirs.push([Math.cos(t), Math.sin(t)]); });
    cands.forEach(function (c) {
      var f = fixed[c.id], b = c.box, cx = (b.x0 + b.x1) / 2, cy = (b.y0 + b.y1) / 2;
      if (f) { out.push({ id: c.id, no: c.no, x: f.x, y: f.y, r: r, leader: leaderFor(f.x, f.y, r, b), overlap: false, manual: true }); return; }
      var best = null;
      [r * 0.35 + 2, r * 1.6 + 2, r * 3 + 2, r * 4.5 + 2].forEach(function (gap, gi) {
        dirs.forEach(function (d, di) {
          var ext = Math.abs(d[0]) * (b.x1 - b.x0) / 2 + Math.abs(d[1]) * (b.y1 - b.y0) / 2;
          var x = cx + d[0] * (ext + gap + r), y = cy + d[1] * (ext + gap + r), pen = gi * 6 + di * 0.4;
          if (x - r < margin || y - r < margin || x + r > pg.w - margin || y + r > pg.h - margin) pen += 1000;
          for (var i = 0; i < placed.length; i++) if (dist(x, y, placed[i].x, placed[i].y) < 2 * r + 2) pen += 500;
          for (var k = 0; k < texts.length; k++) if (circleHitsBox(x, y, r, texts[k])) pen += 200;
          for (var u = 0; u < userBoxes.length; u++) if (inBox(userBoxes[u], x, y)) pen += 300;
          var cross = 0;
          for (var q = 0; q < segs.length && cross < 6; q++) {
            var sg = segs[q];
            if (Math.min(sg.x1, sg.x2) > x + r || Math.max(sg.x1, sg.x2) < x - r || Math.min(sg.y1, sg.y2) > y + r || Math.max(sg.y1, sg.y2) < y - r) continue;
            if (ptSegDist(x, y, sg) < r) cross++;
          }
          pen += cross * 4;
          if (!best || pen < best.pen) best = { x: x, y: y, pen: pen, gi: gi };
        });
      });
      placed.push({ x: best.x, y: best.y, r: r });
      out.push({ id: c.id, no: c.no, x: round(best.x, 2), y: round(best.y, 2), r: r, leader: best.gi > 0 ? leaderFor(best.x, best.y, r, b) : null, overlap: best.pen >= 200 });
    });
    return out;
  }
  function leaderFor(x, y, r, b) {
    var p = nearestOnBox(b, x, y), d = dist(x, y, p[0], p[1]);
    if (d <= r * 1.25) return null;
    return [round(x + (p[0] - x) / d * r, 2), round(y + (p[1] - y) / d * r, 2), round(p[0], 2), round(p[1], 2)];
  }

  /* ── 10. 풍선 읽기 · 지시선 연결 ──────────────────────────── */
  /* 원 안의 작은 정수 = 풍선. s.isolatedBalloons 이면 원 없이 홀로 있는 정수도(확인 필요) */
  function detectBalloons(pg, s) {
    s = s || {};
    var out = [], used = {};
    var ints = (pg.tokens || []).filter(function (t) { return /^\d{1,3}$/.test(t.str) && +t.str > 0; });
    var circles = pg.circles || [];
    ints.forEach(function (t) {
      var best = null;
      circles.forEach(function (c, ci) {
        if (used[ci]) return;
        var d = dist(t.cx, t.cy, c.cx, c.cy);
        if (d < 0.6 * c.r && c.r >= 0.7 * t.h && c.r <= 3.6 * t.h && (t.box.x1 - t.box.x0) < 2.1 * c.r && (!best || d < best.d)) best = { c: c, ci: ci, d: d };
      });
      if (best) { used[best.ci] = 1; out.push({ no: +t.str, cx: best.c.cx, cy: best.c.cy, r: best.c.r, tokenId: t.id, page: pg.page, src: 'circle' }); }
    });
    if (s.isolatedBalloons) {
      var zones = autoExcludeBoxes(pg, s);
      ints.forEach(function (t) {
        if (out.some(function (b) { return b.tokenId === t.id; }) || insideAny(t, zones)) return;
        var alone = (pg.tokens || []).every(function (o) { return o === t || boxDist(o.box, t.box) > 2.5 * t.h; });
        if (alone) out.push({ no: +t.str, cx: t.cx, cy: t.cy, r: t.h * 1.1, tokenId: t.id, page: pg.page, src: 'isolated', check: true });
      });
    }
    return out.sort(function (a, b) { return a.no - b.no; });
  }
  function leaderTarget(b, segs) {
    var tolR = Math.max(1.6, 0.18 * b.r), best = null;
    segs.forEach(function (s, i) {
      [[s.x1, s.y1, s.x2, s.y2], [s.x2, s.y2, s.x1, s.y1]].forEach(function (e) {
        var dn = dist(e[0], e[1], b.cx, b.cy), df = dist(e[2], e[3], b.cx, b.cy);
        if (Math.abs(dn - b.r) <= tolR && df > b.r * 1.4) {
          var L = dist(e[0], e[1], e[2], e[3]);
          if (!best || L > best.L) best = { i: i, x: e[2], y: e[3], L: L, sx: e[0], sy: e[1] };
        }
      });
    });
    if (!best) return null;
    var path = [[best.sx, best.sy], [best.x, best.y]], usedSeg = {}; usedSeg[best.i] = 1;
    var px = best.x, py = best.y;
    for (var hop = 0; hop < 3; hop++) {
      var next = null;
      for (var k = 0; k < segs.length; k++) {
        if (usedSeg[k]) continue;
        var s2 = segs[k];
        if (dist(s2.x1, s2.y1, px, py) < 1.2) next = { k: k, x: s2.x2, y: s2.y2 };
        else if (dist(s2.x2, s2.y2, px, py) < 1.2) next = { k: k, x: s2.x1, y: s2.y1 };
        if (next) break;
      }
      if (!next || dist(next.x, next.y, b.cx, b.cy) < dist(px, py, b.cx, b.cy)) break;
      usedSeg[next.k] = 1; px = next.x; py = next.y; path.push([px, py]);
    }
    return { x: px, y: py, path: path };
  }
  /* 풍선 → 가리키는 검사 대상 후보. 지시선이 있으면 그 끝에서, 없으면 풍선 둘레에서 가장 가까운 후보(겹치지 않게 전체 배정) */
  function linkBalloons(balloons, cands, pg, s) {
    s = s || {};
    var linkMax = toPt(s.linkMaxMm || 15), nearMax = toPt(s.nearMaxMm || 14), hop = toPt(4), segs = pg.segs || [];
    var info = balloons.map(function (b) { return { b: b, lead: leaderTarget(b, segs) }; });
    // 풍선마다 닿을 수 있는 후보와 거리. 지시선이 꺾여 이어지면 앞쪽 꼭짓점일수록 우선(용접 기준선 끝에서 멈추도록)
    var opts = info.map(function (x) {
      var list = [];
      cands.forEach(function (c, ci) {
        var d = null, method;
        if (x.lead) {
          method = 'leader';
          for (var k = 1; k < x.lead.path.length; k++) {
            var dk = ptBoxDist(c.box, x.lead.path[k][0], x.lead.path[k][1]);
            if (dk <= linkMax) { var v = dk + (k - 1) * hop; if (d == null || v < d) d = v; }
          }
          if (d == null) return;
          d = d * 0.5;
        } else { d = Math.max(0, ptBoxDist(c.box, x.b.cx, x.b.cy) - x.b.r); method = 'near'; if (d > nearMax) return; }
        list.push({ ci: ci, d: d, method: method });
      });
      return list.sort(function (p, q) { return p.d - q.d; });
    });
    // 가까운 쪽부터 짝을 짓되, 이미 짝이 있는 후보는 그 풍선이 다른 후보로 옮겨 갈 수 있으면 비켜 줍니다(최대 짝짓기)
    var owner = {}, pick = info.map(function () { return null; });
    function tryAssign(bi, seen) {
      for (var i = 0; i < opts[bi].length; i++) {
        var o = opts[bi][i];
        if (seen[o.ci]) continue;
        seen[o.ci] = 1;
        if (owner[o.ci] == null || tryAssign(owner[o.ci], seen)) { owner[o.ci] = bi; pick[bi] = o; return true; }
      }
      return false;
    }
    info.map(function (x, bi) { return bi; }).sort(function (a, b) { return (opts[a][0] ? opts[a][0].d : 1e9) - (opts[b][0] ? opts[b][0].d : 1e9); })
      .forEach(function (bi) { tryAssign(bi, {}); });
    return info.map(function (x, bi) {
      var r = { no: x.b.no, page: x.b.page, balloon: x.b, leader: x.lead ? x.lead.path : null, target: null, candId: null, method: 'none', dist: null };
      var p = pick[bi];
      if (p && owner[p.ci] === bi) { r.candId = cands[p.ci].id; r.method = p.method; r.dist = round(p.method === 'leader' ? p.d * 2 : p.d, 1); }
      if (x.lead) { var e = x.lead.path[Math.min(x.lead.path.length - 1, 1)]; r.target = [e[0], e[1]]; }
      if (!r.candId) {
        var t = r.target || [x.b.cx, x.b.cy], half = toPt(12);
        r.crop = { x0: t[0] - half, y0: t[1] - half, x1: t[0] + half, y1: t[1] + half };
        r.check = true;
      }
      return r;
    });
  }

  /* ── 11. 일치율(정답 대조) ────────────────────────────────── */
  function catNorm(v) {
    var s = str(v).replace(/\s+/g, '');
    if (!s) return '';
    if (/형상/.test(s)) return '용접 형상';
    if (/기하|GD&?T|geometric/i.test(s)) return '기하공차';
    if (/용접|weld/i.test(s)) return '용접';
    if (/공차|tol/i.test(s)) return '공차';
    if (/치수|dim/i.test(s)) return '치수';
    if (/기타|거칠기|etc/i.test(s)) return '기타';
    return s;
  }
  function numOrNull(v) {
    if (v == null || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var s = normText(v).replace(/^[Ø±R+]/, '').replace(/[^\d.\-+]/g, '');
    var n = parseFloat(s);
    return isFinite(n) ? n : null;
  }
  var TRUTH_FIELDS = {
    no: /^(번호|풍선|풍선번호|no\.?|#|item|순번)$/i, cat: /^(구분|분류|종류구분|category|type|항목구분)$/i, kind: /^(종류|세부|kind|형태)$/i,
    raw: /^(원문|도면\s?표기|spec|사양|표기|규격|치수\s?표기|내용)$/i, nominal: /^(공칭값?|기준값?|기준\s?치수|nominal|설계값|치수)$/i,
    upper: /^(상한(공차|편차)?|\+?\s?공차|upper|상한값?)$/i, lower: /^(하한(공차|편차)?|-\s?공차|lower|하한값?)$/i, unit: /^(단위|unit)$/i, page: /^(쪽|페이지|page|매)$/i
  };
  /* 정답 표(2차원 배열) → 항목. 머리행을 위 10줄에서 찾습니다 */
  function parseTruth(aoa) {
    var rows = aoa || [], hi = -1, map = {};
    for (var i = 0; i < Math.min(rows.length, 10) && hi < 0; i++) {
      var m = {};
      (rows[i] || []).forEach(function (cell, ci) {
        var k = str(cell).replace(/\s+/g, ' ');
        Object.keys(TRUTH_FIELDS).forEach(function (f) { if (m[f] == null && TRUTH_FIELDS[f].test(k)) m[f] = ci; });
      });
      if (m.no != null && (m.raw != null || m.nominal != null || m.cat != null)) { hi = i; map = m; }
    }
    if (hi < 0) return { items: [], warnings: ['머리행(번호 · 구분 · 원문 또는 공칭값)을 찾지 못했습니다.'] };
    var items = [], warnings = [];
    for (var r = hi + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var no = numOrNull(row[map.no]);
      if (no == null) continue;
      var raw = map.raw != null ? str(row[map.raw]) : '';
      var it = { no: no, cat: catNorm(map.cat != null ? row[map.cat] : ''), kind: map.kind != null ? str(row[map.kind]) : '', raw: raw,
        nominal: map.nominal != null ? numOrNull(row[map.nominal]) : null, upper: map.upper != null ? numOrNull(row[map.upper]) : null,
        lower: map.lower != null ? numOrNull(row[map.lower]) : null, unit: map.unit != null ? str(row[map.unit]) : '', page: map.page != null ? numOrNull(row[map.page]) : null };
      if ((!it.cat || (it.nominal == null && it.upper == null)) && raw) {
        var c = classifyToken(raw);
        if (!it.cat && c.cat) it.cat = c.cat;
        if (it.nominal == null && it.upper == null && c.cat) { it.nominal = c.nominal; it.upper = c.upper; it.lower = c.lower; }
      }
      if (it.cat === '용접' && /형상/.test(it.kind + it.raw)) it.cat = '용접 형상';
      if (!it.cat) warnings.push(no + '번: 구분을 알 수 없어 「기타」로 봅니다.');
      items.push(it);
    }
    return { items: items, warnings: warnings, headerRow: hi };
  }
  function cmpCat(it) { var c = catNorm(it.cat); if (c === '용접' && /형상/.test(str(it.kind))) return '용접 형상'; return c || '기타'; }
  function eqNum(a, b, zeroBlank) {
    if (a == null && b == null) return true;
    if (zeroBlank && ((a == null && Math.abs(b) < 1e-9) || (b == null && Math.abs(a) < 1e-9))) return true;
    return a != null && b != null && Math.abs(a - b) < 1e-6;
  }
  function sameValue(p, t) {
    var cp = cmpCat(p), ct = cmpCat(t);
    if (cp !== ct) return false;
    if (ct === '용접 형상') return true;
    return eqNum(p.nominal, t.nominal) && eqNum(p.upper, t.upper, true) && eqNum(p.lower, t.lower, true);
  }
  function itemText(it) {
    var s = it.raw || '';
    if (!s) s = (it.nominal != null ? it.nominal : '') + (it.upper != null || it.lower != null ? ' ' + (it.upper != null ? it.upper : '') + '/' + (it.lower != null ? it.lower : '') : '');
    return cmpCat(it) + ' ' + s;
  }
  /* mode = 'number'(풍선 번호로 짝, 목표 2) · 'value'(값으로 짝, 번호는 다를 수 있음, 목표 1) */
  function scoreResults(pred, truth, mode) {
    pred = pred || []; truth = truth || [];
    var pairs = [], usedP = {}, usedT = {};
    if (mode === 'number') {
      var byNo = {};
      pred.forEach(function (p, i) { if (p.no != null && byNo[p.no] == null) byNo[p.no] = i; });
      truth.forEach(function (t, ti) {
        var pi = byNo[t.no];
        if (pi == null || usedP[pi]) return;
        var p = pred[pi];
        if (!p.cat && p.nominal == null && p.upper == null) return; // 풍선만 찾고 연결 못 함 = 놓침
        usedP[pi] = usedT[ti] = 1; pairs.push({ ti: ti, pi: pi, exact: sameValue(p, t) });
      });
    } else {
      truth.forEach(function (t, ti) {
        for (var i = 0; i < pred.length; i++) if (!usedP[i] && sameValue(pred[i], t)) { usedP[i] = usedT[ti] = 1; pairs.push({ ti: ti, pi: i, exact: true }); return; }
      });
      truth.forEach(function (t, ti) {
        if (usedT[ti]) return;
        for (var i = 0; i < pred.length; i++) {
          if (usedP[i]) continue;
          var p = pred[i], ct = cmpCat(t);
          var close = (t.nominal != null && eqNum(p.nominal, t.nominal)) || (ct === '기하공차' && cmpCat(p) === ct && eqNum(p.upper, t.upper));
          if (close) { usedP[i] = usedT[ti] = 1; pairs.push({ ti: ti, pi: i, exact: false }); return; }
        }
      });
    }
    var tp = pairs.length, nP = pred.filter(function (p) { return p.cat || p.nominal != null; }).length, nT = truth.length;
    var exact = pairs.filter(function (p) { return p.exact; }).length;
    var numSame = pairs.filter(function (p) { return pred[p.pi].no != null && pred[p.pi].no === truth[p.ti].no; }).length;
    var precision = nP ? tp / nP : 0, recall = nT ? tp / nT : 0, f1 = precision + recall ? 2 * precision * recall / (precision + recall) : 0;
    var byCat = {};
    truth.forEach(function (t) { var c = cmpCat(t); byCat[c] = byCat[c] || { truth: 0, found: 0, exact: 0 }; byCat[c].truth++; });
    pairs.forEach(function (p) { var c = cmpCat(truth[p.ti]); byCat[c].found++; if (p.exact) byCat[c].exact++; });
    var misses = truth.map(function (t, i) { return usedT[i] ? null : { no: t.no, text: itemText(t) }; }).filter(Boolean);
    var wrong = pairs.filter(function (p) { return !p.exact; }).map(function (p) { return { no: truth[p.ti].no, truth: itemText(truth[p.ti]), pred: itemText(pred[p.pi]), predNo: pred[p.pi].no }; });
    var extra = pred.map(function (p, i) { return usedP[i] || !(p.cat || p.nominal != null) ? null : { no: p.no, text: itemText(p) }; }).filter(Boolean);
    return { mode: mode, nPred: nP, nTruth: nT, tp: tp, fp: nP - tp, fn: nT - tp, precision: round(precision, 4), recall: round(recall, 4), f1: round(f1, 4),
      exact: exact, valueRate: round(nT ? exact / nT : 0, 4), numberRate: round(tp ? numSame / tp : 0, 4), byCat: byCat, misses: misses, wrong: wrong, extra: extra };
  }

  /* ── 12. 풍선 PDF 쓰기 ────────────────────────────────────── */
  /* pages: [{ w, h, jpeg: Uint8Array|null, imgW, imgH, texts: [{str,x,y,w,h,a}], balloons: [{x,y,r,no,leader}] }]
     원본 쪽 그림(JPEG) 위에 풍선을 벡터로 그리고, 원본 글자는 보이지 않는 글자 층으로 같은 자리에 넣습니다(다시 읽기·검색용). */
  function latin1(s) { var a = new Uint8Array(s.length); for (var i = 0; i < s.length; i++) a[i] = s.charCodeAt(i) & 255; return a; }
  function f2(n) { return (Math.round(n * 100) / 100).toString(); }
  function hex4(n) { return ('000' + n.toString(16).toUpperCase()).slice(-4); }
  function circlePath(cx, cy, r) {
    var k = 0.5523 * r;
    return f2(cx + r) + ' ' + f2(cy) + ' m ' + [[cx + r, cy + k, cx + k, cy + r, cx, cy + r], [cx - k, cy + r, cx - r, cy + k, cx - r, cy], [cx - r, cy - k, cx - k, cy - r, cx, cy - r], [cx + k, cy - r, cx + r, cy - k, cx + r, cy]]
      .map(function (q) { return q.map(f2).join(' ') + ' c'; }).join(' ') + ' h';
  }
  function buildBalloonPdf(pages, opt) {
    opt = opt || {};
    var color = opt.color || [0.8, 0.1, 0.1];
    var objs = [], his = {};
    function add(body) { objs.push(body); return objs.length; }
    var catalog = add(null), pagesId = add(null);
    var fHel = add('<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>');
    var allChars = {};
    pages.forEach(function (p) { (p.texts || []).forEach(function (t) { for (var i = 0; i < t.str.length; i++) { var c = t.str.charCodeAt(i); if (c < 0xD800 || c > 0xDFFF) his[c >> 8] = 1; allChars[c] = 1; } }); });
    var ranges = Object.keys(his).map(function (k) { var hb = +k; return '<' + hex4(hb << 8) + '> <' + hex4((hb << 8) | 255) + '> <' + hex4(hb << 8) + '>'; });
    var cmap = '/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n' +
      ranges.length + ' beginbfrange\n' + ranges.join('\n') + '\nendbfrange\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend';
    var toUni = add({ stream: latin1(cmap) });
    var fDesc = add('<< /Type /FontDescriptor /FontName /GlyphLessFont /Flags 5 /FontBBox [0 0 500 1000] /ItalicAngle 0 /Ascent 1000 /Descent 0 /CapHeight 1000 /StemV 80 >>');
    var cid = add('<< /Type /Font /Subtype /CIDFontType2 /BaseFont /GlyphLessFont /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ' + fDesc + ' 0 R /DW 500 /CIDToGIDMap /Identity >>');
    var fInv = add('<< /Type /Font /Subtype /Type0 /BaseFont /GlyphLessFont /Encoding /Identity-H /DescendantFonts [' + cid + ' 0 R] /ToUnicode ' + toUni + ' 0 R >>');
    var kids = [];
    pages.forEach(function (p) {
      var H = p.h, c = [];
      var img = null;
      if (p.jpeg) {
        img = add({ dict: '/Type /XObject /Subtype /Image /Width ' + p.imgW + ' /Height ' + p.imgH + ' /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode', stream: p.jpeg });
        c.push('q ' + f2(p.w) + ' 0 0 ' + f2(H) + ' 0 0 cm /Im0 Do Q');
      }
      if ((p.texts || []).length) {
        c.push('BT 3 Tr');
        p.texts.forEach(function (t) {
          var n = t.str.length; if (!n || !(t.h > 0)) return;
          var th = -t.a, sx = Math.max(0.05, (t.w || n * t.h * 0.5) / (n * 0.5 * t.h));
          var m = [t.h * Math.cos(th) * sx, t.h * Math.sin(th) * sx, -t.h * Math.sin(th), t.h * Math.cos(th), t.x, H - t.y];
          var hx = ''; for (var i = 0; i < n; i++) { var cc = t.str.charCodeAt(i); hx += hex4(cc >= 0xD800 && cc <= 0xDFFF ? 63 : cc); }
          c.push('/F2 1 Tf ' + m.map(f2).join(' ') + ' Tm <' + hx + '> Tj');
        });
        c.push('ET');
      }
      if (p.segs || p.tris) {
        // 원본의 선(지시선·용접 기호)을 칠하지 않은 경로(n)로 남겨, 이 PDF 를 다시 읽어도 지시선·용접 삼각형을 찾게 합니다
        var pth = [];
        (p.segs || []).forEach(function (sg) { pth.push(f2(sg.x1) + ' ' + f2(H - sg.y1) + ' m ' + f2(sg.x2) + ' ' + f2(H - sg.y2) + ' l n'); });
        (p.tris || []).forEach(function (t) { if (t.pts) pth.push(t.pts.map(function (q, i) { return f2(q[0]) + ' ' + f2(H - q[1]) + (i ? ' l' : ' m'); }).join(' ') + ' h n'); });
        c.push(pth.join('\n'));
      }
      (p.balloons || []).forEach(function (b) {
        var cy = H - b.y, col = color.join(' ');
        c.push('q ' + col + ' RG 1 1 1 rg 0.9 w ' + circlePath(b.x, cy, b.r) + ' B');
        if (b.leader) c.push(f2(b.leader[0]) + ' ' + f2(H - b.leader[1]) + ' m ' + f2(b.leader[2]) + ' ' + f2(H - b.leader[3]) + ' l S');
        var label = String(b.no), fs = Math.max(5, b.r * (label.length > 2 ? 0.8 : 1.05)), tw = 0.556 * fs * label.length;
        c.push('BT ' + col + ' rg /F1 ' + f2(fs) + ' Tf ' + f2(b.x - tw / 2) + ' ' + f2(cy - fs * 0.36) + ' Td (' + label + ') Tj ET Q');
      });
      var content = add({ stream: latin1(c.join('\n')) });
      var pg = add('<< /Type /Page /Parent ' + pagesId + ' 0 R /MediaBox [0 0 ' + f2(p.w) + ' ' + f2(H) + '] /Resources << /Font << /F1 ' + fHel + ' 0 R /F2 ' + fInv + ' 0 R >>' +
        (img ? ' /XObject << /Im0 ' + img + ' 0 R >>' : '') + ' >> /Contents ' + content + ' 0 R >>');
      kids.push(pg);
    });
    objs[catalog - 1] = '<< /Type /Catalog /Pages ' + pagesId + ' 0 R >>';
    objs[pagesId - 1] = '<< /Type /Pages /Kids [' + kids.map(function (k) { return k + ' 0 R'; }).join(' ') + '] /Count ' + kids.length + ' >>';
    var chunks = [latin1('%PDF-1.4\n%\xE2\xE3\xCF\xD3\n')], offs = [], pos = chunks[0].length;
    objs.forEach(function (o, i) {
      offs.push(pos);
      var parts;
      if (typeof o === 'string') parts = [latin1((i + 1) + ' 0 obj\n' + o + '\nendobj\n')];
      else parts = [latin1((i + 1) + ' 0 obj\n<< ' + (o.dict ? o.dict + ' ' : '') + '/Length ' + o.stream.length + ' >>\nstream\n'), o.stream, latin1('\nendstream\nendobj\n')];
      parts.forEach(function (x) { chunks.push(x); pos += x.length; });
    });
    var xref = 'xref\n0 ' + (objs.length + 1) + '\n0000000000 65535 f \n' + offs.map(function (o) { return ('000000000' + o).slice(-10) + ' 00000 n \n'; }).join('') +
      'trailer\n<< /Size ' + (objs.length + 1) + ' /Root ' + catalog + ' 0 R /Info << /Producer (data10-03 balloon tool) >> >>\nstartxref\n' + pos + '\n%%EOF\n';
    chunks.push(latin1(xref));
    var total = chunks.reduce(function (a, b) { return a + b.length; }, 0), out = new Uint8Array(total), k = 0;
    chunks.forEach(function (x) { out.set(x, k); k += x.length; });
    return out;
  }

  /* ── 13. AI 프롬프트(애매한 후보) ─────────────────────────── */
  function nearText(c, toks, k) {
    return (toks || []).filter(function (t) { return c.tokens.indexOf(t.id) < 0 && boxDist(t.box, c.box) < (c.h || 10) * 4; })
      .sort(function (a, b) { return boxDist(a.box, c.box) - boxDist(b.box, c.box); }).slice(0, k || 4).map(function (t) { return t.str; });
  }
  function aiPrompt(list, ctx) {
    ctx = ctx || {};
    var lines = [
      '너는 기계 도면 검사 기준서를 만드는 품질 담당자를 돕는다. 아래는 벡터 PDF 도면에서 규칙으로 뽑은 「검사 대상 후보」 중 판단이 애매한 것들이다.',
      '각 후보가 검사 대상인지, 대상이면 구분(치수 · 공차 · 기하공차 · 용접 · 기타)과 공칭값 · 상한공차 · 하한공차를 판단해줘.',
      '규칙: 표제란 · 주기 · 참고 치수(괄호) · 테두리 구역 번호는 검사 대상이 아니다. 끼워맞춤 기호(H7 · g6 등)는 ISO 286 기준으로 공차를 계산해줘. 확실하지 않으면 「확인 필요」라고 적어줘.',
      '답은 표 한 줄에 후보 하나씩, 이 형식 그대로 써줘:',
      '| ID | 검사대상(예/아니오) | 구분 | 공칭값 | 상한공차 | 하한공차 | 근거 |',
      '',
      '도면: ' + (ctx.name || '(이름 없음)') + (ctx.page ? ' · ' + ctx.page + '쪽' : ''),
      '', '| ID | 도면 글자 | 규칙이 본 구분 | 주변 글자 |', '|---|---|---|---|'
    ];
    list.forEach(function (c) { lines.push('| ' + c.id + ' | ' + c.raw + ' | ' + (c.cat || '?') + (c.kind ? '(' + c.kind + ')' : '') + ' | ' + (c.near || []).join(' / ') + ' |'); });
    return lines.join('\n');
  }
  function parseAiAnswer(text, ids) {
    var out = [], known = {};
    (ids || []).forEach(function (i) { known[i] = 1; });
    str(text).split(/\r?\n/).forEach(function (line) {
      var cells = line.indexOf('|') >= 0 ? line.split('|').map(str).filter(function (x, i, a) { return !(x === '' && (i === 0 || i === a.length - 1)); }) : line.split('\t').map(str);
      if (cells.length < 3) return;
      var id = cells[0];
      if (!known[id]) return;
      var yes = !/아니|no|x|제외/i.test(cells[1]);
      out.push({ id: id, include: yes, cat: catNorm(cells[2]) === '용접 형상' ? '용접' : catNorm(cells[2]), nominal: numOrNull(cells[3]), upper: numOrNull(cells[4]), lower: numOrNull(cells[5]), reason: cells[6] || '' });
    });
    return out;
  }

  /* ── 14. 엑셀 행 · 저장 형식 ──────────────────────────────── */
  var ROW_HEAD = ['번호', '쪽', '구역', '구분', '종류', '원문', '공칭값', '상한공차', '하한공차', '최대 허용값', '최소 허용값', '단위', '수량', '데이텀', 'X(mm)', 'Y(mm)', '풍선 X(mm)', '풍선 Y(mm)', '연결 방식', '확인 필요', '메모', '형상 잘라낼 영역(mm, 왼쪽 위 X·Y · 오른쪽 아래 X·Y)'];
  function limitVal(nom, dev) { return nom != null && dev != null ? round(nom + dev, 4) : null; }
  function itemRow(x) {
    var c = x.cand || {}, b = x.balloon || {}, crop = x.crop || c.crop;
    var cat = c.cat ? c.cat : '연결 못 함';
    return [x.no, x.page, x.zone || '', cat, c.kind || '', c.raw || '', c.nominal, c.upper, c.lower,
      c.cat === '기하공차' ? null : limitVal(c.nominal, c.upper), c.cat === '기하공차' ? null : limitVal(c.nominal, c.lower), c.unit || '', c.count, c.datum || '',
      c.cx != null ? mm(c.cx) : null, c.cy != null ? mm(c.cy) : null, b.x != null ? mm(b.x) : null, b.y != null ? mm(b.y) : null, x.method || '',
      (x.check || c.check) ? '확인 필요' : '', [c.note, x.note].filter(Boolean).join(' · '), crop ? [mm(crop.x0), mm(crop.y0), mm(crop.x1), mm(crop.y1)].join(', ') : ''];
  }
  function toPred(x) { var c = x.cand || {}; return { no: x.no, cat: c.cat === '용접' && c.kind === '용접 형상' ? '용접 형상' : c.cat, kind: c.kind, raw: c.raw, nominal: c.nominal, upper: c.upper, lower: c.lower }; }
  var DEFAULTS = {
    borderPct: 4, autoBorder: true, autoTitle: true, autoNotes: true, includeEtc: true, titleKeywords: DEFAULT_TITLE_KEYWORDS,
    order: 'rows', bands: 6, zoneCols: 8, zoneRows: 6, startNo: 1, balloonR: 4, linkMaxMm: 15, nearMaxMm: 14, isolatedBalloons: false, target: 85
  };
  function cleanSettings(o) {
    var s = clone(DEFAULTS);
    if (!o || typeof o !== 'object') return s;
    ['autoBorder', 'autoTitle', 'autoNotes', 'includeEtc', 'isolatedBalloons'].forEach(function (k) { if (typeof o[k] === 'boolean') s[k] = o[k]; });
    [['borderPct', 0, 15], ['bands', 1, 30], ['zoneCols', 1, 26], ['zoneRows', 1, 20], ['startNo', 1, 9999], ['balloonR', 1.5, 12], ['linkMaxMm', 2, 60], ['nearMaxMm', 2, 60], ['target', 50, 100]]
      .forEach(function (k) { var v = Number(o[k[0]]); if (isFinite(v) && v >= k[1] && v <= k[2]) s[k[0]] = v; });
    if (['rows', 'cols', 'zone'].indexOf(o.order) >= 0) s.order = o.order;
    if (typeof o.titleKeywords === 'string') s.titleKeywords = o.titleKeywords;
    return s;
  }
  var SCHEMA_VERSION = 1;
  function emptyDb() { return { v: SCHEMA_VERSION, settings: clone(DEFAULTS), docs: {}, history: [], cur: '' }; }
  function cleanDoc(d) {
    if (!d || typeof d !== 'object') return null;
    var o = { name: str(d.name), size: Number(d.size) || 0, pages: Number(d.pages) || 0, sample: str(d.sample), opened: str(d.opened) };
    ['excl', 'removed', 'changed', 'added', 'nums', 'bpos', 'linkOv', 'truth'].forEach(function (k) { o[k] = d[k] && typeof d[k] === 'object' ? d[k] : (k === 'truth' ? null : {}); });
    return o;
  }
  function restoreDb(o) {
    var db = emptyDb();
    if (!o || typeof o !== 'object') return db;
    db.settings = cleanSettings(o.settings);
    if (o.docs && typeof o.docs === 'object') Object.keys(o.docs).forEach(function (k) { var d = cleanDoc(o.docs[k]); if (d) db.docs[k] = d; });
    if (Array.isArray(o.history)) db.history = o.history.filter(function (h) { return h && typeof h === 'object' && h.at; }).slice(-200);
    db.cur = db.docs[o.cur] ? o.cur : '';
    return db;
  }
  function docKey(name, size) { return str(name) + '|' + (Number(size) || 0); }

  return {
    PT2MM: PT2MM, mm: mm, toPt: toPt, round: round, box: box, union: union, pad: pad, inBox: inBox, boxDist: boxDist, ptBoxDist: ptBoxDist, boxOverlap: boxOverlap, dist: dist, center: center,
    mul: mul, pdfTextItems: pdfTextItems, pathsFromOpList: pathsFromOpList, findCircles: findCircles, findWeldTriangles: findWeldTriangles,
    normText: normText, groupTokens: groupTokens, classifyToken: classifyToken, isoFit: isoFit, GDT: GDT, CATS: CATS,
    DEFAULT_TITLE_KEYWORDS: DEFAULT_TITLE_KEYWORDS, autoExcludeBoxes: autoExcludeBoxes, isNoteLine: isNoteLine, detectCandidates: detectCandidates, makeCand: makeCand, candId: candId,
    zoneOf: zoneOf, orderCandidates: orderCandidates, placeBalloons: placeBalloons, leaderFor: leaderFor,
    detectBalloons: detectBalloons, leaderTarget: leaderTarget, linkBalloons: linkBalloons,
    catNorm: catNorm, numOrNull: numOrNull, parseTruth: parseTruth, scoreResults: scoreResults, sameValue: sameValue,
    buildBalloonPdf: buildBalloonPdf, aiPrompt: aiPrompt, parseAiAnswer: parseAiAnswer, nearText: nearText,
    ROW_HEAD: ROW_HEAD, itemRow: itemRow, toPred: toPred, DEFAULTS: DEFAULTS, cleanSettings: cleanSettings, emptyDb: emptyDb, restoreDb: restoreDb, docKey: docKey, SCHEMA_VERSION: SCHEMA_VERSION
  };
});

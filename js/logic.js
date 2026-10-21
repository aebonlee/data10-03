/* 검사 기준서 작성 자동화 — 순수 로직 (화면·저장소와 무관, node 에서도 불러 씀)

   좌표는 모두 「PDF 쪽 단위(pt, 1/72 인치) · 왼쪽 위 원점 · 아래로 갈수록 y 증가」입니다. mm = pt × 25.4 / 72.

   1. 도우미·기하            6. 제외 영역(테두리·표제란·개정란·주기)   11. 일치율(정답 대조)
   2. pdf.js 결과 정리        7. 검사 대상 후보                         12. 풍선 PDF 쓰기
   3. 원·용접 삼각형 찾기     8. 번호 순서(줄·열·구역)                  13. AI 프롬프트(애매한 후보)
   4. 글자 조각 → 토큰        9. 풍선 자리 잡기(겹침 피하기)            14. 엑셀 행 · 저장 형식
   5. 토큰 분류(치수·공차·기하공차·용접)   10. 풍선 읽기 · 지시선 연결
*/
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./glyph-ref.js'));
  else root.DRLogic = factory(root.DRGlyphRef);
})(typeof self !== 'undefined' ? self : this, function (GLYPH_REF) {
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
  /* getOperatorList() → 선분·부분 경로(곡선은 꺾은선으로 근사). OPS 는 pdfjsLib.OPS
     paints = 칠하기 · 긋기 한 번마다의 경로 묶음 { pid, kind: 'f'(칠함)|'s'(그음)|'b'(둘 다), fill, stroke(색 'r,g,b'), subs }.
     글자를 윤곽선으로 바꾼 PDF 는 글자 하나하나가 칠한 경로라, 4-1 에서 이 묶음으로 글자를 되읽습니다. */
  function colorKey(args) {
    if (args == null) return null;
    if (typeof args === 'string') { var h = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(args); return h ? [parseInt(h[1], 16), parseInt(h[2], 16), parseInt(h[3], 16)].join(',') : args; }
    var a = Array.prototype.slice.call(args);
    if (a.length === 1) { var g = Math.round(a[0] <= 1 ? a[0] * 255 : a[0]); return [g, g, g].join(','); }
    if (a.length === 4) { var k = a[3]; return [0, 1, 2].map(function (i) { return Math.round(255 * (1 - Math.min(1, a[i] + k))); }).join(','); }
    return a.slice(0, 3).map(function (v) { return Math.round(v <= 1 && a.every(function (q) { return q <= 1; }) ? v * 255 : v); }).join(',');
  }
  function pathsFromOpList(fnArray, argsArray, OPS, vt) {
    var ctm = [1, 0, 0, 1, 0, 0], stack = [], subpaths = [], cur = null, pending = [], paints = [];
    var fill = '0,0,0', stroke = '0,0,0', cstack = [];
    function dev(x, y) { return apply(mul(vt, ctm), x, y); }
    function start(p) { cur = { pts: [p], closed: false, curves: 0, pid: -1 }; subpaths.push(cur); pending.push(cur); }
    var FILLS = [OPS.fill, OPS.eoFill], STROKES = [OPS.stroke, OPS.closeStroke], BOTH = [OPS.fillStroke, OPS.eoFillStroke, OPS.closeFillStroke, OPS.closeEOFillStroke];
    for (var i = 0; i < fnArray.length; i++) {
      var fn = fnArray[i], args = argsArray[i];
      if (fn === OPS.save) { stack.push(ctm.slice()); cstack.push([fill, stroke]); }
      else if (fn === OPS.restore) { if (stack.length) ctm = stack.pop(); if (cstack.length) { var cs = cstack.pop(); fill = cs[0]; stroke = cs[1]; } }
      else if (fn === OPS.transform) ctm = mul(ctm, args);
      else if (fn === OPS.setFillRGBColor || fn === OPS.setFillGray || fn === OPS.setFillCMYKColor) fill = colorKey(args);
      else if (fn === OPS.setStrokeRGBColor || fn === OPS.setStrokeGray || fn === OPS.setStrokeCMYKColor) stroke = colorKey(args);
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
      } else if (FILLS.indexOf(fn) >= 0 || STROKES.indexOf(fn) >= 0 || BOTH.indexOf(fn) >= 0 || fn === OPS.endPath) {
        if (pending.length && fn !== OPS.endPath) {
          var kind = FILLS.indexOf(fn) >= 0 ? 'f' : STROKES.indexOf(fn) >= 0 ? 's' : 'b', pid = paints.length;
          pending.forEach(function (sp) { sp.pid = pid; sp.color = kind === 'f' ? fill : stroke; });
          paints.push({ pid: pid, kind: kind, fill: kind === 's' ? null : fill, stroke: kind === 'f' ? null : stroke, subs: pending });
        }
        pending = []; cur = null;
      }
    }
    var segs = [];
    subpaths.forEach(function (sp) {
      if (sp.curves) return; // 지시선·치수선은 곧은 선
      var p = sp.pts;
      for (var q = 1; q < p.length; q++) segs.push({ x1: p[q - 1][0], y1: p[q - 1][1], x2: p[q][0], y2: p[q][1], pid: sp.pid, color: sp.color });
      if (sp.closed && p.length > 2) segs.push({ x1: p[p.length - 1][0], y1: p[p.length - 1][1], x2: p[0][0], y2: p[0][1], pid: sp.pid, color: sp.color });
    });
    segs = segs.filter(function (s) { return dist(s.x1, s.y1, s.x2, s.y2) > 0.3; });
    return { subpaths: subpaths, segs: segs, paints: paints };
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
      if (r >= 2.5 && r <= 70 && sd / r < 0.08) out.push({ cx: cx, cy: cy, r: r, color: sp.color || null });
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

  /* 선으로 그린 용접 기호(ISO 2553): 기준선(실선) + 바로 옆 점선(식별선). 실선 하나에 짧은 평행 선분이 4개 이상 1~4pt 옆에 늘어서면 기준선으로 봅니다.
     기준선에 한 끝이 닿고 옆으로 솟은 선(필릿 삼각형 · 개선 기호 등)을 기호로 모읍니다. 반환 [{ x1,y1,x2,y2, a, L, dash(점선 쪽 +1/-1), syms:[{side, box}] }] */
  function findWeldRefs(segs) {
    segs = segs || [];
    var near = gridIndex(segs, segBox, 30), refs = [];
    segs.forEach(function (S) {
      var L = dist(S.x1, S.y1, S.x2, S.y2);
      if (L < 20 || L > 400) return;
      var a = Math.atan2(S.y2 - S.y1, S.x2 - S.x1);
      if (a > Math.PI / 2 + 0.01) a -= Math.PI; else if (a <= -Math.PI / 2 - 0.01) a += Math.PI; // 왼→오(세로는 아래→위)로 읽는 방향
      var p1 = local(S.x1, S.y1, a), p2 = local(S.x2, S.y2, a), u0 = Math.min(p1[0], p2[0]), u1 = Math.max(p1[0], p2[0]), v = p1[1];
      var dashes = { '1': 0, '-1': 0 };
      near(Math.min(S.x1, S.x2) - 5, Math.min(S.y1, S.y2) - 5, Math.max(S.x1, S.x2) + 5, Math.max(S.y1, S.y2) + 5).forEach(function (o) {
        if (o === S) return;
        var l = dist(o.x1, o.y1, o.x2, o.y2);
        if (l < 0.8 || l > 8 || l > 0.3 * L) return;
        if (Math.abs(Math.sin(Math.atan2(o.y2 - o.y1, o.x2 - o.x1) - a)) > 0.07) return;
        var m = local((o.x1 + o.x2) / 2, (o.y1 + o.y2) / 2, a), dv = m[1] - v;
        if (Math.abs(dv) >= 0.8 && Math.abs(dv) <= 4.5 && m[0] >= u0 - 2 && m[0] <= u1 + 2) dashes[dv > 0 ? '1' : '-1']++;
      });
      var side = dashes['1'] >= 4 ? 1 : dashes['-1'] >= 4 ? -1 : 0;
      if (!side) return;
      if (refs.some(function (r) { return Math.abs(r.a - a) < 0.05 && Math.abs(r.v - v) < 0.6 && Math.min(r.u1, u1) - Math.max(r.u0, u0) > 0.5 * L; })) return;
      refs.push({ x1: S.x1, y1: S.y1, x2: S.x2, y2: S.y2, a: a, L: L, u0: u0, u1: u1, v: v, dash: side });
    });
    // 기호: 기준선에 닿고(0.8pt 안) 옆으로 2~25pt 솟은 선. 양 끝 1.5pt 안(화살표 · 꼬리)은 뺌
    refs.forEach(function (r) {
      var touch = [];
      near(Math.min(r.x1, r.x2) - 25, Math.min(r.y1, r.y2) - 25, Math.max(r.x1, r.x2) + 25, Math.max(r.y1, r.y2) + 25).forEach(function (o) {
        var A = local(o.x1, o.y1, r.a), B = local(o.x2, o.y2, r.a), dA = A[1] - r.v, dB = B[1] - r.v;
        var on = Math.abs(dA) < 0.8 ? A : Math.abs(dB) < 0.8 ? B : null, off = on === A ? B : A, doff = off[1] - r.v;
        if (!on || Math.abs(doff) < 2 || Math.abs(doff) > 25) return;
        if (on[0] < r.u0 + 1.5 || on[0] > r.u1 - 1.5) return;
        if (Math.abs(doff) < 4.6 && Math.abs(off[0] - on[0]) < 0.5) return; // 점선과 실선 사이 짧은 막대
        touch.push({ u: (on[0] + off[0]) / 2, u0: Math.min(on[0], off[0]), u1: Math.max(on[0], off[0]), side: doff > 0 ? 1 : -1, ext: Math.abs(doff) });
      });
      var groups = [];
      touch.sort(function (p, q) { return p.u - q.u; }).forEach(function (t) {
        var g = groups.filter(function (x) { return x.side === t.side && t.u0 - x.u1 < 4; })[0];
        if (g) { g.u0 = Math.min(g.u0, t.u0); g.u1 = Math.max(g.u1, t.u1); g.ext = Math.max(g.ext, t.ext); g.n++; } else groups.push({ side: t.side, u0: t.u0, u1: t.u1, ext: t.ext, n: 1 });
      });
      // 기호는 선 둘 이상(삼각형 두 변 · V · 네모 …)이 닿고, 폭 30pt 안. 선 하나만 닿는 것은 도면 선이 지나가는 것
      r.syms = groups.filter(function (g) { return g.n >= 2 && g.u1 - g.u0 <= 30 && g.ext >= 3; }).map(function (g) {
        var v0 = g.side > 0 ? r.v : r.v - g.ext, v1 = g.side > 0 ? r.v + g.ext : r.v;
        var pts = [[g.u0, v0], [g.u1, v0], [g.u0, v1], [g.u1, v1]].map(function (q) { return fromLocal(q[0], q[1], r.a); });
        return { side: g.side, u0: g.u0, u1: g.u1, ext: g.ext, box: boxOfPts(pts) };
      });
    });
    return refs.filter(function (r) { return r.syms.length; });
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
    var parts = [], srcs = {};
    members.forEach(function (m) { parts = parts.concat(m.parts || [m.id]); srcs[m.src || 'text'] = 1; });
    return { id: 't' + page + '_' + Math.round((b.x0 + b.x1) / 2) + '_' + Math.round((b.y0 + b.y1) / 2), page: page, str: s.replace(/\s+/g, ' ').trim(),
      a: a, h: h, u0: u0, u1: u1, v: v, box: b, cx: (b.x0 + b.x1) / 2, cy: (b.y0 + b.y1) / 2, parts: parts, src: srcs.text ? 'text' : srcs.outline ? 'outline' : 'symbol' };
  }
  var CELL = /^(Ø?\d*\.?\d+[MLPⓂⓁⓅ]?|[A-Z](-[A-Z])?|CZ|[MLPⓂⓁⓅ]|Ø?\d*\.?\d+\s*CZ)$/;
  var SMALL_NUM = /^([+\-±]\s*\d*\.?\d+|0)$/;
  /* circles(선택) = findCircles 결과. 원 안에 홀로 든 1~3자리 숫자(풍선 · 품번 번호)는 옆 글자와 붙이지 않습니다 */
  function groupTokens(items, page, circles) {
    page = page || (items[0] && items[0].page) || 1;
    var L = items.map(function (it) {
      var p = local(it.x, it.y, it.a), st = normText(it.str), solo = false;
      if (circles && /^\d{1,3}$/.test(st)) {
        var b = it.box || itemBox(it.x, it.y, it.w, it.h, it.a), c0 = center(b);
        solo = circles.some(function (c) { return dist(c0[0], c0[1], c.cx, c.cy) < 0.6 * c.r && c.r >= 0.5 * it.h && c.r <= 3.6 * it.h; });
      }
      return { id: it.id, str: st, h: it.h, a: it.a, u0: p[0], u1: p[0] + it.w, v: p[1], src: it.src || 'text', solo: solo };
    }).filter(function (x) { return x.str; });
    // ① 같은 줄·같은 크기·가까운 조각 붙이기
    var parent = L.map(function (_, i) { return i; });
    function find(i) { while (parent[i] !== i) { parent[i] = parent[parent[i]]; i = parent[i]; } return i; }
    for (var i = 0; i < L.length; i++) for (var j = 0; j < L.length; j++) {
      if (i === j) continue;
      var A = L[i], B = L[j];
      if (A.solo || B.solo) continue;
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
    var toks = Object.keys(groups).map(function (k) { var t = makeToken(groups[k], page); t.solo = groups[k].length === 1 && groups[k][0].solo; return t; });
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
          if (o === t || o.dead || o.solo || cells.indexOf(o) >= 0 || !sameAngle(o.a, t.a)) continue;
          var g2 = o.u0 - lastC.u1;
          if (Math.abs(o.v - t.v) <= 0.35 * t.h && g2 > -0.5 * t.h && g2 < 3 * t.h && CELL.test(o.str.replace(/\s+/g, ''))) { cells.push(o); o.dead = true; grew = true; break; }
        }
      }
      if (cells.length === 1 && t.src === 'symbol') t.dead = true; // 칸 글자를 못 붙인 선 기호(빗금 · 구멍 표시일 수 있음)
      if (cells.length > 1) {
        var nt = makeToken(cells.map(function (c) { return { id: c.id, parts: c.parts, str: c.str, h: c.h, a: c.a, u0: c.u0, u1: c.u1, v: c.v, src: c.src }; }), page);
        Object.assign(t, nt);
      }
    });
    toks = toks.filter(function (t) { return !t.dead; });
    // ③ 위·아래로 쌓인 공차(+0.05 / -0.02)를 앞 치수에 붙임. 작은 글자면 하나도, 같은 크기면 위아래 둘이 쌓였을 때만(윤곽선 도면은 공차를 같은 크기로 씀)
    function attachTail(t, use) {
      use.forEach(function (o) { o.dead = true; });
      var tail = use.length === 2 ? use[0].str + '/' + use[1].str : use[0].str;
      var b = t.box; use.forEach(function (o) { b = union(b, o.box); });
      t.str = t.str + ' ' + tail.replace(/\s+/g, '');
      t.box = b; t.u1 = Math.max(t.u1, Math.max.apply(null, use.map(function (o) { return o.u1; })));
      t.parts = t.parts.concat.apply(t.parts, use.map(function (o) { return o.parts; }));
      t.stacked = use.length === 2;
      t.cx = (b.x0 + b.x1) / 2; t.cy = (b.y0 + b.y1) / 2;
    }
    toks.forEach(function (t) {
      if (t.dead || t.solo || !/\d$/.test(t.str) || SMALL_NUM.test(t.str)) return;
      var smalls = toks.filter(function (o) {
        if (o === t || o.dead || o.solo || !sameAngle(o.a, t.a) || !SMALL_NUM.test(o.str)) return false;
        if (o.h > 1.1 * t.h || o.h < 0.35 * t.h) return false;
        return o.u0 >= t.u1 - 0.3 * t.h && o.u0 <= t.u1 + (o.h > 0.85 * t.h ? 2 : 1.2) * t.h && o.v >= t.v - 1.6 * t.h && o.v <= t.v + 0.5 * t.h;
      }).sort(function (p, q) { return p.v - q.v; });
      var small = smalls.filter(function (o) { return o.h <= 0.85 * t.h && o.v >= t.v - 1.3 * t.h; });
      var same = smalls.filter(function (o) { return o.h > 0.85 * t.h; });
      var use = small.length ? small.slice(0, 2) : null;
      if (!use && same.length >= 2) { var a0 = same[0], b0 = same.filter(function (o) { return o !== a0 && o.v - a0.v > 0.6 * t.h && Math.abs(o.u0 - a0.u0) < 1.2 * t.h; })[0]; if (b0) use = [a0, b0]; }
      if (use) attachTail(t, use);
    });
    // ③-2 같은 줄 「45 +0.02」 위에 같은 크기 「+0.08」 → 45 +0.08/+0.02
    toks.forEach(function (t) {
      var m = /^(.*\d)\s+([+\-]\s*\d*\.?\d+|0)$/.exec(t.str);
      if (t.dead || !m || t.stacked) return;
      var tailU0 = t.u1 - (t.u1 - t.u0) * m[2].length / t.str.length;
      var up = toks.filter(function (o) { return o !== t && !o.dead && !o.solo && sameAngle(o.a, t.a) && SMALL_NUM.test(o.str) && o.h > 0.6 * t.h && o.h < 1.2 * t.h && Math.abs(o.u0 - tailU0) < 1.0 * t.h && o.v < t.v - 0.6 * t.h && o.v > t.v - 1.7 * t.h; })[0];
      if (!up) return;
      up.dead = true; t.str = m[1] + ' ' + up.str.replace(/\s+/g, '') + '/' + m[2].replace(/\s+/g, ''); t.box = union(t.box, up.box); t.parts = t.parts.concat(up.parts); t.stacked = true; t.cx = (t.box.x0 + t.box.x1) / 2; t.cy = (t.box.y0 + t.box.y1) / 2;
    });
    // ④ 따로 떨어진 앞뒤 조각 잇기: 「352」+「±1」, 「6x」+「M12」, 「PCD」+「Ø120」
    toks.forEach(function (t) {
      if (t.dead || t.solo) return;
      var nx = toks.filter(function (o) { return o !== t && !o.dead && !o.solo && sameAngle(o.a, t.a) && Math.abs(o.v - t.v) < 0.35 * t.h && o.u0 - t.u1 > -0.3 * t.h && o.u0 - t.u1 < 1.6 * t.h && Math.abs(o.h - t.h) < 0.25 * t.h; })
        .sort(function (p, q) { return p.u0 - q.u0; })[0];
      if (!nx) return;
      var ok = (/\d$/.test(t.str) && /^±/.test(nx.str)) || (/^\d+\s*[xX]$/.test(t.str) && /^(M|Ø|R|SØ|SR)\s*\d/.test(nx.str)) || (/^PCD$/i.test(t.str) && /^Ø/.test(nx.str));
      if (!ok) return;
      nx.dead = true; t.str = t.str + ' ' + nx.str; t.box = union(t.box, nx.box); t.u1 = Math.max(t.u1, nx.u1); t.parts = t.parts.concat(nx.parts); t.cx = (t.box.x0 + t.box.x1) / 2; t.cy = (t.box.y0 + t.box.y1) / 2;
    });
    return toks.filter(function (t) { return !t.dead; }).map(function (t) { delete t.dead; delete t.solo; return t; });
  }

  /* ── 4-1. 윤곽선 글자 읽기(글자 층 없는 PDF) ────────────────
     CAD 가 글자를 「칠한 윤곽선」 으로 바꿔 내보낸 PDF 는 getTextContent 가 비어 있습니다. 칠한 경로 묶음 하나가 글 한 덩어리이므로,
     ① 바깥 윤곽 + 그 안의 구멍을 글자 하나로 묶고 ② 16×16 격자로 칠해 참조 글자(js/glyph-ref.js)와 모양을 맞춰 보고
     ③ 읽는 방향(0° · ±90° · 180° · 글자 중심이 늘어선 비스듬한 방향)마다 점수를 더해 가장 그럴듯한 방향을 고른 뒤
     ④ 줄을 나누고 글자 사이 간격으로 띄어쓰기를 넣어 getTextContent 와 같은 모양의 글자 조각으로 돌려줍니다. 결과는 groupTokens 로 이어집니다. */
  var GG = 16;
  function polyArea(p) { var a = 0; for (var i = 0; i < p.length; i++) { var q = p[(i + 1) % p.length]; a += p[i][0] * q[1] - q[0] * p[i][1]; } return a / 2; }
  function pip(x, y, poly) { var c = false; for (var i = 0, j = poly.length - 1; i < poly.length; j = i++) { var xi = poly[i][0], yi = poly[i][1], xj = poly[j][0], yj = poly[j][1]; if ((yi > y) !== (yj > y) && x < (xj - xi) * (y - yi) / (yj - yi) + xi) c = !c; } return c; }
  /* 짝홀 규칙으로 칠한 덮임 비율(0~1)을 b 상자에 맞춰 늘린 GG×GG 격자에 */
  function rasterPolys(polys, b) {
    var S = 4, N = GG * S, out = new Float32Array(GG * GG), w = (b.x1 - b.x0) || 1e-6, h = (b.y1 - b.y0) || 1e-6, edges = [];
    polys.forEach(function (p) { for (var i = 0; i < p.length; i++) { var a = p[i], c = p[(i + 1) % p.length]; edges.push([(a[0] - b.x0) / w * N, (a[1] - b.y0) / h * N, (c[0] - b.x0) / w * N, (c[1] - b.y0) / h * N]); } });
    for (var r = 0; r < N; r++) {
      var y = r + 0.5, xs = [];
      for (var e = 0; e < edges.length; e++) { var E = edges[e]; if ((E[1] > y) !== (E[3] > y)) xs.push(E[0] + (y - E[1]) * (E[2] - E[0]) / (E[3] - E[1])); }
      xs.sort(function (p, q) { return p - q; });
      for (var k = 0; k + 1 < xs.length; k += 2) {
        var a0 = Math.max(0, xs[k]), c0 = Math.min(N, xs[k + 1]);
        for (var x = Math.ceil(a0 - 0.5); x + 0.5 <= c0 && x < N; x++) out[Math.floor(r / S) * GG + Math.floor(x / S)] += 1 / (S * S);
      }
    }
    return out;
  }
  function rotGrid(g, k) { // 화면에서 시계 방향 90° × k
    var a = g;
    for (var t = 0; t < ((k % 4) + 4) % 4; t++) { var o = new Float32Array(GG * GG); for (var r = 0; r < GG; r++) for (var c = 0; c < GG; c++) o[c * GG + (GG - 1 - r)] = a[r * GG + c]; a = o; }
    return a;
  }
  function gridDist(a, b) { var s = 0; for (var i = 0; i < a.length; i++) s += Math.abs(a[i] - b[i]); return s / a.length; }
  var GLYPH_T = null;
  function glyphTemplates() {
    if (GLYPH_T) return GLYPH_T;
    GLYPH_T = [];
    Object.keys(GLYPH_REF || {}).forEach(function (key) {
      var polys = GLYPH_REF[key].map(function (f) { var p = []; for (var i = 0; i + 1 < f.length; i += 2) p.push([f[i], -f[i + 1]]); return p; });
      // 바깥 윤곽이 둘 이상인 글자(± = : % i …)는 조각으로 나뉘어 들어오므로 쓰지 않음(±는 따로 합침)
      var outers = polys.filter(function (p) { return !polys.some(function (o) { return o !== p && Math.abs(polyArea(o)) > Math.abs(polyArea(p)) && pip(p[0][0], p[0][1], o); }); });
      if (outers.length !== 1) return;
      var b = boxOfPts([].concat.apply([], polys)); // 이름 「1#h」 「a#dv」 = 같은 글자의 다른 모양
      GLYPH_T.push({ ch: /^(.+)#[a-z]+$/.test(key) ? key.replace(/#[a-z]+$/, '') : key, g: rasterPolys(polys, b), ar: (b.x1 - b.x0) / Math.max(1, b.y1 - b.y0), h: (b.y1 - b.y0) / 100, vtop: b.y0 / 100, vbot: b.y1 / 100, holes: polys.length - 1 });
    });
    return GLYPH_T;
  }
  /* 칠한 경로 묶음 → 글자들(바깥 윤곽 + 구멍) */
  function assembleGlyphs(subs) {
    var S = subs.filter(function (sp) { return sp.pts.length >= 3; }).map(function (sp) { return { pts: sp.pts, b: boxOfPts(sp.pts), a: Math.abs(polyArea(sp.pts)) }; });
    S.forEach(function (s) {
      var best = null;
      S.forEach(function (o) { if (o === s || o.a <= s.a) return; if (o.b.x0 <= s.b.x0 && o.b.y0 <= s.b.y0 && o.b.x1 >= s.b.x1 && o.b.y1 >= s.b.y1 && pip(s.pts[0][0], s.pts[0][1], o.pts) && (!best || o.a < best.a)) best = o; });
      s.parent = best;
    });
    S.forEach(function (s) { var d = 0, q = s.parent; while (q) { d++; q = q.parent; } s.depth = d; });
    return S.filter(function (s) { return s.depth % 2 === 0; }).map(function (o) {
      var polys = [o.pts].concat(S.filter(function (s) { return s.parent === o; }).map(function (s) { return s.pts; }));
      return { polys: polys, b: o.b, holes: polys.length - 1 };
    });
  }
  var STRONG_CH = /[0-9A-HJ-Z]/, COMMON_CH = /[0-9.,+\-()°Øx±]/;
  function glyphUbox(g, a) {
    var pts = [[g.b.x0, g.b.y0], [g.b.x1, g.b.y0], [g.b.x0, g.b.y1], [g.b.x1, g.b.y1]].map(function (p) { return local(p[0], p[1], a); });
    return { u0: Math.min.apply(null, pts.map(function (p) { return p[0]; })), u1: Math.max.apply(null, pts.map(function (p) { return p[0]; })), v0: Math.min.apply(null, pts.map(function (p) { return p[1]; })), v1: Math.max.apply(null, pts.map(function (p) { return p[1]; })) };
  }
  function glyphGridAt(g, a, k) {
    if (k != null) { if (!g.r) g.r = rasterPolys(g.polys, g.b); return rotGrid(g.r, k); }
    var lp = g.polys.map(function (p) { return p.map(function (q) { return local(q[0], q[1], a); }); }); // 비스듬한 방향: 돌려서 다시 칠함
    return rasterPolys(lp, boxOfPts([].concat.apply([], lp)));
  }
  function shapeScore(x, t) { return gridDist(x.gr, t.g) + 0.25 * Math.abs(Math.log((x.ar + 0.02) / (t.ar + 0.02))) + 0.12 * Math.abs(x.g.holes - t.holes); }
  /* 글자 묶음 하나를 방향 a 로 읽기 → { total, lines:[{ m:[글자], cap, base, u0, u1, str }] } */
  function readGlyphsAt(glyphs, a, k, T, pageCap) {
    var gl = glyphs.map(function (g) { var ub = glyphUbox(g, a), w = ub.u1 - ub.u0, h = ub.v1 - ub.v0; return { g: g, ub: ub, gr: glyphGridAt(g, a, k), w: w, h: h, ar: w / Math.max(h, 1e-3) }; });
    gl.forEach(function (x) { var b1 = null; T.forEach(function (t) { var sc = shapeScore(x, t); if (!b1 || sc < b1.s) b1 = { t: t, s: sc }; }); x.s1 = b1; });
    // ± : 「+」 바로 아래에 같은 폭의 가로 막대
    gl.forEach(function (x) {
      if (x.dead || x.s1.t.ch !== '+') return;
      var bar = gl.filter(function (y) { return y !== x && !y.dead && y.ar > 2.2 && y.ub.v0 >= x.ub.v1 - 0.05 * x.h && y.ub.v0 - x.ub.v1 < 0.6 * x.h && Math.abs((y.ub.u0 + y.ub.u1) - (x.ub.u0 + x.ub.u1)) < 0.4 * x.w && y.w > 0.6 * x.w && y.w < 1.4 * x.w; })[0];
      if (bar) { bar.dead = true; x.pm = true; x.ub = { u0: Math.min(x.ub.u0, bar.ub.u0), u1: Math.max(x.ub.u1, bar.ub.u1), v0: x.ub.v0, v1: bar.ub.v1 }; x.h = x.ub.v1 - x.ub.v0; }
    });
    var lines = [];
    gl.filter(function (x) { return !x.dead; }).sort(function (p, q) { return (p.ub.v0 + p.ub.v1) - (q.ub.v0 + q.ub.v1); }).forEach(function (x) {
      var L = lines.filter(function (l) { var lh = l.v1 - l.v0, ov = Math.min(l.v1, x.ub.v1) - Math.max(l.v0, x.ub.v0); return ov > 0.3 * Math.min(lh, x.h) || (x.h < 0.4 * lh && x.ub.v0 >= l.v0 - 0.1 * lh && x.ub.v1 <= l.v1 + 0.3 * lh); })[0];
      if (L) { L.m.push(x); L.v0 = Math.min(L.v0, x.ub.v0); L.v1 = Math.max(L.v1, x.ub.v1); } else lines.push({ m: [x], v0: x.ub.v0, v1: x.ub.v1 });
    });
    var total = 0;
    lines.forEach(function (L) {
      var strong = L.m.filter(function (x) { return STRONG_CH.test(x.s1.t.ch) && x.s1.s < 0.2; });
      function med(arr) { arr.sort(function (p, q) { return p - q; }); return arr[Math.floor(arr.length / 2)]; }
      var cap = strong.length ? med(strong.map(function (x) { return x.h / x.s1.t.h; })) : null;
      var mx = Math.max.apply(null, L.m.map(function (x) { return x.h; }));
      if (!cap) cap = pageCap && mx < 1.7 * pageCap && mx > 0.45 * pageCap ? pageCap : mx;
      var base = strong.length ? med(strong.map(function (x) { return x.ub.v1 - x.s1.t.vbot * cap; })) : L.v1;
      L.m.forEach(function (x) {
        if (x.pm) { x.ch = '±'; x.s = x.s1.s; total += x.s; return; }
        var b2 = null;
        T.forEach(function (t) {
          var sc = shapeScore(x, t) + 0.5 * Math.abs(x.h / cap - t.h) + 0.5 * Math.abs((x.ub.v0 - base) / cap - t.vtop) + (COMMON_CH.test(t.ch) || /[A-Z]/.test(t.ch) ? 0 : 0.03);
          if (!b2 || sc < b2.s) b2 = { t: t, s: sc };
        });
        x.ch = b2.s < 0.3 ? b2.t.ch : '?'; x.s = b2.s; total += Math.min(b2.s, 0.6);
      });
      L.m.sort(function (p, q) { return p.ub.u0 - q.ub.u0; });
      // 숫자 옆 글자 바로잡기: 숫자 사이의 O → 0, l · I → 1, 숫자 사이 기준선의 작은 점은 . 또는 (기준선 아래로 내려가면) ,
      L.m.forEach(function (x, i) {
        var a = L.m[i - 1], b = L.m[i + 1], dig = function (y) { return y && /[0-9]/.test(y.ch); };
        if (dig(a) && dig(b) && x.h < 0.45 * cap && x.ub.v0 > base - 0.4 * cap) x.ch = x.ub.v1 > base + 0.08 * cap ? ',' : '.';
        if ((dig(a) || dig(b)) && x.ch === 'O') x.ch = '0';
        if ((dig(a) || dig(b)) && /[lI|]/.test(x.ch)) x.ch = '1';
      });
      // 띄어쓰기: 좁은 글자(1 · I · l)도 보통 폭(대문자 높이의 0.55)을 가진 것으로 보고, 가운데끼리 거리에서 뺀 틈이
      // 0.3(대문자 높이 기준)을 넘을 때. 자간을 넓힌 글(H O L E)은 그 줄의 보통 틈보다 크게 벌어진 곳만
      var gaps = L.m.map(function (x, i) {
        if (!i) return 0;
        var pv = L.m[i - 1], ea = Math.max(pv.ub.u1 - pv.ub.u0, 0.55 * cap), eb = Math.max(x.ub.u1 - x.ub.u0, 0.55 * cap);
        return ((x.ub.u0 + x.ub.u1) - (pv.ub.u0 + pv.ub.u1)) / 2 - (ea + eb) / 2;
      });
      var gs = gaps.slice(1).sort(function (p, q) { return p - q; }), med = gs.length ? gs[Math.floor(gs.length / 2)] : 0;
      var thr = Math.max(0.3 * cap, L.m.length >= 3 ? 1.8 * med + 0.15 * cap : 0);
      var str0 = '';
      L.m.forEach(function (x, i) {
        var pc = i ? L.m[i - 1].ch : '', digits = /[0-9]/.test(pc) && /[0-9]/.test(x.ch); // 숫자 사이는 자간이 넓어도 한 수(아주 멀 때만 띄움)
        if (i && !/[.,]/.test(pc + x.ch) && gaps[i] > (digits ? Math.max(thr, 1.2 * cap) : thr)) str0 += ' ';
        str0 += x.ch;
      });
      L.cap = cap; L.base = base; L.str = str0;
      L.u0 = Math.min.apply(null, L.m.map(function (x) { return x.ub.u0; })); L.u1 = Math.max.apply(null, L.m.map(function (x) { return x.ub.u1; }));
    });
    return { total: total, lines: lines };
  }
  function principalAngle(glyphs) {
    var sz = glyphs.map(function (g) { return Math.max(g.b.x1 - g.b.x0, g.b.y1 - g.b.y0); }), sorted = sz.slice().sort(function (a, b) { return a - b; }), mid = sorted[Math.floor(sorted.length / 2)] || 0;
    glyphs = glyphs.filter(function (g, i) { return sz[i] >= 0.6 * mid; }); // 쉼표 · 점 · ° 처럼 작은 글자는 줄 방향 계산에서 뺌
    if (glyphs.length < 2) return null;
    var cs = glyphs.map(function (g) { return center(g.b); }), mx = 0, my = 0;
    cs.forEach(function (c) { mx += c[0]; my += c[1]; }); mx /= cs.length; my /= cs.length;
    var sxx = 0, syy = 0, sxy = 0;
    cs.forEach(function (c) { sxx += (c[0] - mx) * (c[0] - mx); syy += (c[1] - my) * (c[1] - my); sxy += (c[0] - mx) * (c[1] - my); });
    var ang = 0.5 * Math.atan2(2 * sxy, sxx - syy), spread = Math.sqrt(Math.max(sxx, syy) / cs.length);
    return spread > 1 ? ang : null;
  }
  function readGlyphGroup(glyphs, T, pageCap) {
    var cands = [[0, 0], [-Math.PI / 2, 1], [Math.PI / 2, 3], [Math.PI, 2]];
    var pa = principalAngle(glyphs);
    if (pa != null) {
      var off = Math.abs(((pa % (Math.PI / 2)) + Math.PI / 2) % (Math.PI / 2));
      if (off > 0.09 && off < Math.PI / 2 - 0.09) [-0.1, -0.075, -0.05, -0.025, 0, 0.025, 0.05, 0.075, 0.1].forEach(function (d) { cands.push([pa + d, null], [pa + Math.PI + d, null]); }); // 비스듬한 글: 중심선 각도 ±6° 안에서 다듬기
    }
    var best = null;
    cands.forEach(function (c, ci) {
      var r = readGlyphsAt(glyphs, c[0], c[1], T, pageCap);
      r.total += ci ? 0.03 * glyphs.length : 0; // 같으면 똑바로 선 글자
      if (!best || r.total < best.total) { best = r; best.a = c[0]; best.k = c[1]; }
    });
    return best;
  }
  /* paints(pathsFromOpList) → { items: getTextContent 꼴 글자 조각, used: {pid:1}, pageCap } */
  function readOutlineText(paints, page, opt) {
    opt = opt || {};
    var T = glyphTemplates();
    if (!T.length) return { items: [], used: {}, pageCap: null };
    var groups = [];
    (paints || []).forEach(function (pt) {
      if (pt.kind !== 'f' || !pt.subs.length || pt.subs.length > 400) return; // 칠하기만 한 경로(풍선은 테두리를 함께 그어 'b')
      var bb = boxOfPts([].concat.apply([], pt.subs.map(function (sp) { return sp.pts; })));
      if (Math.min(bb.x1 - bb.x0, bb.y1 - bb.y0) > 60) return;
      var gl = assembleGlyphs(pt.subs).filter(function (g) { return Math.max(g.b.x1 - g.b.x0, g.b.y1 - g.b.y0) < 60 && Math.max(g.b.x1 - g.b.x0, g.b.y1 - g.b.y0) > 0.6; });
      if (gl.length && gl.length <= 120) groups.push({ pid: pt.pid, gl: gl });
    });
    var caps = [];
    groups.forEach(function (G) {
      G.first = readGlyphGroup(G.gl, T, null);
      G.first.lines.forEach(function (l) { l.m.forEach(function (x) { if (/[0-9]/.test(x.ch) && x.s < 0.15) caps.push(x.h / x.s1.t.h); }); });
    });
    caps.sort(function (a, b) { return a - b; });
    var pageCap = caps.length ? caps[Math.floor(caps.length / 2)] : null;
    var items = [], used = {}, n = 0;
    groups.forEach(function (G) {
      var rd = pageCap ? readGlyphGroup(G.gl, T, pageCap) : G.first, took = false;
      rd.lines.forEach(function (l) {
        var mean = l.m.reduce(function (s, x) { return s + x.s; }, 0) / l.m.length, txt = l.str.replace(/\s+/g, '');
        if (!txt || mean > (rd.k == null ? 0.27 : 0.2) || txt.replace(/\?/g, '').length < Math.max(1, txt.length / 2)) return; // 비스듬한 글은 다시 칠한 격자라 조금 너그럽게
        if (l.m.length === 1 && (l.m[0].s > 0.15 || /[Il|\-_'".,]/.test(l.m[0].ch) || (pageCap && l.cap > 2.5 * pageCap))) return;
        var p = fromLocal(l.u0, l.base, rd.a), hh = l.cap / 0.72;
        items.push({ id: 'o' + page + '_' + (n++), page: page, str: l.str, x: p[0], y: p[1], w: l.u1 - l.u0, h: hh, a: rd.a, box: itemBox(p[0], p[1], l.u1 - l.u0, hh, rd.a), src: 'outline', score: round(mean, 3), pid: G.pid, nGlyph: l.m.length, G: G });
        took = true;
      });
      if (took) used[G.pid] = 1;
    });
    // 뒤집어도 같은 모양(0 · 8 · 1 …)의 한두 글자는 방향을 잘못 잡기 쉬움 — 가까운 긴 글(3글자 이상)과 방향이 반대면 그 방향으로 다시 읽기
    var nearIt = gridIndex(items, function (it) { return it.box; }, 40);
    items.forEach(function (it, ii) {
      if (it.nGlyph > 2 || !it.G || it.G.gl.length > 2) return;
      var R = 4 * (pageCap || it.h);
      var nb = nearIt(it.box.x0 - R, it.box.y0 - R, it.box.x1 + R, it.box.y1 + R).filter(function (o) { return o !== it && o.nGlyph >= 2 && boxDist(o.box, it.box) < R; })
        .sort(function (p, q) { return boxDist(p.box, it.box) - boxDist(q.box, it.box); })[0];
      if (!nb) return;
      var d = Math.abs(((nb.a - it.a) % (2 * Math.PI) + 3 * Math.PI) % (2 * Math.PI) - Math.PI);
      if (Math.abs(d - Math.PI) > 0.1) return; // 반대(180°)가 아님
      var k = Math.abs(nb.a) < 0.01 ? 0 : Math.abs(nb.a + Math.PI / 2) < 0.01 ? 1 : Math.abs(nb.a - Math.PI / 2) < 0.01 ? 3 : Math.abs(Math.abs(nb.a) - Math.PI) < 0.01 ? 2 : null;
      var r2 = readGlyphsAt(it.G.gl, nb.a, k, T, pageCap), l2 = r2.lines.length === 1 ? r2.lines[0] : null;
      if (!l2) return;
      var m2 = l2.m.reduce(function (s0, x) { return s0 + x.s; }, 0) / l2.m.length;
      if (m2 > it.score + 0.06 || /\?/.test(l2.str)) return;
      var p2 = fromLocal(l2.u0, l2.base, nb.a), h2 = l2.cap / 0.72;
      items[ii] = Object.assign({}, it, { str: l2.str, x: p2[0], y: p2[1], w: l2.u1 - l2.u0, h: h2, a: nb.a, box: itemBox(p2[0], p2[1], l2.u1 - l2.u0, h2, nb.a), score: round(m2, 3), flipped: true });
    });
    items.forEach(function (it) { delete it.G; });
    return { items: items, used: used, pageCap: pageCap };
  }

  /* ── 4-2. 선으로 그린 기호(Ø · 기하공차 기호) ────────────────
     윤곽선 PDF 에서 지름 기호 Ø 는 「작은 원 + 가운데를 지나는 비스듬한 선」, 기하공차 틀의 기호는 선 몇 개로 그어져 있어 글자로 읽히지 않습니다.
     원 · 짧은 선의 배치로 Ø ◎ ⌖ ⊥ ∥ ▱ 를 알아보고, Ø 는 바로 오른쪽 글자 앞에 붙이고, 기하공차 기호는 틀 칸 안에 있을 때만 글자 조각으로 더합니다. */
  function strokeShapes(paints) {
    var circles = [], lines = [], polys4 = [], rings = [];
    (paints || []).forEach(function (pt) {
      if (pt.kind === 'f') return;
      pt.subs.forEach(function (sp) {
        var p = sp.pts, n = p.length;
        if (sp.curves >= 2 && n >= 10) {
          var bb = boxOfPts(p), cx = (bb.x0 + bb.x1) / 2, cy = (bb.y0 + bb.y1) / 2, w = bb.x1 - bb.x0, h = bb.y1 - bb.y0, r = (w + h) / 4; // 점이 고르지 않아 평균 대신 상자 가운데
          var sd = Math.sqrt(p.reduce(function (a, q) { var d = dist(q[0], q[1], cx, cy) - r; return a + d * d; }, 0) / n);
          if (r >= 0.8 && r <= 10 && w / h > 0.85 && w / h < 1.18 && sd / r < 0.08 && (sp.closed || dist(p[0][0], p[0][1], p[n - 1][0], p[n - 1][1]) < 0.3 * r)) circles.push({ cx: cx, cy: cy, r: r, pid: sp.pid });
          return;
        }
        if (sp.curves) return;
        var q = dedupePts(p);
        if (q.length >= 6 && dist(p[0][0], p[0][1], p[n - 1][0], p[n - 1][1]) < 0.3) { var rb = boxOfPts(q); if (rb.x1 - rb.x0 < 4 && rb.y1 - rb.y0 < 4 && rb.x1 - rb.x0 > 0.6) rings.push({ cx: (rb.x0 + rb.x1) / 2, cy: (rb.y0 + rb.y1) / 2, r: (rb.x1 - rb.x0 + rb.y1 - rb.y0) / 4 }); }
        var closedP = sp.closed || (p.length > 4 && dist(p[0][0], p[0][1], p[n - 1][0], p[n - 1][1]) < 0.3);
        if (q.length === 4 && closedP) polys4.push({ pts: q, pid: sp.pid, b: boxOfPts(q) });
        for (var k = 1; k < q.length + (sp.closed ? 1 : 0); k++) {
          var A = q[k - 1], B = q[k % q.length], L = dist(A[0], A[1], B[0], B[1]);
          if (L > 0.5 && L < 40) lines.push({ x1: A[0], y1: A[1], x2: B[0], y2: B[1], L: L, ang: Math.atan2(B[1] - A[1], B[0] - A[0]), mx: (A[0] + B[0]) / 2, my: (A[1] + B[1]) / 2, pid: sp.pid });
        }
      });
    });
    circles.filter(function (c) { return c.r < 2; }).forEach(function (c) { rings.push(c); });
    return { circles: circles, lines: lines, polys4: polys4, rings: rings };
  }
  function axisDeg(ang) { var d = Math.abs(ang * 180 / Math.PI) % 180; return d > 90 ? 180 - d : d; } // 0 = 가로, 90 = 세로
  /* 칸 나눔 색인(가까운 것만 빨리 찾기) — boxFn(item) → {x0,y0,x1,y1} */
  function gridIndex(list, boxFn, size) {
    var cells = {};
    list.forEach(function (it, i) {
      var b = boxFn(it), i0 = Math.floor(b.x0 / size), i1 = Math.floor(b.x1 / size), j0 = Math.floor(b.y0 / size), j1 = Math.floor(b.y1 / size);
      if ((i1 - i0 + 1) * (j1 - j0 + 1) > 400) { i1 = i0; j1 = j0; (cells.big = cells.big || []).push(i); return; }
      for (var x = i0; x <= i1; x++) for (var y = j0; y <= j1; y++) (cells[x + ',' + y] = cells[x + ',' + y] || []).push(i);
    });
    return function (x0, y0, x1, y1) {
      var seen = {}, out = [];
      for (var x = Math.floor(x0 / size); x <= Math.floor(x1 / size); x++) for (var y = Math.floor(y0 / size); y <= Math.floor(y1 / size); y++)
        (cells[x + ',' + y] || []).forEach(function (i) { if (!seen[i]) { seen[i] = 1; out.push(list[i]); } });
      (cells.big || []).forEach(function (i) { if (!seen[i]) { seen[i] = 1; out.push(list[i]); } });
      return out;
    };
  }
  function segBox(s) { return { x0: Math.min(s.x1, s.x2), y0: Math.min(s.y1, s.y2), x1: Math.max(s.x1, s.x2), y1: Math.max(s.y1, s.y2) }; }
  /* (cx,cy) 를 둘러싼 틀 칸: 왼쪽 · 오른쪽 세로선과 위 · 아래 가로선 */
  function cellAround(cx, cy, near, maxd) {
    var L = null, R = null, U = null, D = null;
    near(cx - maxd, cy - maxd, cx + maxd, cy + maxd).forEach(function (s) {
      var vert = Math.abs(s.x1 - s.x2) < 0.3, hor = Math.abs(s.y1 - s.y2) < 0.3;
      if (vert && Math.min(s.y1, s.y2) <= cy && Math.max(s.y1, s.y2) >= cy) { var dx = s.x1 - cx; if (dx < 0 && -dx < maxd && (L == null || dx > L)) L = dx; if (dx > 0 && dx < maxd && (R == null || dx < R)) R = dx; }
      if (hor && Math.min(s.x1, s.x2) <= cx && Math.max(s.x1, s.x2) >= cx) { var dy = s.y1 - cy; if (dy < 0 && -dy < maxd && (U == null || dy > U)) U = dy; if (dy > 0 && dy < maxd && (D == null || dy < D)) D = dy; }
    });
    return L != null && R != null && U != null && D != null ? { x0: cx + L, x1: cx + R, y0: cy + U, y1: cy + D } : null;
  }
  function strokeSymbols(paints, segs, items, pageCap) {
    var sh = strokeShapes(paints), out = [], cap = pageCap || 8;
    var lineUsed = {}, circUsed = {};
    sh.lines.forEach(function (l, i) { l.i = i; });
    sh.circles.forEach(function (c, i) { c.i = i; });
    var nearL = gridIndex(sh.lines, segBox, 20), nearC = gridIndex(sh.circles, function (c) { return { x0: c.cx - c.r, y0: c.cy - c.r, x1: c.cx + c.r, y1: c.cy + c.r }; }, 20);
    // Ø · ⌖ · ◎ : 원 + 가운데를 지나는 선 / 같은 중심의 원 둘
    sh.circles.forEach(function (c) {
      if (circUsed[c.i]) return;
      var around = nearC(c.cx - c.r, c.cy - c.r, c.cx + c.r, c.cy + c.r);
      if (around.some(function (o) { return o !== c && dist(o.cx, o.cy, c.cx, c.cy) < 0.15 * o.r && c.r < 0.8 * o.r && c.r > 0.3 * o.r; })) return; // 바깥 원이 따로 있음
      var inner = around.filter(function (o) { return o !== c && dist(o.cx, o.cy, c.cx, c.cy) < 0.15 * c.r && o.r < 0.8 * c.r && o.r > 0.3 * c.r; });
      var through = nearL(c.cx - 2 * c.r, c.cy - 2 * c.r, c.cx + 2 * c.r, c.cy + 2 * c.r).filter(function (l) { return dist(l.mx, l.my, c.cx, c.cy) < 0.4 * c.r && l.L > 1.6 * c.r && l.L < 4.6 * c.r; });
      var hv = through.filter(function (l) { var d = axisDeg(l.ang); return d < 8 || d > 82; });
      var kind = null;
      if (inner.length && !through.length) kind = '◎';
      else if (hv.length >= 2 && hv.some(function (l) { return axisDeg(l.ang) < 8; }) && hv.some(function (l) { return axisDeg(l.ang) > 82; })) kind = '⌖';
      else if (through.length === 1 && axisDeg(through[0].ang) > 40) kind = 'Ø';
      if (!kind) return;
      circUsed[c.i] = 1; inner.forEach(function (o) { circUsed[o.i] = 1; }); through.forEach(function (l) { lineUsed[l.i] = 1; });
      out.push({ kind: kind, cx: c.cx, cy: c.cy, r: c.r, box: { x0: c.cx - c.r, y0: c.cy - c.r * 1.3, x1: c.cx + c.r, y1: c.cy + c.r * 1.3 } });
    });
    // × : 같은 상자를 대각선으로 지나는 짧은 선 둘 · ° : 아주 작은 고리(원 또는 닫힌 꺾은선)
    sh.lines.forEach(function (l) {
      if (lineUsed[l.i] || l.L > 1.6 * cap || l.L < 0.25 * cap) return;
      var d = axisDeg(l.ang);
      if (d < 30 || d > 60) return;
      var m = nearL(l.mx - 1, l.my - 1, l.mx + 1, l.my + 1).filter(function (o) { return o !== l && !lineUsed[o.i] && dist(o.mx, o.my, l.mx, l.my) < 0.12 * l.L && Math.abs(o.L - l.L) < 0.2 * l.L && Math.abs(axisDeg(o.ang) - d) < 12 && Math.abs(Math.sin(o.ang - l.ang)) > 0.85; })[0];
      if (m) { lineUsed[l.i] = lineUsed[m.i] = 1; out.push({ kind: 'x', cx: l.mx, cy: l.my, r: l.L * 0.36, box: pad(boxOfPts([[l.x1, l.y1], [l.x2, l.y2], [m.x1, m.y1], [m.x2, m.y2]]), 0.2) }); }
    });
    sh.rings.forEach(function (c) { if (c.r < 0.25 * cap && c.r > 0.05 * cap) out.push({ kind: '°', cx: c.cx, cy: c.cy, r: c.r, box: { x0: c.cx - c.r, y0: c.cy - c.r, x1: c.cx + c.r, y1: c.cy + c.r } }); });
    // ∥ : 짧은 비스듬한 평행선 꼭 둘(빗금처럼 셋 이상 늘어서면 아님) · ⊥ : 세로선 아래 끝이 같은 길이쯤의 가로선 가운데
    sh.lines.forEach(function (l) {
      if (lineUsed[l.i] || l.L > 3 * cap) return;
      var d = axisDeg(l.ang), nb = nearL(l.mx - 1.5 * l.L, l.my - 1.5 * l.L, l.mx + 1.5 * l.L, l.my + 1.5 * l.L);
      if (d > 45 && d < 80) {
        var par = nb.filter(function (o) { return o !== l && Math.abs(axisDeg(o.ang) - d) < 4 && Math.abs(o.L - l.L) < 0.25 * l.L && Math.abs(o.my - l.my) < 0.25 * l.L && Math.abs(o.mx - l.mx) < 1.5 * l.L; });
        var mate = par.filter(function (o) { return !lineUsed[o.i] && Math.abs(o.mx - l.mx) > 0.15 * l.L && Math.abs(o.mx - l.mx) < 0.8 * l.L; });
        if (mate.length === 1 && par.length === 1) { var m = mate[0]; lineUsed[l.i] = lineUsed[m.i] = 1; out.push({ kind: '∥', cx: (l.mx + m.mx) / 2, cy: (l.my + m.my) / 2, r: l.L / 2, box: pad(boxOfPts([[l.x1, l.y1], [l.x2, l.y2], [m.x1, m.y1], [m.x2, m.y2]]), 0.5) }); }
      } else if (d > 85) {
        var bot = Math.max(l.y1, l.y2), x = l.mx;
        var base = nb.filter(function (o) { return o !== l && !lineUsed[o.i] && axisDeg(o.ang) < 5 && Math.abs(o.my - bot) < 0.4 && Math.abs(o.mx - x) < 0.15 * o.L && o.L > 0.6 * l.L && o.L < 1.6 * l.L; })[0];
        if (base) { lineUsed[l.i] = lineUsed[base.i] = 1; out.push({ kind: '⊥', cx: x, cy: bot - l.L / 2, r: l.L / 2, box: { x0: base.mx - base.L / 2, y0: bot - l.L, x1: base.mx + base.L / 2, y1: bot } }); }
      }
    });
    sh.polys4.forEach(function (p) {
      var e = [0, 1, 2, 3].map(function (k) { var A = p.pts[k], B = p.pts[(k + 1) % 4]; return axisDeg(Math.atan2(B[1] - A[1], B[0] - A[0])); });
      var w = p.b.x1 - p.b.x0, h = p.b.y1 - p.b.y0;
      if (((e[0] < 5 && e[2] < 5 && e[1] > 40 && e[1] < 80) || (e[1] < 5 && e[3] < 5 && e[0] > 40 && e[0] < 80)) && h > 1 && h < 2 * cap && w < 3 * h)
        out.push({ kind: '▱', cx: (p.b.x0 + p.b.x1) / 2, cy: (p.b.y0 + p.b.y1) / 2, r: h / 2, box: p.b });
    });
    // 글자 조각으로: Ø 는 오른쪽 글자 앞에 붙이고, 기하공차 기호는 틀 칸 안에 있을 때만
    var nearI = gridIndex(items || [], function (it) { return it.box; }, 40), nearS = gridIndex(segs || [], segBox, 40), add = [];
    out.forEach(function (sym, si) {
      var best = null;
      nearI(sym.cx - 50, sym.cy - 50, sym.cx + 50, sym.cy + 50).forEach(function (it) {
        var p = local(sym.cx, sym.cy, it.a), q = local(it.x, it.y, it.a), du = q[0] - p[0], dv = (q[1] - 0.36 * it.h) - p[1];
        if (du > -0.2 * it.h && du < (sym.kind === 'Ø' ? 1.3 : 2.6) * it.h && Math.abs(dv) < 0.45 * it.h && (!best || du < best.du)) best = { it: it, du: du };
      });
      if (sym.kind === '°' || sym.kind === 'x') {
        // 왼쪽 글자 끝 바로 뒤(° 는 윗부분), × 는 오른쪽 글자까지 이어 붙임
        var left = null, right = null;
        nearI(sym.cx - 40, sym.cy - 40, sym.cx + 40, sym.cy + 40).forEach(function (it) {
          if (it.dead) return;
          var p = local(sym.cx, sym.cy, it.a), q = local(it.x, it.y, it.a), gapL = p[0] - (q[0] + it.w), gapR = q[0] - p[0];
          var dv = sym.kind === '°' ? (q[1] - 0.72 * it.h) - p[1] : (q[1] - 0.36 * it.h) - p[1];
          if (Math.abs(dv) > (sym.kind === '°' ? 0.35 : 0.3) * it.h) return;
          if (gapL > -0.2 * it.h && gapL < (sym.kind === '°' ? 0.6 : 1.2) * it.h && (!left || gapL < left.g)) left = { it: it, g: gapL };
          if (gapR > 0 && gapR < 1.2 * it.h && (!right || gapR < right.g)) right = { it: it, g: gapR };
        });
        if (!left) return;
        var A = left.it;
        if (sym.kind === '°') { A.str += '°'; var pe = local(sym.box.x1, sym.cy, A.a)[0] - local(A.x, A.y, A.a)[0]; if (pe > A.w) { A.w = pe; A.box = itemBox(A.x, A.y, A.w, A.h, A.a); } return; }
        if (!right || right.it === A) return;
        var B = right.it, qa = local(A.x, A.y, A.a), qb2 = local(B.x, B.y, A.a);
        A.str = A.str + ' x ' + B.str; A.w = Math.max(A.w, qb2[0] + B.w - qa[0]); A.box = itemBox(A.x, A.y, A.w, A.h, A.a); B.dead = true;
        return;
      }
      if (sym.kind === 'Ø') {
        if (!best) return;
        var it = best.it, q0 = local(it.x, it.y, it.a), nu = Math.min(q0[0], local(sym.cx, sym.cy, it.a)[0] - sym.r), np = fromLocal(nu, q0[1], it.a);
        it.w += q0[0] - nu; it.x = np[0]; it.y = np[1]; it.str = 'Ø' + it.str; it.box = itemBox(it.x, it.y, it.w, it.h, it.a);
        return;
      }
      var cell = cellAround(sym.cx, sym.cy, nearS, 2.2 * cap);
      if (!cell || !best || (cell.x1 - cell.x0) > 4 * cap || (cell.y1 - cell.y0) > 3 * cap) return;
      var b0 = best.it, qb = local(b0.x, b0.y, b0.a), ps = fromLocal(local(sym.box.x0, sym.cy, b0.a)[0], qb[1], b0.a), ww = sym.box.x1 - sym.box.x0;
      add.push({ id: 'g' + (b0.page || 1) + '_' + si, page: b0.page || 1, str: sym.kind, x: ps[0], y: ps[1], w: ww, h: b0.h, a: b0.a, box: itemBox(ps[0], ps[1], ww, b0.h, b0.a), src: 'symbol' });
    });
    return { items: (items || []).filter(function (it) { return !it.dead; }).concat(add), symbols: out };
  }

  /* ── 5. 토큰 분류 ─────────────────────────────────────────── */
  function normText(s) {
    return str(s).replace(/[⌀øφΦ∅]/g, 'Ø').replace(/[−–—‐]/g, '-').replace(/[×✕]/g, 'x').replace(/º/g, '°')
      .replace(/＋/g, '+').replace(/\+\s*\/\s*-|\+-(?=\s*\d)/g, '±').replace(/(\d),(\d)/g, '$1.$2').replace(/\s+/g, ' ').trim()
      .replace(/(^|[\s(])(SØ|SR|Ø|M|R)\s+(?=\d)/g, '$1$2').replace(/([+\-±])\s+(?=\d)/g, '$1');
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
    if (/^0\d+$/.test(s)) { r.reason = '앞자리 0 — 표식 번호(02 · 08 …)'; return r; }
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
    if (/^PCD\s*(?=[ØR\d])/i.test(s)) { var pc = classifyToken(s.replace(/^PCD\s*/i, '')); if (pc.cat) { pc.note = (pc.note ? pc.note + ' · ' : '') + 'PCD'; return pc; } }
    if ((m = /^(\d+\s*\/\s*\d+|\d+(?:\.\d+)?)\s*-\s*(\d+)\s*(NPSF|NPTF|NPT|NPS|UNC|UNF|UNEF|BSPT|BSPP|BSP|PT|PF|PS)\b/i.exec(s))) { r.cat = '치수'; r.kind = '나사'; r.unit = ''; r.note = m[1].replace(/\s+/g, '') + '-' + m[2] + ' ' + m[3].toUpperCase(); r.spec = r.note; return r; }
    if (/^\d+\s*[xX×]$/.test(s)) { r.reason = '개수 표기'; return r; }
    if ((m = new RegExp('^(\\d+)\\s*[xX×]\\s*' + NUM + '(?:\\s*\\(\\s*(' + NUM.slice(1, -1) + ')\\s*°?\\s*\\))?$').exec(s))) {
      r.cat = '용접'; r.kind = '단속'; r.count = +m[1]; r.nominal = +m[2]; r.unit = 'mm'; r.spec = m[1] + 'x' + m[2] + (m[3] ? '(' + m[3] + '°)' : '');
      r.note = '개수 ' + m[1] + ' · 길이 ' + m[2] + (m[3] ? ' · ' + m[3] + '°' : ''); return r;
    }
    if (/^C\.?\s?W\.?$/i.test(s) || /^(FULL\s*PEN|CJP|PJP)$/i.test(s)) { r.cat = '용접'; r.kind = '연속 용접'; r.unit = ''; return r; }
    if ((m = new RegExp('^([az])\\s?' + NUM + '$').exec(s))) { r.cat = '용접'; r.kind = m[1] === 'a' ? '목두께' : '각장'; r.nominal = +m[2]; return r; }
    if ((m = /^(\d+)\s*-\s*(\d+)\s*\(\s*(\d+)\s*\)$/.exec(s))) { r.cat = '용접'; r.kind = '단속 용접'; r.note = '개수 ' + m[1] + ' · 길이 ' + m[2] + ' · 피치 ' + m[3]; r.unit = ''; return r; }
    if (/^[▽△◺⊿▷◁▼▲◸◹◿]+$/.test(s)) { r.cat = '용접'; r.kind = '용접 기호 글자'; r.unit = ''; r.check = true; r.note = '표면 기호일 수도 있음'; return r; }
    if ((m = new RegExp('^(Ra|Rz|Ry|Rmax|Rzmax|Ramax)\\s*' + NUM + '$', 'i').exec(s))) { r.cat = '기타'; r.kind = '표면 거칠기'; r.nominal = +m[2]; r.unit = 'μm'; return r; }
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
      // 쪽 짧은 변의 4% 이되 12mm 를 넘지 않게(A0 처럼 큰 도면에서 띠가 도면 안쪽 글자까지 먹지 않도록)
      var bm = Math.min((s.borderPct == null ? 4 : s.borderPct) / 100 * Math.min(w, h), toPt(s.borderMaxMm == null ? 12 : s.borderMaxMm));
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
  /* 도면이 아니라 부품표 · 보고서 쪽: 원(구멍 · 둥근 모서리)이 하나도 없고, 글 조각의 35% 이상이 낱말 */
  function isReportPage(pg) {
    var toks = pg.tokens || [];
    if ((pg.circles || []).length || (pg.refs || []).length || toks.length < 20) return false;
    var words = toks.filter(function (t) { return /[A-Za-z가-힣]{3,}/.test(t.str) && !classifyToken(t.str).cat; }).length;
    return words >= 0.35 * toks.length;
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
  var CATS = ['치수', '공차', '기하공차', '용접', '주기', '기타'];
  function candId(page, b) { return 'c' + page + '_' + Math.round((b.x0 + b.x1) / 2) + '_' + Math.round((b.y0 + b.y1) / 2); }
  function makeCand(page, t, c, extra) {
    var o = { id: candId(page, t.box), page: page, raw: t.str, box: t.box, cx: (t.box.x0 + t.box.x1) / 2, cy: (t.box.y0 + t.box.y1) / 2, h: t.h || 10,
      cat: c.cat, kind: c.kind, nominal: c.nominal, upper: c.upper, lower: c.lower, unit: c.unit, count: c.count, cls: c.cls, datum: c.datum, spec: c.spec || '',
      check: !!c.check, note: c.note || '', src: 'auto', tokens: t.id ? [t.id] : [] };
    return Object.assign(o, extra || {});
  }
  /* pg = { page, w, h, items, tokens, segs, subpaths, tris }, opt = { skip: {토큰id:1}, userBoxes: [], includeEtc } */
  function detectCandidates(pg, s, opt) {
    s = s || {}; opt = opt || {};
    var boxes = autoExcludeBoxes(pg, s).concat((opt.userBoxes || []).map(function (b) { return Object.assign({ reason: '사용자 제외 영역' }, b); }));
    if (s.skipReportPages !== false && isReportPage(pg)) { boxes.push({ x0: 0, y0: 0, x1: pg.w, y1: pg.h, reason: '표 · 보고서 쪽(원 · 곡선 없음, 낱말이 많음)', auto: true }); }
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
    // 선으로 그린 용접 기호(기준선 + 점선)가 있는 도면은 weldCallouts 가 맡고, 삼각형 규칙(칠한 삼각형 · 크기 글자 묶기)은 쓰지 않습니다
    (pg.refs && pg.refs.length ? [] : pg.tris || []).forEach(function (tri) {
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
    weldCallouts(pg, cands, boxes);
    cands.forEach(function (c) { c.std = stdOf(c); });
    return { cands: cands, excluded: excluded, boxes: boxes };
  }

  /* 기준서 구분: 용접 기준서(용접 · 용접 형상 · 용접 부기) / 치수 기준서(그 밖) */
  function stdOf(c) { return c && (c.cat === '용접' || c.cat === '용접 형상' || (c.cat === '주기' && /용접/.test(c.kind || ''))) ? 'weld' : 'dim'; }
  /* 선으로 그린 용접 기호(findWeldRefs) 둘레: 크기 글자가 없는 쪽 기호는 「용접 형상」, 기준선 위 숫자는 용접 크기, 아래 · 위 문장은 「용접 부기」 */
  function weldCallouts(pg, cands, boxes) {
    (pg.refs || []).forEach(function (r) {
      if (insideAny({ cx: (r.x1 + r.x2) / 2, cy: (r.y1 + r.y2) / 2 }, boxes)) return;
      function rel(c) { var p = local(c.cx, c.cy, r.a); return { u: p[0], dv: p[1] - r.v }; }
      var hh = 12;
      var near = cands.filter(function (c) { var q = rel(c); hh = c.h || hh; return q.u > r.u0 - 3 * c.h && q.u < r.u1 + 2 * c.h && Math.abs(q.dv) < 1.6 * c.h; });
      near.forEach(function (c) {
        if (c.cat === '치수' && c.kind === '길이' && c.upper == null && c.lower == null && !c.count) { c.cat = '용접'; c.kind = '용접 크기'; c.note = (c.note ? c.note + ' · ' : '') + '용접 기준선 위 숫자'; c.check = false; }
      });
      r.syms.forEach(function (sy) {
        var sized = near.some(function (c) { var q = rel(c); return c.cat === '용접' && q.dv * sy.side > 0 && q.u > r.u0 - 3 * c.h && q.u < sy.u1 + 1.5 * c.h; });
        if (sized) return;
        var bb = pad(sy.box, 2), id = candId(pg.page, bb);
        if (cands.some(function (c) { return c.id === id; })) return;
        cands.push({ id: id, page: pg.page, raw: '(용접 형상)', box: bb, cx: (bb.x0 + bb.x1) / 2, cy: (bb.y0 + bb.y1) / 2, h: hh, cat: '용접', kind: '용접 형상', nominal: null, upper: null, lower: null, unit: '', count: null, cls: '', datum: '',
          check: true, note: '크기 글자 없는 용접 기호 — 형상 그림으로 확인', src: 'auto', shape: true, crop: pad(sy.box, 14), tokens: [] });
      });
      // 부기: 기준선 아래(또는 위) 3줄 안의 문장. 같은 쪽 줄들을 한 항목으로
      var lines = (pg.tokens || []).filter(function (t) {
        if (!/[A-Za-z가-힣]{2,}/.test(t.str) || classifyToken(t.str).cat || !sameAngle(t.a, r.a)) return false;
        var p = local(t.cx, t.cy, r.a), dv = p[1] - r.v;
        return Math.abs(dv) > 0.3 * t.h && Math.abs(dv) < 6 * t.h && t.box && local(t.box.x0, t.box.y0, r.a)[0] > r.u0 - 4 * t.h && p[0] < r.u1 + 2 * t.h;
      });
      [1, -1].forEach(function (side) {
        var L2 = lines.filter(function (t) { return (local(t.cx, t.cy, r.a)[1] - r.v) * side > 0; }).sort(function (p, q) { return local(p.cx, p.cy, r.a)[1] * side - local(q.cx, q.cy, r.a)[1] * side; });
        // 기준선에서 첫 줄이 1.8줄 안, 그 뒤 줄은 앞 줄과 1.8줄 안일 때만 이어 붙임
        var keep = [], lastV = r.v;
        L2.forEach(function (t) { var vv = local(t.cx, t.cy, r.a)[1]; if (Math.abs(vv - lastV) <= (keep.length ? 1.8 : 2.4) * t.h) { keep.push(t); lastV = vv; } });
        L2 = keep;
        if (!L2.length) return;
        var b = null; L2.forEach(function (t) { b = union(b, t.box); });
        var raw = L2.map(function (t) { return t.str; }).join(' '), id = candId(pg.page, b);
        if (insideAny({ cx: (b.x0 + b.x1) / 2, cy: (b.y0 + b.y1) / 2 }, boxes) || cands.some(function (c) { return c.id === id; })) return;
        cands.push({ id: id, page: pg.page, raw: raw, box: b, cx: (b.x0 + b.x1) / 2, cy: (b.y0 + b.y1) / 2, h: L2[0].h, cat: '주기', kind: '용접 부기', nominal: null, upper: null, lower: null, unit: '', count: null, cls: '', datum: '',
          check: false, note: '용접 기호에 딸린 문장', src: 'auto', spec: normNote(raw), tokens: L2.map(function (t) { return t.id; }) });
      });
    });
  }
  function normNote(s) { return str(s).toUpperCase().replace(/[^A-Z0-9가-힣#]+/g, ' ').trim(); }

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
      if (best) { used[best.ci] = 1; out.push({ no: +t.str, cx: best.c.cx, cy: best.c.cy, r: best.c.r, color: best.c.color || null, tokenId: t.id, page: pg.page, src: 'circle' }); }
    });
    // 글자 층 풍선(사람 · 풍선 도구가 단 것)과 윤곽선 글자 풍선(도면에 원래 그려진 품번 풍선)이 섞이면 글자 층 풍선만
    var srcOf = {}; ints.forEach(function (t) { srcOf[t.id] = t.src || 'text'; });
    if (out.some(function (b) { return srcOf[b.tokenId] === 'text'; }) && out.some(function (b) { return srcOf[b.tokenId] === 'outline'; }))
      out = out.filter(function (b) { return srcOf[b.tokenId] === 'text'; });
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
        if (Math.abs(dn - b.r) <= tolR && df > b.r + Math.max(2, 0.15 * b.r)) {
          // 지시선은 원 둘레에서 바깥으로 곧게(반지름 방향) 뻗음. 풍선 아래를 지나는 도면 선을 지시선으로 잡지 않게
          var L = dist(e[0], e[1], e[2], e[3]), rx = (e[0] - b.cx) / dn, ry = (e[1] - b.cy) / dn, cosr = ((e[2] - e[0]) * rx + (e[3] - e[1]) * ry) / (L || 1);
          if (cosr < 0.75) return;
          // 색이 있는 선(검정이 아닌 지시선)을 먼저, 같은 색이면 긴 것
          var pref = (s.color && s.color !== '0,0,0' ? 2 : 0) + (s.color && b.color && s.color === b.color ? 1 : 0);
          if (!best || pref > best.pref || (pref === best.pref && L > best.L)) best = { i: i, x: e[2], y: e[3], L: L, sx: e[0], sy: e[1], pref: pref };
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
    if (/주기|부기|note/i.test(s)) return '주기';
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
      if (raw) {
        var cr = classifyToken(raw);
        if (!it.kind && cr.cat) it.kind = cr.kind;
        if (cr.datum) it.datum = cr.datum;
        if (cr.spec) it.spec = cr.spec;
        if (it.count == null && cr.count != null) it.count = cr.count;
        if (!it.cat && !cr.cat && /[A-Za-z가-힣]{2,}\s+[A-Za-z가-힣#]/.test(raw)) it.cat = '주기';
      }
      if (it.cat === '주기') it.spec = normNote(raw);
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
  /* 문장 비교: 글자와 숫자만 남겨(띄어쓰기 · 기호 무시) 편집 거리로 닮은 정도 0~1 */
  function noteSim(a, b) {
    var A = normNote(a).replace(/[^A-Z0-9가-힣]/g, ''), B = normNote(b).replace(/[^A-Z0-9가-힣]/g, '');
    if (!A.length || !B.length) return 0;
    var prev = [], cur2 = [], i, j;
    for (j = 0; j <= B.length; j++) prev[j] = j;
    for (i = 1; i <= A.length; i++) {
      cur2 = [i];
      for (j = 1; j <= B.length; j++) cur2[j] = Math.min(prev[j] + 1, cur2[j - 1] + 1, prev[j - 1] + (A[i - 1] === B[j - 1] ? 0 : 1));
      prev = cur2;
    }
    return 1 - prev[B.length] / Math.max(A.length, B.length);
  }
  function sameValue(p, t) {
    var cp = cmpCat(p), ct = cmpCat(t);
    if (cp !== ct) return false;
    if (ct === '용접 형상') return true;
    if (ct === '주기') return noteSim(p.spec || p.raw, t.spec || t.raw) >= 0.8;
    if (p.spec && t.spec && /단속/.test((p.kind || '') + (t.kind || ''))) return p.spec === t.spec;
    return eqNum(p.nominal, t.nominal) && eqNum(p.upper, t.upper, true) && eqNum(p.lower, t.lower, true);
  }
  function itemText(it) {
    var s = it.raw || '';
    if (!s) s = (it.nominal != null ? it.nominal : '') + (it.upper != null || it.lower != null ? ' ' + (it.upper != null ? it.upper : '') + '/' + (it.lower != null ? it.lower : '') : '');
    return cmpCat(it) + ' ' + s;
  }
  /* 검사 기준서의 세 필드 — 항목명(종류) · 기준값 · 공차.
     기하공차: 기준값 = 공차값, 공차 = 데이텀 / 용접 단속: 기준값 = 「개수x길이」 / 용접 부기: 기준값 = 문장 / 용접 형상: 기준값 · 공차 없음 */
  function inspFields(it) {
    it = it || {};
    var c = cmpCat(it), kind = str(it.kind);
    if (!kind || /^(치수|공차)$/.test(kind)) { var k2 = it.raw ? classifyToken(it.raw) : null; kind = (k2 && k2.cat && k2.kind) || kind || c; }
    if (c === '용접 형상') return { name: '용접 형상', base: '', tol: '' };
    if (c === '주기') return { name: '부기', base: normNote(it.spec || it.raw), tol: '' };
    if (c === '기하공차') return { name: kind, base: it.upper, tol: str(it.datum).replace(/\s+/g, ' ') };
    if (c === '용접') return { name: kind, base: /단속/.test(kind) && it.spec ? it.spec : it.nominal, tol: '' };
    return { name: kind, base: it.nominal != null ? it.nominal : (it.spec || ''), tol: [it.upper == null ? 0 : it.upper, it.lower == null ? 0 : it.lower].join('/') };
  }
  function sameField(a, b) {
    if (typeof a === 'number' || typeof b === 'number') return a != null && b != null && Math.abs(a - b) < 1e-6;
    if (a == null || b == null) return a == b; // eslint-disable-line eqeqeq
    var A = String(a), B = String(b);
    if (/[A-Z가-힣]{2,}\s/.test(A) || /[A-Z가-힣]{2,}\s/.test(B)) return noteSim(A, B) >= 0.8; // 부기 문장
    return A === B;
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
    // 수강생 방식 일치율: 항목명 · 기준값 · 공차 필드별(비교 대상 = 정답 행, 못 찾은 행은 세 필드 모두 불일치) + 세 필드가 모두 맞은 행
    var fm = { name: 0, base: 0, tol: 0, full: 0 };
    pairs.forEach(function (pr) {
      var a = inspFields(pred[pr.pi]), b = inspFields(truth[pr.ti]);
      var n1 = a.name === b.name, n2 = sameField(a.base, b.base), n3 = sameField(a.tol, b.tol);
      if (n1) fm.name++; if (n2) fm.base++; if (n3) fm.tol++; if (n1 && n2 && n3) fm.full++;
      pr.fields = [n1, n2, n3];
    });
    var fieldRate = { name: round(nT ? fm.name / nT : 0, 4), base: round(nT ? fm.base / nT : 0, 4), tol: round(nT ? fm.tol / nT : 0, 4) };
    return { mode: mode, nPred: nP, nTruth: nT, tp: tp, fp: nP - tp, fn: nT - tp, precision: round(precision, 4), recall: round(recall, 4), f1: round(f1, 4),
      exact: exact, valueRate: round(nT ? exact / nT : 0, 4), numberRate: round(tp ? numSame / tp : 0, 4), byCat: byCat, misses: misses, wrong: wrong, extra: extra,
      fieldRate: fieldRate, fullRate: round(nT ? fm.full / nT : 0, 4), fullCount: fm.full };
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
  var ROW_HEAD = ['번호', '쪽', '구역', '구분', '종류', '원문', '공칭값', '상한공차', '하한공차', '최대 허용값', '최소 허용값', '단위', '수량', '데이텀', 'X(mm)', 'Y(mm)', '풍선 X(mm)', '풍선 Y(mm)', '연결 방식', '확인 필요', '메모', '형상 잘라낼 영역(mm, 왼쪽 위 X·Y · 오른쪽 아래 X·Y)', '기준서'];
  function limitVal(nom, dev) { return nom != null && dev != null ? round(nom + dev, 4) : null; }
  function itemRow(x) {
    var c = x.cand || {}, b = x.balloon || {}, crop = x.crop || c.crop;
    var cat = c.cat ? c.cat : '연결 못 함';
    return [x.no, x.page, x.zone || '', cat, c.kind || '', c.raw || '', c.nominal, c.upper, c.lower,
      c.cat === '기하공차' ? null : limitVal(c.nominal, c.upper), c.cat === '기하공차' ? null : limitVal(c.nominal, c.lower), c.unit || '', c.count, c.datum || '',
      c.cx != null ? mm(c.cx) : null, c.cy != null ? mm(c.cy) : null, b.x != null ? mm(b.x) : null, b.y != null ? mm(b.y) : null, x.method || '',
      (x.check || c.check) ? '확인 필요' : '', [c.note, x.note].filter(Boolean).join(' · '), crop ? [mm(crop.x0), mm(crop.y0), mm(crop.x1), mm(crop.y1)].join(', ') : '', c.cat ? (stdOf(c) === 'weld' ? '용접' : '치수') : ''];
  }
  function toPred(x) { var c = x.cand || {}; return { no: x.no, cat: c.cat === '용접' && c.kind === '용접 형상' ? '용접 형상' : c.cat, kind: c.kind, raw: c.raw, nominal: c.nominal, upper: c.upper, lower: c.lower, datum: c.datum, spec: c.spec, count: c.count, std: c.std || stdOf(c) }; }
  var DEFAULTS = {
    borderPct: 4, autoBorder: true, autoTitle: true, autoNotes: true, includeEtc: true, titleKeywords: DEFAULT_TITLE_KEYWORDS,
    order: 'rows', bands: 6, zoneCols: 8, zoneRows: 6, startNo: 1, balloonR: 4, linkMaxMm: 15, nearMaxMm: 14, isolatedBalloons: false, target: 85, std: 'all', skipReportPages: true, borderMaxMm: 12
  };
  function cleanSettings(o) {
    var s = clone(DEFAULTS);
    if (!o || typeof o !== 'object') return s;
    ['autoBorder', 'autoTitle', 'autoNotes', 'includeEtc', 'isolatedBalloons', 'skipReportPages'].forEach(function (k) { if (typeof o[k] === 'boolean') s[k] = o[k]; });
    [['borderPct', 0, 15], ['borderMaxMm', 2, 60], ['bands', 1, 30], ['zoneCols', 1, 26], ['zoneRows', 1, 20], ['startNo', 1, 9999], ['balloonR', 1.5, 12], ['linkMaxMm', 2, 60], ['nearMaxMm', 2, 60], ['target', 50, 100]]
      .forEach(function (k) { var v = Number(o[k[0]]); if (isFinite(v) && v >= k[1] && v <= k[2]) s[k[0]] = v; });
    if (['rows', 'cols', 'zone'].indexOf(o.order) >= 0) s.order = o.order;
    if (['all', 'dim', 'weld'].indexOf(o.std) >= 0) s.std = o.std;
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
    mul: mul, pdfTextItems: pdfTextItems, pathsFromOpList: pathsFromOpList, readOutlineText: readOutlineText, strokeSymbols: strokeSymbols, strokeShapes: strokeShapes, _assembleGlyphs: assembleGlyphs, _readGlyphGroup: readGlyphGroup, glyphTemplates: glyphTemplates, findCircles: findCircles, findWeldTriangles: findWeldTriangles,
    normText: normText, groupTokens: groupTokens, classifyToken: classifyToken, isoFit: isoFit, GDT: GDT, CATS: CATS,
    DEFAULT_TITLE_KEYWORDS: DEFAULT_TITLE_KEYWORDS, autoExcludeBoxes: autoExcludeBoxes, isNoteLine: isNoteLine, isReportPage: isReportPage, detectCandidates: detectCandidates, makeCand: makeCand, candId: candId,
    zoneOf: zoneOf, orderCandidates: orderCandidates, placeBalloons: placeBalloons, leaderFor: leaderFor,
    detectBalloons: detectBalloons, findWeldRefs: findWeldRefs, stdOf: stdOf, normNote: normNote, leaderTarget: leaderTarget, linkBalloons: linkBalloons,
    catNorm: catNorm, numOrNull: numOrNull, parseTruth: parseTruth, scoreResults: scoreResults, sameValue: sameValue, inspFields: inspFields,
    buildBalloonPdf: buildBalloonPdf, aiPrompt: aiPrompt, parseAiAnswer: parseAiAnswer, nearText: nearText,
    ROW_HEAD: ROW_HEAD, itemRow: itemRow, toPred: toPred, DEFAULTS: DEFAULTS, cleanSettings: cleanSettings, emptyDb: emptyDb, restoreDb: restoreDb, docKey: docKey, SCHEMA_VERSION: SCHEMA_VERSION
  };
});

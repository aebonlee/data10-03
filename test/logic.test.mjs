// 실행: node test/logic.test.mjs   (설치 없음 — vendor 의 pdf.js · SheetJS 만 씀)
// 실제 회사 도면은 리포에 넣지 않습니다. 가상 글자 조각과 scripts/make-samples.py 로 만든 가상 도면 4종으로 검사합니다.
// 도면 D(2차) = 실제 도면 2장에서 드러난 틀림 유형(글자 층 없는 윤곽선 글자 · 선으로 그린 Ø/×/°/기하공차 기호 · 같은 크기로 쌓인 공차 ·
//              떨어진 ±1 · ISO 용접 기준선 · a 크기 · 개수x길이 · 부기 · 양면 · 형상만 · 품번 풍선 · 빨간 두 토막 지시선)을 흉내 낸 가상 도면.
// 도면 E(3차) = 새 실제 도면 2종에서 드러난 유형(자간 넓힌 글자 층 · 일반공차 ISO 13920 · 한 줄 선 글꼴 · 꺾은선 Φ · ⏤ · 2x · 양쪽 각장 z · J 기호)을
//              지어낸 값으로 흉내 낸 가상 도면. 정답은 수강생 기준서 양식(시트 치수 · 각장).
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const L = require('../js/logic.js');
const F = require('../js/flow.js');
const E = require('../js/ai-endpoint.js');
const Sample = require('../js/sample-pdfs.js');
const XLSX = require('../vendor/xlsx.full.min.js');
const M = await import('../scripts/measure-samples.mjs');

let passed = 0;
const queue = [];
function test(name, fn) { queue.push([name, fn]); }
function section(name) { queue.push([null, name]); }
const S = L.cleanSettings();
const mm = v => v * 72 / 25.4;
// 가상 글자 조각: x,y(기준선, pt), 글자 크기 h, 각도 a
function it(str, x, y, h = 12, a = 0, w) { return { id: 'i' + Math.random().toString(36).slice(2, 8), page: 1, str, x, y, h, a, w: w ?? str.length * h * 0.55, box: null }; }

section('토큰 분류 — 치수·공차');
test('길이 · 지름 · 반지름 · 모따기 · 나사 · 각도 · 개수', () => {
  const c = s => L.classifyToken(s);
  assert.deepEqual([c('120').cat, c('120').kind, c('120').nominal], ['치수', '길이', 120]);
  assert.deepEqual([c('Ø20').kind, c('Ø20').nominal], ['지름', 20]);
  assert.equal(c('⌀20').kind, '지름'); // 다른 지름 기호도 같은 값으로
  assert.deepEqual([c('R5').kind, c('C1').kind, c('SR10').kind], ['반지름', '모따기', '구면 반지름']);
  const t = c('M8x1.25');
  assert.deepEqual([t.cat, t.kind, t.nominal, t.note], ['치수', '나사', 8, '피치 1.25']);
  assert.deepEqual([c('15°').kind, c('15°').unit], ['각도', '°']);
  assert.deepEqual([c('2-Ø8.5').count, c('2-Ø8.5').nominal, c('4xM6').count], [2, 8.5, 4]);
  assert.equal(c('1x45°').kind, '모따기');
});
test('± · 위아래 공차 · 한쪽 0 · 각도 공차 · 쉼표 소수', () => {
  const c = s => L.classifyToken(s);
  assert.deepEqual([c('45±0.1').cat, c('45±0.1').upper, c('45±0.1').lower], ['공차', 0.1, -0.1]);
  assert.deepEqual([c('45 +/-0.1').upper, c('45 +/-0.1').lower], [0.1, -0.1]);
  assert.deepEqual([c('25 +0.05/-0.02').upper, c('25 +0.05/-0.02').lower], [0.05, -0.02]);
  assert.deepEqual([c('20 0/-0.1').upper, c('20 0/-0.1').lower], [0, -0.1]);
  assert.deepEqual([c('40 +0.1/0').upper, c('40 +0.1/0').lower], [0.1, 0]);
  assert.deepEqual([c('30°±0.5°').unit, c('30°±0.5°').upper], ['°', 0.5]);
  assert.equal(c('12,5').nominal, 12.5);
});
test('끼워맞춤 기호 → ISO 286 환산(H7 · g6 · k6 · js) · 모르는 등급은 확인 필요', () => {
  assert.deepEqual(L.isoFit(12, 'H7'), { upper: 0.018, lower: 0 });
  assert.deepEqual(L.isoFit(30, 'g6'), { upper: -0.007, lower: -0.02 });
  assert.deepEqual(L.isoFit(25, 'k6'), { upper: 0.015, lower: 0.002 });
  assert.deepEqual(L.isoFit(50, 'js7'), { upper: 0.0125, lower: -0.0125 });
  assert.deepEqual(L.isoFit(40, 'h7'), { upper: 0, lower: -0.025 });
  assert.equal(L.isoFit(250, 'H7'), null); // 표 밖(180 초과)
  const h = L.classifyToken('Ø12 H7');
  assert.deepEqual([h.cat, h.upper, h.lower, h.cls], ['공차', 0.018, 0, 'H7']);
  const u = L.classifyToken('Ø12 U9');
  assert.equal(u.cat, '공차'); assert.equal(u.upper, null); assert.equal(u.check, true);
  const th = L.classifyToken('M12x1.5-6g');
  assert.deepEqual([th.cat, th.kind, th.upper, th.check], ['공차', '나사', null, false]);
});
test('검사 대상 아님 — 날짜 · 척도 · 매수 · 참고 치수 · 문장 · 품번 · 글자', () => {
  const c = s => L.classifyToken(s);
  for (const s of ['2026-10-20', '1:2', '1/1', '(100)', '(2:1)', '1. UNSPECIFIED CHAMFER C0.5', 'BRK-100', 'SS400', 'A', 'SECTION A-A', '±0.5', 'mm'])
    assert.equal(c(s).cat, null, s + ' → ' + c(s).cat);
  assert.ok(c('6 THRU').note.startsWith('THRU'));
});
test('기하공차 · 용접 글자 · 표면 거칠기', () => {
  const c = s => L.classifyToken(s);
  const g = c('⊥ 0.1 A');
  assert.deepEqual([g.cat, g.kind, g.upper, g.lower, g.datum], ['기하공차', '직각도', 0.1, 0, 'A']);
  const g2 = c('◎ Ø0.02 A-B');
  assert.deepEqual([g2.kind, g2.upper, g2.datum], ['동심도', 0.02, 'A-B']);
  assert.equal(c('⏥ 0.05').kind, '평면도');
  assert.deepEqual([c('C.W').cat, c('C.W').kind], ['용접', '연속 용접']);
  assert.deepEqual([c('a4').kind, c('a4').nominal], ['목두께', 4]);
  assert.equal(c('4-50(100)').kind, '단속 용접');
  assert.deepEqual([c('▽').cat, c('▽').check], ['용접', true]);
  assert.deepEqual([c('Ra 1.6').cat, c('Ra 1.6').nominal, c('Ra 1.6').unit], ['기타', 1.6, 'μm']);
});

section('글자 조각 → 토큰');
test('따로 쓴 「45」 「±0.1」 을 한 토큰으로, 다른 줄은 붙이지 않음', () => {
  const items = L.pdfTextItems([], [1, 0, 0, 1, 0, 0]);
  assert.equal(items.length, 0);
  const t = L.groupTokens([it('45', 100, 100, 12, 0, 13), it('±0.1', 115, 100, 12, 0, 24), it('120', 100, 130)], 1);
  const strs = t.map(x => x.str).sort();
  assert.deepEqual(strs, ['120', '45±0.1']);
});
test('위·아래로 쌓인 작은 공차를 앞 치수에 붙임', () => {
  const t = L.groupTokens([it('25', 100, 100, 12, 0, 13), it('+0.05', 114, 96, 7.5, 0, 20), it('-0.02', 114, 100.5, 7.5, 0, 20)], 1);
  assert.equal(t.length, 1);
  assert.equal(t[0].str, '25 +0.05/-0.02');
  assert.equal(t[0].stacked, true);
  assert.deepEqual([L.classifyToken(t[0].str).upper, L.classifyToken(t[0].str).lower], [0.05, -0.02]);
});
test('기하공차 틀 — 칸이 멀리 떨어진 기호·공차·데이텀을 한 토큰으로', () => {
  const t = L.groupTokens([it('⊥', 100, 100, 11, 0, 8), it('0.1', 122, 100, 11, 0, 15), it('A', 150, 100, 11, 0, 8), it('80', 300, 100)], 1);
  assert.deepEqual(t.map(x => x.str).sort(), ['80', '⊥ 0.1 A']);
});
test('90° 돌린 글자 — pdf.js 변환으로 각도·상자 계산, 같은 방향끼리만 붙임', () => {
  const vt = [1, 0, 0, -1, 0, 800]; // 배율 1 보기 변환(위아래 뒤집기)
  const raw = [{ str: '8', transform: [0, 12, -12, 0, 100, 400], width: 7 }, { str: '0', transform: [0, 12, -12, 0, 100, 407.5], width: 7 }];
  const items = L.pdfTextItems(raw, vt, 1);
  assert.ok(Math.abs(items[0].a + Math.PI / 2) < 1e-9, '진행 방향 = 위쪽');
  const b = items[0].box;
  assert.ok(b.x1 - b.x0 < 14 && b.y1 - b.y0 > 6, '세로로 긴 상자');
  const t = L.groupTokens(items, 1);
  assert.equal(t.length, 1); assert.equal(t[0].str, '80');
});

section('선 · 원 · 용접 삼각형');
const OPS = { save: 1, restore: 2, transform: 3, constructPath: 4, moveTo: 13, lineTo: 14, curveTo: 15, curveTo2: 16, curveTo3: 17, closePath: 18, rectangle: 19, paintFormXObjectBegin: 20, paintFormXObjectEnd: 21 };
function circleOps(cx, cy, r) {
  const k = 0.5523 * r;
  return [[OPS.moveTo, OPS.curveTo, OPS.curveTo, OPS.curveTo, OPS.curveTo, OPS.closePath],
    [cx + r, cy, cx + r, cy + k, cx + k, cy + r, cx, cy + r, cx - k, cy + r, cx - r, cy + k, cx - r, cy, cx - r, cy - k, cx - k, cy - r, cx, cy - r, cx + k, cy - r, cx + r, cy - k, cx + r, cy]];
}
test('곡선 4개로 된 원 → 원, 좌표 변환(CTM) 반영', () => {
  const vt = [1, 0, 0, -1, 0, 800];
  const p = L.pathsFromOpList([OPS.save, OPS.transform, OPS.constructPath, OPS.restore], [null, [2, 0, 0, 2, 10, 10], circleOps(50, 50, 6), null], OPS, vt);
  const cs = L.findCircles(p.subpaths);
  assert.equal(cs.length, 1);
  assert.ok(Math.abs(cs[0].cx - 110) < 0.5 && Math.abs(cs[0].cy - (800 - 110)) < 0.5 && Math.abs(cs[0].r - 12) < 0.5, JSON.stringify(cs[0]));
  assert.equal(p.segs.length, 0, '곡선은 지시선 후보에서 뺌');
});
test('직각 이등변 삼각형 = 용접 기호, 가는 화살촉 · 큰 삼각형은 아님', () => {
  const tri = pts => [[OPS.moveTo, OPS.lineTo, OPS.lineTo, OPS.closePath], pts.flat()];
  const p = L.pathsFromOpList([OPS.constructPath, OPS.constructPath, OPS.constructPath, OPS.constructPath],
    [tri([[0, 0], [11, 0], [0, 11]]), tri([[100, 0], [91.5, 1.3], [91.5, -1.3]]), tri([[200, 0], [300, 0], [200, 100]]), [[OPS.rectangle], [0, 50, 20, 10]]], OPS, [1, 0, 0, 1, 0, 0]);
  const t = L.findWeldTriangles(p.subpaths);
  assert.equal(t.length, 1);
  assert.ok(t[0].box.x1 - t[0].box.x0 < 12);
  assert.equal(p.segs.length, 3 + 3 + 3 + 4);
});

section('제외 영역 · 후보');
function page(tokensSpec, extra = {}) {
  const items = tokensSpec.map(s => it(...s));
  const tokens = L.groupTokens(items, 1);
  return Object.assign({ page: 1, w: 1190, h: 842, items, tokens, segs: [], circles: [], tris: [] }, extra);
}
test('테두리 구역 번호 · 표제란 숫자 · 주기 문장 속 숫자는 제외, 도면 안 치수는 후보', () => {
  const pg = page([
    ['1', 70, 20, 9], ['2', 210, 20, 9], ['120', 400, 300],
    ['TITLE', 800, 760, 8], ['BRACKET', 850, 760], ['SCALE', 800, 790, 8], ['1:2', 850, 790], ['DATE', 950, 790, 8], ['2026-10-20', 1000, 790], ['GENERAL TOLERANCE ±0.5', 800, 740, 8], ['50', 1000, 760],
    ['NOTES', 60, 700, 10], ['1.', 60, 716, 9, 0, 6], ['WELD', 75, 716, 9, 0, 25], ['SIZE', 110, 716, 9, 0, 20], ['6', 140, 716, 9, 0, 5], ['UNLESS', 155, 716, 9, 0, 35],
  ]);
  const r = L.detectCandidates(pg, S);
  assert.deepEqual(r.cands.map(c => c.raw), ['120']);
  const reasons = Object.fromEntries(r.excluded.map(x => [x.raw, x.reason]));
  assert.equal(reasons['1'], '테두리 구역');
  assert.equal(reasons['50'], '표제란');
  assert.ok(['주기란', '주기 문장'].includes(reasons['6']), reasons['6']);
  // 사용자 제외 영역
  const r2 = L.detectCandidates(pg, S, { userBoxes: [{ x0: 380, y0: 280, x1: 460, y1: 310 }] });
  assert.equal(r2.cands.length, 0);
  // 끄면 다시 후보
  const r3 = L.detectCandidates(pg, Object.assign({}, S, { autoBorder: false, autoTitle: false, autoNotes: false }));
  assert.ok(r3.cands.length >= 4);
});
test('용접 삼각형 옆 숫자·C.W 를 한 항목으로, 글자 없는 삼각형은 「용접 형상」', () => {
  const pg = page([['6', 300, 311, 11, 0, 6], ['C.W', 318, 311, 11, 0, 20], ['80', 600, 300]], {
    tris: [{ box: { x0: 307, y0: 300, x1: 318.3 - 3, y1: 311 }, cx: 311, cy: 305 }, { box: { x0: 700, y0: 500, x1: 711, y1: 511 }, cx: 705, cy: 505 }]
  });
  const r = L.detectCandidates(pg, S);
  const w = r.cands.filter(c => c.cat === '용접');
  assert.equal(w.length, 2);
  const a = w.find(c => !c.shape), b = w.find(c => c.shape);
  assert.deepEqual([a.raw, a.kind, a.nominal], ['6 C.W', '필릿', 6]);
  assert.equal(b.kind, '용접 형상'); assert.ok(b.crop && b.crop.x1 - b.crop.x0 > 40);
  assert.ok(r.cands.some(c => c.raw === '80'));
});

section('번호 순서 · 풍선 자리');
const pgO = { w: 1000, h: 600, tokens: [], segs: [] };
const C = (id, cx, cy) => ({ id, cx, cy, box: { x0: cx - 10, y0: cy - 6, x1: cx + 10, y1: cy + 6 } });
test('줄(위→아래, 띠 안 왼→오) · 열 · 구역 순서', () => {
  const cs = [C('a', 800, 50), C('b', 100, 80), C('c', 500, 400), C('d', 120, 420)];
  assert.deepEqual(L.orderCandidates(cs, pgO, { order: 'rows', bands: 3 }).map(c => c.id), ['b', 'a', 'd', 'c']);
  assert.deepEqual(L.orderCandidates(cs, pgO, { order: 'cols', bands: 4 }).map(c => c.id), ['b', 'd', 'c', 'a']);
  assert.deepEqual(L.orderCandidates(cs, pgO, { order: 'zone', zoneCols: 2, zoneRows: 2 }).map(c => c.id), ['b', 'a', 'd', 'c']);
  assert.equal(L.zoneOf(C('x', 990, 590), pgO, { zoneCols: 8, zoneRows: 6 }).label, 'F8');
});
test('풍선끼리 · 글자와 겹치지 않게, 쪽 밖으로 나가지 않게, 옮긴 자리는 그대로', () => {
  const cs = [];
  for (let i = 0; i < 6; i++) cs.push(Object.assign(C('k' + i, 300 + i * 26, 300), { no: i + 1 })); // 촘촘한 치수 6개
  const toks = cs.map(c => ({ box: c.box }));
  const pg = { w: 1000, h: 600, tokens: toks, segs: [] };
  const out = L.placeBalloons(cs, pg, { balloonR: 4 });
  const r = mm(4);
  for (let i = 0; i < out.length; i++) {
    for (let j = i + 1; j < out.length; j++) assert.ok(L.dist(out[i].x, out[i].y, out[j].x, out[j].y) >= 2 * r, '풍선 겹침 ' + i + ',' + j);
    toks.forEach(t => assert.ok(L.ptBoxDist(t.box, out[i].x, out[i].y) >= r - 0.01, '글자 겹침'));
    assert.equal(out[i].overlap, false);
  }
  const edge = L.placeBalloons([Object.assign(C('e', 5, 5), { no: 1 })], { w: 1000, h: 600, tokens: [], segs: [] }, { balloonR: 4 })[0];
  assert.ok(edge.x - r >= 0 && edge.y - r >= 0, '쪽 안');
  const fixed = L.placeBalloons([cs[0]], pg, { balloonR: 4 }, { k0: { x: 50, y: 60 } })[0];
  assert.deepEqual([fixed.x, fixed.y, fixed.manual], [50, 60, true]);
  assert.ok(fixed.leader, '멀리 옮기면 지시선');
});

section('풍선 읽기 · 지시선 연결');
test('원 안 정수 = 풍선, 원 밖 정수(구역 번호)는 아님', () => {
  const pg = page([['3', 200, 204, 10, 0, 6], ['7', 600, 30, 9, 0, 5], ['120', 300, 300]], { circles: [{ cx: 203, cy: 200, r: 11 }, { cx: 500, cy: 500, r: 17 }] });
  const b = L.detectBalloons(pg, S);
  assert.deepEqual(b.map(x => x.no), [3]);
  const iso = L.detectBalloons(pg, Object.assign({}, S, { isolatedBalloons: true }));
  assert.ok(iso.some(x => x.no === 7 && x.src === 'isolated') === false, '테두리 띠 안은 아님');
});
test('지시선 끝 글자 연결 · 꺾인 지시선 · 가까운 글자 · 서로 가로채지 않게(최대 짝짓기)', () => {
  const cands = [C('far', 600, 300), C('near1', 200, 200), C('near2', 240, 160)];
  const balloons = [
    { no: 1, cx: 400, cy: 400, r: 11, page: 1 },   // 지시선 → far
    { no: 2, cx: 238, cy: 185, r: 11, page: 1 },   // near2 가 가장 가깝고 near1 도 닿음 — 먼저 짝을 고르면 near2 를 가져감
    { no: 3, cx: 265, cy: 135, r: 11, page: 1 },   // near2 만 닿음
    { no: 4, cx: 900, cy: 550, r: 11, page: 1 },   // 아무것도 없음
  ];
  const segs = [{ x1: 411, y1: 400, x2: 520, y2: 400 }, { x1: 520, y1: 400, x2: 590, y2: 310 }];
  const res = L.linkBalloons(balloons, cands, { segs }, S);
  const by = Object.fromEntries(res.map(r => [r.no, r]));
  assert.equal(by[1].candId, 'far'); assert.equal(by[1].method, 'leader');
  assert.equal(by[3].candId, 'near2');
  assert.equal(by[2].candId, 'near1', '3번에게 near2 를 비켜 줌');
  assert.equal(by[4].candId, null); assert.ok(by[4].crop && by[4].check);
});

section('정답 대조 · 일치율');
const truthAoa = [['예시 정답'], ['번호', '구분', '원문', '공칭값', '상한공차', '하한공차'], [1, '치수', '120', 120, '', ''], [2, '공차', '45±0.1', '', '', ''], [3, '용접 형상', '', '', '', ''], [4, '기하공차', '⊥ 0.1 A', '', 0.1, 0], ['합계', '', '', '', '', '']];
test('머리행 찾기 · 원문만 있으면 원문을 읽어 값 채우기', () => {
  const t = L.parseTruth(truthAoa);
  assert.equal(t.headerRow, 1);
  assert.equal(t.items.length, 4);
  assert.deepEqual([t.items[1].nominal, t.items[1].upper, t.items[1].lower], [45, 0.1, -0.1]);
  assert.equal(t.items[2].cat, '용접 형상');
  assert.equal(L.parseTruth([['아무거나'], [1, 2]]).items.length, 0);
});
test('번호 기준(목표 2): 정밀도·재현율·F1·값 일치율·놓침·값 다름·더 찾음', () => {
  const truth = L.parseTruth(truthAoa).items;
  const pred = [
    { no: 1, cat: '치수', nominal: 120 }, { no: 2, cat: '공차', nominal: 45, upper: 0.2, lower: -0.2 },
    { no: 4, cat: '기하공차', upper: 0.1, lower: 0 }, { no: 9, cat: '치수', nominal: 5 }, { no: 3, cat: null },
  ];
  const s = L.scoreResults(pred, truth, 'number');
  assert.deepEqual([s.nTruth, s.nPred, s.tp, s.fp, s.fn, s.exact], [4, 4, 3, 1, 1, 2]);
  assert.equal(s.precision, 0.75); assert.equal(s.recall, 0.75); assert.equal(s.f1, 0.75); assert.equal(s.valueRate, 0.5);
  assert.deepEqual(s.misses.map(m => m.no), [3]);
  assert.deepEqual(s.wrong.map(m => m.no), [2]);
  assert.deepEqual(s.extra.map(m => m.no), [9]);
  assert.deepEqual(s.byCat['공차'], { truth: 1, found: 1, exact: 0 });
});
test('값 기준(목표 1): 번호가 달라도 값으로 짝, 번호 같은 비율 따로', () => {
  const truth = L.parseTruth(truthAoa).items;
  const pred = [{ no: 4, cat: '치수', nominal: 120 }, { no: 2, cat: '공차', nominal: 45, upper: 0.1, lower: -0.1 }, { no: 1, cat: '용접 형상' }, { no: 3, cat: '기하공차', upper: 0.1, lower: 0 }];
  const s = L.scoreResults(pred, truth, 'value');
  assert.equal(s.exact, 4); assert.equal(s.valueRate, 1); assert.equal(s.numberRate, 0.25);
});

section('풍선 PDF · AI 프롬프트 · 저장');
const lib = await M.loadPdfjs();
test('풍선 PDF: 글자 층(보이지 않음) + 풍선 + 원본 선 → 다시 읽으면 목표 2 가 목표 1 결과와 같음', async () => {
  for (const d of Sample.drawings) {
    const pages = await M.readPages(lib, Buffer.from(Sample.pdf[d.key + '_plain'], 'base64'));
    const g1 = F.goal1(pages, S, {});
    const pdfPages = pages.map(pg => ({ w: pg.w, h: pg.h, jpeg: null, texts: pg.items, segs: pg.segs, tris: pg.tris,
      balloons: g1.filter(x => x.page === pg.page).map(x => ({ x: x.balloon.x, y: x.balloon.y, r: x.balloon.r, no: x.no, leader: x.balloon.leader })) }));
    const bytes = L.buildBalloonPdf(pdfPages);
    const head = Buffer.from(bytes.slice(0, 8)).toString('latin1');
    assert.equal(head, '%PDF-1.4');
    const back = await M.readPages(lib, bytes);
    const allText = back[0].items.map(i => i.str).join(' ');
    assert.ok(allText.includes('BRACKET') || allText.includes('WELDED') || allText.includes('SHAFT'), '원본 글자 층');
    const g2 = F.goal2(back, S, {});
    assert.equal(g2.length, g1.length, d.name + ' 풍선 수');
    const sc = L.scoreResults(g2.map(L.toPred), g1.map(L.toPred), 'number');
    assert.equal(sc.valueRate, 1, d.name + ' 되읽기 ' + JSON.stringify(sc.wrong.concat(sc.misses)));
  }
});
test('AI 프롬프트에 후보 ID · 글자, 답 표 읽기(모르는 ID 무시)', () => {
  const p = L.aiPrompt([{ id: 'c1_10_20', raw: '6', cat: '치수', kind: '길이', near: ['WELD'] }], { name: '도면A', page: 1 });
  assert.ok(p.includes('c1_10_20') && p.includes('| ID |'));
  const a = L.parseAiAnswer('| ID | 검사대상 | 구분 | 공칭값 | 상한 | 하한 | 근거 |\n|---|\n| c1_10_20 | 아니오 | 치수 | 6 | | | 주기 문장 |\n| zz | 예 | 치수 | 1 | | | |', ['c1_10_20']);
  assert.equal(a.length, 1); assert.equal(a[0].include, false); assert.equal(a[0].nominal, 6);
});
test('저장 형식 되살리기 · 설정 범위 · AI 설정 저장 칸 이름', () => {
  const db = L.restoreDb({ settings: { balloonR: 99, order: 'zone', bands: 3 }, docs: { k: { name: 'a.pdf', excl: { 1: [{ x0: 1, y0: 1, x1: 2, y2: 2 }] } } }, cur: 'k', history: [{ at: '2026-10-20', f1: 1 }, 5] });
  assert.equal(db.settings.balloonR, 4); assert.equal(db.settings.order, 'zone'); assert.equal(db.settings.bands, 3);
  assert.equal(db.cur, 'k'); assert.equal(db.history.length, 1); assert.deepEqual(db.docs.k.removed, {});
  assert.equal(E.STORE_KEY, 'data10-03.ai');
  assert.equal(L.itemRow({ no: 1, page: 1, cand: { cat: '공차', nominal: 45, upper: 0.1, lower: -0.1, cx: 72, cy: 72 } })[9], 45.1);
});

section('2차 — 윤곽선 글자 · 선 기호 · 용접 기준선 · 기준서 구분 · 수강생 일치율');
const pagesD = await M.readPages(lib, Buffer.from(Sample.pdf['D_plain'], 'base64'));
const pagesDb = await M.readPages(lib, Buffer.from(Sample.pdf['D_ball'], 'base64'));
test('도면 D: 글자 층이 없어도 칠한 윤곽선을 글자로 되읽음(돌린 글 · 비스듬한 글 · 「1」 띄어쓰기)', () => {
  const pg = pagesD[0];
  assert.equal(pg.textItems, 0);
  assert.ok(pg.outlineItems > 40, '윤곽선 글자 ' + pg.outlineItems);
  const toks = pg.tokens.map(t => t.str);
  for (const want of ['1240', '27.5°', '1141 ±1', 'a3.5', '4x40']) assert.ok(toks.includes(want), want + ' 없음: ' + toks.join(' | '));
  const t39 = pg.tokens.find(t => t.str === '27.5°');
  assert.ok(Math.abs(t39.a * 180 / Math.PI + 25) < 4, '비스듬한 각도 ' + t39.a);
  assert.ok(Math.abs(Math.abs(pg.tokens.find(t => t.str === '1240').a) - Math.PI / 2) < 0.05, '세로 글');
});
test('선으로 그린 기호: Ø · × · ° 를 글자에 붙이고, 기하공차 기호(⌖ ▱ ∥ ⊥)는 틀 칸 글자와 묶음 · 빗금은 ∥ 아님', () => {
  const toks = pagesD[0].tokens.map(t => t.str);
  for (const want of ['2 x 45°', '⌖ Ø0.3', '▱ 1.5 CZ', '∥ 0.2 B', '⊥ Ø0.2 B']) assert.ok(toks.includes(want), want + ' 없음: ' + toks.join(' | '));
  assert.ok(!toks.includes('∥'), '칸 글자 없는 ∥(빗금 · 빈 네모 속 기호)가 홀로 남음');
  assert.deepEqual([L.classifyToken('▱ 1.5 CZ').kind, L.classifyToken('▱ 1.5 CZ').upper], ['평면도', 1.5]);
});
test('윤곽선 글자: 숫자 사이 기준선 아래로 내려간 작은 점은 쉼표(소수점) — 참조 글꼴에 없는 모양이어도', () => {
  const R = require('../js/glyph-ref.js'), k = 0.09; // 대문자 높이 9pt
  const glyph = (ch, ox) => R[ch].map(f => { const p = []; for (let i = 0; i + 1 < f.length; i += 2) p.push([ox + f[i] * k, 100 - f[i + 1] * k]); return p; });
  const subs = [...glyph('4', 10), [[16.2, 99.5], [17.4, 99.5], [17.0, 102.2], [16.0, 102.2]], ...glyph('5', 18.5)].map(pts => ({ pts, curves: 0, closed: true }));
  const r = L.readOutlineText([{ pid: 0, kind: 'f', fill: '0,0,0', subs }], 1);
  assert.deepEqual(r.items.map(i => i.str), ['4,5']);
  assert.equal(L.classifyToken('4,5').nominal, 4.5);
});
test('부품표 · 보고서 쪽은 통째로 빼고, 앞자리 0 표식 번호(02 · 08)는 치수 아님', () => {
  const words = Array.from({ length: 14 }, (_, i) => ({ str: 'PLATE WA ' + i, cx: 10, cy: 10 + i * 10, box: L.box(0, 0, 1, 1) }));
  const nums = Array.from({ length: 16 }, (_, i) => ({ id: 'n' + i, str: String(i + 10), cx: 300, cy: 10 + i * 10, h: 10, a: 0, box: L.box(290, 5 + i * 10, 310, 12 + i * 10) }));
  const pg = { page: 3, w: 842, h: 595, tokens: words.concat(nums), circles: [], refs: [] };
  assert.equal(L.isReportPage(pg), true);
  assert.equal(L.detectCandidates(pg, L.cleanSettings(), {}).cands.length, 0);
  assert.ok(L.detectCandidates(pg, L.cleanSettings({ skipReportPages: false }), {}).cands.length > 0);
  assert.equal(L.isReportPage(Object.assign({}, pg, { circles: [{ cx: 1, cy: 1, r: 3 }] })), false);
  assert.equal(L.isReportPage(pagesD[0]), false);
  assert.equal(L.classifyToken('02').cat, null); assert.equal(L.classifyToken('0.5').nominal, 0.5);
});
test('A0 처럼 큰 도면: 테두리 띠는 12mm 까지만(안쪽 용접 크기 글자를 테두리로 빼지 않음)', () => {
  const tok = (str, y) => ({ id: str, str, cx: 900, cy: y, h: 10, a: 0, v: y, box: L.box(890, y - 5, 910, y + 5) });
  const pg = { page: 1, w: 3370, h: 2384, tokens: [tok('a3.5', 85), tok('7', 15)], circles: [{ cx: 0, cy: 0, r: 1 }] };
  const c = L.detectCandidates(pg, L.cleanSettings(), {});
  assert.deepEqual(c.cands.map(x => x.raw), ['a3.5']);
  assert.deepEqual(c.excluded.map(x => x.raw), ['7']);
});
test('같은 크기로 쌓인 공차 · 떨어진 ±1 · 「6x」+「M12」 잇기', () => {
  const toks = pagesD[0].tokens.map(t => t.str);
  for (const want of ['Ø45 +0.08/+0.02', '58 +1/0', '352 ±1', '6x M12']) assert.ok(toks.includes(want), want + ' 없음: ' + toks.join(' | '));
  const c = L.classifyToken('Ø45 +0.08/+0.02');
  assert.deepEqual([c.cat, c.kind, c.nominal, c.upper, c.lower], ['공차', '지름', 45, 0.08, 0.02]);
});
test('ISO 용접 기준선(실선 + 점선): a 크기 · 개수x길이 · 부기 따로, 양면은 둘, 크기 없는 기호만 「용접 형상」', () => {
  const pg = pagesD[0];
  assert.equal(pg.refs.length, 3, '기준선 ' + pg.refs.length);
  const c = F.pageCandidates(pg, L.cleanSettings({ std: 'weld' }), {}, { balloons: L.detectBalloons(pg, S) }).cands.filter(x => L.stdOf(x) === 'weld');
  const by = k => c.filter(x => x.kind === k).map(x => x.raw);
  assert.deepEqual(by('목두께').sort(), ['a3.5', 'a6.3', 'a6.3']);
  assert.deepEqual(by('단속'), ['4x40']);
  assert.equal(by('용접 형상').length, 1, '형상 ' + JSON.stringify(by('용접 형상')));
  const note = c.find(x => x.kind === '용접 부기');
  assert.ok(note && L.normNote(note.raw).replace(/\s/g, '') === 'DONOTWELD#12AREA', '부기 ' + (note && note.raw));
  const sx = L.classifyToken('3x40(180°)');
  assert.deepEqual([sx.cat, sx.kind, sx.count, sx.nominal, sx.spec], ['용접', '단속', 3, 40, '3x40(180°)']);
  assert.equal(L.classifyToken('2x').cat, null);
});
test('기준서 구분: 치수 · 용접을 따로 1번부터, 풍선 읽기는 그 기준서 항목에만 연결', () => {
  const gd = F.goal1(pagesD, L.cleanSettings({ std: 'dim' }), {}), gw = F.goal1(pagesD, L.cleanSettings({ std: 'weld' }), {});
  assert.equal(gd.length, 13); assert.equal(gw.length, 6);
  assert.deepEqual(gd.map(x => x.no), Array.from({ length: 13 }, (_, i) => i + 1));
  assert.deepEqual(gw.map(x => x.no), [1, 2, 3, 4, 5, 6]);
  assert.ok(gd.every(x => x.std === 'dim') && gw.every(x => x.std === 'weld'));
  const all = F.goal1(pagesD, S, {});
  assert.equal(all.length, 19);
  const g2w = F.goal2(pagesDb, L.cleanSettings({ std: 'weld' }), {});
  assert.ok(g2w.filter(x => x.cand).every(x => x.std === 'weld'), '용접 기준서 연결에 치수 항목이 섞임');
});
test('풍선 읽기: 글자 층 검사 풍선만(윤곽선 품번 풍선 12 는 뺌), 빨간 두 토막 지시선을 도면 선보다 먼저', () => {
  const bl = L.detectBalloons(pagesDb[0], S);
  assert.equal(bl.length, 19);
  assert.equal(bl.filter(b => b.no === 12).length, 1);
  // 풍선 둘레를 스치는 긴 검정 선 + 가운데서 나뉜 빨간 지시선 → 빨간 선 끝
  const b = { cx: 100, cy: 100, r: 20 };
  const segs = [{ x1: 60, y1: 120, x2: 300, y2: 120, color: '0,0,0' }, { x1: 100, y1: 120, x2: 100, y2: 200, color: '0,0,0' },
    { x1: 130, y1: 70, x2: 114.1, y2: 85.9, color: '255,0,0' }, { x1: 130, y1: 70, x2: 145.9, y2: 54.1, color: '255,0,0' }];
  const lt = L.leaderTarget(b, segs);
  assert.ok(lt && Math.abs(lt.x - 145.9) < 0.5 && Math.abs(lt.y - 54.1) < 0.5, JSON.stringify(lt));
  // 검정 지시선만 있을 때: 원 둘레에 닿아 옆으로 길게 지나가는 선(접선)보다 짧아도 반지름 방향으로 뻗는 선
  const lt2 = L.leaderTarget(b, [{ x1: 100, y1: 80, x2: 260, y2: 80, color: '0,0,0' }, { x1: 114.1, y1: 114.1, x2: 128, y2: 128, color: '0,0,0' }]);
  assert.ok(lt2 && Math.abs(lt2.x - 128) < 0.5 && Math.abs(lt2.y - 128) < 0.5, JSON.stringify(lt2));
});
test('수강생 방식 일치율: 필드(항목명 · 기준값 · 공차) 일치율 · 완전 일치율 — 못 찾은 행은 세 필드 모두 불일치', () => {
  const truth = [{ no: 1, cat: '공차', kind: '길이', raw: '640 ±1', nominal: 640, upper: 1, lower: -1 }, { no: 2, cat: '공차', kind: '지름', raw: 'Ø45 +0.08/+0.02', nominal: 45, upper: 0.08, lower: 0.02 },
    { no: 3, cat: '기하공차', kind: '평면도', raw: '▱ 1.5 CZ', upper: 1.5, lower: 0 }, { no: 4, cat: '주기', kind: '용접 부기', raw: 'DO NOT WELD #12 AREA' },
    { no: 5, cat: '치수', kind: '지름', raw: 'Ø32', nominal: 32 }];
  const pred = [{ no: 1, cat: '공차', kind: '길이', raw: '640 ±1', nominal: 640, upper: 1, lower: -1 }, { no: 2, cat: '공차', kind: '지름', raw: 'Ø45 +0.02', nominal: 45, upper: 0.02, lower: 0 },
    { no: 4, cat: '주기', kind: '용접 부기', raw: 'DO NOTWELD #12AREA' }, { no: 5, cat: '치수', kind: '길이', raw: '32', nominal: 32 }];
  const sc = L.scoreResults(pred, truth, 'number');
  // 1 = 셋 다 맞음 · 2 = 공차 다름 · 3 = 못 찾음 · 4 = 부기 문장(띄어쓰기 · 기호 무시) 맞음 · 5 = 항목명(지름 ↔ 길이)만 다름
  assert.deepEqual(sc.fieldRate, { name: 0.6, base: 0.8, tol: 0.6 });
  assert.equal(sc.fullRate, 0.4); assert.equal(sc.fullCount, 2);
  assert.equal(sc.exact, 3);
});

section('3차 — 새 도면 2종에서 드러난 것(자간 · 일반공차 · 기준서 양식 · 2x · 한 줄 선 글꼴)');
const pagesE = await M.readPages(lib, Buffer.from(Sample.pdf['E_plain'], 'base64'));
const pagesEb = await M.readPages(lib, Buffer.from(Sample.pdf['E_ball'], 'base64'));
const pagesEw = await M.readPages(lib, Buffer.from(Sample.pdf['E_ballw'], 'base64'));
const truthE = M.sampleSheetTruth('E');
test('자간을 넓힌 글자 층(「2 6 2」 「2 7 , 5」 「± 1」)은 빈칸을 빼고 읽음 · 낱말 사이 빈칸은 그대로', () => {
  assert.equal(L.unspace('2 6 2'), '262'); assert.equal(L.unspace('2 7 , 5'), '27,5'); assert.equal(L.unspace('± 1'), '±1');
  assert.equal(L.unspace('SEE VIEW A'), 'SEE VIEW A'); assert.equal(L.unspace('2x M20'), '2x M20');
  const toks = pagesE[0].tokens.map(t => t.str);
  for (const want of ['262', '27.5 ±1']) assert.ok(toks.includes(want), want + ' 없음: ' + toks.join(' | '));
});
test('일반공차: 주기의 ISO 13920-BF 를 찾아 공차 없는 치수의 Tolerance 칸을 채움(2mm 미만 · 공차 적은 치수는 그대로)', () => {
  assert.equal(L.detectGenTol(pagesE), '13920-B');
  assert.equal(L.detectGenTol([{ items: [{ str: 'GENERAL TOLERANCES ISO 2768-mK' }] }]), '2768-m');
  assert.equal(L.detectGenTol([{ items: [{ str: 'GENERAL TOLERANCE ±0.5' }] }]), '');
  const g = (n, a) => L.genTolOf('13920-B', n, a);
  assert.deepEqual([g(16.5), g(47), g(262), g(640), g(1300), g(0), g(1.5), g(28, true)], [1, 2, 2, 3, 4, null, null, 0.75]);
  assert.equal(L.genTolOf('2768-m', 50), 0.3);
  const row = raw => { const c = L.classifyToken(raw); return L.sheetRow(Object.assign({ raw }, c), '13920-B'); };
  assert.deepEqual([row('262').max, row('262').min], [2, -2]);
  assert.deepEqual([row('284 ±1').max, row('284 ±1').min], [1, -1]);
  assert.deepEqual([row('28°').item, row('28°').max], ['Angle', 0.75]);
  assert.equal(L.sheetRow(Object.assign({}, L.classifyToken('262')), '').max, null); // 일반공차를 쓰지 않으면 빈칸
  assert.equal(L.cleanSettings({ genTol: 'none' }).genTol, 'none'); assert.equal(L.cleanSettings({ genTol: 'x' }).genTol, 'auto');
  assert.equal(L.genTolFor(pagesE, L.cleanSettings({ genTol: 'none' })), '');
});
test('기준서 양식 칸: 항목명 · 기준값 · MAX/MIN · 측정 기구를 수강생 기준서처럼', () => {
  const row = raw => { const c = L.classifyToken(raw); return L.sheetRow(Object.assign({ raw }, c), '13920-B'); };
  const pick = r => [r.item, r.std, r.max, r.min, r.eq];
  assert.deepEqual(pick(row('Ø71 +0.046/0')), ['Hole', '71', 0.046, 0, 'CMM']);
  assert.deepEqual(pick(row('47 F8')).slice(0, 2), ['Hole', '47']);
  assert.deepEqual(pick(row('M20')), ['TAP', 'M20x2.5', null, null, 'Plug Gage']);
  assert.deepEqual(pick(row('2x M20 x28')).slice(0, 2), ['TAP', 'M20x2.5']); // x28 = 깊이(피치 아님)
  assert.equal(row('2x M20 x28').count, 2);
  assert.deepEqual(pick(row('⊥ Ø0.6 B')), ['Perpendicularity', '0.6', 0, -0.6, 'CMM']);
  assert.deepEqual(pick(row('Rzmax3.2')), ['Roughness', '3.2', 0, -3.2, 'Roughness\nMeasurement']);
  assert.deepEqual(pick(row('a4.6')), ['Weld', 'a4.6\n(07x07)', null, null, 'Leg Length Gage']);
  assert.deepEqual(pick(row('a6.6')).slice(1, 2), ['a6.6\n(09x09)']);
  assert.deepEqual(pick(row('z6 z8')).slice(0, 2), ['Weld', 'z6 z8']);
  assert.deepEqual(pick(row('3x25(180°)')).slice(0, 2), ['Weld', '3x25']);
  assert.deepEqual(pick(row('1.5 x 45°')).slice(0, 2), ['Chamfer', '1.5x45°']);
  assert.deepEqual(pick(L.sheetRow({ cat: '용접 형상', kind: '용접 형상' })), ['Weld', '', null, null, 'Visual']);
  assert.deepEqual(pick(L.sheetRow({ cat: '주기', kind: '용접 부기', raw: 'NO PAINT  ZONE B' })).slice(0, 2), ['NO PAINT ZONE B', '']);
  assert.equal(L.normStd('a4,6\n(07x07)'), L.normStd('a4.6 (7x7)')); assert.equal(L.normStd('M24x3.0'), L.normStd('M24x3'));
});
test('정답 기준서 양식 읽기: 쪽마다 되풀이되는 머리행 · Tolerance 아래 MAX/MIN · 시트별 치수/용접, 기준서 칸으로 대조', () => {
  const head = [['No.', 'Inspection \nItem', 'Standard', 'Tolerance', '', 'Inspection\nEquipment'], ['', '', '', 'MAX', 'MIN', '']];
  const aoa = [['Standard Inspection Report For Dimension'], ...head, [1, 'Dimension', '47', 1, -1, 'CMM'], [2, 'Flatness', '1.5', 0, -1.5, 'CMM'], [], ['Standard Inspection Report For Dimension'], ...head, [3, 'Hole', '71', 0.046, 0, 'CMM'], [4, 'TAP', 'M20x2.5', '-', '-', 'Plug Gage']];
  const t = L.parseSheetTruth(aoa, '치수');
  assert.deepEqual(t.map(x => x.no), [1, 2, 3, 4]);
  assert.deepEqual(t.map(x => x.cat), ['공차', '기하공차', '공차', '치수']);
  const w = L.parseSheetTruth([...head, [1, 'Weld', 'a4.6\n(07x07)', '-', '-', 'Leg Length\nGage'], [2, 'Weld', '', '-', '-', 'Visual'], [3, 'ZONE B AREA', '', '-', '-', 'Visual'], [4, 'Weld', '3x25', '-', '-', 'Steel Ruler']], '각장');
  assert.deepEqual(w.map(x => [x.std, x.cat, x.kind]), [['weld', '용접', '목두께'], ['weld', '용접 형상', '용접 형상'], ['weld', '주기', '용접 부기'], ['weld', '용접', '단속']]);
  assert.equal(L.parseSheetTruth([['번호', '구분', '원문'], [1, '치수', '120']]), null); // 1 · 2차 정답 양식은 그대로 parseTruth 로
  // 치수 · 용접 정답이 섞이면(번호가 기준서마다 1번부터) 기준서 + 번호로 짝
  const pred = [{ no: 1, std: 'dim', cat: '공차', kind: '길이', raw: '47 ±1', nominal: 47, upper: 1, lower: -1 }, { no: 1, std: 'weld', cat: '용접', kind: '목두께', raw: 'a4.6', nominal: 4.6 },
    { no: 2, std: 'dim', cat: '기하공차', kind: '평면도', raw: '▱ 1.5', upper: 1.5, lower: 0 }, { no: 2, std: 'weld', cat: '용접 형상', kind: '용접 형상' }, { no: 3, std: 'weld', cat: '주기', kind: '용접 부기', raw: 'ZONE  B AREA' }];
  const sc = L.scoreResults(pred, t.concat(w), 'number');
  assert.equal(sc.sheetMode, true);
  assert.equal(sc.fullCount, 5); assert.deepEqual(sc.misses.map(m => m.no), [3, 4, 4]);
  assert.deepEqual(sc.fieldRate, { name: 0.625, base: 0.625, tol: 0.625 });
});
test('「2x」 = 같은 Spec 실측값 2개 — 옆 항목의 개수로 붙이고, 기준서 엑셀에 측정값 칸을 그 수만큼 비워 둠', () => {
  const tok = (id, str, x, y) => ({ id, str, cx: x + 10, cy: y, h: 10, a: 0, box: L.box(x, y - 8, x + 20, y + 2) });
  const pg = { page: 1, w: 842, h: 595, tokens: [tok('t1', '2x', 200, 300), tok('t2', 'Ø29.5', 224, 300), tok('t3', '3x', 400, 100), tok('t4', '83.5', 520, 100)], circles: [{ cx: 1, cy: 1, r: 3 }], refs: [] };
  const c = L.detectCandidates(pg, L.cleanSettings(), {}).cands;
  assert.equal(c.find(x => x.raw === 'Ø29.5').count, 2);
  assert.equal(c.find(x => x.raw === '83.5').count, null); // 멀리 떨어진 「3x」 는 붙이지 않음
  const items = [{ no: 1, page: 1, cand: { cat: '치수', kind: '지름', raw: 'Ø29.5', nominal: 29.5, count: 2 } }, { no: 2, page: 1, cand: { cat: '용접', kind: '목두께', raw: 'a4.6', nominal: 4.6 } }];
  const aoa = L.sheetAoa(items, '13920-B');
  assert.deepEqual(aoa[0].slice(0, 9), ['No.', 'Inspection Item', 'Standard', 'Tolerance MAX', 'Tolerance MIN', 'Inspection Equipment', "Q'ty(개수)", '측정값 1', '측정값 2']);
  assert.deepEqual(aoa[1].slice(0, 9), [1, 'Hole', '29.5', 1, -1, 'CMM', 2, '', '']);
  assert.deepEqual(aoa[2].slice(0, 9), [2, 'Weld', 'a4.6\n(07x07)', '-', '-', 'Leg Length Gage', 1, '', '-']);
  const goal2 = F.goal2(pagesEb, L.cleanSettings({ std: 'dim' }), {});
  assert.deepEqual(goal2.filter(x => x.cand && x.cand.count > 1).map(x => x.no), [4, 7]);
});
test('한 줄 선 글꼴: 선으로 그은 글자를 읽음(참조와 다른 꼴) · 글자 층 앞 선 글자 「M」 · 빗금 · 평행선 · 되풀이 X 는 글로 만들지 않음', () => {
  // 시험용 선 글꼴(도구 참조와도, 예시 도면 E 의 글꼴과도 다른 꼴): 대문자 높이 10 → 8pt
  const G = { '3': [[[0.4, 9], [1.6, 10], [4.6, 10], [5.8, 8.8], [5.8, 6.6], [4.4, 5.3], [2.4, 5.3]], [[4.4, 5.3], [6, 3.8], [6, 1.4], [4.6, 0], [1.4, 0], [0.2, 1.3]]], '7': [[[0, 10], [6, 10], [2.4, 0]]], '0': [[[1.2, 0], [4.8, 0], [6, 2.4], [6, 7.6], [4.8, 10], [1.2, 10], [0, 7.6], [0, 2.4], [1.2, 0]]] };
  const k = 0.8, paints = [];
  const draw = (s, x0, y0, rot) => { let ox = 0; for (const ch of s) { const subs = G[ch].map(pl => ({ pts: pl.map(([px, py]) => { const u = ox + px * k, v = -py * k; return rot ? [x0 + v, y0 - u] : [x0 + u, y0 + v]; }), curves: 0, closed: false })); paints.push({ pid: paints.length, kind: 's', stroke: '128,128,128', subs }); ox += (6 + 2.2) * k; } };
  draw('370', 100, 200, false); draw('73', 300, 400, true);
  for (let i = 0; i < 6; i++) paints.push({ pid: paints.length, kind: 's', subs: [{ pts: [[500 + i * 4, 100], [506 + i * 4, 90]], curves: 0 }] }); // 빗금
  for (let i = 0; i < 3; i++) paints.push({ pid: paints.length, kind: 's', subs: [{ pts: [[600 + i * 3.5, 100], [600 + i * 3.5, 92]], curves: 0 }] }); // 짧은 세로 평행선(「111」 아님)
  for (let i = 0; i < 3; i++) paints.push({ pid: paints.length, kind: 's', subs: [{ pts: [[700 + i * 8, 100], [706 + i * 8, 92]], curves: 0 }, { pts: [[700 + i * 8, 92], [706 + i * 8, 100]], curves: 0 }] }); // 빗금 교차(「XXX」 아님)
  const r = L.readStrokeText(paints, [], 8, 1);
  assert.deepEqual(r.items.map(i => i.str).sort(), ['370', '73']);
  const v = r.items.find(i => i.str === '73');
  assert.ok(Math.abs(Math.abs(v.a) - Math.PI / 2) < 0.01, '세로 선 글자 ' + v.a);
  // 글자 층 「20」 바로 앞의 선 글자 「M」(가운데가 중간까지만 내려오는 꼴)
  const text = [{ id: 'i1', page: 1, str: '20', x: 210, y: 300, w: 13, h: 12, a: 0, box: L.box(210, 290.4, 223, 302.6) }];
  const mP = [{ pid: 0, kind: 's', subs: [{ pts: [[200.6, 300], [200.6, 291.4], [204.2, 296.0], [207.8, 291.4], [207.8, 300]], curves: 0 }] }];
  assert.equal(L.readStrokeText(mP, text, 8.6, 1).items[0].str, 'M20');
  const toks = pagesE[0].tokens.map(t => t.str);
  for (const want of ['146', 'a4', '2x M20']) assert.ok(toks.includes(want), want + ' 없음: ' + toks.join(' | '));
  assert.ok(toks.some(t => L.normNote(t).replace(/\s/g, '') === 'DONOTGRIND'), '선 글꼴 부기');
});
test('선 기호: 꺾은선 원 Φ · ⏤ 진직도 — 기하공차 기호는 칸 글자도 틀 안이고 두 칸이 한 변을 함께 쓸 때만(구멍 중심 표시 옆 치수는 아님)', () => {
  const toks = pagesE[0].tokens.map(t => t.str);
  for (const want of ['Ø71', '⏤ 0.3']) assert.ok(toks.some(t => t.startsWith(want)), want + ' 없음: ' + toks.join(' | '));
  assert.deepEqual([L.classifyToken('⏤ 0.3').kind, L.classifyToken('⏤ 0.3').upper], ['진직도', 0.3]);
  // 원 + 십자(⌖ 모양 구멍 중심 표시) 바로 옆의 세로 치수 「146」 — 틀 칸이 없으니 기호가 아님
  const c = { cx: 100, cy: 200, r: 3 }, circ = []; for (let i = 0; i <= 24; i++) circ.push([c.cx + c.r * Math.cos(i * Math.PI / 12), c.cy + c.r * Math.sin(i * Math.PI / 12)]);
  const paints = [{ pid: 0, kind: 's', subs: [{ pts: circ, curves: 0, closed: true }, { pts: [[94, 200], [106, 200]], curves: 0 }, { pts: [[100, 194], [100, 206]], curves: 0 }] }];
  // 기호 둘레는 칸처럼 막혀 있고(88~112 × 190~210) 글자도 칸 안이지만(84~118 × 160~190) 두 칸이 변을 함께 쓰지 않음
  const segs = [[88, 188, 88, 212], [112, 188, 112, 212], [88, 210, 112, 210], [84, 150, 84, 190], [118, 150, 118, 190], [84, 160, 118, 160], [84, 190, 118, 190], [94, 200, 106, 200], [100, 194, 100, 206]].map(([x1, y1, x2, y2]) => ({ x1, y1, x2, y2 }));
  const items = [{ id: 'i1', page: 1, str: '146', x: 104, y: 192, w: 20, h: 12, a: -Math.PI / 2, box: L.box(94, 172, 107, 192) }];
  const r = L.strokeSymbols(paints, segs, items, 8.6);
  assert.ok(!r.items.some(i => i.src === 'symbol'), JSON.stringify(r.items.map(i => i.str)));
});
test('용접 기준선은 화살표 · 꼬리가 비스듬히 붙은 선만(숨은선이 붙은 외형선 아님) · J 기호(세로선 + 호)는 형상 · 꺾은선 원 안 품번은 후보 아님', () => {
  const pg = pagesE[0];
  assert.equal(pg.refs.length, 3, '기준선 ' + pg.refs.length);
  const c = F.pageCandidates(pg, L.cleanSettings({ std: 'weld' }), {}, { balloons: [] }).cands.filter(x => L.stdOf(x) === 'weld');
  assert.equal(c.filter(x => x.kind === '용접 형상').length, 1, JSON.stringify(c.map(x => x.raw)));
  assert.deepEqual(c.filter(x => x.kind === '각장').map(x => x.spec), ['z6 z8']);
  const all = L.detectCandidates(pg, L.cleanSettings(), {});
  assert.ok(!all.cands.some(x => x.raw === '5'), '품번 5 가 후보');
  assert.ok(all.excluded.some(x => x.raw === '5' && /원 안/.test(x.reason)));
});
test('「⊥」 를 두 토막씩 그은 선 글꼴 기호도 하나로 · 짧은 색 지시선(가운데서 나뉜 첫 토막)도 지시선', () => {
  // 틀: 칸1(기호) 100~117 · 칸2(글자) 117~150, 높이 17
  const segs = [[100, 300, 150, 300], [100, 317, 150, 317], [100, 300, 100, 317], [117, 300, 117, 317], [150, 300, 150, 317]].map(([x1, y1, x2, y2]) => ({ x1, y1, x2, y2 }));
  const paints = [{ pid: 0, kind: 's', subs: [{ pts: [[103, 314.5], [108.5, 314.5], [114, 314.5]], curves: 0 }, { pts: [[108.5, 314.5], [108.5, 308.5], [108.5, 302.5]], curves: 0 }] }];
  const items = [{ id: 'i1', page: 1, str: '0,1', x: 121, y: 312.9, w: 16, h: 11.8, a: 0, box: L.box(121, 303.5, 137, 315.5) }];
  const r = L.strokeSymbols(paints, segs, items, 8.5);
  assert.ok(r.items.some(i => i.str === '⊥'), JSON.stringify(r.items.map(i => i.str)));
  const lt = L.leaderTarget({ cx: 100, cy: 100, r: 21.6 }, [{ x1: 120.6, y1: 100, x2: 122.7, y2: 100, color: '255,0,0' }, { x1: 122.7, y1: 100, x2: 124.8, y2: 100, color: '255,0,0' }]);
  assert.ok(lt && lt.x > 122, JSON.stringify(lt));
  // ◎ 를 꺾은선 원 둘(곡선 없음)로 그린 것
  const ring = (cx, cy, r) => { const p = []; for (let i = 0; i <= 32; i++) p.push([cx + r * Math.cos(i * Math.PI / 16), cy + r * Math.sin(i * Math.PI / 16)]); return { pts: p, curves: 0, closed: false }; };
  const r2 = L.strokeSymbols([{ pid: 0, kind: 's', subs: [ring(108.5, 308.5, 5), ring(108.5, 308.5, 2.8)] }], segs, [{ id: 'i2', page: 1, str: '0,1', x: 121, y: 312.9, w: 16, h: 11.8, a: 0, box: L.box(121, 303.5, 137, 315.5) }], 8.5);
  assert.ok(r2.items.some(i => i.str === '◎'), JSON.stringify(r2.items.map(i => i.str)));
});
test('용접 기준선 옆 숫자: 기준선과 같은 방향 · 기준선 위쪽 구간의 숫자만 용접 크기(세로 치수 · 기준선 왼쪽 치수는 그대로)', () => {
  const tok = (id, str, x, y, a) => ({ id, str, cx: x, cy: y, h: 12, a, box: a ? L.box(x - 6, y - 8, x + 6, y + 8) : L.box(x - 8, y - 6, x + 8, y + 6) });
  const ref = { x1: 300, y1: 400, x2: 400, y2: 400, a: 0, L: 100, u0: 300, u1: 400, v: 400, dash: 1, tip: true, syms: [{ side: -1, u0: 360, u1: 372, ext: 11, box: L.box(360, 389, 372, 400) }] };
  const pg = { page: 1, w: 1200, h: 800, tokens: [tok('a', '6', 330, 393, 0), tok('b', '43.5', 340, 392, -Math.PI / 2), tok('c', '41', 270, 393, 0)], circles: [{ cx: 1, cy: 1, r: 3 }], refs: [ref] };
  const c = L.detectCandidates(pg, L.cleanSettings(), {}).cands;
  assert.deepEqual(c.map(x => [x.raw, x.kind]).sort(), [['41', '길이'], ['43.5', '길이'], ['6', '용접 크기']]);
});
test('용접 기호: 떨어진 두 세로선(플러그 · 슬롯)도 기호 하나', () => {
  const segs = [{ x1: 100, y1: 300, x2: 200, y2: 300 }, { x1: 100, y1: 300, x2: 86, y2: 314 }]; // 기준선 + 비스듬한 화살표 선
  for (let x = 103; x < 196; x += 6) segs.push({ x1: x, y1: 302.3, x2: x + 3, y2: 302.3 }); // 점선
  segs.push({ x1: 148, y1: 300, x2: 148, y2: 283 }, { x1: 156, y1: 300, x2: 156, y2: 283 }); // 8pt 떨어진 두 세로선(위 막대 없음)
  const refs = L.findWeldRefs(segs);
  assert.equal(refs.length, 1); assert.equal(refs[0].syms.length, 1);
  assert.ok(Math.abs(refs[0].syms[0].u1 - refs[0].syms[0].u0 - 8) < 0.5, JSON.stringify(refs[0].syms));
});
test('풍선 연결: 가까운 짝을 먼저(옆 풍선이 먼 글자를 두고 가까운 짝을 빼앗지 않음)', () => {
  // 1번 풍선의 지시선은 a 바로 옆에서 끝나고, 2번 풍선의 지시선은 a · b 사이에서 끝나 치수선을 타고 내려감
  const cands = [{ id: 'a', box: L.box(130, 128, 144, 150) }, { id: 'b', box: L.box(158, 123, 172, 146) }];
  const balloons = [{ no: 1, page: 1, cx: 105, cy: 138, r: 20 }, { no: 2, page: 1, cx: 118, cy: 93, r: 20 }];
  const segs = [{ x1: 125, y1: 138.4, x2: 128.7, y2: 138.6, color: '255,0,0' }, { x1: 132.6, y1: 104.7, x2: 145.3, y2: 113.6, color: '255,0,0' }, { x1: 145.3, y1: 113.6, x2: 145.2, y2: 165, color: '0,0,0' }];
  const lk = L.linkBalloons(balloons, cands, { page: 1, segs }, L.cleanSettings());
  assert.deepEqual(lk.map(x => [x.no, x.candId]), [[1, 'a'], [2, 'b']]);
});
test('도면 E 전체: 수강생 기준서 양식 정답으로 목표 1 · 목표 2(치수 · 용접 풍선 PDF 따로) 완전 일치율', async () => {
  const runs = [[pagesE, 'dim', 'value'], [pagesEb, 'dim', 'number'], [pagesE, 'weld', 'value'], [pagesEw, 'weld', 'number']];
  for (const [pages, std, mode] of runs) {
    const s = L.cleanSettings({ std });
    const items = mode === 'number' ? F.goal2(pages, s, {}) : F.goal1(pages, s, {});
    const sc = L.scoreResults(items.map(L.toPred), truthE.filter(t => t.std === std), mode);
    assert.equal(sc.fullRate, 1, std + ' ' + mode + ' ' + JSON.stringify(sc.wrong.concat(sc.misses)));
    assert.equal(sc.precision, 1, std + ' ' + mode + ' 더 찾음 ' + JSON.stringify(sc.extra));
  }
});

section('예시 도면 5종 — 전체 흐름(pdf.js)');
test('정답 엑셀과 js 안 정답이 같음', () => {
  for (const d of Sample.drawings) {
    if (Sample.truthSheets && Sample.truthSheets[d.key]) { // 도면 E: 기준서 양식 — 엑셀 두 시트 = js 의 시트별 정답
      const x = M.readTruth(path.join(root, 'samples', d.name + '_정답.xlsx')), j = M.sampleSheetTruth(d.key);
      assert.equal(x.length, j.length); assert.equal(x.length, 11);
      x.forEach((t, i) => { assert.equal(t.no, j[i].no); assert.equal(t.std, j[i].std); assert.deepEqual(L.sameSheetFields(t.sheet, j[i].sheet), [true, true, true], d.name + ' ' + t.no); });
      continue;
    }
    const fromXlsx = M.readTruth(path.join(root, 'samples', d.name + '_정답.xlsx'));
    assert.equal(fromXlsx.length, Sample.truth[d.key].length);
    fromXlsx.forEach((t, i) => { assert.equal(t.no, Sample.truth[d.key][i].no); assert.equal(L.sameValue(t, Sample.truth[d.key][i]), true, d.name + ' ' + t.no); });
  }
});
test('목표 1(풍선 없음)·목표 2(풍선 있음) 일치율 — 예시에서 85% 이상, 선으로 그린 「8」 은 못 읽는 것이 정상', async () => {
  const got = {};
  for (const d of Sample.drawings) {
    if (Sample.truthSheets && Sample.truthSheets[d.key]) continue; // 도면 E 는 아래 3차 절에서
    for (const [suf, ball] of [['_풍선없음.pdf', false], ['_풍선있음.pdf', true]]) {
      const r = await M.measure(lib, path.join(root, 'samples', d.name + suf), Sample.truth[d.key], ball);
      got[d.key + (ball ? '2' : '1')] = r.score;
      assert.ok(r.score.valueRate >= 0.85, d.name + suf + ' ' + r.score.valueRate);
      assert.equal(r.score.precision, 1, d.name + suf + ' 엉뚱한 것을 찾음 ' + JSON.stringify(r.score.extra));
    }
  }
  assert.deepEqual(got.C1.misses.map(m => m.no), [12]);
  assert.deepEqual(got.C2.misses.map(m => m.no), [12]);
  assert.equal(got.B2.byCat['용접 형상'].exact, 1);
  assert.equal(got.B2.byCat['용접'].exact, 3);
});

for (const [name, fn] of queue) {
  if (name === null) { console.log('\n' + fn); continue; }
  try { await fn(); passed++; console.log('  ok  ' + name); }
  catch (e) { console.error('  FAIL ' + name + '\n       ' + e.message); process.exitCode = 1; }
}
console.log('\n' + passed + '개 통과' + (process.exitCode ? ' — 실패 있음' : ''));

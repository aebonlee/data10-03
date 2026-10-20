// 실행: node test/logic.test.mjs   (설치 없음 — vendor 의 pdf.js · SheetJS 만 씀)
// 실제 회사 도면은 아직 받지 않았습니다. 가상 글자 조각과 scripts/make-samples.py 로 만든 가상 도면 3종으로 검사합니다.
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

section('예시 도면 3종 — 전체 흐름(pdf.js)');
test('정답 엑셀과 js 안 정답이 같음', () => {
  for (const d of Sample.drawings) {
    const fromXlsx = M.readTruth(path.join(root, 'samples', d.name + '_정답.xlsx'));
    assert.equal(fromXlsx.length, Sample.truth[d.key].length);
    fromXlsx.forEach((t, i) => { assert.equal(t.no, Sample.truth[d.key][i].no); assert.equal(L.sameValue(t, Sample.truth[d.key][i]), true, d.name + ' ' + t.no); });
  }
});
test('목표 1(풍선 없음)·목표 2(풍선 있음) 일치율 — 예시에서 85% 이상, 선으로 그린 「8」 은 못 읽는 것이 정상', async () => {
  const got = {};
  for (const d of Sample.drawings) {
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

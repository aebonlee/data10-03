// 헤드리스 브라우저로 주요 흐름 점검 (선택 — playwright-core 가 있는 PC 에서만)
//   PLAYWRIGHT_CORE=/경로/node_modules/playwright-core node scripts/check-browser.mjs
// file:// 로 index.html 을 열어 예시 도면 A(풍선 없음) → 후보 확인 · 글자 눌러 추가 · 제외 영역 그리기·지우기 → 목표 1 풍선 배치 · 끌어 옮기기
// → 풍선 PDF · 엑셀 내려받아 다시 읽기 → 일치율(예시 정답) · 기록 → 예시 도면 B·C(풍선 있음) 목표 2 → 도면 D(윤곽선 · 기준서 구분)
// → 도면 E(선 글꼴 · 기준서 양식 엑셀 · 2x · 일반공차) → 설정 순서로 누르고,
// JS 오류와 폭별 가로 넘침을 잽니다. BREAK=1 이면 일부러 넓은 상자를 넣어 넘침 점검기가 잡는지 봅니다.
import { createRequire } from 'node:module';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const { chromium } = require(process.env.PLAYWRIGHT_CORE || 'playwright-core');
const XLSX = require(path.join(root, 'vendor/xlsx.full.min.js'));
const M = await import('./measure-samples.mjs');
const L = require(path.join(root, 'js/logic.js'));
const F = require(path.join(root, 'js/flow.js'));
const url = 'file://' + path.join(root, 'index.html');
const shots = process.env.SHOTS; // 스크린숏 저장 폴더(선택)

const errors = [], problems = [];
const browser = await chromium.launch();
const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1440, height: 1000 } });
const page = await ctx.newPage();
page.on('pageerror', e => errors.push('pageerror: ' + e.message));
page.on('console', m => { if (m.type() === 'error') errors.push('console: ' + m.text()); });
function ok(cond, msg) { if (!cond) problems.push(msg); else console.log('  ok  ' + msg); }
async function go(hash) { await page.evaluate(h => { location.hash = h; }, hash); await page.waitForTimeout(250); }
async function clickText(sel, text) { await page.locator(sel, { hasText: text }).first().click(); await page.waitForTimeout(250); }
async function waitCanvas() { await page.waitForSelector('.sheet canvas', { timeout: 15000 }); }
async function shot(name) { if (shots) await page.screenshot({ path: path.join(shots, name + '.png'), fullPage: false }); }
async function dl(fn) { const [d] = await Promise.all([page.waitForEvent('download'), fn()]); const f = path.join(root, 'samples', '.check-' + Date.now() + path.extname(d.suggestedFilename())); await d.saveAs(f); const b = fs.readFileSync(f); fs.unlinkSync(f); return { name: d.suggestedFilename(), bytes: b }; }
// 도면 좌표(pt) → 화면 좌표
async function client(ptx, pty) {
  return page.evaluate(([x, y]) => {
    const s = document.querySelector('svg.ov'), vb = s.viewBox.baseVal;
    let r = s.getBoundingClientRect(), cy = r.top + y / vb.height * r.height;
    if (cy < 80 || cy > innerHeight - 40) { window.scrollBy(0, cy - innerHeight / 2); r = s.getBoundingClientRect(); }
    return [r.left + x / vb.width * r.width, r.top + y / vb.height * r.height];
  }, [ptx, pty]);
}
const rowsOnPage = () => page.locator('main section.card table.list tbody tr').count();

await page.goto(url);
if (process.env.BREAK) await page.addStyleTag({ content: '.page-head { min-width: 900px; }' });
ok((await page.textContent('main')).includes('일하는 순서'), '대시보드');
await clickText('button', '풍선 없음 → 목표 1');
await page.waitForFunction(() => location.hash === '#/drawing' && document.querySelector('main').textContent.includes('검사 대상 후보'), null, { timeout: 15000 });
await waitCanvas();
ok(await page.locator('#sampleBanner').isVisible(), '예시 도면 띠');
const n0 = await page.locator('svg.ov rect.cand').count();
ok(n0 === 14, '도면 A 후보 14개 표시(' + n0 + ')');
await shot('01_drawing');

// 글자 눌러 추가 — 참고 치수 (100) 은 규칙이 뺀 것
await clickText('.seg button', '글자 눌러 추가');
const tok = await page.evaluate(() => { const r = [...document.querySelectorAll('svg.ov rect.tok')].find(x => x.querySelector('title').textContent.startsWith('(100)')); return r ? [+r.getAttribute('x') + 3, +r.getAttribute('y') + 3] : null; });
ok(!!tok, '추가 모드에서 (100) 글자 네모');
if (tok) { const c = await client(...tok); await page.mouse.click(c[0], c[1]); await page.waitForTimeout(300); }
ok(await page.locator('svg.ov rect.cand').count() === 15, '글자 눌러 후보 추가 → 15개');
await page.locator('table.list tbody tr', { hasText: '(100)' }).locator('button', { hasText: '빼기' }).click();
await page.waitForTimeout(200);
ok(await page.locator('svg.ov rect.cand').count() === 14, '추가한 후보 빼기 → 14개');

// 제외 영역 그리기 — 「120」 둘레
await clickText('.seg button', '제외 영역 그리기');
const t120 = await page.evaluate(() => { const r = [...document.querySelectorAll('svg.ov rect.cand')].find(x => x.querySelector('title').textContent.endsWith('· 120')); return [+r.getAttribute('x'), +r.getAttribute('y'), +r.getAttribute('width'), +r.getAttribute('height')]; });
await client(t120[0], t120[1]); const a = await client(t120[0] - 6, t120[1] - 6), b = await client(t120[0] + t120[2] + 6, t120[1] + t120[3] + 6);
await page.mouse.move(a[0], a[1]); await page.mouse.down(); await page.mouse.move((a[0] + b[0]) / 2, (a[1] + b[1]) / 2); await page.mouse.move(b[0], b[1]); await page.mouse.up();
await page.waitForTimeout(300);
ok(await page.locator('svg.ov rect.cand').count() === 13, '제외 영역 안 「120」 빠짐 → 13개');
ok(await page.locator('svg.ov rect.zone.user').count() === 1, '직접 그린 제외 영역 1개');
const z = await page.evaluate(() => { const r = document.querySelector('svg.ov rect.zone.user'); return [+r.getAttribute('x') + 2, +r.getAttribute('y') + 2]; });
const zc = await client(...z); await page.mouse.click(zc[0], zc[1]); await page.waitForTimeout(200);
await page.locator('#dialog button', { hasText: '지우기' }).click(); await page.waitForTimeout(300);
ok(await page.locator('svg.ov rect.cand').count() === 14, '제외 영역 지우면 다시 14개');
await clickText('.seg button', '보기·고치기');
ok((await page.textContent('main')).includes('애매한 후보 AI 에게 묻기'), 'AI 프롬프트 칸');

// 목표 1
await go('#/balloon'); await waitCanvas();
ok(await page.locator('svg.ov g.bl').count() === 14, '풍선 14개');
ok(await page.locator('svg.ov g.bl.ovl').count() === 0, '겹친 풍선 0개');
await shot('02_balloon');
const g = await page.evaluate(() => { const e = document.querySelector('svg.ov g.bl'); const m = /translate\(([\d.]+) ([\d.]+)\)/.exec(e.getAttribute('transform')); return [+m[1], +m[2]]; });
await client(...g); const g0 = await client(...g), g1 = await client(g[0] + 30, g[1] + 20);
await page.mouse.move(g0[0], g0[1]); await page.mouse.down(); await page.mouse.move(g1[0], g1[1], { steps: 4 }); await page.mouse.up(); await page.waitForTimeout(300);
const moved = await page.evaluate(() => { const st = window.__app.state(); return Object.keys(st.db.docs[st.cur.key].bpos).length; });
ok(moved === 1, '풍선 끌어 옮기기 저장');
const pdf = await dl(() => clickText('button', '풍선 PDF 내려받기'));
const lib = await M.loadPdfjs();
const back = await M.readPages(lib, pdf.bytes);
const g2 = F.goal2(back, L.cleanSettings(), {});
const sc = L.scoreResults(g2.map(L.toPred), require(path.join(root, 'js/sample-pdfs.js')).truth.A, 'value');
ok(back.length === 1 && g2.length === 14 && sc.valueRate === 1, '내려받은 풍선 PDF 다시 읽기: 풍선 ' + g2.length + '개 · 값 일치 ' + sc.valueRate + ' · ' + Math.round(pdf.bytes.length / 1024) + 'KB');
const x1 = await dl(() => clickText('button', '엑셀 내려받기'));
const wb = XLSX.read(x1.bytes, { type: 'buffer' });
ok(wb.SheetNames.join(',') === '치수 기준서,치수 기준서 상세,용접 기준서,용접 기준서 상세,도면 정보,제외된 글자', '엑셀 시트 6개(치수 · 용접 기준서 + 상세)');
const rows = XLSX.utils.sheet_to_json(wb.Sheets['치수 기준서 상세'], { header: 1 });
ok(rows.length === 15 && rows[0][0] === '번호' && rows.slice(1).every(r => r[3]) && rows[1][22] === '치수', '치수 기준서 상세 14행 + 머리 · 기준서 칸');
const rowsS = XLSX.utils.sheet_to_json(wb.Sheets['치수 기준서'], { header: 1 });
ok(rowsS.length === 15 && rowsS[0].slice(0, 7).join('|') === L.SHEET_HEAD.join('|') && rowsS.slice(1).every(r => r[1]), '치수 기준서 = 수강생 기준서 칸(No. · Inspection Item · Standard · MAX · MIN · 기구 · 개수)');
const png = await dl(() => clickText('button', '이 쪽 PNG'));
ok(png.bytes.slice(1, 4).toString() === 'PNG', 'PNG 내려받기');

// 일치율
await go('#/accuracy');
await clickText('button', '예시 정답 불러오기');
const accText = await page.textContent('main');
ok(accText.includes('100.0%') && accText.includes('목표 85% 이상'), '목표 1 일치율 100%');
await clickText('button', '측정 기록에 남기기');
ok(await page.locator('.trend svg circle').count() === 1, '측정 기록 1건 · 추이 그래프');
await shot('03_accuracy');
const sx = await dl(() => clickText('button', '대조 결과 엑셀'));
ok(XLSX.read(sx.bytes, { type: 'buffer' }).SheetNames.join(',') === '요약,구분별,틀린 항목', '대조 결과 엑셀');

// 후보 값 고치기 → 일치율이 내려가는지(측정이 고침을 반영)
await go('#/drawing'); await waitCanvas();
await page.locator('table.list tbody tr', { hasText: 'R5' }).first().click(); await page.waitForTimeout(200);
await page.locator('.edit-form label', { hasText: '공칭값' }).locator('input').fill('6');
await clickText('.edit-form button', '저장');
await go('#/accuracy');
ok((await page.textContent('main')).includes('92.9%'), '값을 틀리게 고치면 일치율 92.9%');

// 목표 2 — 도면 B(용접), C(선으로 그린 치수)
await go('#/home');
await page.locator('.sample-card', { hasText: '예시 도면 B' }).locator('button', { hasText: '풍선 있음' }).click();
await page.waitForFunction(() => location.hash === '#/extract' && document.querySelectorAll('svg.ov g.bl').length > 0, null, { timeout: 15000 });
await waitCanvas();
ok(await page.locator('svg.ov g.bl').count() === 13 && await page.locator('svg.ov g.bl.miss').count() === 0, '도면 B 풍선 13개 모두 연결');
const wtext = await page.textContent('main table.list');
ok(wtext.includes('용접 형상') && wtext.includes('6 C.W') && wtext.includes('단속 필릿'), '용접 · 용접 형상 행');
await page.waitForTimeout(800);
ok(await page.locator('main table.list img.thumb').count() >= 1, '용접 형상 그림 미리보기');
await shot('04_extract_B');
const x2 = await dl(() => clickText('button', '엑셀 내려받기'));
const r2 = XLSX.utils.sheet_to_json(XLSX.read(x2.bytes, { type: 'buffer' }).Sheets['용접 기준서 상세'], { header: 1 });
const shapeRow = r2.find(r => r[4] === '용접 형상');
ok(shapeRow && /\d/.test(shapeRow[21] || ''), '엑셀에 용접 형상 잘라낼 영역 좌표');
await go('#/accuracy'); await clickText('button', '예시 정답 불러오기');
ok((await page.textContent('main')).includes('100.0%'), '도면 B 목표 2 일치율 100%');
// 연결 바꾸기: 1번 풍선 → 다른 후보
await go('#/extract'); await waitCanvas();
await page.locator('svg.ov g.bl[data-b="1:1"]').click({ force: true }); await page.waitForTimeout(200);
await clickText('aside button', '연결 비우기');
ok(await page.locator('svg.ov g.bl.miss').count() === 1, '연결 비우기');
await clickText('aside button', '자동 연결로');
ok(await page.locator('svg.ov g.bl.miss').count() === 0, '자동 연결로 되돌리기');

await go('#/home');
await page.locator('.sample-card', { hasText: '예시 도면 C' }).locator('button', { hasText: '풍선 있음' }).click();
await page.waitForFunction(() => document.querySelectorAll('svg.ov g.bl').length > 0, null, { timeout: 15000 });
ok(await page.locator('svg.ov g.bl.miss').count() === 1, '도면 C: 선으로 그린 「8」 풍선 1개는 연결 못 함');
await go('#/accuracy'); await clickText('button', '예시 정답 불러오기');
ok((await page.textContent('main')).includes('91.7%'), '도면 C 목표 2 일치율 91.7%');

// 2차: 도면 D — 글자 층 없는 윤곽선 도면, 기준서 구분(치수 · 용접)
await go('#/home');
await page.locator('.sample-card', { hasText: '예시 도면 D' }).locator('button', { hasText: '풍선 없음' }).click();
await page.waitForFunction(() => location.hash === '#/drawing' && document.querySelector('main').textContent.includes('검사 대상 후보'), null, { timeout: 30000 });
ok((await page.textContent('main')).includes('윤곽선을 글자 모양과 맞춰 되읽었습니다'), '도면 D: 윤곽선 글자 되읽기 안내');
await go('#/balloon'); await waitCanvas();
await clickText('button', '용접 기준서');
ok((await page.locator('svg.ov g.bl').count()) === 6, '용접 기준서만: 풍선 6개');
await clickText('button', '치수 기준서');
ok((await page.locator('svg.ov g.bl').count()) === 13, '치수 기준서만: 풍선 13개');
await clickText('button', '전체');
await go('#/accuracy'); await clickText('button', '예시 정답 불러오기');
const dacc = await page.textContent('main');
ok(dacc.includes('완전 일치율') && dacc.includes('필드 일치율 · 항목명'), '수강생 방식 일치율(필드 · 완전) 표시');

// 3차: 도면 E — 한 줄 선 글꼴 · 자간 넓힌 글 · 일반공차 · 2x, 정답은 수강생 기준서 양식(치수 · 각장 시트)
await go('#/home');
await page.locator('.sample-card', { hasText: '예시 도면 E' }).locator('button', { hasText: '풍선 없음' }).click();
await page.waitForFunction(() => location.hash === '#/drawing' && document.querySelector('main').textContent.includes('검사 대상 후보'), null, { timeout: 30000 });
await go('#/balloon'); await waitCanvas();
await clickText('button', '전체');
const xe = await dl(() => clickText('button', '엑셀 내려받기'));
const wbe = XLSX.read(xe.bytes, { type: 'buffer' });
const re = XLSX.utils.sheet_to_json(wbe.Sheets['치수 기준서'], { header: 1, defval: '' });
const tapRow = re.find(r => r[1] === 'TAP'), holeRow = re.find(r => r[1] === 'Hole' && r[2] === '11');
ok(tapRow && tapRow[2] === 'M20x2.5' && tapRow[6] === 2 && tapRow[7] === '' && tapRow[8] === '' && (tapRow.length <= 9 || re[0][9] !== '측정값 3' || tapRow[9] === '-'), '「2x M20」 → TAP M20x2.5 · 개수 2 · 측정값 칸 2개');
ok(holeRow && holeRow[3] === 0.018 && holeRow[4] === 0, 'Ø11 H7 → Hole 11 · MAX 0.018 · MIN 0');
ok(re.some(r => r[1] === 'Dimension' && r[2] === '262' && r[3] === 2 && r[4] === -2), '자간 넓힌 「2 6 2」 + 일반공차 ISO 13920-BF → 262 ±2');
ok(re.some(r => r[1] === 'Dimension' && r[2] === '146'), '파란 선 글꼴 「146」');
const we = XLSX.utils.sheet_to_json(wbe.Sheets['용접 기준서'], { header: 1, defval: '' });
ok(we.some(r => r[1] === 'Weld' && /^a4\n\(06x06\)$/.test(r[2])) && we.some(r => r[1] === 'Weld' && r[2] === '' && r[5] === 'Visual'), '용접 기준서: 선 글꼴 a4 → a4 (06x06) · 형상 = Weld / Visual');
ok(XLSX.utils.sheet_to_json(wbe.Sheets['도면 정보'], { header: 1 }).some(r => r[0] === '일반공차' && /13920-B/.test(r[1])), '도면 정보에 찾은 일반공차');
await go('#/accuracy'); await clickText('button', '예시 정답 불러오기');
let eacc = await page.textContent('main');
ok(eacc.includes('완전 일치율') && eacc.includes('100.0%'), '도면 E 목표 1(전체) 기준서 양식 정답 대조');
await go('#/home');
await page.locator('.sample-card', { hasText: '예시 도면 E' }).locator('button', { hasText: '용접 풍선' }).click();
await page.waitForFunction(() => location.hash === '#/extract' && document.querySelectorAll('svg.ov g.bl').length > 0, null, { timeout: 30000 });
ok(await page.locator('svg.ov g.bl').count() === 4 && await page.locator('svg.ov g.bl.miss').count() === 0, '도면 E 용접 풍선 4개 모두 연결(기준서 = 용접)');
await go('#/accuracy'); await clickText('button', '예시 정답 불러오기');
eacc = await page.textContent('main');
ok(/완전 일치율\s*100\.0%/.test(eacc.replace(/\s+/g, ' ')) || eacc.includes('100.0%'), '도면 E 목표 2(용접) 100%');

await go('#/settings');
ok((await page.textContent('main')).includes('AI 연결 설정'), '설정 화면 · AI 연결 설정');
ok((await page.textContent('main')).includes('일반공차'), '설정 화면 · 일반공차');

// 폭별 가로 넘침 — 안쪽 가로 스크롤 상자(overflow auto) 안은 의도된 것이라 뺍니다
await go('#/home');
await page.locator('.sample-card', { hasText: '예시 도면 A' }).locator('button', { hasText: '풍선 없음' }).click();
await page.waitForTimeout(1500);
for (const w of [360, 390, 768, 1024, 1280, 1440]) {
  await page.setViewportSize({ width: w, height: 900 });
  for (const r of ['#/home', '#/drawing', '#/balloon', '#/extract', '#/accuracy', '#/settings']) {
    await go(r);
    const bad = await page.evaluate(() => {
      const cw = document.documentElement.clientWidth, out = [];
      for (const el of document.querySelectorAll('main *, .topbar *')) {
        if (el.closest('svg')) continue;
        const b = el.getBoundingClientRect();
        if (!b.width || !b.height) continue;
        let p = el.parentElement, inner = false;
        while (p) { const ox = getComputedStyle(p).overflowX; if (ox === 'auto' || ox === 'scroll' || ox === 'hidden') { inner = true; break; } p = p.parentElement; }
        if (!inner && b.right > cw + 1) out.push(el.tagName + '.' + el.className + ' ' + Math.round(b.right) + '>' + cw);
      }
      return { out: out.slice(0, 3), page: document.documentElement.scrollWidth > cw + 1 };
    });
    if (bad.page || bad.out.length) problems.push(w + 'px ' + r + ' 넘침: ' + bad.out.join(' | '));
    if (w === 390 && r === '#/balloon') await shot('05_390_balloon');
  }
}
console.log('  폭 6가지 × 화면 6개 넘침 점검 끝');

await browser.close();
if (errors.length) problems.push(...errors);
if (problems.length) { console.error('\n문제 ' + problems.length + '건\n' + problems.join('\n')); process.exitCode = 1; }
else console.log('\n브라우저 점검 통과');

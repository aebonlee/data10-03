// 일치율 측정(명령줄) — 브라우저 화면 「일치율 측정」과 같은 로직을 node 에서 돌립니다.
//   예시 도면 전부:       node scripts/measure-samples.mjs
//   내 도면 하나:         node scripts/measure-samples.mjs 도면.pdf 정답.xlsx [--balloon] [--std=dim|weld]   (--balloon = 이미 풍선이 있는 도면 → 목표 2,
//                         정답이 수강생 기준서 양식이면 --std 로 치수 · 용접 시트를 고름)
//   자세히(놓친 것 목록):  ... --detail
// pdf.js 는 vendor/pdfjs 의 것을 씁니다(설치 불필요).
import { createRequire } from 'node:module';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const L = require(path.join(root, 'js/logic.js'));
const F = require(path.join(root, 'js/flow.js'));
const XLSX = require(path.join(root, 'vendor/xlsx.full.min.js'));

export async function loadPdfjs() {
  const w = console.warn, l = console.log;
  console.warn = () => {}; console.log = () => {}; // 그림 그리기용 canvas 가 없다는 안내(글자·선만 쓰므로 무관)
  globalThis.pdfjsWorker = require(path.join(root, 'vendor/pdfjs/pdf.worker.min.js'));
  const lib = require(path.join(root, 'vendor/pdfjs/pdf.min.js'));
  console.warn = w; console.log = l;
  return lib;
}
export async function readPages(lib, bytes) {
  const pdf = await lib.getDocument({ data: new Uint8Array(bytes), isEvalSupported: false, disableFontFace: true, verbosity: 0 }).promise;
  return F.extractAll(pdf, lib.OPS);
}
export function readTruth(file, std) {
  const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer' });
  // 수강생 기준서 양식(No. · Inspection Item · Standard · Tolerance)이면 시트마다 읽어 합침 — std(dim · weld)를 주면 그 기준서만
  let sheet = [];
  wb.SheetNames.forEach(n => { const st = L.parseSheetTruth(XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: '' }), n); if (st) sheet = sheet.concat(st); });
  if (sheet.length) return std ? sheet.filter(t => t.std === std) : sheet;
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: '' });
  return L.parseTruth(aoa).items;
}
export async function measure(lib, pdfFile, truthItems, ballooned, settings) {
  const s = L.cleanSettings(settings);
  const pages = await readPages(lib, fs.readFileSync(pdfFile));
  const items = ballooned ? F.goal2(pages, s, {}) : F.goal1(pages, s, {});
  const score = L.scoreResults(items.map(L.toPred), truthItems, ballooned ? 'number' : 'value');
  if (score.sheetMode) console.log('  기준서 양식 — 필드 일치율 항목명 ' + pct(score.fieldRate.name) + ' · 기준값 ' + pct(score.fieldRate.base) + ' · 공차 ' + pct(score.fieldRate.tol) + ' · 완전 일치율 ' + pct(score.fullRate));
  return { items, score, pages };
}
export function sampleSheetTruth(key) {
  const Sample = require(path.join(root, 'js/sample-pdfs.js')), sh = (Sample.truthSheets || {})[key] || {};
  return Object.keys(sh).flatMap(n => L.parseSheetTruth(sh[n], n) || []);
}
function pct(x) { return (x * 100).toFixed(1) + '%'; }
function line(name, sc) {
  return [name.padEnd(28), ('정답 ' + sc.nTruth).padEnd(8), ('찾음 ' + sc.nPred).padEnd(8), '정밀도 ' + pct(sc.precision), ' 재현율 ' + pct(sc.recall), ' F1 ' + pct(sc.f1), ' 값 일치율 ' + pct(sc.valueRate),
    sc.mode === 'value' ? ' 번호 같음 ' + pct(sc.numberRate) : ''].join('');
}
function detail(sc) {
  sc.misses.forEach(m => console.log('     놓침  ' + m.no + '번 ' + m.text));
  sc.wrong.forEach(m => console.log('     값 다름 ' + m.no + '번 정답 「' + m.truth + '」 · 도구 「' + m.pred + '」'));
  sc.extra.forEach(m => console.log('     더 찾음 ' + (m.no ?? '') + ' ' + m.text));
}

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  const args = process.argv.slice(2), det = args.includes('--detail');
  const lib = await loadPdfjs();
  const files = args.filter(a => !a.startsWith('--'));
  if (files.length >= 2) {
    const std = (args.find(a => a.startsWith('--std=')) || '').slice(6) || undefined; // --std=dim | --std=weld (기준서 양식 정답의 시트 고르기)
    const r = await measure(lib, files[0], readTruth(files[1], std), args.includes('--balloon'), std ? { std } : undefined);
    console.log(line(path.basename(files[0]), r.score)); if (det) detail(r.score);
  } else {
    const Sample = require(path.join(root, 'js/sample-pdfs.js'));
    let sumT = 0, sumE = 0, sumTP = 0, sumP = 0;
    for (const d of Sample.drawings) {
      const sheet = Sample.truthSheets && Sample.truthSheets[d.key]; // 도면 E: 수강생 기준서 양식 — 치수 · 용접 기준서를 따로 잼
      const runs = sheet ? [['_풍선없음.pdf', false, 'dim'], ['_풍선있음.pdf', true, 'dim'], ['_풍선없음.pdf', false, 'weld'], ['_풍선있음_용접.pdf', true, 'weld']] : [['_풍선없음.pdf', false], ['_풍선있음.pdf', true]];
      for (const [suffix, ball, std] of runs) {
        const f = path.join(root, 'samples', d.name + suffix);
        const truth = sheet ? sampleSheetTruth(d.key).filter(t => t.std === std) : Sample.truth[d.key];
        const r = await measure(lib, f, truth, ball, std ? { std } : undefined);
        console.log((ball ? '목표2 ' : '목표1 ') + line(d.name + suffix.replace('.pdf', '') + (std ? '(' + (std === 'dim' ? '치수' : '용접') + ')' : ''), r.score));
        if (det) detail(r.score);
        sumT += r.score.nTruth; sumE += r.score.exact; sumTP += r.score.tp; sumP += r.score.nPred;
      }
    }
    console.log('전체: 정답 ' + sumT + '개 중 값까지 맞음 ' + sumE + '개 (' + pct(sumE / sumT) + ') · 정밀도 ' + pct(sumTP / sumP) + ' · 재현율 ' + pct(sumTP / sumT));
  }
}

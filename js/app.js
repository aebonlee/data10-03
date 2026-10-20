/* Balloon Inspector — 화면 (대시보드 · 도면 열기·인식 · 목표 1 풍선 배치 · 목표 2 풍선 읽기 · 일치율 측정 · 설정)
   계산은 js/logic.js · js/flow.js, 저장은 js/store.js, AI 연결은 js/ai-endpoint.js · js/ai-panel.js 에 있습니다. */
(function () {
  'use strict';
  var L = window.DRLogic, F = window.DRFlow, S = window.DRStore, SM = window.DRSample, XLSX = window.XLSX;
  var CAT_COLOR = { '치수': '#1f5fa8', '공차': '#8a2be2', '기하공차': '#0f7b6c', '용접': '#c2410c', '기타': '#6b7280' };

  var db = S.loadDb();
  var cur = null;        // { key, name, size, pdf, pages: [쪽 자료], sample }
  var canvases = {};     // '쪽@폭' → 그린 canvas
  var ui = { page: 1, zoom: 1, mode: 'view', sel: null, selB: null, filter: 'all', aiText: '', aiParsed: null, accMode: '', loading: '' };

  // ── 작은 도구 ─────────────────────────────────────────────
  function h(tag, attrs) {
    var el = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v == null || v === false) return;
      if (k === 'class') el.className = v;
      else if (k === 'text') el.textContent = v;
      else if (k === 'value') el.value = v;
      else if (k.slice(0, 2) === 'on' && typeof v === 'function') el.addEventListener(k.slice(2), v);
      else if (v === true) el.setAttribute(k, '');
      else el.setAttribute(k, v);
    });
    for (var i = 2; i < arguments.length; i++) add(el, arguments[i]);
    return el;
  }
  function add(el, c) {
    if (c == null || c === false) return;
    if (Array.isArray(c)) { c.forEach(function (x) { add(el, x); }); return; }
    el.appendChild(typeof c === 'string' || typeof c === 'number' ? document.createTextNode(String(c)) : c);
  }
  var SVGNS = 'http://www.w3.org/2000/svg';
  function sv(tag, attrs) {
    var el = document.createElementNS(SVGNS, tag);
    Object.keys(attrs || {}).forEach(function (k) {
      var v = attrs[k]; if (v == null) return;
      if (k.slice(0, 2) === 'on' && typeof v === 'function') el.addEventListener(k.slice(2), v); else el.setAttribute(k, v);
    });
    for (var i = 2; i < arguments.length; i++) { var c = arguments[i]; if (c == null) continue; el.appendChild(typeof c === 'string' || typeof c === 'number' ? document.createTextNode(String(c)) : c); }
    return el;
  }
  function fmt(n) { return n == null || n === '' ? '' : String(L.round(+n, 4)); }
  function pct(x) { return x == null ? '-' : (Math.round(x * 1000) / 10).toFixed(1) + '%'; }
  function toast(msg, isError) {
    var t = document.getElementById('toast');
    t.textContent = msg; t.className = 'toast' + (isError ? ' error' : ''); t.hidden = false;
    clearTimeout(toast._t); toast._t = setTimeout(function () { t.hidden = true; }, 4200);
  }
  function save() {
    if (!S.saveDb(db)) {
      document.getElementById('storeBanner').hidden = false;
      if (S.lastError() === 'quota') toast('브라우저 저장 공간이 찼습니다. 설정에서 JSON 백업을 받아 두세요.', true);
    }
  }
  function stamp() { var d = new Date(); function p(n) { return (n < 10 ? '0' : '') + n; } return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes()); }
  function nowText() { var d = new Date(); function p(n) { return (n < 10 ? '0' : '') + n; } return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' ' + p(d.getHours()) + ':' + p(d.getMinutes()); }
  function field(label, input, hint, cls) { return h('label', { class: 'field' + (cls ? ' ' + cls : '') }, h('span', null, label), input, hint ? h('small', { class: 'hint' }, hint) : null); }
  function select(options, value, attrs) {
    var s = h('select', attrs || {});
    options.forEach(function (o) { var v = Array.isArray(o) ? o[0] : o, t = Array.isArray(o) ? o[1] : o; s.appendChild(h('option', { value: v, selected: String(v) === String(value) }, t)); });
    return s;
  }
  function openDialog(title, content, actions) {
    var d = document.getElementById('dialog');
    document.getElementById('dialogTitle').textContent = title;
    var c = document.getElementById('dialogContent'); c.textContent = ''; add(c, content);
    var a = document.getElementById('dialogActions'); a.textContent = ''; add(a, actions);
    if (!d.open) d.showModal();
  }
  function closeDialog() { var d = document.getElementById('dialog'); if (d.open) d.close(); }
  function confirmBox(title, text, okLabel, onOk, danger) {
    openDialog(title, h('p', { class: 'pre' }, text), [
      h('button', { type: 'button', class: 'btn', onclick: closeDialog }, '그만두기'),
      h('button', { type: 'button', class: 'btn ' + (danger ? 'btn-danger' : 'btn-primary'), onclick: function () { closeDialog(); onOk(); } }, okLabel)]);
  }
  function download(name, blob) {
    var a = h('a', { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  }
  function writeXlsx(name, sheets, widths) {
    var wb = XLSX.utils.book_new();
    Object.keys(sheets).forEach(function (n) {
      var ws = XLSX.utils.aoa_to_sheet(sheets[n]), w = (widths && widths[n]) || [], cols = 0;
      sheets[n].forEach(function (r) { cols = Math.max(cols, r.length); });
      ws['!cols'] = Array.from({ length: cols }, function (_, i) { return { wch: w[i] || 12 }; });
      XLSX.utils.book_append_sheet(wb, ws, n.slice(0, 31));
    });
    download(name, new Blob([XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  }
  function copyText(s) {
    function fallback() {
      var ta = h('textarea', { style: 'position:fixed;left:-9999px;top:0' }); ta.value = s; document.body.appendChild(ta); ta.select();
      var ok = false; try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
      ta.remove(); toast(ok ? '복사했습니다' : '복사하지 못했습니다. 글상자를 눌러 직접 복사하세요', !ok);
    }
    if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(s).then(function () { toast('복사했습니다'); }, fallback); else fallback();
  }
  function pageHead(title, sub, right) { return h('div', { class: 'page-head' }, h('div', { class: 'ph-text' }, h('h1', null, title), sub ? h('p', { class: 'note' }, sub) : null), right || null); }
  function kpi(k, v, unit, cls) { return h('div', { class: 'kpi' + (cls ? ' ' + cls : '') }, h('div', { class: 'k' }, k), h('div', { class: 'v' }, v, unit ? h('small', null, unit) : null)); }
  function safeName(s) { return String(s || '도면').replace(/\.pdf$/i, '').replace(/[\\/:*?"<>|\s]+/g, '_').slice(0, 60); }
  function catDot(c) { return h('span', { class: 'cat-dot', style: 'background:' + (CAT_COLOR[c] || '#9aa3ae') }); }

  // ── 문서 상태 ─────────────────────────────────────────────
  function doc() {
    if (!cur) return null;
    var d = db.docs[cur.key];
    if (!d) { d = db.docs[cur.key] = L.restoreDb({ docs: { k: { name: cur.name } } }).docs.k; }
    d.name = cur.name; d.size = cur.size; d.pages = cur.pages.length; d.sample = cur.sample || '';
    return d;
  }
  function pg(n) { return cur && cur.pages[(n || ui.page) - 1]; }
  function settings() { return db.settings; }
  function pageCands(n) { return F.pageCandidates(pg(n), settings(), doc(), { balloons: L.detectBalloons(pg(n), settings()) }); }
  function allCands() { var o = []; cur.pages.forEach(function (p) { o = o.concat(pageCands(p.page).cands); }); return o; }
  function g1() { return F.goal1(cur.pages, settings(), doc()); }
  function g2() { return F.goal2(cur.pages, settings(), doc()); }
  function needDoc(main, title, sub) {
    if (cur) return false;
    main.appendChild(pageHead(title, sub));
    main.appendChild(h('section', { class: 'card' }, h('p', null, '아직 연 도면이 없습니다. 「도면 열기·인식」에서 PDF 를 고르거나, 예시 도면으로 흐름을 먼저 볼 수 있습니다.'),
      h('div', { class: 'btn-row' }, h('a', { class: 'btn btn-primary', href: '#/drawing' }, '도면 열기로'), h('button', { type: 'button', class: 'btn', onclick: function () { openSample('A', false); } }, '예시 도면 A 열기'))));
    return true;
  }

  // ── pdf.js 불러오기 ───────────────────────────────────────
  // file:// 에서는 브라우저가 Worker 를 막으므로 worker 스크립트를 일반 스크립트로 먼저 넣어 화면 스레드에서 돌게 합니다(data09-08 과 같은 방식).
  // 파일은 바이트로만 넘기고 글꼴표 주소를 주지 않아 밖으로 요청이 나가지 않습니다.
  var pdfLoading = null;
  function withPdf(cb) {
    if (window.pdfjsLib && window.pdfjsWorker) { cb(window.pdfjsLib); return; }
    if (!pdfLoading) {
      pdfLoading = [];
      var addS = function (src, next) {
        var s = document.createElement('script'); s.src = src; s.onload = next;
        s.onerror = function () { toast('PDF 읽기 도구(vendor/pdfjs)를 불러오지 못했습니다. 폴더째 받았는지 확인해 주세요.', true); pdfLoading = null; };
        document.head.appendChild(s);
      };
      addS('vendor/pdfjs/pdf.worker.min.js', function () {
        addS('vendor/pdfjs/pdf.min.js', function () {
          window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdfjs/pdf.worker.min.js';
          var q = pdfLoading; pdfLoading = null; q.forEach(function (f) { f(window.pdfjsLib); });
        });
      });
    }
    pdfLoading.push(cb);
  }
  function b64bytes(b64) { var s = atob(b64), a = new Uint8Array(s.length); for (var i = 0; i < s.length; i++) a[i] = s.charCodeAt(i); return a; }
  function openBytes(bytes, name, key, sample, quiet) {
    ui.loading = name; render();
    withPdf(function (lib) {
      lib.getDocument({ data: bytes.slice(), isEvalSupported: false, disableFontFace: true, verbosity: 0 }).promise.then(function (pdf) {
        return F.extractAll(pdf, lib.OPS).then(function (pages) {
          cur = { key: key, name: name, size: bytes.length, pdf: pdf, pages: pages, sample: sample || '', bytes: bytes };
          cur.ballooned = F.looksBallooned(pages, settings());
          canvases = {}; ui.page = 1; ui.sel = null; ui.selB = null; ui.loading = ''; ui.aiParsed = null; ui.accMode = '';
          var d = doc(); d.opened = nowText(); db.cur = key; save();
          if (!sample) S.putPdf(key, bytes);
          if (!quiet) {
            var nText = pages.reduce(function (a, p) { return a + p.items.length; }, 0);
            if (!nText) toast('이 PDF 에는 글자 층이 없습니다(스캔·그림 또는 글자를 선으로 바꾼 출력). 1단계로는 읽을 수 없고 2단계(OCR) 대상입니다.', true);
            else toast(name + ' — ' + pages.length + '쪽, 글자 ' + nText + '조각을 읽었습니다.' + (cur.ballooned ? ' 번호 풍선이 있는 도면으로 보여 「목표 2 · 풍선 읽기」로 갑니다.' : ''));
            location.hash = cur.ballooned ? '#/extract' : '#/drawing';
          }
          render();
        });
      }).catch(function (e) { ui.loading = ''; render(); toast('PDF 를 열지 못했습니다: ' + e.message, true); });
    });
  }
  function openFile(file) {
    if (!/\.pdf$/i.test(file.name) && file.type !== 'application/pdf') { toast('PDF 파일만 열 수 있습니다. 그림(JPG·PNG)·스캔 도면은 2단계(OCR) 대상입니다.', true); return; }
    var r = new FileReader();
    r.onerror = function () { toast('파일을 읽지 못했습니다.', true); };
    r.onload = function () { openBytes(new Uint8Array(r.result), file.name, L.docKey(file.name, file.size), '', false); };
    r.readAsArrayBuffer(file);
  }
  function sampleMeta(key) { return SM.drawings.filter(function (d) { return d.key === key; })[0]; }
  function openSample(key, ball, quiet) {
    var m = sampleMeta(key); if (!m) return;
    var name = m.name + (ball ? '_풍선있음' : '_풍선없음') + '.pdf', skey = 'sample:' + key + (ball ? '_ball' : '_plain');
    openBytes(b64bytes(SM.pdf[key + (ball ? '_ball' : '_plain')]), name, skey, key, quiet);
  }
  function restoreLast() {
    var k = db.cur; if (!k) return;
    var m = /^sample:([A-Z])_(ball|plain)$/.exec(k);
    if (m) { openSample(m[1], m[2] === 'ball', true); return; }
    S.getPdf(k).then(function (bytes) { if (bytes) openBytes(bytes, db.docs[k].name, k, '', true); });
  }

  // ── 도면 보기(그림 + 겹쳐 그리기) ─────────────────────────
  function pageCanvas(n, cssW, cb) {
    var p = pg(n), dpr = Math.min(window.devicePixelRatio || 1, 2), W = Math.min(5000, Math.round(cssW * dpr)), key = n + '@' + W;
    if (canvases[key]) { cb(canvases[key]); return; }
    cur.pdf.getPage(n).then(function (page) {
      var vp = page.getViewport({ scale: W / p.w }), c = document.createElement('canvas');
      c.width = Math.round(vp.width); c.height = Math.round(vp.height);
      var g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
      return page.render({ canvasContext: g, viewport: vp }).promise.then(function () { canvases[key] = c; cb(c); });
    }).catch(function (e) { toast('도면을 그리지 못했습니다: ' + e.message, true); });
  }
  function ptOf(e, svg) {
    var r = svg.getBoundingClientRect(), p = pg();
    return [(e.clientX - r.left) / r.width * p.w, (e.clientY - r.top) / r.height * p.h];
  }
  function rectEl(b, cls, extra) { return sv('rect', Object.assign({ x: b.x0, y: b.y0, width: Math.max(0.5, b.x1 - b.x0), height: Math.max(0.5, b.y1 - b.y0), class: cls }, extra || {})); }
  /* opts: { draw(svg, p), onDown(e, pt, svg), legend } */
  function viewer(opts) {
    var p = pg(), n = cur.pages.length;
    var bar = h('div', { class: 'viewer-bar' },
      h('div', { class: 'seg', role: 'group', 'aria-label': '쪽' },
        h('button', { type: 'button', disabled: ui.page <= 1, onclick: function () { ui.page--; ui.sel = ui.selB = null; render(); } }, '이전 쪽'),
        h('button', { type: 'button', disabled: true }, ui.page + ' / ' + n + '쪽'),
        h('button', { type: 'button', disabled: ui.page >= n, onclick: function () { ui.page++; ui.sel = ui.selB = null; render(); } }, '다음 쪽')),
      h('span', { class: 'grow' }),
      h('div', { class: 'seg', role: 'group', 'aria-label': '확대' }, [[1, '맞춤'], [1.5, '150%'], [2, '200%'], [3, '300%']].map(function (z) {
        return h('button', { type: 'button', 'aria-pressed': String(ui.zoom === z[0]), onclick: function () { ui.zoom = z[0]; render(); } }, z[1]);
      })));
    var svg = sv('svg', { class: 'ov', viewBox: '0 0 ' + p.w + ' ' + p.h, preserveAspectRatio: 'none', role: 'img', 'aria-label': '도면 위 표시' });
    var sheet = h('div', { class: 'sheet', style: 'width:' + (ui.zoom * 100) + '%;aspect-ratio:' + p.w + ' / ' + p.h }, h('div', { class: 'loading' }, '도면 그리는 중…'), svg);
    var box = h('div', { class: 'viewer' + (ui.zoom === 1 ? ' fit' : '') }, sheet);
    opts.draw(svg, p);
    if (opts.onDown) svg.addEventListener('pointerdown', function (e) { opts.onDown(e, ptOf(e, svg), svg); });
    requestAnimationFrame(function () {
      var w = sheet.clientWidth; if (!w) return;
      pageCanvas(ui.page, w, function (c) { var l = sheet.querySelector('.loading'); if (l) l.remove(); if (!sheet.querySelector('canvas')) sheet.insertBefore(c, svg); });
    });
    return h('div', null, bar, box, opts.legend ? legend() : null);
  }
  function legend() {
    return h('ul', { class: 'legend-ov' }, L.CATS.map(function (c) { return h('li', null, h('i', { class: 'lg-' + c }), c); }),
      h('li', null, h('i', { class: 'lg-zone' }), '제외 영역(회색 = 자동 · 주황 = 직접 그림)'), h('li', null, '점선 테두리 = 확인 필요'));
  }
  function drawCands(svg, list, opt) {
    opt = opt || {};
    list.forEach(function (c) {
      var r = rectEl(L.pad(c.box, 1.5), 'cand k-' + (c.cat || '기타') + (c.check ? ' chk' : '') + (ui.sel === c.id ? ' sel' : '') + (opt.faint ? ' hidden' : ''), { 'data-id': c.id });
      r.appendChild(sv('title', null, (c.cat || '') + ' · ' + c.raw));
      svg.appendChild(r);
    });
  }
  function drawZones(svg, boxes) {
    boxes.forEach(function (b, i) {
      if (b.reason === '테두리 구역') return;
      var r = rectEl(b, 'zone' + (b.auto ? '' : ' user'), b.auto ? null : { 'data-zone': i });
      r.appendChild(sv('title', null, b.reason || '제외 영역')); svg.appendChild(r);
    });
    var bm = (settings().borderPct || 0) / 100 * Math.min(pg().w, pg().h);
    if (settings().autoBorder !== false && bm > 0) svg.appendChild(sv('path', { class: 'zone', 'fill-rule': 'evenodd', d: 'M0 0H' + pg().w + 'V' + pg().h + 'H0Z M' + bm + ' ' + bm + 'V' + (pg().h - bm) + 'H' + (pg().w - bm) + 'V' + bm + 'Z' }, sv('title', null, '테두리 구역')));
  }
  function dragRect(e, svg, start, done) {
    e.preventDefault();
    var r = sv('rect', { class: 'drag-box', x: start[0], y: start[1], width: 0, height: 0 }); svg.appendChild(r);
    function mv(ev) { var q = ptOf(ev, svg), b = L.box(start[0], start[1], q[0], q[1]); r.setAttribute('x', b.x0); r.setAttribute('y', b.y0); r.setAttribute('width', b.x1 - b.x0); r.setAttribute('height', b.y1 - b.y0); }
    function up(ev) {
      window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up);
      var q = ptOf(ev, svg), b = L.box(start[0], start[1], q[0], q[1]); r.remove();
      if (b.x1 - b.x0 < 4 || b.y1 - b.y0 < 4) { toast('끌어서 네모를 그려 주세요.'); return; }
      done({ x0: L.round(b.x0, 1), y0: L.round(b.y0, 1), x1: L.round(b.x1, 1), y1: L.round(b.y1, 1) });
    }
    window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up);
  }
  function cropUrl(n, b, cb) {
    var p = pg(n), W = Math.min(4200, p.w * 3);
    pageCanvas(n, W, function (c) {
      var k = c.width / p.w, x = Math.max(0, b.x0 * k), y = Math.max(0, b.y0 * k), w = Math.min(c.width - x, (b.x1 - b.x0) * k), hh = Math.min(c.height - y, (b.y1 - b.y0) * k);
      var o = document.createElement('canvas'); o.width = Math.max(1, Math.round(w)); o.height = Math.max(1, Math.round(hh));
      o.getContext('2d').drawImage(c, x, y, w, hh, 0, 0, o.width, o.height);
      cb(o);
    });
  }

  // ── 후보 고치기 ───────────────────────────────────────────
  function findCand(id) { return allCands().filter(function (c) { return c.id === id; })[0]; }
  function isManual(c) { return c && c.src === 'manual'; }
  function updateCand(c, patch) {
    var d = doc();
    if (isManual(c)) {
      var list = d.added[c.page] || [];
      list.forEach(function (a) { if (a.id === c.id) Object.assign(a, patch); });
    } else d.changed[c.id] = Object.assign(d.changed[c.id] || {}, patch);
    save();
  }
  function removeCand(c) {
    var d = doc();
    if (isManual(c)) d.added[c.page] = (d.added[c.page] || []).filter(function (a) { return a.id !== c.id; });
    else d.removed[c.id] = 1;
    if (ui.sel === c.id) ui.sel = null;
    save(); render();
  }
  function addManual(page, base) {
    var d = doc(), id = 'm' + page + '_' + Date.now().toString(36);
    var c = Object.assign({ id: id, page: page, src: 'manual', tokens: [], check: false, note: '직접 추가' }, base, { id: id });
    (d.added[page] = d.added[page] || []).push(c);
    ui.sel = id; save(); render();
    return c;
  }
  function addFromToken(t) {
    var c = L.classifyToken(t.str);
    var m = L.makeCand(ui.page, t, c.cat ? c : Object.assign(c, { cat: '치수', check: true, note: '규칙이 대상 아님으로 본 글자 — 값을 확인해 주세요' }), { src: 'manual' });
    delete m.id;
    addManual(ui.page, m);
    toast('「' + t.str + '」 을 후보로 넣었습니다.');
  }
  function editPanel(c) {
    if (!c) return h('div', { class: 'note' }, h('p', null, '도면 위 네모를 누르면 여기서 구분·값을 고칩니다.'),
      h('ul', { class: 'list-plain' }, h('li', null, '「글자 눌러 추가」: 규칙이 놓친 글자를 눌러 후보로 넣습니다.'), h('li', null, '「제외 영역 그리기」: 표제란·부품표처럼 검사 대상이 아닌 곳을 끌어서 네모로 막습니다.'), h('li', null, '「형상 영역 그리기」: 글자 없이 형상만 있는 용접 기호를 네모로 감싸 「용접 형상」 항목으로 넣습니다.')));
    var f = {};
    function inp(k, v) { f[k] = h('input', { value: v == null ? '' : v, inputmode: k === 'raw' || k === 'kind' || k === 'note' ? null : 'decimal' }); return f[k]; }
    var catSel = select(L.CATS.map(function (x) { return [x, x]; }), c.cat);
    return h('div', { class: 'edit-form' },
      h('p', null, catDot(c.cat), h('strong', null, c.raw), ' ', c.check ? h('span', { class: 'badge warn' }, '확인 필요') : null, ' ', isManual(c) ? h('span', { class: 'badge muted' }, '직접 추가') : null, c.edited ? h('span', { class: 'badge muted' }, '고침') : null),
      c.note ? h('p', { class: 'note' }, c.note) : null,
      field('구분', catSel), field('종류', inp('kind', c.kind)), field('원문', inp('raw', c.raw)),
      h('div', { class: 'form-grid three' }, field('공칭값', inp('nominal', c.nominal)), field('상한공차', inp('upper', c.upper)), field('하한공차', inp('lower', c.lower))),
      field('단위', inp('unit', c.unit)),
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
          var num = function (k) { var v = f[k].value.trim(); return v === '' ? null : L.numOrNull(v); };
          updateCand(c, { cat: catSel.value, kind: f.kind.value.trim(), raw: f.raw.value.trim() || c.raw, nominal: num('nominal'), upper: num('upper'), lower: num('lower'), unit: f.unit.value.trim(), check: false });
          toast('고쳤습니다.'); render();
        } }, '저장'),
        h('button', { type: 'button', class: 'btn', onclick: function () { var r = L.classifyToken(f.raw.value); if (!r.cat) { toast('원문을 치수 형식으로 읽지 못했습니다.', true); return; } f.nominal.value = fmt(r.nominal); f.upper.value = fmt(r.upper); f.lower.value = fmt(r.lower); catSel.value = r.cat; f.kind.value = r.kind; } }, '원문 다시 읽기'),
        h('button', { type: 'button', class: 'btn btn-danger', onclick: function () { removeCand(c); } }, '후보에서 빼기')));
  }

  // ── 라우팅 · 메뉴 ─────────────────────────────────────────
  function routeParts() { return location.hash.replace(/^#\/?/, '').split('/'); }
  function route() { return routeParts()[0] || 'home'; }
  var SCREENS = { home: renderHome, drawing: renderDrawing, balloon: renderBalloon, extract: renderExtract, accuracy: renderAccuracy, settings: renderSettings };
  function render() {
    var r = SCREENS[route()] ? route() : 'home', main = document.getElementById('main');
    document.body.setAttribute('data-route', r);
    document.body.classList.toggle('mode-draw', ui.mode === 'excl' || ui.mode === 'shape');
    document.body.classList.toggle('mode-add', ui.mode === 'add');
    document.querySelectorAll('.sidebar a[data-route]').forEach(function (a) { if (a.getAttribute('data-route') === r) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current'); });
    document.getElementById('sampleBanner').hidden = !(cur && cur.sample);
    document.getElementById('docText').textContent = cur ? cur.name + ' · ' + cur.pages.length + '쪽' : '열린 도면 없음';
    var y = window.scrollY;
    main.textContent = '';
    if (ui.loading) { main.appendChild(h('section', { class: 'card' }, h('p', null, '「' + ui.loading + '」 을 읽는 중입니다…'))); return; }
    SCREENS[r](main);
    window.scrollTo(0, y);
  }
  window.addEventListener('hashchange', function () { setMenu(false); ui.mode = 'view'; ui.sel = ui.selB = null; render(); window.scrollTo(0, 0); });
  var sidebar = document.getElementById('sidebar'), backdrop = document.getElementById('sideBackdrop'), menuBtn = document.getElementById('menuBtn');
  function setMenu(open) { document.body.classList.toggle('menu-open', open); backdrop.hidden = !open; menuBtn.setAttribute('aria-expanded', String(open)); }
  menuBtn.addEventListener('click', function () { setMenu(!document.body.classList.contains('menu-open')); });
  backdrop.addEventListener('click', function () { setMenu(false); menuBtn.focus(); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && document.body.classList.contains('menu-open')) { setMenu(false); menuBtn.focus(); } });
  sidebar.addEventListener('click', function (e) { if (e.target.closest('a')) setMenu(false); });
  var lastW = window.innerWidth;
  window.addEventListener('resize', function () { clearTimeout(render._r); render._r = setTimeout(function () { if (Math.abs(window.innerWidth - lastW) > 40) { lastW = window.innerWidth; if (cur) render(); } }, 250); });

  // ── 대시보드 ──────────────────────────────────────────────
  function sampleCards() {
    return h('div', { class: 'samples' }, SM.drawings.map(function (d) {
      var desc = { A: '브래킷 — 치수 · 쌓인 공차 · 끼워맞춤 · 기하공차 · 나사 · 각도', B: '용접 구조물 — 용접 기호(크기 · 형상만) · 평행도', C: '축 — 끼워맞춤 · 동심도 · 흔들림 · 선으로 그린 치수 하나' }[d.key];
      return h('div', { class: 'sample-card' }, h('strong', null, '예시 도면 ' + d.key), h('span', { class: 'note' }, desc),
        h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn btn-sm2', onclick: function () { openSample(d.key, false); } }, '풍선 없음 → 목표 1'),
          h('button', { type: 'button', class: 'btn btn-sm2', onclick: function () { openSample(d.key, true); } }, '풍선 있음 → 목표 2')));
    }));
  }
  function trendChart(list, target) {
    var W = 560, H = 180, padL = 34, padB = 22, n = list.length;
    var x = function (i) { return padL + (n <= 1 ? (W - padL - 10) / 2 : i * (W - padL - 10) / (n - 1)); }, y = function (v) { return 10 + (1 - v) * (H - padB - 10); };
    var s = sv('svg', { viewBox: '0 0 ' + W + ' ' + H, role: 'img', 'aria-label': '값 일치율 추이(측정 기록 ' + n + '건, 목표 ' + target + '%)' });
    [0, 0.5, 0.85, 1].forEach(function (v) { s.appendChild(sv('line', { class: v === target / 100 ? 'tg' : 'ax', x1: padL, x2: W - 6, y1: y(v), y2: y(v) })); s.appendChild(sv('text', { x: 2, y: y(v) + 4 }, Math.round(v * 100) + '%')); });
    if (Math.abs(target / 100 - 0.85) > 1e-9) s.appendChild(sv('line', { class: 'tg', x1: padL, x2: W - 6, y1: y(target / 100), y2: y(target / 100) }));
    if (n) {
      s.appendChild(sv('polyline', { class: 'ln', points: list.map(function (r, i) { return x(i) + ',' + y(r.valueRate); }).join(' ') }));
      list.forEach(function (r, i) { s.appendChild(sv('circle', { class: 'pt', cx: x(i), cy: y(r.valueRate), r: 4 }, sv('title', null, r.at + ' · ' + r.doc + ' · ' + pct(r.valueRate)))); });
    }
    return h('div', { class: 'trend' }, s);
  }
  function renderHome(main) {
    main.appendChild(h('div', { class: 'hero' }, h('div', { class: 'hero-text' }, h('h1', null, '도면 검사 기준서 작성 자동화'),
      h('p', null, '번호 풍선이 없는 PDF 도면에는 풍선을 자동으로 달고(목표 1), 풍선이 있는 도면에서는 풍선별 치수·공차·용접 정보를 엑셀로 뽑습니다(목표 2). 무엇보다 먼저 정답과 대조해 지금 일치율을 잽니다.')),
      h('span', { class: 'hero-chip' }, '목표 일치율 ' + settings().target + '%')));
    var flow = [['#/drawing', '도면 열기·인식', '벡터 PDF 의 글자로 검사 대상 후보 찾기'], ['#/drawing', '후보 확인·고치기', '제외 영역 · 글자 눌러 추가 · 값 고치기'],
      ['#/balloon', '목표 1 · 풍선 배치', '번호 순서 · 겹침 피하기 · 풍선 PDF'], ['#/extract', '목표 2 · 풍선 읽기', '풍선 번호 → 지시선 · 가까운 글자 연결'], ['#/accuracy', '일치율 측정', '정답 엑셀과 대조 · 기록']];
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '일하는 순서'), h('ol', { class: 'flow' }, flow.map(function (f, i) {
      return h('li', null, h('a', { href: f[0] }, h('span', { class: 'flow-n' }, String(i + 1)), h('span', { class: 'flow-t' }, h('strong', null, f[1]), h('span', null, f[2]))));
    }))));
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '예시 도면(가상)으로 먼저 보기'),
      h('p', { class: 'note' }, '도면 3종을 직접 만들어 넣었습니다. 같은 도면의 「풍선 없음」·「풍선 있음」 PDF 와 사람이 정리했다고 가정한 정답 엑셀이 한 벌입니다(samples 폴더). 표제란 숫자 · 테두리 구역 번호 · 주기 문장 · 참고 치수 같은 함정을 일부러 넣었습니다.'),
      sampleCards()));
    if (cur) {
      var cs = allCands(), chk = cs.filter(function (c) { return c.check; }).length;
      main.appendChild(h('section', { class: 'card' }, h('div', { class: 'card-head' }, h('h2', null, '열린 도면'), h('a', { class: 'dp-more', href: cur.ballooned ? '#/extract' : '#/balloon' }, cur.ballooned ? '목표 2 로' : '목표 1 로')),
        h('div', { class: 'kpis' }, kpi('도면', cur.name.replace(/\.pdf$/i, ''), '', 'sm'), kpi('쪽 수', String(cur.pages.length), '쪽'), kpi('검사 대상 후보', String(cs.length), '개'), kpi('확인 필요', String(chk), '개'),
          kpi('풍선', cur.ballooned ? '있음' : '없음', ''))));
    }
    var hist = db.history.slice(-30);
    main.appendChild(h('section', { class: 'card' }, h('div', { class: 'card-head' }, h('h2', null, '일치율 기록'), h('a', { class: 'dp-more', href: '#/accuracy' }, '측정하러 가기')),
      hist.length ? trendChart(hist, settings().target) : h('p', { class: 'note' }, '아직 기록이 없습니다. 「일치율 측정」에서 정답 엑셀과 대조한 뒤 「측정 기록에 남기기」를 누르면 여기에 쌓입니다. 규칙을 고칠 때마다 같은 도면으로 다시 재면 85% 에 얼마나 다가갔는지 보입니다.')));
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '1단계에서 할 수 있는 것과 못 하는 것'), h('ul', { class: 'limits' },
      h('li', null, h('strong', null, '벡터 PDF 만 읽습니다. '), 'CAD 에서 PDF 로 내보내 글자를 고를 수 있는 도면이어야 합니다. 스캔 그림 · 글자를 선으로 바꾼 PDF 는 글자 층이 없어 읽지 못합니다(2단계 OCR).'),
      h('li', null, h('strong', null, '규칙(휴리스틱)으로 찾습니다. '), '치수 형식 · 표제란 낱말 · 원과 지시선의 기하로 판단하므로, 회사 도면 양식이 다르면 설정의 낱말·거리 기준을 맞춰야 합니다. 그래서 정답과 대조하는 측정부터 합니다.'),
      h('li', null, h('strong', null, '풍선 PDF 는 원본 쪽을 고해상도 그림으로 넣고 그 위에 풍선을 벡터로 그립니다. '), '원본 글자는 보이지 않는 글자 층으로 같은 자리에 넣어 검색 · 다시 읽기가 됩니다. 원본 벡터를 그대로 두고 덧그리는 편집은 2단계입니다.'),
      h('li', null, h('strong', null, '애매한 후보는 AI 에게 물을 수 있습니다(선택). '), '프롬프트를 복사해 회사가 허용한 AI 에 붙여 넣는 방식이 기본이고, 키는 이 도구에 없습니다.'))));
  }

  // ── 도면 열기·인식 ────────────────────────────────────────
  function openCard() {
    var input = h('input', { type: 'file', accept: '.pdf,application/pdf', 'aria-label': 'PDF 도면 고르기', onchange: function (e) { if (e.target.files[0]) openFile(e.target.files[0]); } });
    var drop = h('label', { class: 'drop' }, input,
      h('strong', null, 'PDF 도면을 끌어 놓거나 눌러서 고르기'), h('span', { class: 'note' }, 'CAD 에서 내보낸 벡터 PDF · 여러 쪽 가능 · 파일은 이 브라우저 안에서만 읽습니다'));
    drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('over'); });
    drop.addEventListener('dragleave', function () { drop.classList.remove('over'); });
    drop.addEventListener('drop', function (e) { e.preventDefault(); drop.classList.remove('over'); if (e.dataTransfer.files[0]) openFile(e.dataTransfer.files[0]); });
    return h('section', { class: 'card' }, h('h2', null, '도면 열기'), drop, h('details', { class: 'kw-box' }, h('summary', null, '예시 도면 열기'), sampleCards()));
  }
  function docInfo() {
    var p = pg(), pc = pageCands(ui.page), nText = cur.pages.reduce(function (a, q) { return a + q.items.length; }, 0);
    var out = h('div', null);
    if (!nText) add(out, h('div', { class: 'alert error' }, '이 PDF 에는 글자 층이 없습니다. 스캔 그림이거나 CAD 가 글자를 선으로 바꿔 내보낸 파일입니다. 1단계 규칙으로는 읽을 수 없습니다 — CAD 에서 「글자를 글자로(TrueType 유지)」 내보내기를 해 보시고, 안 되면 2단계(OCR) 대상입니다.'));
    else if (!p.items.length) add(out, h('div', { class: 'alert warn' }, '이 쪽에는 글자 층이 없습니다.'));
    if (cur.ballooned) add(out, h('div', { class: 'alert info' }, '이 도면에는 번호 풍선이 이미 있습니다(원 안의 숫자 ' + L.detectBalloons(p, settings()).length + '개, 이 쪽). 풍선별 정보를 뽑으려면 ', h('a', { href: '#/extract' }, '목표 2 · 풍선 읽기'), '로 가세요. 아래 후보 목록에서는 풍선 숫자를 뺐습니다.'));
    add(out, h('div', { class: 'kpis' }, kpi('이 쪽 글자 조각', String(p.items.length), '개'), kpi('붙인 토큰', String(p.tokens.length), '개'), kpi('검사 대상 후보', String(pc.cands.length), '개'),
      kpi('제외됨', String(pc.excluded.length), '개'), kpi('확인 필요', String(pc.cands.filter(function (c) { return c.check; }).length), '개'), kpi('선 · 원 · 용접 기호', p.segs.length + ' · ' + p.circles.length + ' · ' + p.tris.length, '', 'sm')));
    return out;
  }
  function renderDrawing(main) {
    main.appendChild(pageHead('도면 열기·인식', '벡터 PDF 의 글자 층에서 치수 · 공차 · 기하공차 · 용접 표기를 찾아 검사 대상 후보로 만듭니다. 틀린 것은 고치고, 놓친 것은 눌러서 더합니다.'));
    main.appendChild(openCard());
    if (!cur) return;
    var p = pg(), pc = pageCands(ui.page), d = doc();
    main.appendChild(h('section', { class: 'card' }, h('div', { class: 'card-head' }, h('h2', null, cur.name), h('span', { class: 'note' }, cur.pages.length + '쪽 · ' + L.mm(p.w) + ' × ' + L.mm(p.h) + 'mm')), docInfo()));
    var modes = [['view', '보기·고치기'], ['add', '글자 눌러 추가'], ['excl', '제외 영역 그리기'], ['shape', '형상 영역 그리기']];
    var modeBar = h('div', { class: 'seg', role: 'group', 'aria-label': '도구' }, modes.map(function (m) { return h('button', { type: 'button', 'aria-pressed': String(ui.mode === m[0]), onclick: function () { ui.mode = m[0]; render(); } }, m[1]); }));
    var v = viewer({
      legend: true,
      draw: function (svg) {
        drawZones(svg, pc.boxes.filter(function (b) { return b.auto; }).concat((d.excl[ui.page] || []).map(function (b) { return Object.assign({ reason: '직접 그린 제외 영역' }, b); })));
        pc.excluded.forEach(function (x) { svg.appendChild(rectEl(L.pad(x.box, 1), 'excl-tok', null)); });
        if (ui.mode === 'add') {
          var used = {}; pc.cands.forEach(function (c) { (c.tokens || []).forEach(function (t) { used[t] = 1; }); });
          p.tokens.forEach(function (t) { if (!used[t.id]) { var r = rectEl(L.pad(t.box, 1), 'tok', { 'data-tok': t.id }); r.appendChild(sv('title', null, t.str + ' — 눌러서 후보로')); svg.appendChild(r); } });
        }
        drawCands(svg, pc.cands);
      },
      onDown: function (e, pt, svg) {
        var t = e.target;
        if (ui.mode === 'excl' || ui.mode === 'shape') {
          if (t.getAttribute && t.getAttribute('data-zone') != null && ui.mode === 'excl') {
            var zi = +t.getAttribute('data-zone') - pc.boxes.filter(function (b) { return b.auto; }).length;
            confirmBox('제외 영역 지우기', '직접 그린 제외 영역을 지울까요? 그 안의 글자가 다시 후보가 됩니다.', '지우기', function () { (d.excl[ui.page] || []).splice(zi, 1); save(); render(); });
            return;
          }
          dragRect(e, svg, pt, function (b) {
            if (ui.mode === 'excl') { (d.excl[ui.page] = d.excl[ui.page] || []).push(b); save(); render(); toast('제외 영역을 더했습니다. 그 안의 글자는 후보에서 빠집니다.'); }
            else addManual(ui.page, { raw: '(용접 형상)', cat: '용접', kind: '용접 형상', box: b, cx: (b.x0 + b.x1) / 2, cy: (b.y0 + b.y1) / 2, h: 10, nominal: null, upper: null, lower: null, unit: '', shape: true, crop: L.pad(b, 6), check: false, note: '직접 그린 형상 영역' });
          });
          return;
        }
        if (t.getAttribute && t.getAttribute('data-tok')) { var tok = p.tokens.filter(function (x) { return x.id === t.getAttribute('data-tok'); })[0]; if (tok) addFromToken(tok); return; }
        if (t.getAttribute && t.getAttribute('data-id')) { ui.sel = t.getAttribute('data-id'); render(); return; }
        if (ui.sel) { ui.sel = null; render(); }
      }
    });
    var selC = ui.sel ? pc.cands.filter(function (c) { return c.id === ui.sel; })[0] : null;
    main.appendChild(h('div', { class: 'work' }, h('section', { class: 'card' }, h('div', { class: 'viewer-bar' }, modeBar), v),
      h('aside', { class: 'card side-panel', 'aria-label': '고르기' }, h('h2', null, selC ? '후보 고치기' : '도움말'), editPanel(selC))));
    main.appendChild(candTable(pc));
    main.appendChild(aiCard());
  }
  function candTable(pc) {
    var list = pc.cands.filter(function (c) { return ui.filter === 'all' || (ui.filter === 'check' ? c.check : c.cat === ui.filter); })
      .sort(function (a, b) { return a.cy - b.cy || a.cx - b.cx; });
    var tb = h('tbody');
    list.forEach(function (c, i) {
      var catSel = select(L.CATS, c.cat, { class: 'type-sel', 'aria-label': '구분', onchange: function (e) { updateCand(c, { cat: e.target.value }); render(); } });
      catSel.addEventListener('click', function (e) { e.stopPropagation(); });
      tb.appendChild(h('tr', { class: 'clickable' + (ui.sel === c.id ? ' sel' : '') + (c.check ? ' row-warn' : ''), tabindex: '0', onclick: function () { ui.sel = c.id; render(); window.scrollTo(0, 0); } },
        h('td', { class: 'num' }, String(i + 1)), h('td', null, catSel), h('td', { class: 'nw' }, c.kind || ''), h('td', { class: 'wide' }, c.raw),
        h('td', { class: 'num' }, fmt(c.nominal)), h('td', { class: 'num' }, fmt(c.upper)), h('td', { class: 'num' }, fmt(c.lower)), h('td', null, c.unit || ''),
        h('td', null, c.check ? h('span', { class: 'badge warn' }, '확인 필요') : (c.src === 'manual' ? h('span', { class: 'badge muted' }, '직접') : '')),
        h('td', { class: 'note' }, c.note || ''),
        h('td', null, h('button', { type: 'button', class: 'btn btn-sm2 btn-danger', onclick: function (e) { e.stopPropagation(); removeCand(c); } }, '빼기'))));
    });
    var filt = h('div', { class: 'chips' }, [['all', '전체'], ['check', '확인 필요']].concat(L.CATS.map(function (c) { return [c, c]; })).map(function (f) {
      return h('button', { type: 'button', class: 'chip-btn', 'aria-pressed': String(ui.filter === f[0]), style: ui.filter === f[0] ? 'border-color:var(--primary);background:#eaf1fa' : null, onclick: function () { ui.filter = f[0]; render(); } }, f[1]);
    }));
    var d = doc();
    return h('section', { class: 'card' }, h('div', { class: 'card-head' }, h('h2', null, '이 쪽의 검사 대상 후보 ' + pc.cands.length + '개'), filt),
      list.length ? h('div', { class: 'table-wrap' }, h('table', { class: 'list' }, h('thead', null, h('tr', null, ['#', '구분', '종류', '원문', '공칭', '상한', '하한', '단위', '', '메모', ''].map(function (t) { return h('th', null, t); }))), tb))
        : h('p', { class: 'note' }, '조건에 맞는 후보가 없습니다.'),
      pc.removed.length ? h('details', { class: 'kw-box' }, h('summary', null, '뺀 후보 ' + pc.removed.length + '개'), h('ul', { class: 'list-plain' }, pc.removed.map(function (c) {
        return h('li', null, c.raw + ' (' + c.cat + ') ', h('button', { type: 'button', class: 'btn btn-sm2', onclick: function () { delete d.removed[c.id]; save(); render(); } }, '되살리기'));
      }))) : null,
      pc.excluded.length ? h('details', { class: 'kw-box' }, h('summary', null, '규칙이 제외한 글자 ' + pc.excluded.length + '개(표제란 · 테두리 · 주기 등)'), h('ul', { class: 'list-plain' }, pc.excluded.map(function (x) {
        return h('li', null, x.raw + ' — ' + x.reason + ' ', h('button', { type: 'button', class: 'btn btn-sm2', onclick: function () { addFromToken(x.token); } }, '후보로 넣기'));
      }))) : null);
  }
  function aiCard() {
    var cs = allCands().filter(function (c) { return c.check; }).slice(0, 40);
    var ids = cs.map(function (c) { return c.id; });
    var prompt = L.aiPrompt(cs.map(function (c) { return Object.assign({}, c, { near: L.nearText(c, pg(c.page).tokens) }); }), { name: cur.name });
    var ans = h('textarea', { class: 'prompt-box answer', 'aria-label': 'AI 답 붙여 넣기', placeholder: 'AI 의 답(표)을 여기에 붙여 넣으세요' }); ans.value = ui.aiText;
    ans.addEventListener('input', function () { ui.aiText = ans.value; });
    var pre = null;
    if (ui.aiParsed) {
      pre = ui.aiParsed.length ? h('div', { class: 'table-wrap' }, h('table', { class: 'list' }, h('thead', null, h('tr', null, ['ID', '검사 대상', '구분', '공칭', '상한', '하한', '근거'].map(function (t) { return h('th', null, t); }))),
        h('tbody', null, ui.aiParsed.map(function (a) { return h('tr', null, h('td', { class: 'mono' }, a.id), h('td', null, a.include ? '예' : '아니오'), h('td', null, a.cat), h('td', { class: 'num' }, fmt(a.nominal)), h('td', { class: 'num' }, fmt(a.upper)), h('td', { class: 'num' }, fmt(a.lower)), h('td', null, a.reason)); }))))
        : h('p', { class: 'alert warn' }, '답에서 후보 ID 가 있는 줄을 찾지 못했습니다. 표 형식(| ID | … |) 그대로인지 확인해 주세요.');
    }
    function readAns(text) { ui.aiText = text; ui.aiParsed = L.parseAiAnswer(text, ids); render(); }
    return h('section', { class: 'card' }, h('h2', null, '애매한 후보 AI 에게 묻기 (선택)'),
      cs.length ? [h('p', { class: 'note' }, '「확인 필요」 후보 ' + cs.length + '개를 주변 글자와 함께 묻는 프롬프트입니다. 도면 글자 일부가 들어가니 회사가 허용한 AI 에만 붙여 넣으세요. 답은 사람이 확인한 뒤 반영됩니다.'),
        h('textarea', { class: 'prompt-box', readonly: true, 'aria-label': 'AI 프롬프트' }, prompt),
        h('div', { class: 'btn-row', style: 'margin:8px 0' }, h('button', { type: 'button', class: 'btn', onclick: function () { copyText(prompt); } }, '프롬프트 복사'),
          window.AIPanel.sendButton('설정한 AI 서버로 보내기', function () { return prompt; }, readAns, { toast: toast, settingsHref: '#/settings' })),
        ans, h('div', { class: 'btn-row', style: 'margin-top:8px' }, h('button', { type: 'button', class: 'btn', onclick: function () { readAns(ans.value); } }, '답 읽기'),
          ui.aiParsed && ui.aiParsed.length ? h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
            var byId = {}; allCands().forEach(function (c) { byId[c.id] = c; });
            ui.aiParsed.forEach(function (a) {
              var c = byId[a.id]; if (!c) return;
              if (!a.include) { var dd = doc(); if (isManual(c)) dd.added[c.page] = (dd.added[c.page] || []).filter(function (x) { return x.id !== c.id; }); else dd.removed[c.id] = 1; return; }
              updateCand(c, { cat: L.CATS.indexOf(a.cat) >= 0 ? a.cat : c.cat, nominal: a.nominal != null ? a.nominal : c.nominal, upper: a.upper, lower: a.lower, check: false, note: 'AI 판단 반영: ' + a.reason });
            });
            save(); toast(ui.aiParsed.length + '건을 반영했습니다.'); ui.aiParsed = null; ui.aiText = ''; render();
          } }, '후보에 반영') : null), pre]
        : h('p', { class: 'note' }, '지금은 「확인 필요」 후보가 없습니다.'));
  }

  // ── 목표 1 · 풍선 배치 ────────────────────────────────────
  function balloonOptions() {
    var s = settings();
    function num(k, label, min, max, step, hint) {
      return field(label, h('input', { type: 'number', min: min, max: max, step: step || 1, value: s[k], onchange: function (e) { var o = {}; o[k] = +e.target.value; db.settings = L.cleanSettings(Object.assign({}, s, o)); save(); render(); } }), hint);
    }
    return h('section', { class: 'card' }, h('h2', null, '번호 매기는 방식'),
      h('div', { class: 'form-grid three' },
        field('순서', select([['rows', '줄 — 위 → 아래, 줄 안에서 왼 → 오'], ['cols', '열 — 왼 → 오, 열 안에서 위 → 아래'], ['zone', '구역 — A1, A2 … (도면 구역 칸)']], s.order, { onchange: function (e) { db.settings.order = e.target.value; save(); render(); } })),
        s.order === 'zone' ? num('zoneCols', '구역 열 수(1, 2, 3 …)', 1, 26) : num('bands', '띠 수(쪽을 몇 줄로 나눌지)', 1, 30, 1, '띠가 많을수록 위아래로 촘촘히 끊어 읽습니다'),
        s.order === 'zone' ? num('zoneRows', '구역 행 수(A, B, C …)', 1, 20) : num('startNo', '시작 번호', 1, 9999),
        num('balloonR', '풍선 반지름(mm)', 1.5, 12, 0.5)),
      h('div', { class: 'btn-row', style: 'margin-top:12px' },
        h('button', { type: 'button', class: 'btn', onclick: function () { doc().nums = {}; save(); render(); toast('번호를 순서대로 다시 매겼습니다.'); } }, '번호 다시 매기기'),
        h('button', { type: 'button', class: 'btn', onclick: function () { doc().bpos = {}; save(); render(); toast('옮긴 풍선을 자동 자리로 되돌렸습니다.'); } }, '풍선 자리 다시 잡기')));
  }
  function balloonG(b, cls, extra) {
    var g = sv('g', Object.assign({ class: 'bl ' + (cls || ''), transform: 'translate(' + b.x + ' ' + b.y + ')' }, extra || {}));
    g.appendChild(sv('circle', { cx: 0, cy: 0, r: b.r }));
    g.appendChild(sv('text', { x: 0, y: 0.5, 'font-size': Math.max(5, b.r * (String(b.no).length > 2 ? 0.8 : 1.05)) }, String(b.no)));
    return g;
  }
  function renderBalloon(main) {
    if (needDoc(main, '목표 1 · 풍선 배치', '풍선 없는 도면의 검사 대상에 번호 풍선을 자동으로 답니다.')) return;
    var items = g1(), d = doc();
    var here = items.filter(function (x) { return x.page === ui.page; }), ovl = items.filter(function (x) { return x.balloon && x.balloon.overlap; }).length;
    main.appendChild(pageHead('목표 1 · 풍선 배치', '검사 대상 후보에 순서대로 번호를 매기고, 글자 · 다른 풍선 · 도면 선과 겹치지 않는 자리에 풍선을 놓습니다. 풍선은 끌어서 옮길 수 있습니다.',
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: exportBalloonPdf }, '풍선 PDF 내려받기'),
        h('button', { type: 'button', class: 'btn', onclick: function () { exportPng(items); } }, '이 쪽 PNG'),
        h('button', { type: 'button', class: 'btn', onclick: function () { exportExcel(items, 1); } }, '엑셀 내려받기'))));
    if (cur.ballooned) main.appendChild(h('div', { class: 'alert warn' }, '이 도면에는 풍선이 이미 있습니다. 풍선별 정보를 뽑으려면 ', h('a', { href: '#/extract' }, '목표 2'), ' 를 쓰세요(여기서는 원래 풍선 숫자를 빼고 새로 답니다).'));
    main.appendChild(h('div', { class: 'kpis' }, kpi('풍선', String(items.length), '개'), kpi('이 쪽', String(here.length), '개'), kpi('확인 필요', String(items.filter(function (x) { return x.cand.check; }).length), '개'), kpi('겹침 남음', String(ovl), '개', ovl ? 'bad' : '')));
    main.appendChild(balloonOptions());
    var selIt = ui.selB ? items.filter(function (x) { return x.cand.id === ui.selB; })[0] : null;
    var v = viewer({
      draw: function (svg) {
        drawCands(svg, here.map(function (x) { return x.cand; }));
        here.forEach(function (x) {
          var b = x.balloon; if (!b) return;
          if (b.leader) svg.appendChild(sv('line', { class: 'lead', x1: b.leader[0], y1: b.leader[1], x2: b.leader[2], y2: b.leader[3] }));
          svg.appendChild(balloonG({ x: b.x, y: b.y, r: b.r, no: x.no }, (b.overlap ? 'ovl' : '') + (ui.selB === x.cand.id ? ' sel' : ''), { 'data-b': x.cand.id }));
        });
      },
      onDown: function (e, pt, svg) {
        var g = e.target.closest && e.target.closest('[data-b]');
        if (!g) { var r = e.target.getAttribute && e.target.getAttribute('data-id'); if (r) { ui.selB = r; render(); } return; }
        e.preventDefault();
        var id = g.getAttribute('data-b'), it = here.filter(function (x) { return x.cand.id === id; })[0], moved = false, start = pt, last = [it.balloon.x, it.balloon.y];
        function mv(ev) { var q = ptOf(ev, svg); if (L.dist(q[0], q[1], start[0], start[1]) > 2) moved = true; last = [it.balloon.x + q[0] - start[0], it.balloon.y + q[1] - start[1]]; g.setAttribute('transform', 'translate(' + last[0] + ' ' + last[1] + ')'); }
        function up() {
          window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up);
          if (moved) { d.bpos[id] = { x: L.round(last[0], 1), y: L.round(last[1], 1) }; save(); }
          ui.selB = id; render();
        }
        window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up);
      }
    });
    var panel;
    if (selIt) {
      var noIn = h('input', { type: 'number', min: 1, value: selIt.no, class: 'num-in', 'aria-label': '번호' });
      panel = h('div', { class: 'edit-form' }, h('p', null, catDot(selIt.cand.cat), h('strong', null, selIt.no + '번 · ' + selIt.cand.raw)), h('p', { class: 'note' }, selIt.cand.cat + (selIt.cand.kind ? ' · ' + selIt.cand.kind : '') + ' · 구역 ' + selIt.zone),
        field('번호', noIn, '같은 번호가 있으면 그 풍선의 번호와 바꿉니다'),
        h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
            var n = Math.round(+noIn.value); if (!(n >= 1)) return;
            var other = items.filter(function (x) { return x.no === n && x.cand.id !== selIt.cand.id; })[0];
            items.forEach(function (x) { d.nums[x.cand.id] = x.no; });
            if (other) d.nums[other.cand.id] = selIt.no;
            d.nums[selIt.cand.id] = n; save(); render();
          } }, '번호 저장'),
          d.bpos[selIt.cand.id] ? h('button', { type: 'button', class: 'btn', onclick: function () { delete d.bpos[selIt.cand.id]; save(); render(); } }, '자동 자리로') : null),
        selIt.note ? h('p', { class: 'alert warn' }, selIt.note) : null);
    } else panel = h('p', { class: 'note' }, '풍선을 누르면 번호를 바꿀 수 있고, 끌면 자리를 옮깁니다. 주황 점선 풍선은 빈자리를 못 찾아 다른 글자와 겹친 것입니다. 후보 자체를 고치려면 「도면 열기·인식」에서 합니다.');
    main.appendChild(h('div', { class: 'work' }, h('section', { class: 'card' }, v), h('aside', { class: 'card side-panel' }, h('h2', null, selIt ? '풍선 ' + selIt.no : '도움말'), panel)));
    main.appendChild(resultTable(items, 1));
  }
  function resultTable(items, goal) {
    var tb = h('tbody');
    items.forEach(function (x) {
      var c = x.cand || {}, shapeCell = h('td', null);
      var crop = x.crop || c.crop;
      if (crop && (c.shape || !x.cand)) {
        var img = h('img', { class: 'thumb', alt: x.no + '번 형상 그림' });
        cropUrl(x.page, crop, function (cv) { img.src = cv.toDataURL('image/png'); });
        add(shapeCell, img, h('button', { type: 'button', class: 'btn btn-sm2', onclick: function (e) { e.stopPropagation(); cropUrl(x.page, crop, function (cv) { cv.toBlob(function (b) { download(safeName(cur.name) + '_' + x.no + '번_형상.png', b); }); }); } }, '그림 받기'));
      }
      tb.appendChild(h('tr', { class: 'clickable' + ((x.check || c.check) ? ' row-warn' : '') + ((goal === 1 ? ui.selB === c.id : ui.selB === x.page + ':' + x.no) ? ' sel' : ''), tabindex: '0',
        onclick: function () { ui.page = x.page; ui.selB = goal === 1 ? c.id : x.page + ':' + x.no; render(); window.scrollTo(0, 0); } },
        h('td', { class: 'num' }, String(x.no)), h('td', { class: 'num' }, String(x.page)), h('td', null, x.zone || ''), h('td', { class: 'nw' }, c.cat ? [catDot(c.cat), c.cat] : h('span', { class: 'badge warn' }, '연결 못 함')),
        h('td', { class: 'nw' }, c.kind || ''), h('td', { class: 'wide' }, c.raw || ''), h('td', { class: 'num' }, fmt(c.nominal)), h('td', { class: 'num' }, fmt(c.upper)), h('td', { class: 'num' }, fmt(c.lower)),
        h('td', null, x.method || ''), h('td', null, (x.check || c.check) ? h('span', { class: 'badge warn' }, '확인 필요') : ''), shapeCell, h('td', { class: 'note' }, [c.note, x.note].filter(Boolean).join(' · '))));
    });
    return h('section', { class: 'card' }, h('h2', null, (goal === 1 ? '풍선 번호별 검사 항목 ' : '풍선별 정보 ') + items.length + '개'),
      items.length ? h('div', { class: 'table-wrap' }, h('table', { class: 'list' }, h('thead', null, h('tr', null, ['번호', '쪽', '구역', '구분', '종류', '원문', '공칭', '상한', '하한', '연결', '', '형상', '메모'].map(function (t) { return h('th', null, t); }))), tb))
        : h('p', { class: 'note' }, goal === 1 ? '검사 대상 후보가 없습니다.' : '풍선을 찾지 못했습니다.'));
  }

  // ── 목표 2 · 풍선 읽기 ────────────────────────────────────
  function renderExtract(main) {
    if (needDoc(main, '목표 2 · 풍선 읽기', '풍선이 있는 도면에서 풍선별 치수·공차·용접 정보를 뽑습니다.')) return;
    var items = g2(), d = doc(), s = settings();
    var here = items.filter(function (x) { return x.page === ui.page; });
    var pc = pageCands(ui.page);
    var miss = items.filter(function (x) { return !x.cand; }).length;
    main.appendChild(pageHead('목표 2 · 풍선 읽기', '원 안의 숫자를 풍선으로 읽고, 지시선 끝이나 가장 가까운 치수 · 공차 · 기하공차 · 용접 표기와 연결합니다. 글자 없이 형상만 가리키면 그 둘레를 잘라 「형상」 그림으로 남깁니다.',
      h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { exportExcel(items, 2); } }, '엑셀 내려받기'), h('a', { class: 'btn', href: '#/accuracy' }, '일치율 재기'))));
    main.appendChild(h('div', { class: 'kpis' }, kpi('찾은 풍선', String(items.length), '개'), kpi('연결됨', String(items.length - miss), '개'), kpi('연결 못 함', String(miss), '개', miss ? 'bad' : ''),
      kpi('확인 필요', String(items.filter(function (x) { return x.check; }).length), '개')));
    if (!items.length) main.appendChild(h('div', { class: 'alert warn' }, '원 안의 숫자(풍선)를 찾지 못했습니다. 풍선이 원 없이 숫자만 있거나, 원이 곡선이 아닌 다각형이면 아래 설정을 켜 보세요. 풍선이 없는 도면이면 ', h('a', { href: '#/balloon' }, '목표 1'), ' 을 쓰세요.'));
    main.appendChild(h('section', { class: 'card' }, h('label', { class: 'check' }, h('input', { type: 'checkbox', checked: s.isolatedBalloons, onchange: function (e) { db.settings.isolatedBalloons = e.target.checked; save(); render(); } }),
      h('span', null, '원 없이 홀로 있는 숫자도 풍선으로 보기(확인 필요로 표시 — 치수 숫자를 잘못 잡을 수 있음)')),
      h('p', { class: 'note' }, '연결 거리: 지시선 끝에서 ' + s.linkMaxMm + 'mm · 지시선이 없으면 풍선 둘레에서 ' + s.nearMaxMm + 'mm 안의 글자. 바꾸려면 설정에서. 풍선을 누르고 도면의 네모를 누르면 연결을 직접 바꿉니다.')));
    var selIt = ui.selB ? here.filter(function (x) { return x.page + ':' + x.no === ui.selB; })[0] : null;
    var v = viewer({
      draw: function (svg) {
        drawCands(svg, pc.cands);
        here.forEach(function (x) {
          if (x.leader) svg.appendChild(sv('polyline', { class: 'lead', points: x.leader.map(function (q) { return q.join(','); }).join(' ') }));
          if (x.cand) svg.appendChild(sv('line', { class: 'link', x1: x.balloon.x, y1: x.balloon.y, x2: x.cand.cx, y2: x.cand.cy }));
          if (!x.cand && x.crop) svg.appendChild(rectEl(x.crop, 'crop'));
          svg.appendChild(balloonG({ x: x.balloon.x, y: x.balloon.y, r: x.balloon.r, no: x.no }, (x.cand ? '' : 'miss') + (ui.selB === x.page + ':' + x.no ? ' sel' : ''), { 'data-b': x.page + ':' + x.no }));
        });
      },
      onDown: function (e) {
        var g = e.target.closest && e.target.closest('[data-b]');
        if (g) { ui.selB = g.getAttribute('data-b'); render(); return; }
        var id = e.target.getAttribute && e.target.getAttribute('data-id');
        if (id && ui.selB) { d.linkOv[ui.selB] = id; save(); toast(ui.selB.split(':')[1] + '번 풍선의 연결을 바꿨습니다.'); render(); }
      }
    });
    var panel = selIt ? h('div', null, h('p', null, h('strong', null, selIt.no + '번 풍선 '), selIt.cand ? [catDot(selIt.cand.cat), selIt.cand.raw] : h('span', { class: 'badge warn' }, '연결 못 함')),
      h('p', { class: 'note' }, '연결 방식: ' + selIt.method + (selIt.note ? ' · ' + selIt.note : '')),
      h('p', { class: 'note' }, '도면에서 다른 네모를 누르면 이 풍선의 연결을 그 글자로 바꿉니다.'),
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn', onclick: function () { d.linkOv[ui.selB] = 'none'; save(); render(); } }, '연결 비우기'),
        d.linkOv[ui.selB] ? h('button', { type: 'button', class: 'btn', onclick: function () { delete d.linkOv[ui.selB]; save(); render(); } }, '자동 연결로') : null,
        selIt.cand ? h('a', { class: 'btn', href: '#/drawing', onclick: function () { setTimeout(function () { ui.sel = selIt.cand.id; render(); }, 0); } }, '값 고치기') : null))
      : h('p', { class: 'note' }, '풍선을 누르면 연결을 확인 · 변경합니다. 빨간 점선 풍선은 가리키는 글자를 못 찾은 것이고, 빨간 점선 네모는 형상 그림으로 잘라 둘 영역입니다.');
    main.appendChild(h('div', { class: 'work' }, h('section', { class: 'card' }, v), h('aside', { class: 'card side-panel' }, h('h2', null, selIt ? '풍선 ' + selIt.no : '도움말'), panel)));
    main.appendChild(resultTable(items, 2));
  }

  // ── 내보내기 ──────────────────────────────────────────────
  function infoSheet(goal, items) {
    var s = settings();
    return [['항목', '값'], ['도면 파일', cur.name], ['쪽 수', cur.pages.length], ['작업', goal === 1 ? '목표 1 — 풍선 자동 배치(풍선 없는 도면)' : '목표 2 — 풍선별 정보 추출(풍선 있는 도면)'],
      ['만든 시각', nowText()], ['인식 방식', '1단계 규칙 인식 — 벡터 PDF 글자 층 + 원·지시선·용접 삼각형 기하. 스캔 그림은 읽지 못함'],
      ['항목 수', items.length], ['확인 필요', items.filter(function (x) { return x.check || (x.cand && x.cand.check); }).length], ['용접 형상', items.filter(function (x) { return x.cand && x.cand.shape; }).length],
      ['번호 순서', goal === 1 ? ({ rows: '줄', cols: '열', zone: '구역' }[s.order]) : '도면의 풍선 번호'], ['좌표', 'mm, 쪽 왼쪽 위가 (0, 0)'],
      ['주의', '이 표는 자동 인식 결과 초안입니다. 「확인 필요」 행과 연결 못 함 행은 반드시 사람이 확인해 주세요.']].concat(cur.sample ? [['예시', '가상 도면 — 실제 도면이 아님']] : []);
  }
  function exportExcel(items, goal) {
    var rows = [L.ROW_HEAD].concat(items.map(L.itemRow));
    var sheets = { '검사항목': rows, '도면 정보': infoSheet(goal, items) };
    var ex = [['쪽', '글자', '제외 이유']];
    cur.pages.forEach(function (p) { pageCands(p.page).excluded.forEach(function (x) { ex.push([p.page, x.raw, x.reason]); }); });
    sheets['제외된 글자'] = ex;
    writeXlsx(safeName(cur.name) + (goal === 1 ? '_풍선배치_' : '_풍선별정보_') + stamp() + '.xlsx', sheets,
      { '검사항목': [6, 5, 6, 9, 11, 18, 9, 9, 9, 10, 10, 6, 6, 8, 8, 8, 10, 10, 10, 9, 30, 30], '도면 정보': [14, 70], '제외된 글자': [5, 30, 14] });
  }
  function drawBalloonsOn(ctx, k, list) {
    list.forEach(function (x) {
      var b = x.balloon; if (!b) return;
      ctx.strokeStyle = '#c81e1e'; ctx.fillStyle = '#fff'; ctx.lineWidth = Math.max(1, 0.9 * k);
      if (b.leader) { ctx.beginPath(); ctx.moveTo(b.leader[0] * k, b.leader[1] * k); ctx.lineTo(b.leader[2] * k, b.leader[3] * k); ctx.stroke(); }
      ctx.beginPath(); ctx.arc(b.x * k, b.y * k, b.r * k, 0, Math.PI * 2); ctx.fill(); ctx.stroke();
      var fs = Math.max(5, b.r * (String(x.no).length > 2 ? 0.8 : 1.05)) * k;
      ctx.fillStyle = '#c81e1e'; ctx.font = 'bold ' + fs + 'px Helvetica, Arial, sans-serif'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText(String(x.no), b.x * k, b.y * k + fs * 0.04);
    });
  }
  function exportPng(items) {
    var p = pg(), W = Math.min(5000, Math.round(p.w * 3));
    pageCanvas(ui.page, W, function (c) {
      var o = document.createElement('canvas'); o.width = c.width; o.height = c.height;
      var ctx = o.getContext('2d'); ctx.drawImage(c, 0, 0);
      drawBalloonsOn(ctx, c.width / p.w, items.filter(function (x) { return x.page === ui.page; }));
      o.toBlob(function (b) { download(safeName(cur.name) + '_' + ui.page + '쪽_풍선.png', b); });
    });
  }
  function exportBalloonPdf() {
    var items = g1(), out = [], i = 0;
    toast('풍선 PDF 를 만드는 중입니다…');
    (function next() {
      if (i >= cur.pages.length) {
        var bytes = L.buildBalloonPdf(out);
        download(safeName(cur.name) + '_풍선_' + stamp() + '.pdf', new Blob([bytes], { type: 'application/pdf' }));
        toast('풍선 PDF 를 내려받았습니다(' + Math.round(bytes.length / 1024) + 'KB). 원본 쪽은 그림 + 보이지 않는 글자 층, 풍선은 벡터입니다.');
        return;
      }
      var p = cur.pages[i], W = Math.min(4200, Math.round(p.w * 2.6));
      pageCanvas(p.page, W, function (c) {
        var jpg = b64bytes(c.toDataURL('image/jpeg', 0.9).split(',')[1]);
        out.push({ w: p.w, h: p.h, jpeg: jpg, imgW: c.width, imgH: c.height, texts: p.items, segs: p.segs, tris: p.tris,
          balloons: items.filter(function (x) { return x.page === p.page && x.balloon; }).map(function (x) { return { x: x.balloon.x, y: x.balloon.y, r: x.balloon.r, no: x.no, leader: x.balloon.leader }; }) });
        i++; next();
      });
    })();
  }

  // ── 일치율 측정 ───────────────────────────────────────────
  function truthTemplate() {
    writeXlsx('정답_양식.xlsx', { '검사항목': [['번호', '구분', '종류', '원문', '공칭값', '상한공차', '하한공차', '단위', '쪽'], [1, '치수', '길이', '120', 120, '', '', 'mm', 1], [2, '공차', '길이', '45±0.1', 45, 0.1, -0.1, 'mm', 1], [3, '기하공차', '직각도', '⊥ 0.1 A', '', 0.1, 0, 'mm', 1], [4, '용접', '필릿', '6', 6, '', '', 'mm', 1], [5, '용접 형상', '', '', '', '', '', '', 1]],
      '안내': [['정답 엑셀 작성 안내'], ['사람이 풍선을 단 결과(또는 지금 쓰시는 검사 기준서)를 이 칸에 맞춰 적어 주세요.'], ['번호 = 풍선 번호. 구분 = 치수 · 공차 · 기하공차 · 용접 · 용접 형상 · 기타.'], ['공칭값 · 상한공차 · 하한공차가 비어 있으면 원문을 읽어 채웁니다(예: 45±0.1).'], ['열 이름이 조금 달라도(No · Spec · 기준값 · 상한 · 하한 등) 알아봅니다.']] },
      { '검사항목': [6, 10, 10, 16, 9, 9, 9, 6, 5], '안내': [80] });
  }
  function readTruthFile(file, cb) {
    var r = new FileReader();
    r.onload = function () {
      try {
        var wb = XLSX.read(new Uint8Array(r.result), { type: 'array' }), best = null;
        wb.SheetNames.forEach(function (n) {
          var t = L.parseTruth(XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, defval: '' }));
          if (t.items.length && (!best || t.items.length > best.items.length)) best = Object.assign(t, { sheet: n });
        });
        if (!best) { toast('정답 표의 머리행(번호 · 구분 · 원문 또는 공칭값)을 찾지 못했습니다. 「정답 양식」을 받아 맞춰 주세요.', true); return; }
        cb(best);
      } catch (e) { toast('엑셀을 읽지 못했습니다: ' + e.message, true); }
    };
    r.readAsArrayBuffer(file);
  }
  function renderAccuracy(main) {
    if (needDoc(main, '일치율 측정', '정답 엑셀과 도구 결과를 대조합니다.')) return;
    var d = doc(), s = settings();
    var mode = ui.accMode || (cur.ballooned ? 'number' : 'value');
    main.appendChild(pageHead('일치율 측정', '사람이 만든 정답(풍선 번호 · 구분 · 값)과 도구의 결과를 대조해, 지금 몇 % 인지 숫자로 봅니다. 규칙을 고칠 때마다 같은 도면으로 다시 재서 기록합니다.',
      h('button', { type: 'button', class: 'btn', onclick: truthTemplate }, '정답 양식 내려받기')));
    var tIn = h('input', { type: 'file', accept: '.xlsx,.xls,.csv', 'aria-label': '정답 엑셀 고르기', onchange: function (e) {
      var f = e.target.files[0]; if (!f) return;
      readTruthFile(f, function (t) { d.truth = { name: f.name, sheet: t.sheet, items: t.items, warnings: t.warnings, at: nowText() }; save(); render(); toast('정답 ' + t.items.length + '개를 읽었습니다.'); });
    } });
    var truthCard = h('section', { class: 'card' }, h('h2', null, '1. 정답'),
      d.truth ? h('p', null, h('strong', null, d.truth.name), ' · ' + d.truth.items.length + '개 · ' + d.truth.at, d.truth.warnings && d.truth.warnings.length ? h('span', { class: 'badge warn' }, ' 주의 ' + d.truth.warnings.length) : null)
        : h('p', { class: 'note' }, '아직 정답이 없습니다. 같은 도면을 사람이 정리한 엑셀(번호 · 구분 · 원문 또는 공칭값 · 상한 · 하한)을 올려 주세요.'),
      h('div', { class: 'btn-row' }, h('label', { class: 'btn file-btn' }, d.truth ? '정답 바꾸기' : '정답 엑셀 올리기', tIn),
        cur.sample ? h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { d.truth = { name: sampleMeta(cur.sample).name + '_정답.xlsx(예시)', items: SM.truth[cur.sample], at: nowText(), warnings: [] }; save(); render(); } }, '예시 정답 불러오기') : null,
        d.truth ? h('button', { type: 'button', class: 'btn btn-danger', onclick: function () { d.truth = null; save(); render(); } }, '정답 지우기') : null));
    main.appendChild(truthCard);
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '2. 무엇을 잴까'), h('div', { class: 'seg', role: 'group', 'aria-label': '측정 대상' },
      h('button', { type: 'button', 'aria-pressed': String(mode === 'value'), onclick: function () { ui.accMode = 'value'; render(); } }, '목표 1 — 자동 배치(값으로 짝)'),
      h('button', { type: 'button', 'aria-pressed': String(mode === 'number'), onclick: function () { ui.accMode = 'number'; render(); } }, '목표 2 — 풍선 읽기(번호로 짝)')),
      h('p', { class: 'note', style: 'margin-top:8px' }, mode === 'value' ? '풍선 없는 도면에서 도구가 찾은 항목을 정답과 값(구분 · 공칭 · 상하한)으로 짝짓습니다. 도구의 번호 순서는 사람과 다를 수 있어 「번호 같음」은 따로 셉니다.'
        : '풍선 있는 도면에서 같은 번호끼리 짝짓고, 그 풍선에 연결된 값이 정답과 같은지 봅니다.')));
    if (!d.truth) return;
    var items = mode === 'value' ? g1() : g2();
    var sc = L.scoreResults(items.map(L.toPred), d.truth.items, mode);
    var tgt = s.target / 100, ok = sc.valueRate >= tgt;
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '3. 결과'),
      h('div', { class: 'kpis' }, kpi('값 일치율(정답 중 값까지 맞음)', pct(sc.valueRate), '', ok ? 'good' : 'bad'), kpi('정밀도', pct(sc.precision)), kpi('재현율', pct(sc.recall)), kpi('F1', pct(sc.f1)),
        mode === 'value' ? kpi('번호 같음', pct(sc.numberRate)) : null, kpi('정답 · 찾음 · 맞게 찾음', sc.nTruth + ' · ' + sc.nPred + ' · ' + sc.tp, '', 'sm')),
      h('div', { class: 'meter', role: 'img', 'aria-label': '값 일치율 ' + pct(sc.valueRate) + ', 목표 ' + s.target + '%' }, h('span', { style: 'width:' + (sc.valueRate * 100) + '%;' + (ok ? 'background:#1f9d55' : '') }), h('i', { style: 'left:' + s.target + '%' })),
      h('p', { class: 'note', style: 'margin-top:6px' }, ok ? '목표 ' + s.target + '% 이상입니다.' : '목표 ' + s.target + '% 까지 ' + Math.max(0, Math.ceil(tgt * sc.nTruth) - sc.exact) + '개 더 맞아야 합니다.'),
      h('details', { class: 'kw-box' }, h('summary', null, '숫자 읽는 법'), h('ul', { class: 'list-plain' },
        h('li', null, '정밀도 = 도구가 찾은 것 중 정답에 있는 비율(엉뚱한 것을 얼마나 덜 잡았나)'), h('li', null, '재현율 = 정답 중 도구가 찾은 비율(얼마나 덜 놓쳤나)'),
        h('li', null, 'F1 = 정밀도와 재현율의 조화평균'), h('li', null, '값 일치율 = 정답 중 구분 · 공칭값 · 상한 · 하한까지 모두 같은 비율 — 「인식 일치율 85%」 는 이 숫자로 봅니다'))),
      h('div', { class: 'btn-row', style: 'margin-top:10px' }, h('input', { id: 'histMemo', placeholder: '메모(예: 표제란 낱말 추가 후)', style: 'min-height:44px;padding:8px 10px;border:1px solid var(--line);border-radius:8px;flex:1 1 220px', 'aria-label': '기록 메모' }),
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
          db.history.push({ at: nowText(), doc: cur.name, mode: mode, nTruth: sc.nTruth, nPred: sc.nPred, precision: sc.precision, recall: sc.recall, f1: sc.f1, valueRate: sc.valueRate, memo: (document.getElementById('histMemo') || {}).value || '' });
          save(); render(); toast('측정 기록에 남겼습니다.');
        } }, '측정 기록에 남기기'),
        h('button', { type: 'button', class: 'btn', onclick: function () { exportScore(sc, mode); } }, '대조 결과 엑셀'))));
    var bc = h('tbody');
    Object.keys(sc.byCat).forEach(function (c) { var b = sc.byCat[c]; bc.appendChild(h('tr', null, h('td', null, catDot(c === '용접 형상' ? '용접' : c), c), h('td', { class: 'num' }, String(b.truth)), h('td', { class: 'num' }, String(b.found)), h('td', { class: 'num' }, String(b.exact)), h('td', { class: 'num' }, pct(b.truth ? b.exact / b.truth : 0)))); });
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '구분별'), h('div', { class: 'table-wrap' }, h('table', { class: 'list' }, h('thead', null, h('tr', null, ['구분', '정답', '찾음', '값까지 맞음', '값 일치율'].map(function (t) { return h('th', null, t); }))), bc))));
    function lst(title, arr, fn) { return h('div', null, h('h3', null, title + ' ' + arr.length + '개'), arr.length ? h('ul', { class: 'list-plain' }, arr.map(function (x) { return h('li', null, fn(x)); })) : h('p', { class: 'note' }, '없음')); }
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '틀린 것 살펴보기'), h('div', { class: 'grid-2' },
      lst('놓친 정답', sc.misses, function (m) { return m.no + '번 · ' + m.text; }),
      lst('값이 다른 것', sc.wrong, function (m) { return m.no + '번 · 정답 「' + m.truth + '」 / 도구' + (m.predNo != null && m.predNo !== m.no ? '(' + m.predNo + '번)' : '') + ' 「' + m.pred + '」'; })),
      lst('정답에 없는데 도구가 찾은 것', sc.extra, function (m) { return (m.no != null ? m.no + '번 · ' : '') + m.text; }),
      h('p', { class: 'note' }, '놓친 것은 「도면 열기·인식」에서 글자 눌러 추가 · 제외 영역 확인으로, 더 찾은 것은 제외 영역 · 표제란 낱말 설정으로 줄입니다. 반복되는 틀림은 규칙을 고칠 근거가 됩니다.')));
    var hist = db.history.filter(function (r) { return r.doc === cur.name; });
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '이 도면의 측정 기록 ' + hist.length + '건'),
      hist.length ? [trendChart(hist, s.target), h('div', { class: 'table-wrap', style: 'margin-top:10px' }, h('table', { class: 'list' }, h('thead', null, h('tr', null, ['시각', '대상', '값 일치율', '정밀도', '재현율', 'F1', '메모', ''].map(function (t) { return h('th', null, t); }))),
        h('tbody', null, hist.map(function (r) { return h('tr', null, h('td', { class: 'nw' }, r.at), h('td', null, r.mode === 'value' ? '목표 1' : '목표 2'), h('td', { class: 'num' }, pct(r.valueRate)), h('td', { class: 'num' }, pct(r.precision)), h('td', { class: 'num' }, pct(r.recall)), h('td', { class: 'num' }, pct(r.f1)), h('td', null, r.memo || ''),
          h('td', null, h('button', { type: 'button', class: 'btn btn-sm2 btn-danger', onclick: function () { db.history.splice(db.history.indexOf(r), 1); save(); render(); } }, '지우기'))); }))))]
        : h('p', { class: 'note' }, '아직 없습니다.')));
  }
  function exportScore(sc, mode) {
    var sum = [['항목', '값'], ['도면', cur.name], ['정답 파일', doc().truth.name], ['대상', mode === 'value' ? '목표 1(값으로 짝)' : '목표 2(번호로 짝)'], ['측정 시각', nowText()],
      ['정답 수', sc.nTruth], ['도구가 찾은 수', sc.nPred], ['맞게 찾음', sc.tp], ['값까지 맞음', sc.exact], ['정밀도', sc.precision], ['재현율', sc.recall], ['F1', sc.f1], ['값 일치율', sc.valueRate], ['목표', settings().target / 100]];
    var cat = [['구분', '정답', '찾음', '값까지 맞음']]; Object.keys(sc.byCat).forEach(function (c) { var b = sc.byCat[c]; cat.push([c, b.truth, b.found, b.exact]); });
    var bad = [['종류', '정답 번호', '정답', '도구 번호', '도구']];
    sc.misses.forEach(function (m) { bad.push(['놓침', m.no, m.text, '', '']); });
    sc.wrong.forEach(function (m) { bad.push(['값 다름', m.no, m.truth, m.predNo, m.pred]); });
    sc.extra.forEach(function (m) { bad.push(['더 찾음', '', '', m.no, m.text]); });
    writeXlsx(safeName(cur.name) + '_일치율_' + stamp() + '.xlsx', { '요약': sum, '구분별': cat, '틀린 항목': bad }, { '요약': [16, 40], '구분별': [12, 8, 8, 12], '틀린 항목': [8, 9, 28, 9, 28] });
  }

  // ── 설정 ──────────────────────────────────────────────────
  function renderSettings(main) {
    var s = settings();
    main.appendChild(pageHead('설정', '인식 규칙의 기준값입니다. 회사 도면 양식에 맞게 바꾼 뒤, 일치율을 다시 재 보세요.'));
    function chk(k, label) { return h('label', { class: 'check' }, h('input', { type: 'checkbox', checked: s[k], onchange: function (e) { db.settings[k] = e.target.checked; save(); toast('저장했습니다.'); } }), h('span', null, label)); }
    function num(k, label, min, max, step, hint) { return field(label, h('input', { type: 'number', min: min, max: max, step: step || 1, value: s[k], onchange: function (e) { var o = {}; o[k] = +e.target.value; db.settings = L.cleanSettings(Object.assign({}, s, o)); save(); toast('저장했습니다.'); } }), hint); }
    var kw = h('textarea', { rows: 4 }, s.titleKeywords);
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '제외 영역 자동 찾기'),
      chk('autoBorder', '테두리 띠(구역 번호 1, 2, 3 · A, B, C) 안의 글자 빼기'), num('borderPct', '테두리 띠 폭(쪽 짧은 변의 %)', 0, 15, 0.5),
      chk('autoTitle', '표제란 · 개정란 빼기(아래 낱말이 2개 이상 모인 오른쪽 아래 · 오른쪽 위 영역)'),
      field('표제란 낱말(쉼표로 구분)', kw, '회사 표제란에 쓰는 칸 이름을 더하세요(예: 도면번호, 품번, 재질)'),
      h('button', { type: 'button', class: 'btn', onclick: function () { db.settings.titleKeywords = kw.value; save(); toast('저장했습니다.'); } }, '낱말 저장'),
      chk('autoNotes', '주기(NOTES) 문장 속 숫자 빼기'), chk('includeEtc', '표면 거칠기(Ra) 등 「기타」도 후보로')));
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '풍선 · 연결 · 목표'), h('div', { class: 'form-grid three' },
      num('balloonR', '풍선 반지름(mm)', 1.5, 12, 0.5), num('linkMaxMm', '지시선 끝 ~ 글자 최대 거리(mm)', 2, 60), num('nearMaxMm', '지시선 없는 풍선 ~ 글자 최대 거리(mm)', 2, 60),
      num('target', '목표 값 일치율(%)', 50, 100), num('zoneCols', '구역 열 수', 1, 26), num('zoneRows', '구역 행 수', 1, 20))));
    main.appendChild(window.AIPanel.settingsCard({ toast: toast }));
    var fin = h('input', { type: 'file', accept: '.json,application/json', 'aria-label': '백업 파일 고르기', onchange: function (e) {
      var f = e.target.files[0]; if (!f) return; var r = new FileReader();
      r.onload = function () { try { db = L.restoreDb(JSON.parse(r.result)); save(); toast('되살렸습니다. 도면 PDF 는 백업에 없으니 다시 열어 주세요.'); render(); } catch (er) { toast('백업 파일을 읽지 못했습니다.', true); } };
      r.readAsText(f);
    } });
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '백업 · 지우기'),
      h('p', { class: 'note' }, '설정 · 후보 고침 · 풍선 위치 · 정답 · 측정 기록을 JSON 으로 받습니다(도면 PDF 파일은 넣지 않음). 지금 ' + S.sizeKb() + 'KB.'),
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn', onclick: function () { download('data10-03_백업_' + stamp() + '.json', new Blob([JSON.stringify(db, null, 1)], { type: 'application/json' })); } }, 'JSON 백업 받기'),
        h('label', { class: 'btn file-btn' }, 'JSON 되살리기', fin),
        h('button', { type: 'button', class: 'btn', onclick: function () { db.settings = L.cleanSettings(); save(); render(); toast('규칙 설정을 처음 값으로 되돌렸습니다.'); } }, '규칙 설정 처음 값으로'),
        h('button', { type: 'button', class: 'btn btn-danger', onclick: function () {
          confirmBox('모두 지우기', '이 브라우저에 저장한 설정 · 고침 · 정답 · 기록 · 도면 파일을 모두 지웁니다.', '모두 지우기', function () { S.clearDb(); S.clearPdf(); db = L.emptyDb(); cur = null; save(); render(); toast('모두 지웠습니다.'); }, true);
        } }, '모두 지우기'))));
  }

  if (!S.available()) document.getElementById('storeBanner').hidden = false;
  render();
  restoreLast();
  window.__app = { openSample: openSample, state: function () { return { cur: cur, db: db, ui: ui }; }, g1: function () { return g1(); }, g2: function () { return g2(); } };
})();

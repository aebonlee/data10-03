/* 브라우저 저장소
   · 설정·후보 고침·풍선 위치·정답·측정 기록 → localStorage (막혀 있거나 가득 차면 이번 창 메모리로만)
   · 도면 PDF 파일 자체 → IndexedDB (용량이 커서). 막혀 있으면 이번 창에서만 기억합니다.
   데이터는 이 브라우저에만 있습니다(서버로 보내지 않음). */
(function (root) {
  'use strict';
  var KEY_DB = 'data10-03.db', IDB = 'data10-03', STORE = 'pdf';
  var memory = {}, memPdf = {}, ok = true, lastError = '';
  function get(k) {
    try { var v = root.localStorage.getItem(k); return v == null && memory[k] != null ? memory[k] : v; }
    catch (e) { ok = false; return memory[k] == null ? null : memory[k]; }
  }
  function set(k, v) {
    memory[k] = v;
    try { root.localStorage.setItem(k, v); lastError = ''; return true; }
    catch (e) { ok = false; lastError = (e && (e.name === 'QuotaExceededError' || e.code === 22)) ? 'quota' : 'blocked'; return false; }
  }
  function del(k) { delete memory[k]; try { root.localStorage.removeItem(k); } catch (e) { ok = false; } }

  var dbp = null;
  function idb() {
    if (dbp) return dbp;
    dbp = new Promise(function (res, rej) {
      try {
        var rq = root.indexedDB.open(IDB, 1);
        rq.onupgradeneeded = function () { rq.result.createObjectStore(STORE); };
        rq.onsuccess = function () { res(rq.result); };
        rq.onerror = function () { rej(rq.error); };
      } catch (e) { rej(e); }
    });
    return dbp;
  }
  function tx(mode, fn) {
    return idb().then(function (d) {
      return new Promise(function (res, rej) {
        var t = d.transaction(STORE, mode), s = t.objectStore(STORE), r = fn(s);
        t.oncomplete = function () { res(r && r.result); };
        t.onerror = function () { rej(t.error); };
      });
    });
  }

  root.DRStore = {
    loadDb: function () {
      var raw = get(KEY_DB);
      if (!raw) return root.DRLogic.emptyDb();
      try { return root.DRLogic.restoreDb(JSON.parse(raw)); } catch (e) { return root.DRLogic.emptyDb(); }
    },
    saveDb: function (db) { return set(KEY_DB, JSON.stringify(db)); },
    clearDb: function () { del(KEY_DB); },
    lastError: function () { return lastError; },
    sizeKb: function () { var v = get(KEY_DB); return v ? Math.round(v.length * 2 / 1024) : 0; },
    available: function () { get(KEY_DB); return ok; },
    /* PDF 바이트 — Promise */
    putPdf: function (key, bytes) {
      memPdf[key] = bytes;
      return tx('readwrite', function (s) { return s.put(bytes, key); }).then(function () { return true; }, function () { return false; });
    },
    getPdf: function (key) {
      if (memPdf[key]) return Promise.resolve(memPdf[key]);
      return tx('readonly', function (s) { return s.get(key); }).then(function (v) { return v || null; }, function () { return null; });
    },
    delPdf: function (key) { delete memPdf[key]; return tx('readwrite', function (s) { return s.delete(key); }).catch(function () {}); },
    clearPdf: function () { memPdf = {}; return tx('readwrite', function (s) { return s.clear(); }).catch(function () {}); }
  };
})(window);

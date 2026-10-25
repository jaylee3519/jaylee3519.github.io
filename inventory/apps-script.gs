/**
 * 매장 재고관리 - 구글 시트 서버 (Google Apps Script)
 *
 * 설치: 구글 시트 > 확장 프로그램 > Apps Script 에 이 파일 내용을 "전부" 붙여넣고
 *       아래 TOKEN 을 원하는 비밀번호로 바꾼 뒤 "웹 앱"으로 배포하세요.
 *       자세한 순서는 inventory/README.md 참고.
 */
const TOKEN = 'CHANGE-ME'; // 앱 설정의 "비밀번호"와 똑같이 입력

const PRODUCT_SHEET = '상품';
const STOCK_SHEET = '위치별재고';
const LOG_SHEET = '기록';
const HEADS = {
  [PRODUCT_SHEET]: ['코드', '상품명', '규격', '총수량', '최소수량', '메모', '수정일시'],
  [STOCK_SHEET]: ['코드', '상품명', '위치', '수량', '수정일시'],
  [LOG_SHEET]: ['일시', '코드', '상품명', '위치', '변동', '위치재고', '총재고', '구분', '메모', '작업자'],
};

function doGet() {
  return json_({ ok: true, data: '재고관리 서버가 동작 중입니다.' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: '잘못된 요청입니다.' });
  }
  if (TOKEN === 'CHANGE-ME') return json_({ ok: false, error: 'Apps Script 코드의 TOKEN 값을 먼저 바꿔주세요.' });
  if (req.token !== TOKEN) return json_({ ok: false, error: '비밀번호가 맞지 않습니다.' });

  // 여러 직원이 동시에 수정해도 수량이 꼬이지 않도록 잠금
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    if (req.action === 'log') return json_({ ok: true, data: readLog_(req) });
    const db = loadDb_();
    const result = Core.run(db, req.action, req);
    saveChanges_(db, result.changes);
    return json_({ ok: true, data: result.data });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

// ---------- 시트 읽기 ----------

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    const head = HEADS[name];
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
    // 바코드 앞자리 0이 사라지지 않도록 코드 열을 텍스트로
    sh.getRange(name === LOG_SHEET ? 'B:B' : 'A:A').setNumberFormat('@');
  }
  return sh;
}

function values_(name) {
  const sh = sheet_(name);
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, HEADS[name].length).getValues() : [];
}

function loadDb_() {
  const products = values_(PRODUCT_SHEET).map((r, i) => ({
    _row: i + 2, code: str_(r[0]), name: str_(r[1]), spec: str_(r[2]),
    minQty: r[4] === '' ? '' : Number(r[4]), memo: str_(r[5]), updatedAt: iso_(r[6]),
  })).filter((p) => p.code);
  const stocks = values_(STOCK_SHEET).map((r, i) => ({
    _row: i + 2, code: str_(r[0]), location: str_(r[2]) || '미지정', qty: Number(r[3]) || 0, updatedAt: iso_(r[4]),
  })).filter((s) => s.code);
  return { products, stocks };
}

function readLog_(req) {
  const sh = sheet_(LOG_SHEET);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const limit = Math.min(Number(req.limit) || 100, 500);
  const code = str_(req.code);
  const take = Math.min(n, code ? 5000 : limit);
  return sh.getRange(sh.getLastRow() - take + 1, 1, take, HEADS[LOG_SHEET].length).getValues()
    .reverse()
    .map((r) => ({
      at: iso_(r[0]), code: str_(r[1]), name: str_(r[2]), location: str_(r[3]), delta: Number(r[4]) || 0,
      qtyAfter: Number(r[5]) || 0, total: Number(r[6]) || 0, type: str_(r[7]), memo: str_(r[8]), user: str_(r[9]),
    }))
    .filter((x) => !code || x.code === code)
    .slice(0, limit);
}

// ---------- 시트 쓰기 (바뀐 행만) ----------

function saveChanges_(db, ch) {
  const totals = {};
  db.stocks.forEach((s) => { totals[s.code] = (totals[s.code] || 0) + (Number(s.qty) || 0); });
  const names = {};
  db.products.forEach((p) => { names[p.code] = p.name; });

  writeRows_(PRODUCT_SHEET, db.products, ch.products, ch.removedProducts,
    (p) => [p.code, p.name, p.spec, totals[p.code] || 0, p.minQty, p.memo, date_(p.updatedAt)]);
  writeRows_(STOCK_SHEET, db.stocks, ch.stocks, ch.removedStocks,
    (s) => [s.code, names[s.code] || '', s.location, Number(s.qty) || 0, date_(s.updatedAt)]);

  if (ch.logs.length) {
    const sh = sheet_(LOG_SHEET);
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 2, ch.logs.length, 1).setNumberFormat('@');
    sh.getRange(start, 1, ch.logs.length, HEADS[LOG_SHEET].length).setValues(ch.logs.map((l) =>
      [date_(l.at), l.code, l.name, l.location, l.delta, l.qtyAfter, l.total, l.type, l.memo, l.user]));
  }

  // 바코드를 연결해 코드가 바뀌면 지난 기록도 새 코드로
  if (ch.rename) {
    const sh = sheet_(LOG_SHEET);
    const n = sh.getLastRow() - 1;
    if (n > 0) {
      const col = sh.getRange(2, 2, n, 1);
      const v = col.getValues();
      let hit = false;
      v.forEach((r) => { if (str_(r[0]) === ch.rename.from) { r[0] = ch.rename.to; hit = true; } });
      if (hit) col.setValues(v);
    }
  }
}

function writeRows_(name, all, dirty, removed, toRow) {
  if (!dirty.size && !removed.length) return;
  const sh = sheet_(name);
  const width = HEADS[name].length;

  // 많이 바뀌면(가져오기 등) 시트 전체를 한 번에 다시 씀
  if (dirty.size + removed.length > 40) {
    const n = sh.getLastRow() - 1;
    if (n > 0) sh.getRange(2, 1, n, width).clearContent();
    if (all.length) {
      sh.getRange(2, 1, all.length, 1).setNumberFormat('@');
      sh.getRange(2, 1, all.length, width).setValues(all.map(toRow));
    }
    return;
  }

  dirty.forEach((o) => {
    if (!o._row) return;
    sh.getRange(o._row, 1).setNumberFormat('@');
    sh.getRange(o._row, 1, 1, width).setValues([toRow(o)]);
  });
  removed.filter((o) => o._row).map((o) => o._row).sort((a, b) => b - a).forEach((r) => sh.deleteRow(r));
  const fresh = [...dirty].filter((o) => !o._row);
  if (fresh.length) {
    const start = sh.getLastRow() + 1;
    sh.getRange(start, 1, fresh.length, 1).setNumberFormat('@');
    sh.getRange(start, 1, fresh.length, width).setValues(fresh.map(toRow));
  }
}

function str_(v) { return String(v == null ? '' : v).trim(); }
function iso_(v) { return v instanceof Date ? v.toISOString() : str_(v); }
function date_(v) { const d = new Date(v); return v && !isNaN(d) ? d : v || ''; }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ===== 아래는 core.js (앱과 함께 쓰는 재고 처리 로직) - 직접 고치지 말고 core.js를 고친 뒤 build.sh 실행 =====

/**
 * 재고 처리 핵심 로직 (앱의 '이 기기에만' 모드와 구글 시트 서버가 함께 사용)
 *
 * db = { products: [{code, name, spec, minQty, memo, updatedAt}],
 *        stocks:   [{code, location, qty, updatedAt}] }
 * Core.run(db, action, req) → { data, changes }
 * changes: 저장해야 할 변경 내용 (서버는 바뀐 행만 시트에 씀)
 */
var Core = (function () {
  function clean(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); }
  function num(v) { return Number(String(v == null ? '' : v).replace(/,/g, '')); }
  function nowIso() { return new Date().toISOString(); }
  function fail(msg) { throw new Error(msg); }
  function newCode() { return 'P-' + Date.now().toString(36).toUpperCase() + Math.random().toString(36).slice(2, 5).toUpperCase(); }

  function run(db, action, req) {
    req = req || {};
    var ch = { products: new Set(), stocks: new Set(), removedProducts: [], removedStocks: [], logs: [], rename: null };
    var at = nowIso();

    function product(code) {
      code = clean(code);
      for (var i = 0; i < db.products.length; i++) if (db.products[i].code === code) return db.products[i];
      return null;
    }
    function mustProduct(code) { return product(code) || fail('등록되지 않은 상품입니다: ' + code); }
    function stocksOf(code) { return db.stocks.filter(function (s) { return s.code === code; }); }
    function total(code) { return stocksOf(code).reduce(function (a, s) { return a + (Number(s.qty) || 0); }, 0); }
    function stock(code, location, create) {
      location = clean(location) || '미지정';
      for (var i = 0; i < db.stocks.length; i++) {
        var s = db.stocks[i];
        if (s.code === code && s.location === location) return s;
      }
      if (!create) return null;
      var n = { code: code, location: location, qty: 0, updatedAt: at };
      db.stocks.push(n);
      ch.stocks.add(n);
      return n;
    }
    function view(p) {
      var st = stocksOf(p.code).map(function (s) { return { location: s.location, qty: Number(s.qty) || 0, updatedAt: s.updatedAt }; });
      return { code: p.code, name: p.name, spec: p.spec, minQty: p.minQty, memo: p.memo, updatedAt: p.updatedAt,
        qty: st.reduce(function (a, s) { return a + s.qty; }, 0), stocks: st };
    }
    function touch(p) { p.updatedAt = at; ch.products.add(p); }
    function setQty(p, s, qty, type, memo) {
      var delta = qty - (Number(s.qty) || 0);
      s.qty = qty; s.updatedAt = at; ch.stocks.add(s); touch(p);
      log(p, s.location, delta, s.qty, type, memo);
    }
    function log(p, location, delta, after, type, memo) {
      ch.logs.push({ at: at, code: p.code, name: p.name + (p.spec ? ' ' + p.spec : ''), location: location || '',
        delta: delta, qtyAfter: after, total: total(p.code), type: type, memo: clean(memo), user: clean(req.user) });
    }
    function fields(p, d) {
      if (d.spec !== undefined) p.spec = clean(d.spec);
      if (d.memo !== undefined) p.memo = clean(d.memo);
      if (d.minQty !== undefined) p.minQty = clean(d.minQty) === '' ? '' : num(d.minQty);
    }
    function pickLocation(p, location) {
      location = clean(location);
      if (location) return location;
      var st = stocksOf(p.code);
      if (st.length === 1) return st[0].location;
      if (!st.length) return '미지정';
      return fail('위치를 선택하세요.');
    }
    function qtyOf(v, label) {
      var n = num(v);
      if (!isFinite(n) || clean(v) === '') fail((label || '수량') + '이 올바르지 않습니다.');
      return n;
    }
    function rename(p, code) {
      if (product(code)) fail('이미 등록된 코드입니다: ' + code);
      var from = p.code;
      stocksOf(from).forEach(function (s) { s.code = code; ch.stocks.add(s); });
      p.code = code;
      ch.rename = { from: from, to: code };
    }
    function addProduct(d, type) {
      var code = clean(d.code) || newCode();
      var name = clean(d.name);
      if (!name) fail('상품명을 입력하세요.');
      if (product(code)) fail('이미 등록된 코드입니다: ' + code);
      var p = { code: code, name: name, spec: '', minQty: '', memo: '', updatedAt: at };
      fields(p, d);
      db.products.push(p); ch.products.add(p);
      var list = d.stocks && d.stocks.length ? d.stocks : [{ location: d.location, qty: d.qty }];
      list.forEach(function (x) {
        var s = stock(code, x.location, true);
        var q = num(x.qty);
        s.qty = (Number(s.qty) || 0) + (isFinite(q) ? q : 0);
        if (x.updatedAt) s.updatedAt = x.updatedAt;
      });
      stocksOf(code).forEach(function (s) { log(p, s.location, s.qty, s.qty, type); });
      return p;
    }

    var data;
    switch (action) {
      case 'list': {
        var by = {};
        db.stocks.forEach(function (s) { (by[s.code] = by[s.code] || []).push(s); });
        data = { products: db.products.map(function (p) {
          var st = (by[p.code] || []).map(function (s) { return { location: s.location, qty: Number(s.qty) || 0, updatedAt: s.updatedAt }; });
          return { code: p.code, name: p.name, spec: p.spec, minQty: p.minQty, memo: p.memo, updatedAt: p.updatedAt,
            qty: st.reduce(function (a, s) { return a + s.qty; }, 0), stocks: st };
        }) };
        break;
      }

      case 'save': {
        var d = req.product || {};
        var orig = clean(req.originalCode);
        if (!orig) { data = view(addProduct(d, '신규등록')); break; }
        var p = mustProduct(orig);
        var code = clean(d.code) || orig;
        if (!clean(d.name)) fail('상품명을 입력하세요.');
        if (code !== orig) rename(p, code);
        if (clean(d.name) !== p.name) { p.name = clean(d.name); stocksOf(p.code).forEach(function (s) { ch.stocks.add(s); }); }
        fields(p, d);
        touch(p);
        data = view(p);
        break;
      }

      case 'adjust': {
        var p = mustProduct(req.code);
        var delta = qtyOf(req.delta);
        if (delta === 0) fail('수량을 입력하세요.');
        var s = stock(p.code, pickLocation(p, req.location), true);
        setQty(p, s, (Number(s.qty) || 0) + delta, req.type || (delta > 0 ? '입고' : '출고'), req.memo);
        data = view(p);
        break;
      }

      case 'set': {
        var p = mustProduct(req.code);
        var s = stock(p.code, pickLocation(p, req.location), true);
        setQty(p, s, qtyOf(req.qty), '실사', req.memo);
        data = view(p);
        break;
      }

      case 'move': {
        var p = mustProduct(req.code);
        var q = qtyOf(req.qty);
        if (q <= 0) fail('옮길 수량을 입력하세요.');
        var from = stock(p.code, req.from, false) || fail('출발 위치를 찾을 수 없습니다.');
        var toName = clean(req.to) || fail('옮길 위치를 입력하세요.');
        if (toName === from.location) fail('같은 위치로는 옮길 수 없습니다.');
        var to = stock(p.code, toName, true);
        from.qty = (Number(from.qty) || 0) - q; to.qty = (Number(to.qty) || 0) + q;
        from.updatedAt = to.updatedAt = at; ch.stocks.add(from); ch.stocks.add(to); touch(p);
        log(p, from.location, -q, from.qty, '이동', '→ ' + to.location + (req.memo ? ' · ' + req.memo : ''));
        log(p, to.location, q, to.qty, '이동', from.location + ' →' + (req.memo ? ' · ' + req.memo : ''));
        data = view(p);
        break;
      }

      case 'removeLocation': {
        var p = mustProduct(req.code);
        var s = stock(p.code, req.location, false) || fail('위치를 찾을 수 없습니다.');
        if (Number(s.qty) !== 0) fail('수량이 0인 위치만 뺄 수 있습니다. 먼저 이동하거나 실사로 0을 맞추세요.');
        db.stocks.splice(db.stocks.indexOf(s), 1);
        ch.stocks.delete(s); ch.removedStocks.push(s);
        touch(p);
        log(p, s.location, 0, 0, '위치삭제');
        data = view(p);
        break;
      }

      case 'delete': {
        var p = mustProduct(req.code);
        log(p, '', -total(p.code), 0, '삭제');
        stocksOf(p.code).forEach(function (s) { db.stocks.splice(db.stocks.indexOf(s), 1); ch.stocks.delete(s); ch.removedStocks.push(s); });
        db.products.splice(db.products.indexOf(p), 1);
        ch.products.delete(p); ch.removedProducts.push(p);
        data = { code: p.code };
        break;
      }

      case 'import': {
        var added = 0, updated = 0;
        (req.products || []).forEach(function (d) {
          var p = product(d.code);
          if (!p) { addProduct(d, '가져오기'); added++; return; }
          if (clean(d.name) && clean(d.name) !== p.name) { p.name = clean(d.name); stocksOf(p.code).forEach(function (s) { ch.stocks.add(s); }); }
          fields(p, d);
          touch(p);
          (d.stocks || []).forEach(function (x) {
            var s = stock(p.code, x.location, true);
            var q = num(x.qty);
            if (clean(x.qty) !== '' && isFinite(q) && q !== Number(s.qty)) setQty(p, s, q, '가져오기');
            if (x.updatedAt) s.updatedAt = x.updatedAt;
          });
          updated++;
        });
        data = { added: added, updated: updated };
        break;
      }

      default:
        fail('알 수 없는 작업: ' + action);
    }
    return { data: data, changes: ch };
  }

  return { run: run, clean: clean, newCode: newCode };
})();

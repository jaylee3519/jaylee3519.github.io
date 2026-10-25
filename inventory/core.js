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

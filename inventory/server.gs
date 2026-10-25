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

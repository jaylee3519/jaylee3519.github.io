/**
 * 매장 재고관리 - 구글 시트 서버 (Google Apps Script)
 *
 * 설치: 구글 시트 > 확장 프로그램 > Apps Script 에 이 파일 내용을 그대로 붙여넣고
 *       아래 TOKEN 을 원하는 비밀번호로 바꾼 뒤 "웹 앱"으로 배포하세요.
 *       자세한 순서는 inventory/README.md 참고.
 */
const TOKEN = 'CHANGE-ME'; // 앱 설정의 "비밀번호"와 똑같이 입력

const PRODUCT_SHEET = '재고';
const LOG_SHEET = '기록';
const P_HEAD = ['코드', '상품명', '수량', '위치', '최소수량', '메모', '수정일시'];
const L_HEAD = ['일시', '코드', '상품명', '변동', '변동후수량', '구분', '메모', '작업자'];

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
  const fn = ACTIONS[req.action];
  if (!fn) return json_({ ok: false, error: '알 수 없는 작업: ' + req.action });

  // 여러 직원이 동시에 수정해도 수량이 꼬이지 않도록 잠금
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    return json_({ ok: true, data: fn(req) });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    lock.releaseLock();
  }
}

const ACTIONS = {
  list: () => ({ products: readProducts_().map(strip_) }),

  save: (req) => {
    const d = req.product || {};
    const code = clean_(d.code);
    const name = String(d.name || '').trim();
    if (!code || !name) throw new Error('코드와 상품명은 꼭 입력해야 합니다.');
    const orig = clean_(req.originalCode) || code;
    const all = readProducts_();
    if (code !== orig && all.some((x) => x.code === code)) throw new Error('이미 등록된 코드입니다: ' + code);

    let p = all.find((x) => x.code === orig);
    if (!p) {
      if (all.some((x) => x.code === code)) throw new Error('이미 등록된 코드입니다: ' + code);
      p = { row: 0, code, name, qty: num_(d.qty) || 0, location: '', minQty: '', memo: '' };
      applyFields_(p, d);
      p.updatedAt = new Date();
      writeProduct_(p);
      addLogs_([logRow_(p, p.qty, '신규등록', '', req.user)]);
      return strip_(p);
    }
    p.code = code;
    p.name = name;
    applyFields_(p, d);
    p.updatedAt = new Date();
    writeProduct_(p);
    return strip_(p);
  },

  adjust: (req) => {
    const delta = num_(req.delta);
    if (!isFinite(delta) || delta === 0) throw new Error('수량이 올바르지 않습니다.');
    return change_(req, (p) => p.qty + delta, req.type || (delta > 0 ? '입고' : '출고'));
  },

  set: (req) => {
    const qty = num_(req.qty);
    if (!isFinite(qty)) throw new Error('수량이 올바르지 않습니다.');
    return change_(req, () => qty, '실사');
  },

  delete: (req) => {
    const p = find_(req.code);
    if (!p) throw new Error('등록되지 않은 코드입니다: ' + req.code);
    sheet_(PRODUCT_SHEET).deleteRow(p.row);
    addLogs_([logRow_(p, 0, '삭제', '', req.user)]);
    return { code: p.code };
  },

  log: (req) => {
    const sh = sheet_(LOG_SHEET);
    const n = sh.getLastRow() - 1;
    if (n < 1) return [];
    const limit = Math.min(num_(req.limit) || 100, 500);
    const code = clean_(req.code);
    const take = Math.min(n, code ? 5000 : limit);
    const rows = sh.getRange(sh.getLastRow() - take + 1, 1, take, L_HEAD.length).getValues();
    return rows
      .reverse()
      .map((r) => ({
        at: iso_(r[0]), code: String(r[1]), name: String(r[2]), delta: Number(r[3]) || 0,
        qtyAfter: Number(r[4]) || 0, type: String(r[5]), memo: String(r[6]), user: String(r[7]),
      }))
      .filter((x) => !code || x.code === code)
      .slice(0, limit);
  },

  // CSV 가져오기 / 기기 데이터 업로드: 한 번에 시트에 기록
  import: (req) => {
    const sh = sheet_(PRODUCT_SHEET);
    const n = sh.getLastRow() - 1;
    const vals = n > 0 ? sh.getRange(2, 1, n, P_HEAD.length).getValues() : [];
    const index = new Map();
    vals.forEach((r, i) => { const c = clean_(r[0]); if (c) index.set(c, i); });
    const now = new Date();
    const logs = [];
    let added = 0, updated = 0;

    (req.products || []).forEach((d) => {
      const code = clean_(d.code);
      if (!code) return;
      const qty = d.qty === '' || d.qty == null ? NaN : num_(d.qty);
      let i = index.get(code);
      let p;
      if (i == null) {
        p = { code, name: String(d.name || '').trim() || code, qty: isFinite(qty) ? qty : 0, location: '', minQty: '', memo: '' };
        applyFields_(p, d);
        p.updatedAt = now;
        index.set(code, vals.length);
        vals.push(productRow_(p));
        logs.push(logRow_(p, p.qty, '가져오기', '', req.user, now));
        added++;
      } else {
        p = rowToProduct_(vals[i], i + 2);
        if (String(d.name || '').trim()) p.name = String(d.name).trim();
        applyFields_(p, d);
        if (isFinite(qty) && qty !== p.qty) {
          const delta = qty - p.qty;
          p.qty = qty;
          logs.push(logRow_(p, delta, '가져오기', '', req.user, now));
        }
        p.updatedAt = now;
        vals[i] = productRow_(p);
        updated++;
      }
    });

    if (vals.length) {
      sh.getRange(2, 1, vals.length, 1).setNumberFormat('@');
      sh.getRange(2, 1, vals.length, P_HEAD.length).setValues(vals);
    }
    addLogs_(logs);
    return { added, updated };
  },
};

// ---------- 내부 함수 ----------

function change_(req, calc, type) {
  const p = find_(req.code);
  if (!p) throw new Error('등록되지 않은 코드입니다: ' + req.code);
  const next = calc(p);
  const delta = next - p.qty;
  p.qty = next;
  p.updatedAt = new Date();
  writeProduct_(p);
  addLogs_([logRow_(p, delta, type, req.memo, req.user)]);
  return strip_(p);
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    const head = name === PRODUCT_SHEET ? P_HEAD : L_HEAD;
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
    // 바코드 앞자리 0이 사라지지 않도록 코드 열을 텍스트로
    sh.getRange(name === PRODUCT_SHEET ? 'A:A' : 'B:B').setNumberFormat('@');
  }
  return sh;
}

function readProducts_() {
  const sh = sheet_(PRODUCT_SHEET);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, P_HEAD.length).getValues()
    .map((r, i) => rowToProduct_(r, i + 2))
    .filter((p) => p.code);
}

function rowToProduct_(r, row) {
  return {
    row, code: clean_(r[0]), name: String(r[1]), qty: Number(r[2]) || 0, location: String(r[3]),
    minQty: r[4] === '' ? '' : Number(r[4]), memo: String(r[5]), updatedAt: iso_(r[6]),
  };
}

function productRow_(p) {
  return [p.code, p.name, p.qty, p.location, p.minQty, p.memo, p.updatedAt];
}

function find_(code) {
  code = clean_(code);
  return readProducts_().find((p) => p.code === code);
}

function writeProduct_(p) {
  const sh = sheet_(PRODUCT_SHEET);
  if (!p.row) p.row = sh.getLastRow() + 1;
  sh.getRange(p.row, 1).setNumberFormat('@');
  sh.getRange(p.row, 1, 1, P_HEAD.length).setValues([productRow_(p)]);
}

function applyFields_(p, d) {
  if (d.location != null) p.location = String(d.location).trim();
  if (d.memo != null) p.memo = String(d.memo).trim();
  if (d.minQty !== undefined) p.minQty = d.minQty === '' || d.minQty == null ? '' : num_(d.minQty);
}

function logRow_(p, delta, type, memo, user, at) {
  return [at || new Date(), p.code, p.name, delta, p.qty, type, memo || '', user || ''];
}

function addLogs_(rows) {
  if (!rows.length) return;
  const sh = sheet_(LOG_SHEET);
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 2, rows.length, 1).setNumberFormat('@');
  sh.getRange(start, 1, rows.length, L_HEAD.length).setValues(rows);
}

function strip_(p) {
  const o = Object.assign({}, p);
  delete o.row;
  if (o.updatedAt instanceof Date) o.updatedAt = o.updatedAt.toISOString();
  return o;
}

function clean_(v) { return String(v == null ? '' : v).trim(); }
function num_(v) { return Number(String(v).replace(/,/g, '')); }
function iso_(v) { return v instanceof Date ? v.toISOString() : String(v || ''); }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

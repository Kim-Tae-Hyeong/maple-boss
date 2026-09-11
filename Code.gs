/**
 * 메이플 보스 일정표 - Apps Script API
 * 구글 시트 > 확장 프로그램 > Apps Script 에 붙여넣으세요.
 */

// ▼ 꼭 바꾸세요. 페이지에서 편집할 때 쓰는 비밀번호예요.
//   Apps Script 편집기에 붙여넣은 쪽에만 진짜 비밀번호를 넣고, 저장소에는 올리지 마세요.
const ADMIN_PASSWORD = '여기에-비밀번호';

const SHEET_NAME = 'schedule';
const HEADERS = ['id', 'date', 'time', 'boss', 'difficulty', 'character', 'party', 'memo', 'status', 'updatedAt'];

/** 처음 한 번만 실행: 시트와 헤더를 만들어요. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sh.getRange(1, 1, sh.getMaxRows(), HEADERS.length).setNumberFormat('@'); // 날짜/시간이 자동 변환되지 않게 텍스트로
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
}

/* ---------- 읽기: 누구나 ---------- */
function doGet() {
  try {
    return json_({ ok: true, items: readAll_() });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

/* ---------- 쓰기: 비밀번호 필요 ---------- */
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: '요청 형식이 올바르지 않아요' });
  }
  if (body.password !== ADMIN_PASSWORD) {
    Utilities.sleep(800); // 비밀번호 무작위 대입 속도 늦추기
    return json_({ ok: false, error: '비밀번호가 틀렸어요' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_();
    const item = body.item || {};

    switch (body.action) {
      case 'verify':
        return json_({ ok: true });

      case 'create': {
        const row = toRow_(Object.assign({}, item, { id: Utilities.getUuid().slice(0, 8) }));
        sh.getRange(sh.getLastRow() + 1, 1, 1, HEADERS.length).setNumberFormat('@').setValues([row]);
        break;
      }

      case 'update': {
        const r = findRow_(sh, item.id);
        if (!r) return json_({ ok: false, error: '일정을 찾을 수 없어요. 새로고침해 주세요' });
        sh.getRange(r, 1, 1, HEADERS.length).setNumberFormat('@').setValues([toRow_(item)]);
        break;
      }

      case 'delete': {
        const r = findRow_(sh, item.id);
        if (!r) return json_({ ok: false, error: '이미 삭제된 일정이에요' });
        sh.deleteRow(r);
        break;
      }

      default:
        return json_({ ok: false, error: '알 수 없는 요청이에요' });
    }
    return json_({ ok: true, items: readAll_() });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 내부 함수 ---------- */
function getSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('시트가 없어요. Apps Script에서 setup()을 먼저 실행하세요');
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function toRow_(item) {
  const text = (v, max) => (v == null ? '' : String(v).trim().slice(0, max || 200));
  const date = text(item.date);
  const time = text(item.time);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('날짜 형식은 2026-09-12 처럼 입력해 주세요');
  if (time && !/^\d{2}:\d{2}$/.test(time)) throw new Error('시간 형식은 20:00 처럼 입력해 주세요');
  if (!text(item.boss)) throw new Error('보스 이름을 입력해 주세요');

  const now = Utilities.formatDate(new Date(), tz_(), "yyyy-MM-dd HH:mm");
  const values = {
    id: text(item.id), date: date, time: time,
    boss: text(item.boss, 50), difficulty: text(item.difficulty, 20),
    character: text(item.character, 50), party: text(item.party, 100),
    memo: text(item.memo, 500), status: text(item.status) || '예정', updatedAt: now
  };
  return HEADERS.map(h => values[h]);
}

function findRow_(sh, id) {
  if (!id) return 0;
  const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 1;
  }
  return 0;
}

function tz_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
}

// 시트에서 직접 입력해서 날짜/시간이 Date로 바뀐 경우도 문자열로 맞춰요
function cell_(v, key) {
  if (v instanceof Date) {
    if (key === 'date') return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
    if (key === 'time') return Utilities.formatDate(v, tz_(), 'HH:mm');
    return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd HH:mm');
  }
  return String(v).trim();
}

function readAll_() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (r[1] === '' || r[1] == null) continue; // 날짜 없는 행은 건너뜀
    // 시트에서 직접 추가해 id가 비어 있으면 자동으로 채움
    if (r[0] === '') {
      r[0] = Utilities.getUuid().slice(0, 8);
      sh.getRange(i + 1, 1).setNumberFormat('@').setValue(r[0]);
    }
    const o = {};
    HEADERS.forEach((h, j) => (o[h] = cell_(r[j] == null ? '' : r[j], h)));
    if (!o.status) o.status = '예정';
    items.push(o);
  }
  return items;
}

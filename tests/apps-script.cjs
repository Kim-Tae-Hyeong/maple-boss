// 실행: node tests/apps-script.cjs (추가 패키지 필요 없음)
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
const headers = ['id', 'date', 'time', 'boss', 'difficulty', 'character', 'party', 'memo', 'status', 'updatedAt'];
const row = (id, date = '2026-12-28', time = '17:00') =>
  [id, date, time, '루시드', '하드', '본캐', '2', '메모', '예정', ''];

function fixture(code, initial = []) {
  const rows = [headers, ...initial].map(r => [...r]);
  const calls = { reads: 0, idReads: 0, writes: 0, timezone: 0, sheet: 0, locks: 0 };
  let sequence = 0;
  const sheet = {
    getLastRow: () => rows.length,
    getDataRange: () => sheet.getRange(1, 1, rows.length, headers.length),
    getRange(r, c, n = 1, w = 1) {
      const range = {
        setNumberFormat: () => range,
        getValues() {
          calls.reads++;
          if (c === 1 && w === 1) calls.idReads++;
          return rows.slice(r - 1, r - 1 + n).map(a => a.slice(c - 1, c - 1 + w));
        },
        setValues(values) {
          calls.writes++;
          assert.equal(values.length, n);
          values.forEach((a, i) => {
            assert.equal(a.length, w);
            rows[r - 1 + i] ??= Array(10).fill('');
            a.forEach((v, j) => { rows[r - 1 + i][c - 1 + j] = v; });
          });
          return range;
        },
        setValue: value => range.setValues([[value]])
      };
      return range;
    },
    getRangeList(addresses) {
      return { setValue(value) {
        calls.writes++;
        addresses.forEach(address => {
          const [, col, r] = /^([IJ])(\d+)$/.exec(address);
          rows[Number(r) - 1][col === 'I' ? 8 : 9] = value;
        });
      } };
    },
    deleteRow: r => rows.splice(r - 1, 1)
  };
  const context = vm.createContext({
    Date,
    SpreadsheetApp: { getActiveSpreadsheet: () => ({
      getSheetByName: () => { calls.sheet++; return sheet; },
      getSpreadsheetTimeZone: () => { calls.timezone++; return 'Asia/Seoul'; }
    }) },
    LockService: { getScriptLock: () => ({
      waitLock: () => { calls.locks++; }, releaseLock() {}
    }) },
    Utilities: {
      getUuid: () => String(++sequence).padStart(8, '0'), sleep() {},
      formatDate(date, timezone, format) {
        assert.equal(timezone, 'Asia/Seoul');
        const s = new Date(date.getTime() + 9 * 3600000).toISOString();
        return format === 'yyyy-MM-dd' ? s.slice(0, 10)
          : format === 'HH:mm' ? s.slice(11, 16) : '2026-10-19 17:00';
      }
    },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: value => ({
      setMimeType: () => JSON.parse(value)
    }) }
  });
  vm.runInContext(code, context);
  return { rows, calls, get: () => context.doGet(),
    post: (action, item, password = '여기에-비밀번호') => context.doPost({
      postData: { contents: JSON.stringify({ action, item, password }) }
    }) };
}

function check(code = source) {
  let f = fixture(code);
  assert.equal(f.post('verify').ok, true);
  assert.equal(f.calls.sheet, 0);
  assert.equal(f.calls.locks, 0);
  assert.equal(f.post('verify', null, 'wrong').ok, false);

  f = fixture(code, Array.from({ length: 12 }, (_, i) => row('id' + i)));
  const ids = [0, 2, 4, 6, 8, 10].map(i => 'id' + i);
  assert.equal(f.post('complete', { ids: [...ids, ids[0]] }).ok, true);
  assert.equal(f.calls.idReads, 1);
  assert.equal(f.calls.writes, 2);
  f.rows.slice(1).forEach((r, i) => {
    assert.equal(r[8], i % 2 === 0 ? '완료' : '예정');
    assert.deepEqual(r.slice(0, 8), row('id' + i).slice(0, 8));
  });
  const writes = f.calls.writes;
  assert.equal(f.post('complete', { ids: ['id0', 'missing'] }).ok, false);
  assert.equal(f.calls.writes, writes);

  f = fixture(code, [row('a')]);
  assert.equal(f.post('complete', { ids: ['a'] }).ok, true);
  assert.equal(f.calls.writes, 1);
  const copy = f.post('copy', { ids: ['a'] });
  assert.equal(copy.ok, true);
  assert.equal(copy.items[1].date, '2027-01-04');
  assert.equal(copy.items[1].time, '17:00');
  assert.equal(copy.items[1].status, '예정');
  const made = f.post('create', { date: '2026-12-28', boss: '루시드' });
  assert.ok(made.createdId);
  assert.equal(f.post('update', { id: made.createdId, date: '2026-12-29', boss: '윌' }).ok, true);
  assert.equal(f.post('delete', { id: made.createdId }).ok, true);
  assert.equal(f.post('delete', { id: made.createdId }).ok, true);

  f = fixture(code, Array.from({ length: 100 }, (_, i) =>
    row('id' + i, new Date('2026-10-18T15:00:00Z'), new Date('2026-10-19T08:00:00Z'))));
  const read = f.get();
  assert.equal(read.items.length, 100);
  assert.equal(read.items[0].date, '2026-10-19');
  assert.equal(read.items[0].time, '17:00');
  assert.equal(f.calls.timezone, 1);
  assert.equal(f.calls.reads, 1);

  f = fixture(code, [row(''), row('a')]);
  assert.ok(f.get().items[0].id);
  assert.equal(f.get().items[0].id, f.rows[1][0]);
  assert.equal(f.calls.timezone, 0);
  assert.equal(fixture(code).get().items.length, 0);
  console.log('PASS: 인증, 일괄 완료, 연도 경계 복사, 생성/수정/삭제, 날짜 변환, ID 자동 입력, 호출 횟수');
}

if (require.main === module) check();
function measure(code = source) {
  const completed = fixture(code, Array.from({ length: 10 }, (_, i) => row('id' + i)));
  assert.equal(completed.post('complete', { ids: Array.from({ length: 10 }, (_, i) => 'id' + i) }).ok, true);
  const dates = fixture(code, Array.from({ length: 100 }, (_, i) =>
    row('id' + i, new Date('2026-10-18T15:00:00Z'), new Date('2026-10-19T08:00:00Z'))));
  assert.equal(dates.get().items[0].date, '2026-10-19');
  return { complete10: completed.calls, read100Dates: dates.calls };
}
module.exports = { fixture, row, check, measure };

// 실행: node tests/frontend-actions.cjs. 실제 API를 호출하지 않는 상호작용 테스트.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const script = html.match(/<script>([\s\S]*?)<\/script>/)[1];
const items = ['a', 'b'].map((id, i) => ({
  id, date: '2026-12-28', time: '17:00', boss: ['검은 마법사', '윌'][i],
  difficulty: '하드', character: '본캐', party: '2', memo: '', status: '예정'
}));

async function fixture() {
  const elements = new Map();
  const timers = new Map();
  const requests = [];
  let timerId = 0;
  let resolvePost, rejectPost;
  function get(id) {
    if (!elements.has(id)) elements.set(id, {
      id, innerHTML: '', textContent: '', open: false, hidden: false,
      disabled: false, isConnected: true, dataset: {}, listeners: {}, attributes: {},
      classList: { add() {}, remove() {}, toggle() {} },
      setAttribute(k, v) { this.attributes[k] = v; },
      addEventListener(type, fn) { this.listeners[type] = fn; },
      querySelectorAll: () => [], focus() {},
      showModal() { this.open = true; },
      close() { this.open = false; this.listeners.close?.(); }
    });
    return elements.get(id);
  }
  const context = vm.createContext({
    document: { getElementById: get, querySelectorAll: () => [], addEventListener() {} },
    localStorage: { getItem: () => null, removeItem() {} },
    AbortController,
    setTimeout: (fn, ms) => { timers.set(++timerId, { fn, ms }); return timerId; },
    clearTimeout: id => timers.delete(id), setInterval() {},
    fetch(url, options = {}) {
      requests.push(options.body ? JSON.parse(options.body) : { action: 'GET' });
      if (!options.body) return Promise.resolve({ json: async () => ({ ok: true, items }) });
      return new Promise((resolve, reject) => {
        resolvePost = data => resolve({ json: async () => data });
        rejectPost = reject;
        options.signal?.addEventListener('abort', () => {
          const error = new Error('Timed out'); error.name = 'AbortError'; reject(error);
        }, { once: true });
      });
    }
  });
  vm.runInContext(script, context);
  await new Promise(resolve => setImmediate(resolve));
  vm.runInContext("state.password = 'test-password'", context);
  return {
    get, requests, context, timers,
    click: id => get(id).listeners.click(),
    posts: () => requests.filter(r => r.action !== 'GET'),
    resolve: data => resolvePost(data), reject: err => rejectPost(err),
    open: action => context.handleItemAction({ target: { closest: () => ({
      dataset: { act: action, ids: 'a,b' }, isConnected: true, focus() {}
    }) } })
  };
}

async function check() {
  for (const action of ['copy', 'complete']) {
    const f = await fixture();
    await f.open(action);
    assert.equal(f.get('actionDlg').open, true);
    assert.equal(f.posts().length, 0, '확인 전에는 쓰기 요청이 없어야 함');
    assert.match(f.get('actionBosses').textContent, /검은 마법사.*윌/);
    if (action === 'copy') assert.match(f.get('actionTo').textContent, /2027-01-04.*17:00/);
    f.click('actionCancel');
    assert.equal(f.get('actionDlg').open, false);
    assert.equal(f.posts().length, 0, '취소하면 쓰기 요청이 없어야 함');

    await f.open(action);
    const pending = f.click('actionConfirm');
    await f.click('actionConfirm');
    assert.equal(f.posts().length, 1, '연속 확인에도 한 번만 호출');
    assert.equal(f.posts()[0].action, action);
    assert.deepEqual(f.posts()[0].item.ids, ['a', 'b']);
    assert.equal(f.get('actionProgress').hidden, false);
    assert.equal(f.get('actionCancel').disabled, true);
    let prevented = false;
    f.get('actionDlg').listeners.cancel({ preventDefault: () => { prevented = true; } });
    assert.equal(prevented, true, '처리 중 Escape로 닫히지 않아야 함');
    [...f.timers.values()].find(t => t.ms === 8000).fn();
    assert.match(f.get('actionProgressText').textContent, /응답을 기다리고/);
    const resultItems = items.map(x => ({ ...x, status: action === 'complete' ? '완료' : '예정' }));
    f.resolve({ ok: true, items: resultItems });
    await pending;
    assert.equal(f.get('actionDlg').open, false);
    assert.equal(f.get('actionProgress').hidden, true);
    assert.equal(f.get('actionDlg').attributes['aria-busy'], 'false');
    assert.equal(vm.runInContext('itemActionPending', f.context), false);
    assert.equal(f.posts().length, 1);
  }

  for (const failure of ['server', 'network', 'timeout']) {
    const f = await fixture();
    await f.open('copy');
    const pending = f.click('actionConfirm');
    if (failure === 'server') f.resolve({ ok: false, error: '비밀번호가 틀렸어요' });
    else if (failure === 'timeout') [...f.timers.values()].find(t => t.ms === 45000).fn();
    else f.reject(new TypeError('Failed to fetch'));
    await pending;
    assert.equal(f.get('actionDlg').open, true);
    assert.equal(f.get('actionError').hidden, false);
    assert.equal(f.get('actionProgress').hidden, true);
    assert.equal(f.get('actionConfirm').hidden, true);
    assert.equal(f.get('actionCancel').disabled, false);
    assert.equal(f.posts().length, 1, '응답 실패를 자동 재전송하지 않아야 함');
    if (failure === 'server') {
      assert.match(f.get('actionError').textContent, /비밀번호가 틀렸어요/);
      assert.equal(vm.runInContext('state.password', f.context), null);
    } else assert.match(f.get('actionError').textContent, /이미 반영됐을 수/);
    const gets = f.requests.filter(r => r.action === 'GET').length;
    f.click('actionCancel');
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(f.requests.filter(r => r.action === 'GET').length, gets + 1);
    assert.equal(f.posts().length, 1);
  }
  console.log('PASS: 확인·취소, 단일 API, 중복 클릭, 진행 표시, 인증·네트워크·시간 초과 복구');
}
if (require.main === module) check().catch(error => { console.error(error); process.exitCode = 1; });
module.exports = { check };

// Project_24 数量変更後の全派生シート同期テスト
// 実行: GASアップスクリプト直下で node tests/project24_all_sheet_sync.test.js
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const SYNC_FILE = path.join(ROOT, 'Project_24', '全シート同期.js');

let failures = 0;
function test(name, fn) {
  try {
    fn();
    console.log(`PASS ${name}`);
  } catch (error) {
    failures++;
    console.error(`FAIL ${name}: ${error.message}`);
  }
}

function loadSyncContext() {
  const props = new Map();
  const alerts = [];
  const context = {
    console,
    Date,
    JSON,
    PropertiesService: {
      getDocumentProperties: () => ({
        getProperty: key => props.has(key) ? props.get(key) : null,
        setProperty: (key, value) => props.set(key, String(value)),
        deleteProperty: key => props.delete(key)
      })
    },
    SpreadsheetApp: {
      getActive: () => ({ toast: () => {} }),
      getUi: () => ({
        ButtonSet: { OK: 'OK' },
        alert: (...args) => alerts.push(args)
      })
    },
    EMS在庫を更新_本体_: () => {},
    引当実行_本体_: () => {
      props.set('引当_整合状態', JSON.stringify({ ts: Date.now(), 要確認: 0, 台帳版: 'v1' }));
      return { success: true };
    },
    取り置き初期登録を作成本体_: () => {},
    キャンセル戻し確認を更新本体_: () => {},
    Yahoo戻し候補を更新_: () => {}
  };
  vm.createContext(context);
  if (fs.existsSync(SYNC_FILE)) vm.runInContext(fs.readFileSync(SYNC_FILE, 'utf8'), context);
  context.__props = props;
  context.__alerts = alerts;
  return context;
}

test('中央同期関数を提供する', () => {
  const context = loadSyncContext();
  assert.strictEqual(typeof context.引当_数値変更後全同期_, 'function');
});

test('中央同期はEMS更新→②→管理シートの順に各1回実行する', () => {
  const context = loadSyncContext();
  const calls = [];
  context.EMS在庫を更新_本体_ = () => calls.push('ems');
  context.引当実行_本体_ = options => {
    calls.push('allocate');
    assert.strictEqual(options.skipManagementRefresh, true);
    context.__props.set('引当_整合状態', JSON.stringify({ ts: Date.now(), 要確認: 0, 台帳版: 'v1' }));
    return { success: true };
  };
  context.取り置き初期登録を作成本体_ = () => calls.push('register');
  context.キャンセル戻し確認を更新本体_ = () => calls.push('return');
  context.Yahoo戻し候補を更新_ = () => calls.push('yahoo');

  const result = context.引当_数値変更後全同期_({ 理由: 'test', EMS更新: true, 完了表示: false });

  assert.strictEqual(result.success, true);
  assert.deepStrictEqual(calls, ['ems', 'allocate', 'register', 'return', 'yahoo']);
  assert.strictEqual(JSON.parse(context.__props.get('引当_全シート同期状態')).status, 'synced');
});

test('EMS更新なしでは②と管理シートだけを各1回実行する', () => {
  const context = loadSyncContext();
  const calls = [];
  context.EMS在庫を更新_本体_ = () => calls.push('ems');
  context.引当実行_本体_ = () => {
    calls.push('allocate');
    context.__props.set('引当_整合状態', JSON.stringify({ ts: Date.now(), 要確認: 0, 台帳版: 'v1' }));
    return { success: true };
  };
  context.取り置き初期登録を作成本体_ = () => calls.push('register');
  context.キャンセル戻し確認を更新本体_ = () => calls.push('return');
  context.Yahoo戻し候補を更新_ = () => calls.push('yahoo');

  const result = context.引当_数値変更後全同期_({ 理由: 'test', 完了表示: false });

  assert.strictEqual(result.success, true);
  assert.deepStrictEqual(calls, ['allocate', 'register', 'return', 'yahoo']);
});

test('管理シート失敗時は整合状態を削除して同期済みにしない', () => {
  const context = loadSyncContext();
  context.取り置き初期登録を作成本体_ = () => { throw new Error('register failed'); };

  const result = context.引当_数値変更後全同期_({ 理由: 'test', 完了表示: false });

  assert.strictEqual(result.success, false);
  assert.strictEqual(context.__props.has('引当_整合状態'), false);
  const state = JSON.parse(context.__props.get('引当_全シート同期状態'));
  assert.strictEqual(state.status, 'failed');
  assert.match(state.error, /register failed/);
});

test('②が完了結果を返さない時は同期失敗にする', () => {
  const context = loadSyncContext();
  context.引当実行_本体_ = () => ({ success: false });

  const result = context.引当_数値変更後全同期_({ 理由: 'test', 完了表示: false });

  assert.strictEqual(result.success, false);
  assert.strictEqual(context.__props.has('引当_整合状態'), false);
  assert.strictEqual(JSON.parse(context.__props.get('引当_全シート同期状態')).status, 'failed');
});

if (failures) process.exit(1);
console.log('ALL TESTS PASSED');

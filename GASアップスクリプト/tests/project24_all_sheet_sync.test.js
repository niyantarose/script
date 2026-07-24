// Project_24 数量変更後の全派生シート同期テスト
// 実行: GASアップスクリプト直下で node tests/project24_all_sheet_sync.test.js
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const SYNC_FILE = path.join(ROOT, 'Project_24', '全シート同期.js');

function readProject(file) {
  return fs.readFileSync(path.join(ROOT, 'Project_24', file), 'utf8');
}

function functionBody(source, name) {
  const marker = `function ${name}(`;
  const start = source.indexOf(marker);
  assert.notStrictEqual(start, -1, `${name} が見つかりません`);
  const next = source.indexOf('\nfunction ', start + marker.length);
  return source.slice(start, next < 0 ? source.length : next);
}

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

const syncEntrypoints = [
  ['取り置き台帳.js', '取り置き登録を反映本体_'],
  ['取り置き台帳.js', 'キャンセル戻し確認を確定本体_'],
  ['取り置き台帳.js', 'キャンセル戻しをYahoo反映済みにする本体_'],
  ['取り置き台帳.js', '選択した取り置きを手動解除本体_'],
  ['取り置き台帳.js', 'orphanBulkRelease'],
  ['受注明細個別ボタン.js', '選択行を個別引当_本体_'],
  ['受注明細個別ボタン.js', '選択行の引当キャンセル_本体_'],
  ['現物確認移行.js', '現物確認移行を反映本体_'],
  ['消込台帳.js', '注文をキャンセル扱い本体_'],
  ['消込台帳.js', '消込台帳_発送済みCSV取込本体_'],
  ['消込台帳.js', '消込台帳のCSV処理済をクリア本体_'],
  ['消込台帳.js', '消込台帳を更新本体_'],
  ['P列自動記入.js', 'P列を書き直す本体_'],
  ['P列自動記入.js', '便の引当をやり直す本体_'],
  ['引当.js', '更新してから引当_本体_'],
  ['引当.js', '取込_実行_']
];

syncEntrypoints.forEach(([file, name]) => {
  test(`${name}は数量変更後に全同期を1回呼ぶ`, () => {
    const body = functionBody(readProject(file), name);
    assert.strictEqual((body.match(/引当_数値変更後全同期_\s*\(/g) || []).length, 1);
  });
});

test('低レベル台帳保存関数は全同期を呼ばない', () => {
  const source = readProject('取り置き台帳.js');
  assert.doesNotMatch(functionBody(source, '取り置き台帳_保存_'), /引当_数値変更後全同期_/);
  assert.doesNotMatch(functionBody(source, 'EMS在庫移動台帳_保存_'), /引当_数値変更後全同期_/);
});

test('全件再計算は管理シート失敗も反映失敗として停止する', () => {
  const body = functionBody(readProject('全件再計算.js'), '全件再計算を反映本体_');
  assert.match(body, /引当実行_本体_\s*\(\s*\{[^}]*requireAllSheetSync\s*:\s*true/s);
});

if (failures) process.exit(1);
console.log('ALL TESTS PASSED');

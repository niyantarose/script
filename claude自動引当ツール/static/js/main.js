// ── チェックボックス管理 ────────────────────────────────────────────
function toggleAll(prefix, checked) {
  document.querySelectorAll('.row-chk-' + prefix).forEach(function(chk) {
    chk.checked = checked;
    chk.closest('tr').classList.toggle('selected', checked);
  });
  updateCount(prefix);
}

function clearAll(prefix) {
  document.querySelectorAll('.row-chk-' + prefix).forEach(function(chk) {
    chk.checked = false;
    chk.closest('tr').classList.remove('selected');
  });
  var allChk = document.getElementById('chk-all-' + prefix);
  if (allChk) allChk.checked = false;
  updateCount(prefix);
}

function updateCount(prefix) {
  var checked = document.querySelectorAll('.row-chk-' + prefix + ':checked').length;
  var total   = document.querySelectorAll('.row-chk-' + prefix).length;
  var el = document.getElementById('sel-count-' + prefix);
  if (el) el.textContent = checked + '件選択中';
  document.querySelectorAll('.row-chk-' + prefix).forEach(function(chk) {
    chk.closest('tr').classList.toggle('selected', chk.checked);
  });
  var allChk = document.getElementById('chk-all-' + prefix);
  if (allChk) allChk.checked = (checked === total && total > 0);
}

function getCheckedIds(prefix) {
  var ids = [];
  document.querySelectorAll('.row-chk-' + prefix + ':checked').forEach(function(chk) {
    ids.push(parseInt(chk.dataset.id));
  });
  return ids;
}

// ── トースト通知 ─────────────────────────────────────────────────────
function showToast(msg, isError) {
  var t = document.getElementById('toast');
  if (!t) return;
  t.textContent = msg;
  t.style.background = isError ? '#dc2626' : '#1a1a1a';
  t.classList.add('show');
  setTimeout(function() { t.classList.remove('show'); }, 2500);
}

// ── セル保存（インライン編集） ────────────────────────────────────────
function saveCell(input) {
  var td    = input.closest('.cell-edit');
  var url   = td.dataset.url;
  var field = td.dataset.field;
  var id    = td.dataset.id;
  var val   = input.value;
  var dispVal = (input.tagName === 'SELECT')
    ? input.options[input.selectedIndex].text : val;

  fetch(url + '/' + id, {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({field: field, value: val})
  })
  .then(function(r) { return r.json(); })
  .then(function(data) {
    if (data.ok) {
      var valEl = td.querySelector('.cell-val');
      if (valEl) valEl.textContent = dispVal;
      hideCellInput(td);
      showToast('保存しました');
    } else {
      showToast('保存失敗: ' + (data.error || ''), true);
      hideCellInput(td);
    }
  })
  .catch(function() { showToast('通信エラー', true); hideCellInput(td); });
}

function showCellInput(td) {
  var valEl   = td.querySelector('.cell-val');
  var inputEl = td.querySelector('.cell-input');
  if (!inputEl) return;
  if (valEl) valEl.classList.add('hidden');
  inputEl.classList.add('active');
  inputEl.focus();
}

function hideCellInput(td) {
  var valEl   = td.querySelector('.cell-val');
  var inputEl = td.querySelector('.cell-input');
  if (valEl) valEl.classList.remove('hidden');
  if (inputEl) inputEl.classList.remove('active');
}

// ── ステータスセレクト（常時表示型） ─────────────────────────────────
function saveStatus(sel, baseUrl, id) {
  fetch(baseUrl + '/' + id, {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({field: 'status', value: sel.value})
  })
  .then(function(r) { return r.json(); })
  .then(function(data) {
    if (data.ok) showToast('ステータス更新しました');
    else showToast('更新失敗', true);
  })
  .catch(function() { showToast('通信エラー', true); });
}

// ── データ取込（汎用） ───────────────────────────────────────────────
function importData(url, label) {
  if (!confirm(label + ' を取込みますか？')) return;
  var btn = event.currentTarget;
  var orig = btn.textContent;
  btn.disabled = true;
  btn.textContent = '取込中...';
  fetch(url, {method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({})})
  .then(function(r) { return r.json(); })
  .then(function(data) {
    if (data.status === 'error') {
      showToast('エラー: ' + data.message, true);
    } else {
      var msg = data.message || (data.imported + '件取込みました');
      if (data.filename) msg += ' (' + data.filename + ')';
      showToast(msg);
      if ((data.imported || 0) > 0) setTimeout(function() { location.reload(); }, 1200);
    }
  })
  .catch(function() { showToast('通信エラー', true); })
  .finally(function() { btn.disabled = false; btn.textContent = orig; });
}

// ── インライン編集の初期化 ────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', function() {
  document.querySelectorAll('.cell-edit').forEach(function(td) {
    var valEl   = td.querySelector('.cell-val');
    var inputEl = td.querySelector('.cell-input');
    if (!valEl || !inputEl) return;

    valEl.addEventListener('click', function() { showCellInput(td); });

    if (inputEl.tagName === 'SELECT') {
      inputEl.addEventListener('change', function() { saveCell(inputEl); });
    } else {
      inputEl.addEventListener('blur',    function() { saveCell(inputEl); });
      inputEl.addEventListener('keydown', function(e) {
        if (e.key === 'Enter')  { e.preventDefault(); saveCell(inputEl); }
        if (e.key === 'Escape') { hideCellInput(td); }
      });
    }
  });
});

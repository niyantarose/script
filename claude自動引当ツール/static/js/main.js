// チェックボックス管理
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
  var total = document.querySelectorAll('.row-chk-' + prefix).length;
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

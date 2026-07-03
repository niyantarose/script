function onOpen() {
  Excel同期メニューを追加_();
  インボイスメニューを追加_();

  SpreadsheetApp.getUi()
    .createMenu('EMS自社用')

    // EMSリスト用
.addItem('No・Box・色・罫線を全体更新', 'EMS_全体更新')
.addItem('EMS発送日から到着予定日を補完', 'EMS_到着予定日を補完')
.addItem('EMS番号ごとの罫線を更新', 'EMS_EMS番号ごとに罫線を更新')
.addItem('選択発送日のEMS番号ごとに罫線を更新', 'EMS_選択発送日_EMS番号ごとに罫線を更新')
.addItem('選択行の発送日グループを更新', 'EMS_選択行の発送日だけ更新')
.addItem('M列EMS番号の空白を削除', 'EMS_EMS番号の空白を削除')
    .addSeparator()
    .addItem('M:Nの色だけクリア', 'EMS_色だけクリア')

    // 発注シート用
    .addSeparator()
.addItem('チェック行をEMSリストへ送る', '発注_チェック行をEMSリストへ送る')
.addItem('発注チェックボックス・罫線を全体更新', '発注_チェックボックスと罫線を全体更新')
.addItem('チェックボックスを同期', 'syncCheckboxes')
.addItem('発注シートの罫線を更新', '発注_番号あり行に格子罫線')
.addItem('消込判定の色を更新', 'manualUpdateKeshikomiColors')

    .addToUi();

  if (typeof 最終データ行へ移動_自動 === 'function') {
    最終データ行へ移動_自動();
  }

  if (typeof refreshEmsCalendarTodayHighlight_ === 'function') {
    refreshEmsCalendarTodayHighlight_();
  }
}

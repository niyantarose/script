function テスト_フォルダ権限チェック() {
  const ids = {
    'Yahoo在庫フォルダ': フォルダID_Yahoo在庫,
    'Amazon在庫フォルダ': フォルダID_Amazon在庫,
    'Qoo10在庫フォルダ': フォルダID_Qoo10在庫,
  };

  for (var name in ids) {
    var id = ids[name];
    try {
      var folder = DriveApp.getFolderById(id);
      Logger.log(name + ' OK: ' + folder.getName());
    } catch (e) {
      Logger.log(name + ' NG: ' + id + ' / ' + e);
    }
  }
}

function テスト_SheetsAPI() {
  const ssId = getSsId_();
  try {
    const res = Sheets.Spreadsheets.get(ssId, { fields: 'spreadsheetId' });
    Logger.log('SheetsAPI OK: ' + JSON.stringify(res));
  } catch (e) {
    Logger.log('SheetsAPI NG: ' + e);
  }
}
function テスト_SheetsAPI() {
  const ssId = getSsId_();
  try {
    const res = Sheets.Spreadsheets.get(ssId, { fields: 'spreadsheetId' });
    Logger.log('SheetsAPI OK: ' + JSON.stringify(res));
  } catch (e) {
    Logger.log('SheetsAPI NG: ' + e);
  }
}

function テスト_DriveFilesAPI() {
  try {
    const res = Drive.Files.list({
      q: "'" + フォルダID_Yahoo在庫 + "' in parents and trashed = false",
      maxResults: 1,
      fields: 'items(id,title)'
    });
    Logger.log('DriveFiles OK: ' + JSON.stringify(res));
  } catch (e) {
    Logger.log('DriveFiles NG: ' + e);
  }
}

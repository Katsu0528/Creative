function processSeikaChanges() {
  var PROCESS_FOLDER_ID = '1AzfpjGwdPKgMzKZTpNRYgoZ2GrngBeht';
  var originalSs = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = originalSs.getSheetByName('成果変更用');
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('成果変更用 sheet not found.');
    return;
  }

  var processSs = getProcessSpreadsheet(PROCESS_FOLDER_ID);
  if (!processSs) {
    SpreadsheetApp.getUi().alert('処理用 spreadsheet not found in folder.');
    return;
  }
  var processSheet = processSs.getSheets()[0];

  var logSheet = originalSs.getSheetByName('処理ログ') || originalSs.insertSheet('処理ログ');
  logSheet.appendRow([new Date(), '処理開始']);

  // Step 1: Build lookup maps from source sheet
  var srcData = sourceSheet.getRange(3, 1, sourceSheet.getLastRow() - 2, 4).getValues();
  var approve = {};
  var cancel = {};
  srcData.forEach(function(row) {
    if (row[0] && row[1]) approve[row[0] + '\u0000' + row[1]] = true;
    if (row[2] && row[3]) cancel[row[2] + '\u0000' + row[3]] = true;
  });

  var lastRow = processSheet.getLastRow();
  if (lastRow > 2) {
    var range = processSheet.getRange(3, 19, lastRow - 2, 19); // S to AK
    var values = range.getValues();
    var changed = 0;
    for (var i = 0; i < values.length; i++) {
      var key = values[i][0] + '\u0000' + values[i][7];
      if (approve[key]) {
        values[i][18] = '承認';
        changed++;
      } else if (cancel[key]) {
        values[i][18] = 'キャンセル';
        changed++;
      }
    }
    range.setValues(values);
    logSheet.appendRow([new Date(), changed + ' row(s) updated in 処理用']);
  } else {
    logSheet.appendRow([new Date(), '処理用 sheet had no data']);
  }

  // Step 2: create DL sheet in original spreadsheet
  var dlSheet = originalSs.getSheetByName('DL');
  if (dlSheet) originalSs.deleteSheet(dlSheet);
  dlSheet = originalSs.insertSheet('DL');

  var indices = [1,4,8,10,22,23,24,26,37,47];
  var dlValues = [];
  var procRange = processSheet.getRange(1, 1, lastRow, processSheet.getMaxColumns()).getValues();
  for (var r = 0; r < procRange.length; r++) {
    var row = [];
    for (var j = 0; j < indices.length; j++) {
      row.push(procRange[r][indices[j]-1]);
    }
    dlValues.push(row);
  }
  dlSheet.getRange(1, 1, dlValues.length, indices.length).setValues(dlValues);
  logSheet.appendRow([new Date(), 'DL sheet created']);

  // Step 3: download DL sheet as Shift_JIS
  downloadCsvDlShiftJis();
  logSheet.appendRow([new Date(), 'DL CSV exported']);

  // Step 4: delete DL sheet
  originalSs.deleteSheet(dlSheet);
  logSheet.appendRow([new Date(), 'DL sheet deleted']);
}


function getProcessSpreadsheet(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  return null;
}

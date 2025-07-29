// Delete sheets whose names start with "(SJIS)action_log_raw_".
// The previous behaviour removed any sheet not listed in keepNames, but the
// new requirement is to only target these raw log sheets.
function deleteOtherSheets(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (name.indexOf('(SJIS)action_log_raw_') === 0) {
      ss.deleteSheet(sheet);
    }
  });
}

// Standalone helper to remove "(SJIS)action_log_raw_" sheets from the
// active spreadsheet.
function cleanupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  deleteOtherSheets(ss);
}


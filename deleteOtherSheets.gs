function deleteOtherSheets(ss, keepNames) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  keepNames = keepNames || [];
  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (keepNames.indexOf(name) === -1) {
      ss.deleteSheet(sheet);
    }
  });
}

// Standalone helper to remove sheets except the default ones.
function cleanupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var keepNames = ['シート1', '抽出結果'];
  deleteOtherSheets(ss, keepNames);
}


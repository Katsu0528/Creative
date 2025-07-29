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


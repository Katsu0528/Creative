function deleteSheetByName(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
}

// Example: delete the sheet created by parseMultiFormatData
function deleteExtractResultSheet() {
  deleteSheetByName('抽出結果');
}

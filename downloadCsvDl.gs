function downloadCsvDlShiftJis() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('DL');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('DL sheet not found.');
    return;
  }
  var data = sheet.getDataRange().getValues();
  var csvContent = '';
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    for (var j = 0; j < row.length; j++) {
      csvContent += row[j];
      if (j < row.length - 1) {
        csvContent += ',';
      }
    }
    csvContent += '\n';
  }
  var blob = Utilities.newBlob(csvContent, 'text/csv', 'DL.csv').setContentTypeFromExtension();
  // Convert the UTF-8 blob to Shift_JIS so the CSV opens correctly in
  // environments expecting that encoding.
  var sjisBlob = convertToShiftJis(blob);
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput('<a href="' + sjisBlob.getBlob().getDataUrl() + '" target="_blank">Download</a>'),
    'Download DL CSV (Shift_JIS)'
  );
}

function convertToShiftJis(blob) {
  // Utilities.newBlob().setDataFromString allows specifying the charset
  // directly without relying on the external Encoding library.
  return Utilities.newBlob('', 'text/csv', blob.getName())
    .setDataFromString(blob.getDataAsString(), 'Shift_JIS');
}

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
  var sjisBlob = convertToShiftJis(blob);
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput('<a href="' + sjisBlob.getBlob().getDataUrl() + '" target="_blank">Download</a>'), 'Download DL CSV (Shift_JIS)');
}

function convertToShiftJis(blob) {
  var uint8Array = new Uint8Array(blob.getBytes());
  var sjisArray = Encoding.convert(uint8Array, {to: 'SJIS', from: 'UTF8'});
  var sjisBlob = Utilities.newBlob(sjisArray, 'text/csv', blob.getName());
  return sjisBlob;
}

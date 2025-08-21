function copyNextMonthSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var pattern = /^(\d{4})年(\d{1,2})月対応_(受領|請求発行|データ格納)$/;
  var latestDate = null;
  var templates = {};
  sheets.forEach(function(sheet) {
    var m = sheet.getName().match(pattern);
    if (m) {
      var year = parseInt(m[1], 10);
      var month = parseInt(m[2], 10);
      var type = m[3];
      var d = new Date(year, month - 1);
      if (!latestDate || d.getTime() > latestDate.getTime()) {
        latestDate = d;
      }
      templates[type] = sheet;
    }
  });
  if (!latestDate) return null;
  var nextYear = latestDate.getFullYear();
  var nextMonth = latestDate.getMonth() + 2;
  if (nextMonth > 12) {
    nextYear++;
    nextMonth = 1;
  }
  var ym = nextYear + '年' + nextMonth + '月対応';
  function copy(sheet, suffix) {
    var copied = sheet.copyTo(ss);
    copied.setName(ym + '_' + suffix);
    ss.setActiveSheet(copied);
    return copied;
  }
  var newReceive = templates['受領'] && copy(templates['受領'], '受領');
  if (newReceive) newReceive.getRange('A1').setValue(ym);
  var newInvoice = templates['請求発行'] && copy(templates['請求発行'], '請求発行');
  if (newInvoice) {
    newInvoice.getRange('A:E').clearContent();
    newInvoice.getRange('G:G').clearContent();
  }
  var newData = templates['データ格納'] && copy(templates['データ格納'], 'データ格納');
  if (newData) {
    newData.getRange('O:S').clearContent();
    newData.getRange('W:AA').clearContent();
  }
  return ym + '_データ格納';
}

function createNextMonthAndSummarize() {
  var dataSheetName = copyNextMonthSheets();
  if (dataSheetName) {
    summarizeResultsByAgency(dataSheetName);
  }
}

function summarizeAdsFromFolder() {
  var folderId = '1zKNeMn3FDbkEt4AMDyLeAbYwka0PKrsq';
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var master = SpreadsheetApp.getActiveSpreadsheet();

  while (files.hasNext()) {
    var file = files.next();
    try {
      var sourceSs = SpreadsheetApp.open(file);
      var sourceSheet = sourceSs.getSheets()[0];
      var data = sourceSheet.getDataRange().getValues();
      if (data.length < 2) {
        file.setTrashed(true);
        continue;
      }

      var dataSheet = master.insertSheet(file.getName());
      dataSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

      var summarySheet = master.insertSheet(file.getName() + '_summary');
      summarySheet.getRange(1, 1, 1, 3).setValues([[
        '広告',
        '件数',
        '成果報酬額合計'
      ]]);

      var lastRow = dataSheet.getLastRow();
      if (lastRow < 2) {
        file.setTrashed(true);
        continue;
      }

      var ads = dataSheet.getRange(2, 3, lastRow - 1, 1).getValues().flat();
      var uniqueAds = Array.from(new Set(ads.filter(String)));
      uniqueAds.sort();

      var rows = [];
      uniqueAds.forEach(function(ad) {
        rows.push([
          ad,
          '=COUNTIF(' + dataSheet.getName() + '!C2:C, "' + ad + '")',
          '=SUMIF(' + dataSheet.getName() + '!C2:C, "' + ad + '", ' + dataSheet.getName() + '!F2:F)'
        ]);
      });

      if (rows.length > 0) {
        summarySheet.getRange(2, 1, rows.length, 3).setValues(rows);
        var totalRow = rows.length + 2;
        summarySheet.getRange(totalRow, 1).setValue('合計');
        summarySheet.getRange(totalRow, 2).setFormula('=SUM(B2:B' + (totalRow - 1) + ')');
        summarySheet.getRange(totalRow, 3).setFormula('=SUM(C2:C' + (totalRow - 1) + ')');
      }

      file.setTrashed(true);
    } catch (e) {
      Logger.log('Error processing file ' + file.getName() + ': ' + e);
      file.setTrashed(true);
    }
  }
}

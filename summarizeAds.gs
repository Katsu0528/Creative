function summarizeAdsFromFolder() {
  var folderId = '1zKNeMn3FDbkEt4AMDyLeAbYwka0PKrsq';
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var master = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Starting summarization for folder: ' + folderId);

  while (files.hasNext()) {
    var file = files.next();
    try {
      Logger.log('Processing file: ' + file.getName());
      var sourceSs = SpreadsheetApp.open(file);
      var sourceSheet = sourceSs.getSheets()[0];
      var data = sourceSheet.getDataRange().getValues();
      if (data.length < 2) {
        Logger.log('No data found in ' + file.getName() + ' - moving to trash');
        file.setTrashed(true);
        continue;
      }

      var dataSheet = master.insertSheet(file.getName());
      dataSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      Logger.log("Copied data from " + file.getName() + " to " + dataSheet.getName());

      var summarySheet = master.insertSheet(file.getName() + '_summary');
      summarySheet.getRange(1, 1, 1, 3).setValues([[
        '広告',
        '件数',
        '成果報酬額合計'
      ]]);
      Logger.log("Created summary sheet: " + summarySheet.getName());

      var lastRow = dataSheet.getLastRow();
      if (lastRow < 2) {
        Logger.log("No rows to summarize in " + dataSheet.getName());
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
      Logger.log("Built " + rows.length + " summary rows for " + file.getName());

      if (rows.length > 0) {
        summarySheet.getRange(2, 1, rows.length, 3).setValues(rows);
        var totalRow = rows.length + 2;
        summarySheet.getRange(totalRow, 1).setValue('合計');
        summarySheet.getRange(totalRow, 2).setFormula('=SUM(B2:B' + (totalRow - 1) + ')');
        summarySheet.getRange(totalRow, 3).setFormula('=SUM(C2:C' + (totalRow - 1) + ')');
      }

      Logger.log("Finished processing " + file.getName() + " - moving to trash");
      file.setTrashed(true);
    } catch (e) {
      Logger.log('Error processing file ' + file.getName() + ': ' + e);
      file.setTrashed(true);
    }
  }
  Logger.log("Summarization complete");
}

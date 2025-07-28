function summarizeAdsFromFolder() {
  var folderId = '1zKNeMn3FDbkEt4AMDyLeAbYwka0PKrsq';
  var masterId = '1iOVL9ojJZvaXU_bRx7vRVyWPQsTTE8_RAw7a_ivXca8';

  Logger.log('Starting summarization for folder: ' + folderId);

  var folder = DriveApp.getFolderById(folderId);

  // Count files first to give better feedback
  var countIter = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var fileCount = 0;
  while (countIter.hasNext()) {
    fileCount++;
    countIter.next();
  }
  Logger.log('Found ' + fileCount + ' spreadsheet file(s)');

  if (fileCount === 0) {
    Logger.log('No files to process. Exiting.');
    Logger.log('Summarization complete');
    return;
  }

  // Reset iterator for actual processing
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var master = SpreadsheetApp.openById(masterId);
  var processedCount = 0;

  while (files.hasNext()) {
    var file = files.next();
    Logger.log('Processing file: ' + file.getName());
    try {
      var sourceSs = SpreadsheetApp.open(file);
      var sourceSheet = sourceSs.getSheets()[0];
      var data = sourceSheet.getDataRange().getValues();
      Logger.log('Read ' + data.length + ' row(s) from ' + file.getName());
      if (data.length < 2) {
        Logger.log('Skipping ' + file.getName() + ' due to insufficient rows');
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
        Logger.log('No data rows in ' + dataSheet.getName());
        file.setTrashed(true);
        continue;
      }

      var ads = dataSheet.getRange(2, 3, lastRow - 1, 1).getValues().flat();
      var uniqueAds = Array.from(new Set(ads.filter(String)));
      uniqueAds.sort();
      Logger.log('Found ' + uniqueAds.length + ' unique ad(s)');

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
      } else {
        Logger.log('No valid ad rows found in ' + dataSheet.getName());
      }

      file.setTrashed(true);
      processedCount++;
      Logger.log('Finished processing ' + file.getName());
    } catch (e) {
      Logger.log('Error processing file ' + file.getName() + ': ' + e);
      file.setTrashed(true);
    }
  }

  Logger.log('Processed ' + processedCount + ' file(s).');
  Logger.log('Summarization complete');
}

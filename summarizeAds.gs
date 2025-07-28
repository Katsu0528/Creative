function summarizeAdsFromFolder() {
  var folderId = '1zKNeMn3FDbkEt4AMDyLeAbYwka0PKrsq';
  var masterId = '1iOVL9ojJZvaXU_bRx7vRVyWPQsTTE8_RAw7a_ivXca8';

  Logger.log('Starting summarization for folder: ' + folderId);

  var folder = DriveApp.getFolderById(folderId);

  // Count convertible files including Excel and document formats
  var countIter = folder.getFiles();
  var fileCount = 0;
  while (countIter.hasNext()) {
    var mime = countIter.next().getMimeType();
    if (mime === MimeType.GOOGLE_SHEETS ||
        mime === MimeType.MICROSOFT_EXCEL ||
        mime === MimeType.MICROSOFT_EXCEL_LEGACY ||
        mime === MimeType.GOOGLE_DOCS ||
        mime === MimeType.MICROSOFT_WORD ||
        mime === MimeType.MICROSOFT_WORD_LEGACY) {
      fileCount++;
    }
  }
  Logger.log('Found ' + fileCount + ' spreadsheet file(s)');

  if (fileCount === 0) {
    Logger.log('No files to process. Exiting.');
    Logger.log('Summarization complete');
    return;
  }

  // Reset iterator for actual processing
  var files = folder.getFiles();
  var master = SpreadsheetApp.openById(masterId);
  var processedCount = 0;

  while (files.hasNext()) {
    var file = files.next();
    var mime = file.getMimeType();
    if (mime !== MimeType.GOOGLE_SHEETS &&
        mime !== MimeType.MICROSOFT_EXCEL &&
        mime !== MimeType.MICROSOFT_EXCEL_LEGACY &&
        mime !== MimeType.GOOGLE_DOCS &&
        mime !== MimeType.MICROSOFT_WORD &&
        mime !== MimeType.MICROSOFT_WORD_LEGACY) {
      Logger.log('Skipping unsupported file: ' + file.getName());
      file.setTrashed(true);
      continue;
    }
    Logger.log('Processing file: ' + file.getName());
    try {
      var sourceSs;
      if (mime === MimeType.GOOGLE_SHEETS) {
        sourceSs = SpreadsheetApp.open(file);
      } else if (mime === MimeType.MICROSOFT_EXCEL ||
                 mime === MimeType.MICROSOFT_EXCEL_LEGACY) {
        var resource = {title: file.getName(), mimeType: MimeType.GOOGLE_SHEETS};
        var converted = Drive.Files.copy(resource, file.getId());
        sourceSs = SpreadsheetApp.openById(converted.id);
        DriveApp.getFileById(converted.id).setTrashed(true);
      } else {
        Logger.log('Attempting to convert document to spreadsheet: ' + file.getName());
        var resource = {title: file.getName(), mimeType: MimeType.GOOGLE_SHEETS};
        try {
          var converted = Drive.Files.copy(resource, file.getId());
          sourceSs = SpreadsheetApp.openById(converted.id);
          DriveApp.getFileById(converted.id).setTrashed(true);
        } catch (e) {
          Logger.log('Failed to convert document ' + file.getName() + ': ' + e);
          file.setTrashed(true);
          continue;
        }
      }
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
      summarySheet.getRange(1, 1, 1, 4).setValues([[
        '広告',
        '単価',
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
      var totalCount = 0;
      var totalAmount = 0;
      uniqueAds.forEach(function(ad) {
        var count = 0;
        var amount = 0;
        for (var i = 1; i < data.length; i++) {
          var row = data[i];
          if (row[2] === ad) {
            count++;
            var val = parseFloat(String(row[5]).replace(/,/g, '')) || 0;
            amount += val;
          }
        }
        var unitPrice = count > 0 ? amount / count : 0;
        rows.push([ad, unitPrice, count, amount]);
        totalCount += count;
        totalAmount += amount;
      });

      if (rows.length > 0) {
        summarySheet.getRange(2, 1, rows.length, 4).setValues(rows);
        var totalRow = rows.length + 2;
        summarySheet.getRange(totalRow, 1, 1, 4).setValues([['合計', '', totalCount, totalAmount]]);
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

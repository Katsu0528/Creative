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

      // Determine the column positions dynamically based on the header row
      var header = data[0];
      var adCol = header.indexOf('広告');
      if (adCol === -1) {
        Logger.log('広告 column not found in ' + file.getName());
        file.setTrashed(true);
        continue;
      }
      var unitCol;
      if (adCol === 2) { // C column pattern
        unitCol = 5;     // F column
      } else if (adCol === 7) { // H column pattern
        unitCol = 22;    // W column
      } else {
        Logger.log('Unexpected 広告 column position in ' + file.getName() + ': ' + adCol);
        file.setTrashed(true);
        continue;
      }

      var adPriceMap = {};
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        var ad = row[adCol];
        if (!ad) {
          continue;
        }
        var rawUnit = row[unitCol];
        var unit = typeof rawUnit === 'number'
          ? rawUnit
          : parseFloat(String(rawUnit).replace(/[¥￥,円]/g, '').trim()) || 0;
        var key = ad + '\u0000' + unit;
        if (!adPriceMap[key]) {
          adPriceMap[key] = { ad: ad, unit: unit, count: 0 };
        }
        adPriceMap[key].count++;
      }

      var rows = [];
      var totalCount = 0;
      for (var key in adPriceMap) {
        var record = adPriceMap[key];
        rows.push([record.ad, record.unit, record.count]);
        totalCount += record.count;
      }

      rows.sort(function(a, b) {
        if (a[0] === b[0]) {
          return a[1] - b[1];
        }
        return a[0] < b[0] ? -1 : 1;
      });

      if (rows.length > 0) {
        summarySheet.getRange(2, 1, rows.length, 3).setValues(rows);
        var formulas = [];
        for (var i = 0; i < rows.length; i++) {
          formulas.push([`=B${i + 2}*C${i + 2}`]);
        }
        summarySheet.getRange(2, 4, rows.length, 1).setFormulas(formulas);
        var totalRow = rows.length + 2;
        summarySheet.getRange(totalRow, 1, 1, 3).setValues([[
          '合計',
          '',
          totalCount
        ]]);
        summarySheet.getRange(totalRow, 4).setFormula(`=SUM(D2:D${totalRow - 1})`);
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

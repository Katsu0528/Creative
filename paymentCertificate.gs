'use strict';

var PAYMENT_CERTIFICATE_FOLDER_ID = '1QUuRgmqtaPjaidCtGOYUxxbRFIzCwi0G';

var PAYMENT_DATES_BY_MONTH = {
  '2024-12': new Date(2025, 0, 31),
  '2025-1': new Date(2025, 1, 28),
  '2025-2': new Date(2025, 2, 31),
  '2025-3': new Date(2025, 3, 30),
  '2025-4': new Date(2025, 4, 30),
  '2025-5': new Date(2025, 5, 30),
  '2025-6': new Date(2025, 6, 31),
  '2025-7': new Date(2025, 7, 29),
  '2025-8': new Date(2025, 8, 30),
  '2025-9': new Date(2025, 9, 31),
  '2025-10': new Date(2025, 10, 28),
  '2025-11': new Date(2025, 11, 30)
};

function generatePaymentCertificate() {
  var folder = DriveApp.getFolderById(PAYMENT_CERTIFICATE_FOLDER_ID);
  var files = folder.getFiles();
  var summary = {};
  var totals = {};
  var fileCount = 0;
  var sheetCount = 0;
  var rowCount = 0;
  var skippedEmptyItem = 0;
  var skippedInvalidDate = 0;

  while (files.hasNext()) {
    var file = files.next();
    var mimeType = file.getMimeType();
    if (mimeType !== MimeType.GOOGLE_SHEETS && mimeType !== MimeType.CSV) {
      continue;
    }
    fileCount += 1;
    if (mimeType === MimeType.GOOGLE_SHEETS) {
      var spreadsheet = SpreadsheetApp.openById(file.getId());
      spreadsheet.getSheets().forEach(function(sheet) {
        sheetCount += 1;
        var values = sheet.getDataRange().getValues();
        var results = processPaymentCertificateRows_(values, summary, totals);
        rowCount += results.rowCount;
        skippedEmptyItem += results.skippedEmptyItem;
        skippedInvalidDate += results.skippedInvalidDate;
      });
    } else {
      var csvText = file.getBlob().getDataAsString('Shift_JIS');
      var csvRows = Utilities.parseCsv(csvText);
      sheetCount += 1;
      var csvResults = processPaymentCertificateRows_(csvRows, summary, totals);
      rowCount += csvResults.rowCount;
      skippedEmptyItem += csvResults.skippedEmptyItem;
      skippedInvalidDate += csvResults.skippedInvalidDate;
    }
  }

  Logger.log('支払証明書集計: files=%s, sheets=%s, rows=%s, skippedEmptyItem=%s, skippedInvalidDate=%s', fileCount, sheetCount, rowCount, skippedEmptyItem, skippedInvalidDate);

  var doc = DocumentApp.create('支払証明書_' + formatToday_());
  var body = doc.getBody();
  var today = new Date();
  var timezone = Session.getScriptTimeZone();

  var dateLine = Utilities.formatDate(today, timezone, 'yyyy年M月d日');
  appendParagraph_(body, dateLine, DocumentApp.HorizontalAlignment.RIGHT);
  appendParagraph_(body, '宛名欄として様', DocumentApp.HorizontalAlignment.LEFT);
  appendParagraph_(body, '〒141-0031', DocumentApp.HorizontalAlignment.RIGHT);
  appendParagraph_(body, '東京都品川区西五反田1−3−8', DocumentApp.HorizontalAlignment.RIGHT);
  appendParagraph_(body, '五反田PLACE　2F', DocumentApp.HorizontalAlignment.RIGHT);
  appendParagraph_(body, '株式会社OTONARI', DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph('');
  appendParagraph_(body, '支払証明書', DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('');
  appendParagraph_(body, '下記の支払を行いましたことを、本状にて証明いたします。', DocumentApp.HorizontalAlignment.LEFT);
  body.appendParagraph('');

  var table = body.appendTable();
  var headerRow = table.appendTableRow();
  headerRow.appendTableCell('支払日');
  headerRow.appendTableCell('金額');
  headerRow.appendTableCell('内容');

  Object.keys(totals)
    .sort(compareYearMonth_)
    .forEach(function(key) {
      var total = totals[key];
      if (!total) {
        return;
      }
      var parts = key.split('-');
      var year = Number(parts[0]);
      var month = Number(parts[1]);
      var paymentDay = paymentDateForMonth_(year, month);
      var items = summary[key];

      if (!paymentDay) {
        Logger.log('支払証明書集計: 支払日未設定 year=%s, month=%s', year, month);
        return;
      }
      var row = table.appendTableRow();
      row.appendTableCell(formatDate_(paymentDay, timezone));
      row.appendTableCell(formatCurrency_(total));
      row.appendTableCell(buildDetailText_(items));
    });

  DriveApp.getFileById(doc.getId()).moveTo(folder);
  return doc.getUrl();
}

function processPaymentCertificateRows_(rows, summary, totals) {
  var rowCount = 0;
  var skippedEmptyItem = 0;
  var skippedInvalidDate = 0;

  rows.forEach(function(row) {
    rowCount += 1;
    var item = row[2];
    if (!item) {
      skippedEmptyItem += 1;
      return;
    }
    var amount = parseAmount_(row[6]);
    var day = parseDate_(row[10]);
    if (!day) {
      skippedInvalidDate += 1;
      return;
    }
    var key = day.getFullYear() + '-' + (day.getMonth() + 1);
    if (!summary[key]) {
      summary[key] = {};
    }
    if (!summary[key][item]) {
      summary[key][item] = { amount: 0, count: 0 };
    }
    summary[key][item].amount += amount;
    summary[key][item].count += 1;
    totals[key] = (totals[key] || 0) + amount;
  });

  return {
    rowCount: rowCount,
    skippedEmptyItem: skippedEmptyItem,
    skippedInvalidDate: skippedInvalidDate
  };
}

function parseAmount_(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }
  if (typeof value === 'number') {
    return value;
  }
  var cleaned = String(value).replace(/[\s,]/g, '').replace(/[¥円]/g, '');
  var amount = Number(cleaned);
  return isNaN(amount) ? 0 : amount;
}

function parseDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }
  if (value === null || value === undefined || value === '') {
    return null;
  }
  var text = String(value).trim();
  var parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }
  var match = text.match(/(\d{4})[\/\.年-](\d{1,2})[\/\.月-](\d{1,2})/);
  if (!match) {
    var monthMatch = text.match(/(\d{4})[\/\.年-](\d{1,2})/);
    if (!monthMatch) {
      return null;
    }
    var yearOnly = Number(monthMatch[1]);
    var monthOnly = Number(monthMatch[2]);
    var monthDate = new Date(yearOnly, monthOnly - 1, 1);
    return isNaN(monthDate.getTime()) ? null : monthDate;
  }
  var year = Number(match[1]);
  var month = Number(match[2]);
  var day = Number(match[3]);
  var date = new Date(year, month - 1, day);
  return isNaN(date.getTime()) ? null : date;
}

function paymentDateForMonth_(year, month) {
  var key = year + '-' + month;
  if (PAYMENT_DATES_BY_MONTH[key]) {
    return PAYMENT_DATES_BY_MONTH[key];
  }
  return null;
}

function formatCurrency_(amount) {
  var rounded = Math.round(amount);
  return rounded.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '円';
}

function formatDate_(date, timezone) {
  return Utilities.formatDate(date, timezone, 'yyyy年M月d日');
}

function formatToday_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
}

function appendParagraph_(body, text, alignment) {
  var paragraph = body.appendParagraph(text);
  paragraph.setAlignment(alignment);
  return paragraph;
}

function buildDetailText_(items) {
  var lines = [];
  Object.keys(items)
    .sort()
    .forEach(function(name) {
      var detail = items[name];
      if (!detail || (!detail.amount && !detail.count)) {
        return;
      }
      lines.push(name + '：' + detail.count + '件　' + formatCurrency_(detail.amount));
    });
  return lines.join('\n');
}

function compareYearMonth_(a, b) {
  var partsA = a.split('-');
  var partsB = b.split('-');
  var yearA = Number(partsA[0]);
  var yearB = Number(partsB[0]);
  if (yearA !== yearB) {
    return yearA - yearB;
  }
  return Number(partsA[1]) - Number(partsB[1]);
}

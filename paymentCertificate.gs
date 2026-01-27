'use strict';

var PAYMENT_CERTIFICATE_FOLDER_ID = '1QUuRgmqtaPjaidCtGOYUxxbRFIzCwi0G';

var PAYMENT_DATES_BY_MONTH = {
  '2024-12': new Date(2025, 1, 31),
  '2025-1': new Date(2025, 2, 28),
  '2025-2': new Date(2025, 3, 31),
  '2025-3': new Date(2025, 4, 30),
  '2025-4': new Date(2025, 5, 30),
  '2025-5': new Date(2025, 6, 30),
  '2025-6': new Date(2025, 7, 31),
  '2025-7': new Date(2025, 8, 29),
  '2025-8': new Date(2025, 9, 30),
  '2025-9': new Date(2025, 10, 31),
  '2025-10': new Date(2025, 11, 28),
  '2025-11': new Date(2025, 12, 30)
};

function generatePaymentCertificate() {
  var folder = DriveApp.getFolderById(PAYMENT_CERTIFICATE_FOLDER_ID);
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var summary = {};
  var totals = {};

  while (files.hasNext()) {
    var file = files.next();
    var spreadsheet = SpreadsheetApp.openById(file.getId());
    spreadsheet.getSheets().forEach(function(sheet) {
      var values = sheet.getDataRange().getValues();
      values.forEach(function(row) {
        var item = row[2];
        if (!item) {
          return;
        }
        var amount = parseAmount_(row[6]);
        var day = parseDate_(row[10]);
        if (!day) {
          return;
        }
        var key = day.getFullYear() + '-' + (day.getMonth() + 1);
        if (!summary[key]) {
          summary[key] = {};
        }
        summary[key][item] = (summary[key][item] || 0) + amount;
        totals[key] = (totals[key] || 0) + amount;
      });
    });
  }

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
  var header = table.appendTableRow();
  header.appendTableCell('支払日');
  header.appendTableCell('金額');
  header.appendTableCell('内容');

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
      var detailLines = Object.keys(items)
        .sort()
        .map(function(name) {
          var itemTotal = items[name];
          if (!itemTotal) {
            return null;
          }
          return name + '：' + formatCurrency_(itemTotal);
        })
        .filter(function(line) {
          return line;
        });

      var row = table.appendTableRow();
      row.appendTableCell(formatDate_(paymentDay, timezone));
      row.appendTableCell(formatCurrency_(total));
      row.appendTableCell(detailLines.join('\n'));
    });

  DriveApp.getFileById(doc.getId()).moveTo(folder);
  return doc.getUrl();
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
    return null;
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
  return new Date(year, month, 0);
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

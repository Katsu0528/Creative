var TARGET_SPREADSHEET_ID = '1qkae2jGCUlykwL-uTf0_eaBGzon20RCC-wBVijyvm8s';
var DATE_SPREADSHEET_ID = '13zQMfgfYlec1BOo0LwWZUerQD9Fm0Fkzav8Z20d5eDE';
var DATE_SHEET_NAME = '日付';

function classifyResultsByClientSheet(records, startDate, endDate) {
  var validRange =
    startDate instanceof Date && !isNaN(startDate) &&
    endDate instanceof Date && !isNaN(endDate);
  if (!validRange) {
    var dateSheet = SpreadsheetApp.openById(DATE_SPREADSHEET_ID)
      .getSheetByName(DATE_SHEET_NAME);
    if (dateSheet) {
      startDate = dateSheet.getRange('B2').getValue();
      endDate = dateSheet.getRange('C2').getValue();
      validRange =
        startDate instanceof Date && !isNaN(startDate) &&
        endDate instanceof Date && !isNaN(endDate);
    }
  }
  if (!validRange) {
    Logger.log('classifyResultsByClientSheet: invalid date range');
    return {};
  }
  startDate.setHours(0, 0, 0, 0);
  endDate.setHours(0, 0, 0, 0);

  var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  var clientSheet = ss.getSheetByName('クライアント情報');
  if (!clientSheet) {
    SpreadsheetApp.getUi().alert('クライアント情報シートが見つかりません');
    return {};
  }

  var data = clientSheet.getDataRange().getValues();
  var advMap = {};
  for (var i = 1; i < data.length; i++) {
    var advId = data[i][14];      // O column
    if (!advId) continue;
    var adName = data[i][0] || '__DEFAULT__';
    var state = data[i][13];      // N column
    (advMap[advId] = advMap[advId] || {})[adName] = state;
  }

  var result = {};
  var notFound = [];
  if (!Array.isArray(records)) records = [];

  for (var r = 0; r < records.length; r++) {
    var rec = records[r];
    var advId = rec.advertiserId || rec.advertiser ||
                rec.advertiser_name || rec.advertiserName || '';
    var advName = rec.advertiser_name || rec.advertiserName || advId;
    var ad = rec.ad || rec.ad_name || rec.adName || '';
    var states = advMap[advId] || {};
    var state = states[ad] || states['__DEFAULT__'];
    if (!state) {
      notFound.push([advName, ad]);
      continue;
    }

    var unix = state === '発生' ? rec.regist_unix : rec.apply_unix;
    var str = state === '発生' ? rec.regist : rec.apply;
    var d = unix ? new Date(Number(unix) * 1000)
                 : (str ? new Date(String(str).replace(' ', 'T')) : null);
    if (!d || d < startDate || d > endDate) continue;

    var entry = result[advId] || (result[advId] = {generated: [], confirmed: []});
    (state === '発生' ? entry.generated : entry.confirmed).push(rec);
  }

  if (notFound.length) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('クライアント情報シートに記載がない成果が ' + notFound.length + ' 件あります。');
    var missSheet = ss.getSheetByName('該当無し') || ss.insertSheet('該当無し');
    missSheet.clearContents();
    missSheet.getRange(1, 1, 1, 2).setValues([['広告主名', '広告名']]);
    missSheet.getRange(2, 1, notFound.length, 2).setValues(notFound);
  }

  return result;
}

function processUniqueAdvertiserAds(sheet) {
  var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  sheet = sheet && typeof sheet.getRange === 'function' ? sheet : ss.getActiveSheet();
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var values = sheet.getRange(2, 22, lastRow - 1, 2).getValues();
  var seen = {};
  for (var i = 0; i < values.length; i++) {
    var adv = values[i][0];
    var ad = values[i][1];
    if (!adv && !ad) continue;
    var key = adv + '\u0000' + ad;
    if (seen[key]) continue;
    seen[key] = true;
    Logger.log('Processing advertiser=' + adv + ', ad=' + ad);
    // ここで広告主と広告名ごとの処理を行う
  }
}

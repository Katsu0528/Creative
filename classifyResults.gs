var TARGET_SPREADSHEET_ID = '1qkae2jGCUlykwL-uTf0_eaBGzon20RCC-wBVijyvm8s';
var DATE_SPREADSHEET_ID = '13zQMfgfYlec1BOo0LwWZUerQD9Fm0Fkzav8Z20d5eDE';
var DATE_SHEET_ID = 0;

function classifyResultsByClientSheet(records, startDate, endDate) {
  if (!(startDate instanceof Date) || !(endDate instanceof Date)) {
    var dateSs = SpreadsheetApp.openById(DATE_SPREADSHEET_ID);
    var dateSheet = dateSs.getSheetById(DATE_SHEET_ID);
    startDate = dateSheet.getRange('B2').getValue();
    endDate = dateSheet.getRange('C2').getValue();
  }
  if (!(startDate instanceof Date) || isNaN(startDate) || !(endDate instanceof Date) || isNaN(endDate)) {
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
  var mapByAdv = {};      // advertiser -> state
  var mapByAdvAd = {};    // advertiser -> (ad -> state)

  for (var i = 1; i < data.length; i++) {
    var adName = data[i][0];       // A column
    var advertiser = data[i][1];   // B column
    var state = data[i][13];       // N column (index 13)
    if (!advertiser) continue;
    if (adName) {
      if (!mapByAdvAd[advertiser]) mapByAdvAd[advertiser] = {};
      mapByAdvAd[advertiser][adName] = state;
    } else {
      mapByAdv[advertiser] = state;
    }
  }

  var result = {};  // advertiser -> {generated: [], confirmed: []}
  var notFound = [];

  // Guard against undefined or non-array records to avoid runtime errors.
  if (!Array.isArray(records)) {
    Logger.log('classifyResultsByClientSheet: records is not an array');
    records = [];
  }

  records.forEach(function(rec) {
    var adv = rec.advertiser || rec.advertiser_name || rec.advertiserName || '';
    var ad = rec.ad || rec.ad_name || rec.adName || '';
    var state = (mapByAdvAd[adv] && mapByAdvAd[adv][ad]) || mapByAdv[adv];
    if (!state) {
      notFound.push([adv, ad]);
      return;
    }

    var unix = state === '発生' ? rec.regist_unix : rec.apply_unix;
    var str = state === '発生' ? rec.regist : rec.apply;
    var d = null;
    if (unix) d = new Date(Number(unix) * 1000);
    else if (str) d = new Date(String(str).replace(' ', 'T'));
    if (!d || d < startDate || d > endDate) return;

    if (!result[adv]) result[adv] = { generated: [], confirmed: [] };
    if (state === '発生') result[adv].generated.push(rec);
    else result[adv].confirmed.push(rec);
  });

  if (notFound.length > 0) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('クライアント情報シートに記載がない成果が ' + notFound.length + ' 件あります。');
    var missSheet = ss.getSheetByName('該当無し') || ss.insertSheet('該当無し');
    missSheet.clearContents();
    missSheet.getRange(1, 1, 1, 2).setValues([['広告主名', '広告名']]);
    missSheet.getRange(2, 1, notFound.length, 2).setValues(notFound);
  }

  return result;
}

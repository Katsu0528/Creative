function summarizeResultsByAgency() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheets()[0];
  var start = inputSheet.getRange('B2').getValue();
  var end = inputSheet.getRange('C2').getValue();
  if (!(start instanceof Date) || !(end instanceof Date)) {
    SpreadsheetApp.getUi().alert('B2/C2 に日付が入力されていません。');
    return;
  }

  var baseUrl = 'https://otonari-asp.com/api/v1/m';
  var accessKey = 'agqnoournapf';
  var secretKey = '1kvu9dyv1alckgocc848socw';
  baseUrl = baseUrl.replace(/\/+$/, '');
  var headers = { 'X-Auth-Token': accessKey + ':' + secretKey };

  var params = [
    'apply_unix=between_date',
    'apply_unix_A_Y=' + start.getFullYear(),
    'apply_unix_A_M=' + (start.getMonth() + 1),
    'apply_unix_A_D=' + start.getDate(),
    'apply_unix_B_Y=' + end.getFullYear(),
    'apply_unix_B_M=' + (end.getMonth() + 1),
    'apply_unix_B_D=' + end.getDate(),
    'limit=500',
    'offset=0'
  ];

  var records = [];
  var offset = 0;
  while (true) {
    params[params.length - 1] = 'offset=' + offset;
    var url = baseUrl + '/action_log_raw/search?' + params.join('&');
    var response;
    try {
      response = UrlFetchApp.fetch(url, { method: 'get', headers: headers });
    } catch (e) {
      SpreadsheetApp.getUi().alert('API取得に失敗しました: ' + e);
      return;
    }
    var json = JSON.parse(response.getContentText());
    if (json.records && json.records.length) {
      records = records.concat(json.records);
    }
    var count = json.header && json.header.count ? json.header.count : 0;
    if (records.length >= count) {
      break;
    }
    offset += json.records.length;
  }

  var advertiserMap = {};
  var userMap = {};
  var promotionMap = {};

  function getAdvertiserName(id) {
    if (!id) return '';
    if (advertiserMap[id]) return advertiserMap[id];
    try {
      var url = baseUrl + '/advertiser/search?id=' + encodeURIComponent(id);
      var res = UrlFetchApp.fetch(url, { method: 'get', headers: headers });
      var json = JSON.parse(res.getContentText());
      var rec = Array.isArray(json.records) ? json.records[0] : json.records;
      advertiserMap[id] = (rec && (rec.company || rec.name)) || id;
    } catch (e) {
      advertiserMap[id] = id;
    }
    return advertiserMap[id];
  }

  function getUserName(id) {
    if (!id) return '';
    if (userMap[id]) return userMap[id];
    try {
      var url = baseUrl + '/user/search?id=' + encodeURIComponent(id);
      var res = UrlFetchApp.fetch(url, { method: 'get', headers: headers });
      var json = JSON.parse(res.getContentText());
      var rec = Array.isArray(json.records) ? json.records[0] : json.records;
      userMap[id] = (rec && rec.name) || id;
    } catch (e) {
      userMap[id] = id;
    }
    return userMap[id];
  }

  function getPromotionName(id) {
    if (!id) return '';
    if (promotionMap[id]) return promotionMap[id];
    try {
      var url = baseUrl + '/promotion/search?id=' + encodeURIComponent(id);
      var res = UrlFetchApp.fetch(url, { method: 'get', headers: headers });
      var json = JSON.parse(res.getContentText());
      var rec = Array.isArray(json.records) ? json.records[0] : json.records;
      promotionMap[id] = (rec && rec.name) || id;
    } catch (e) {
      promotionMap[id] = id;
    }
    return promotionMap[id];
  }

  var summary = {};
  records.forEach(function(rec) {
    var agency = getAdvertiserName(rec.advertiser || '');
    var manager = getUserName(rec.user || '');
    var ad = getPromotionName(rec.promotion || '');
    var unit = Number(rec.gross_action_cost || 0);
    var key = agency + '\u0000' + manager + '\u0000' + ad;
    if (!summary[key]) {
      summary[key] = { agency: agency, manager: manager, ad: ad, unit: unit, count: 0 };
    }
    summary[key].count++;
  });

  var outSheet = ss.getSheetByName('シート2') || ss.getSheetByName('Sheet2');
  if (!outSheet) {
    outSheet = ss.insertSheet('シート2');
  }
  outSheet.clearContents();
  outSheet.getRange(1, 1, 1, 5).setValues([[
    '会社名・担当者名',
    '広告名',
    '件数',
    'グロス単価',
    '金額'
  ]]);

  var rows = [];
  for (var k in summary) {
    var s = summary[k];
    rows.push([
      s.agency + ' ' + s.manager,
      s.ad,
      s.count,
      s.unit,
      s.count * s.unit
    ]);
  }

  if (rows.length > 0) {
    outSheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
}

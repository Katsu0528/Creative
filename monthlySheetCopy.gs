function monthlySheetCopy() {
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
  baseUrl = baseUrl.replace(/\/+$, '');
  var headers = { 'X-Auth-Token': accessKey + ':' + secretKey };

  var params = [
    'approve_unix=between_date',
    'approve_unix_A_Y=' + start.getFullYear(),
    'approve_unix_A_M=' + (start.getMonth() + 1),
    'approve_unix_A_D=' + start.getDate(),
    'approve_unix_B_Y=' + end.getFullYear(),
    'approve_unix_B_M=' + (end.getMonth() + 1),
    'approve_unix_B_D=' + end.getDate(),
    'state[]=2',
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

  var sheetName = '月次コピー';
  var outSheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  outSheet.clearContents();

  if (records.length === 0) {
    return;
  }

  var keys = Object.keys(records[0]);
  outSheet.getRange(1, 1, 1, keys.length).setValues([keys]);
  var rows = records.map(function(rec) {
    return keys.map(function(k) { return rec[k]; });
  });
  outSheet.getRange(2, 1, rows.length, keys.length).setValues(rows);
}

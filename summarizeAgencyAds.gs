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

  var advertiserMap = {};
  var userMap = {};
  var promotionMap = {};
  var mediaMap = {};

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

  function getMediaName(id) {
    if (!id) return '';
    if (mediaMap[id]) return mediaMap[id];
    try {
      var url = baseUrl + '/media/search?id=' + encodeURIComponent(id);
      var res = UrlFetchApp.fetch(url, { method: 'get', headers: headers });
      var json = JSON.parse(res.getContentText());
      var rec = Array.isArray(json.records) ? json.records[0] : json.records;
      mediaMap[id] = (rec && rec.name) || id;
    } catch (e) {
      mediaMap[id] = id;
    }
    return mediaMap[id];
  }

  var summary = {};
  var summary3 = {};
  records.forEach(function(rec) {
    var agency = getAdvertiserName(rec.advertiser || '');
    var manager = getUserName(rec.user || '');
    var ad = getPromotionName(rec.promotion || '');
    var affiliate = getMediaName(rec.media || '');
    var grossUnit = Number(rec.gross_action_cost || 0);
    var netUnit = Number(rec.net_action_cost || 0);
    var subject = rec.subject || '';
    var grossReward = Number(rec.gross_reward || 0);
    var netReward = Number(rec.net_reward || 0);

    var key3 = affiliate + '\u0000' + subject + '\u0000' + agency + '\u0000' + grossReward + '\u0000' + netReward + '\u0000' + ad;
    if (!summary3[key3]) {
      summary3[key3] = {
        affiliate: affiliate,
        subject: subject,
        advertiser: agency,
        grossReward: grossReward,
        netReward: netReward,
        ad: ad,
        confirmedCount: 0,
        confirmedGross: 0
      };
    }
    if (Number(rec.state) === 2) {
      summary3[key3].confirmedCount++;
      summary3[key3].confirmedGross += grossReward;
    }

    if (Number(rec.state) !== 2) return;
    var key = agency + '\u0000' + manager + '\u0000' + ad + '\u0000' + affiliate;
    if (!summary[key]) {
      summary[key] = {
        agency: agency,
        manager: manager,
        ad: ad,
        affiliate: affiliate,
        grossUnit: grossUnit,
        netUnit: netUnit,
        count: 0
      };
    }
    summary[key].count++;
  });

  var outSheet = ss.getSheetByName('シート2') || ss.getSheetByName('Sheet2');
  if (!outSheet) {
    outSheet = ss.insertSheet('シート2');
  }
  outSheet.clearContents();
  outSheet.getRange(1, 1, 1, 7).setValues([[
    '会社名・担当者名',
    '広告名',
    'アフィリエイター',
    '件数',
    'グロス単価',
    'ネット単価',
    '金額'
  ]]);

  var rows = [];
  for (var k in summary) {
    var s = summary[k];
    rows.push([
      s.agency + ' ' + s.manager,
      s.ad,
      s.affiliate,
      s.count,
      s.grossUnit,
      s.netUnit,
      s.count * s.grossUnit
    ]);
  }

  if (rows.length > 0) {
    outSheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }

  var outSheet3 = ss.getSheetByName('シート3') || ss.getSheetByName('Sheet3');
  if (!outSheet3) {
    outSheet3 = ss.insertSheet('シート3');
  }
  outSheet3.clearContents();
  outSheet3.getRange(1, 1, 1, 8).setValues([[
    'アフィリエイター',
    '成果名',
    '広告主',
    '成果報酬額（グロス）[円]',
    '成果報酬額（ネット）[円]',
    '広告',
    '確定成果数[件]',
    '確定成果額（グロス）[円]'
  ]]);

  var rows3 = [];
  for (var k3 in summary3) {
    var s3 = summary3[k3];
    rows3.push([
      s3.affiliate,
      s3.subject,
      s3.advertiser,
      s3.grossReward,
      s3.netReward,
      s3.ad,
      s3.confirmedCount,
      s3.confirmedGross
    ]);
  }

  if (rows3.length > 0) {
    outSheet3.getRange(2, 1, rows3.length, 8).setValues(rows3);
  }
}

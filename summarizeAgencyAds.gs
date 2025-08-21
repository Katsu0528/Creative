function summarizeResultsByAgency(targetSheetName) {
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
    'state[]=1',
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
        generatedCount: 0,
        generatedGross: 0,
        confirmedCount: 0,
        confirmedGross: 0
      };
    }
    if (Number(rec.state) === 1) {
      summary3[key3].generatedCount++;
      summary3[key3].generatedGross += grossReward;
    } else if (Number(rec.state) === 2) {
      summary3[key3].confirmedCount++;
      summary3[key3].confirmedGross += grossReward;
    }

    if (Number(rec.state) !== 1) return;
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

  var summarySheet = ss.getSheetByName(targetSheetName) || ss.getSheetByName('2025年8月対応_データ格納');
  if (summarySheet) {
    summarySheet.getRange(1, 15, 1, 5).setValues([[
      'アフィリエイター',
      '成果名',
      '広告主',
      '成果報酬額（グロス）[円]',
      '成果報酬額（ネット）[円]'
    ]]);
    summarySheet.getRange(1, 23, 1, 5).setValues([[
      '広告',
      '発生成果数[件]',
      '発生成果額（グロス）[円]',
      '確定成果数[件]',
      '確定成果額（グロス）[円]'
    ]]);

    var rowsLeft = [];
    var rowsRight = [];
    for (var k3 in summary3) {
      var s3 = summary3[k3];
      rowsLeft.push([
        s3.affiliate,
        s3.subject,
        s3.advertiser,
        s3.grossReward,
        s3.netReward
      ]);
      rowsRight.push([
        s3.ad,
        s3.generatedCount,
        s3.generatedGross,
        s3.confirmedCount,
        s3.confirmedGross
      ]);
    }
    if (rowsLeft.length > 0) {
      summarySheet.getRange(2, 15, rowsLeft.length, 5).setValues(rowsLeft);
      summarySheet.getRange(2, 23, rowsRight.length, 5).setValues(rowsRight);
    }
  }
}

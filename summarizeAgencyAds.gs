var SPREADSHEET_ID = '1qkae2jGCUlykwL-uTf0_eaBGzon20RCC-wBVijyvm8s';

function summarizeApprovedResultsByAgency(targetSheetName) {
  Logger.log('summarizeApprovedResultsByAgency: start' + (targetSheetName ? ' target=' + targetSheetName : ''));
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var inputSheet = ss.getSheets()[0];
  var start = inputSheet.getRange('B2').getValue();
  var end = inputSheet.getRange('C2').getValue();
  if (!(start instanceof Date) || !(end instanceof Date)) {
    SpreadsheetApp.getUi().alert('B2/C2 に日付が入力されていません。');
    Logger.log('summarizeApprovedResultsByAgency: invalid date range');
    return;
  }
  Logger.log('summarizeApprovedResultsByAgency: fetching records from ' + start + ' to ' + end);

  var baseUrl = 'https://otonari-asp.com/api/v1/m';
  var accessKey = 'agqnoournapf';
  var secretKey = '1kvu9dyv1alckgocc848socw';
  baseUrl = baseUrl.replace(/\/+$/, '');
  var headers = { 'X-Auth-Token': accessKey + ':' + secretKey };

  var params = [
    'approval_unix=between_date',
    'approval_unix_A_Y=' + start.getFullYear(),
    'approval_unix_A_M=' + (start.getMonth() + 1),
    'approval_unix_A_D=' + start.getDate(),
    'approval_unix_B_Y=' + end.getFullYear(),
    'approval_unix_B_M=' + (end.getMonth() + 1),
    'approval_unix_B_D=' + end.getDate(),
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
      Logger.log('summarizeApprovedResultsByAgency: API fetch failed');
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
  Logger.log('summarizeApprovedResultsByAgency: fetched ' + records.length + ' record(s)');

  var advertiserMap = {};
  var userMap = {};
  var promotionMap = {};
  var mediaMap = {};

  var advertiserSet = {};
  var userSet = {};
  var promotionSet = {};
  var mediaSet = {};

  records.forEach(function(rec) {
    if (rec.advertiser) advertiserSet[rec.advertiser] = true;
    if (rec.user) userSet[rec.user] = true;
    if (rec.promotion) promotionSet[rec.promotion] = true;
    if (rec.media) mediaSet[rec.media] = true;
  });

  function fetchNames(ids, endpoint, map, nameResolver) {
    for (var i = 0; i < ids.length; i += 100) {
      var batch = ids.slice(i, i + 100);
      var requests = batch.map(function(id) {
        return { url: baseUrl + '/' + endpoint + '/search?id=' + encodeURIComponent(id), method: 'get', headers: headers };
      });
      var responses = UrlFetchApp.fetchAll(requests);
      responses.forEach(function(res, idx) {
        var id = batch[idx];
        try {
          var json = JSON.parse(res.getContentText());
          var rec = Array.isArray(json.records) ? json.records[0] : json.records;
          map[id] = nameResolver(rec) || id;
        } catch (e) {
          map[id] = id;
        }
      });
    }
  }

  fetchNames(Object.keys(advertiserSet), 'advertiser', advertiserMap, function(rec) {
    return rec && (rec.company || rec.name);
  });
  fetchNames(Object.keys(userSet), 'user', userMap, function(rec) {
    return rec && rec.name;
  });
  fetchNames(Object.keys(promotionSet), 'promotion', promotionMap, function(rec) {
    return rec && rec.name;
  });
  fetchNames(Object.keys(mediaSet), 'media', mediaMap, function(rec) {
    return rec && rec.name;
  });

  var summary = {};
  var summary3 = {};
  records.forEach(function(rec) {
    var agency = rec.advertiser ? (advertiserMap[rec.advertiser] || rec.advertiser) : '';
    var manager = rec.user ? (userMap[rec.user] || rec.user) : '';
    var ad = rec.promotion ? (promotionMap[rec.promotion] || rec.promotion) : '';
    var affiliate = rec.media ? (mediaMap[rec.media] || rec.media) : '';
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
  Logger.log('summarizeApprovedResultsByAgency: wrote ' + rows.length + ' row(s) to ' + outSheet.getName());

  var summarySheet = null;
  if (targetSheetName) {
    summarySheet = ss.getSheetByName(targetSheetName);
    Logger.log('summarizeApprovedResultsByAgency: using target sheet ' + targetSheetName);
    if (!summarySheet) {
      Logger.log('summarizeApprovedResultsByAgency: target sheet not found');
    }
  } else {
    var latestDate = null;
    var pattern = /^(\d{4})年(\d{1,2})月対応_データ格納$/;
    ss.getSheets().forEach(function(sheet) {
      var match = sheet.getName().match(pattern);
      if (match) {
        var d = new Date(parseInt(match[1], 10), parseInt(match[2], 10) - 1);
        if (!latestDate || d.getTime() > latestDate.getTime()) {
          latestDate = d;
          summarySheet = sheet;
        }
      }
    });
    if (summarySheet) {
      Logger.log('summarizeApprovedResultsByAgency: detected latest data sheet ' + summarySheet.getName());
    } else {
      var fallbackName = Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyy年M月対応_データ格納');
      summarySheet = ss.getSheetByName(fallbackName);
      Logger.log('summarizeApprovedResultsByAgency: fallback to sheet ' + (summarySheet ? summarySheet.getName() : 'none'));
    }
  }
  if (summarySheet) {
    Logger.log('summarizeApprovedResultsByAgency: writing summary data to ' + summarySheet.getName());
    summarySheet.getRange(1, 15, 1, 5).setValues([[
      'アフィリエイター',
      '成果名',
      '広告主',
      '成果報酬額（グロス）[円]',
      '成果報酬額（ネット）[円]'
    ]]);
    summarySheet.getRange(1, 23, 1, 3).setValues([[
      '広告',
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
        s3.confirmedCount,
        s3.confirmedGross
      ]);
    }
    if (rowsLeft.length > 0) {
      summarySheet.getRange(2, 15, rowsLeft.length, 5).setValues(rowsLeft);
      summarySheet.getRange(2, 23, rowsRight.length, 3).setValues(rowsRight);
    }
    Logger.log('summarizeApprovedResultsByAgency: wrote ' + rowsLeft.length + ' row(s) to ' + summarySheet.getName());
  }
  Logger.log('summarizeApprovedResultsByAgency: complete');
}

var SPREADSHEET_ID = '1qkae2jGCUlykwL-uTf0_eaBGzon20RCC-wBVijyvm8s';
var PROGRESS_KEY = 'SUMMARY_PROGRESS';
var TOTAL_STEPS = 7;

function showProgress_() {
  setProgress_(0, '処理開始', 0, TOTAL_STEPS);
  try {
    var ui = SpreadsheetApp.getUi();
    var html = HtmlService.createHtmlOutput(
      '<html><body>' +
        '<progress id="p" max="100" value="0" style="width:100%"></progress>' +
        '<div id="status" style="text-align:center;margin-top:4px;font-family:sans-serif;"></div>' +
        '<script>' +
          '(function poll(){google.script.run.withSuccessHandler(function(v){' +
          'document.getElementById("p").value=v.value;' +
          'var t=v.message||"";' +
          'if(v.total){t+=" ("+v.current+"/"+v.total+")";}' +
          'document.getElementById("status").innerText=t;' +
          'if(v.value<100){setTimeout(poll,500);}else{google.script.host.close();}' +
          '}).getProgress();})();' +
        '</script>' +
      '</body></html>'
    );
    ui.showModelessDialog(html, '処理中');
  } catch (e) {
    Logger.log('showProgress_: UI not available: ' + e);
  }
}

function setProgress_(v, message, current, total) {
  var data = {
    value: v,
    message: message || '',
    current: current || 0,
    total: total || 0
  };
  PropertiesService.getScriptProperties().setProperty(PROGRESS_KEY, JSON.stringify(data));
}

function getProgress() {
  var prop = PropertiesService.getScriptProperties().getProperty(PROGRESS_KEY);
  return prop ? JSON.parse(prop) : { value: 0, message: '', current: 0, total: 0 };
}

function alertUi_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log('alertUi_: ' + message);
  }
}

function summarizeApprovedResultsByAgency(targetSheetName) {
  Logger.log('summarizeApprovedResultsByAgency: start' + (targetSheetName ? ' target=' + targetSheetName : ''));
  showProgress_();
  try {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var dateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var start = dateSheet.getRange('B2').getValue();
  var end = dateSheet.getRange('C2').getValue();
  if (!(start instanceof Date) || !(end instanceof Date)) {
    alertUi_('B2/C2 に日付が入力されていません。');
    Logger.log('summarizeApprovedResultsByAgency: invalid date range');
    setProgress_(100, 'エラー: 日付が正しく入力されていません', 0, TOTAL_STEPS);
    return;
  }
  start.setHours(0, 0, 0, 0);
  end.setHours(23, 59, 59, 999);
  Logger.log('summarizeApprovedResultsByAgency: fetching records from ' + start + ' to ' + end);
  setProgress_(10, '期間チェック完了', 1, TOTAL_STEPS);

  var baseUrl = 'https://otonari-asp.com/api/v1/m';
  var accessKey = 'agqnoournapf';
  var secretKey = '1kvu9dyv1alckgocc848socw';
  baseUrl = baseUrl.replace(/\/+$/, '');
  var headers = { 'X-Auth-Token': accessKey + ':' + secretKey };
  function fetchRecords(dateField, states) {
    var apiEnd = new Date(end.getTime());
    apiEnd.setDate(apiEnd.getDate() + 1);
    Logger.log('fetchRecords: ' + dateField + ' を between_date で ' + start + ' ～ ' + end +
      '（API検索は翌日0時まで）' +
      (states && states.length ? '、state=' + states.join(',') : '、state 指定なし') + ' の条件で検索');
    var params = [
      dateField + '=between_date',
      dateField + '_A_Y=' + start.getFullYear(),
      dateField + '_A_M=' + (start.getMonth() + 1),
      dateField + '_A_D=' + start.getDate(),
      dateField + '_B_Y=' + apiEnd.getFullYear(),
      dateField + '_B_M=' + (apiEnd.getMonth() + 1),
      dateField + '_B_D=' + apiEnd.getDate(),
      'limit=500'
    ];
    if (states) {
      states.forEach(function(s) {
        params.push('state[]=' + s);
      });
    }
    var baseParams = params.join('&');
    var url = baseUrl + '/action_log_raw/search?' + baseParams + '&offset=0';
    var response;
    for (var attempt = 0; attempt < 3; attempt++) {
        try {
          response = UrlFetchApp.fetch(url, { method: 'get', headers: headers });
          break;
        } catch (e) {
          if (attempt === 2) {
            alertUi_('API取得に失敗しました: ' + e);
            Logger.log('summarizeApprovedResultsByAgency: API fetch failed at ' + url);
            return null;
          }
          Utilities.sleep(1000 * Math.pow(2, attempt));
        }
    }
    var json = JSON.parse(response.getContentText());
    var result = json.records && json.records.length ? json.records : [];
    var count = json.header && json.header.count ? json.header.count : result.length;
    var fetched = result.length;
    if (fetched < count) {
      var requests = [];
      for (var offset = fetched; offset < count; offset += 500) {
        requests.push({
          url: baseUrl + '/action_log_raw/search?' + baseParams + '&offset=' + offset,
          method: 'get',
          headers: headers
        });
      }
      var responses = UrlFetchApp.fetchAll(requests);
      responses.forEach(function(res) {
        try {
          var j = JSON.parse(res.getContentText());
          if (j.records && j.records.length) {
            result = result.concat(j.records);
          }
        } catch (e) {}
      });
    }
    return result;
  }

  function fetchGeneratedRecords() {
    return fetchRecords('regist_unix');
  }

  function fetchConfirmedRecords() {
    // 確定成果は承認日時で抽出
    return fetchRecords('apply_unix', ['2']);
  }

  function filterRecords(records, unixField, dateField) {
    return records.filter(function(rec) {
      var d = null;
      if (rec[unixField]) {
        d = new Date(Number(rec[unixField]) * 1000);
      } else if (rec[dateField]) {
        d = new Date(rec[dateField]);
      }
      return d && d.getTime() >= start.getTime() && d.getTime() <= end.getTime();
    });
  }

  var generatedRecords = fetchGeneratedRecords();
  if (generatedRecords === null) { setProgress_(100, 'エラー: 発生成果の取得に失敗しました', 2, TOTAL_STEPS); return; }
  var generatedFetched = generatedRecords.length;
  generatedRecords = filterRecords(generatedRecords, 'regist_unix', 'regist_at');
  Logger.log('発生成果の取得ロジック: regist_unix または regist_at が期間内のレコードを対象。API取得件数=' + generatedFetched + '件、フィルタ後=' + generatedRecords.length + '件');
  setProgress_(30, '発生成果取得完了', 2, TOTAL_STEPS);
  var confirmedRecords = fetchConfirmedRecords();
  if (confirmedRecords === null) { setProgress_(100, 'エラー: 確定成果の取得に失敗しました', 3, TOTAL_STEPS); return; }
  var confirmedFetched = confirmedRecords.length;
  confirmedRecords = filterRecords(confirmedRecords, 'apply_unix', 'apply_at');
  Logger.log('確定成果の取得ロジック: apply_unix または apply_at が期間内で state=2 のレコードを対象。API取得件数=' + confirmedFetched + '件、フィルタ後=' + confirmedRecords.length + '件');
  setProgress_(50, '確定成果取得完了', 3, TOTAL_STEPS);
  var records = generatedRecords.concat(confirmedRecords);
  Logger.log('summarizeApprovedResultsByAgency: fetched ' + generatedRecords.length + ' generated record(s) and ' + confirmedRecords.length + ' confirmed record(s)');

  var advertiserMap = {};
  var userMap = {};
  var promotionMap = {};
  var promotionAdvertiserMap = {};
  var mediaMap = {};
  var mediaInfoMap = {};

  var advertiserSet = {};
  var userSet = {};
  var promotionSet = {};
  var mediaSet = {};

  records.forEach(function(rec) {
    if (rec.advertiser || rec.advertiser === 0) advertiserSet[rec.advertiser] = true;
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

  function fetchPromotions(ids) {
    for (var i = 0; i < ids.length; i += 100) {
      var batch = ids.slice(i, i + 100);
      var requests = batch.map(function(id) {
        return { url: baseUrl + '/promotion/search?id=' + encodeURIComponent(id), method: 'get', headers: headers };
      });
      var responses = UrlFetchApp.fetchAll(requests);
      responses.forEach(function(res, idx) {
        var id = batch[idx];
        try {
          var json = JSON.parse(res.getContentText());
          var rec = Array.isArray(json.records) ? json.records[0] : json.records;
          promotionMap[id] = rec && rec.name;
          if (rec && (rec.advertiser || rec.advertiser === 0)) {
            promotionAdvertiserMap[id] = rec.advertiser;
            advertiserSet[rec.advertiser] = true;
          }
        } catch (e) {
          promotionMap[id] = id;
        }
      });
    }
  }

  fetchPromotions(Object.keys(promotionSet));

  fetchNames(Object.keys(advertiserSet), 'advertiser', advertiserMap, function(rec) {
    return rec && (rec.company || rec.name);
  });

  // メディア情報を取得し、会社名と担当者IDを保持
  fetchNames(Object.keys(mediaSet), 'media', mediaInfoMap, function(rec) {
    if (!rec) return { company: '', user: '' };
    if (rec.user) userSet[rec.user] = true;
    return { company: rec.name || '', user: rec.user || '' };
  });

  // メディアの担当者を含めたユーザー情報を取得
  fetchNames(Object.keys(userSet), 'user', userMap, function(rec) {
    return rec && rec.name;
  });

  // 会社名と担当者名を結合したアフィリエイター名のマップを作成
  Object.keys(mediaInfoMap).forEach(function(id) {
    var info = mediaInfoMap[id];
    var person = info.user ? (userMap[info.user] || '') : '';
    mediaMap[id] = info.company && person ? info.company + ' ' + person : (info.company || person);
  });
  setProgress_(60, 'マスタ情報取得完了', 4, TOTAL_STEPS);

  var adListSheet = ss.getSheetByName('【毎月更新】広告一覧');
  if (!adListSheet) {
    adListSheet = ss.insertSheet('【毎月更新】広告一覧');
  }
  adListSheet.clearContents();
  adListSheet.getRange(1, 1, 1, 2).setValues([[
    '広告名',
    '広告主名'
  ]]);
  var adRows = [];
  Object.keys(promotionMap).forEach(function(pid) {
    var adName = promotionMap[pid];
    var advId = promotionAdvertiserMap[pid];
    var advertiserName = (advId || advId === 0) ? (advertiserMap[advId] || advId) : '';
    adRows.push([adName, advertiserName]);
  });
  if (adRows.length > 0) {
    adListSheet.getRange(2, 1, adRows.length, 2).setValues(adRows);
  }
  Logger.log('summarizeApprovedResultsByAgency: wrote ' + adRows.length + ' row(s) to 【毎月更新】広告一覧');
  setProgress_(70, '広告一覧作成完了', 5, TOTAL_STEPS);

  var summary = {};
  var summary3 = {};
  var summaryByAd = {};
  generatedRecords.forEach(function(rec) {
    var advId = (rec.advertiser || rec.advertiser === 0) ? rec.advertiser : promotionAdvertiserMap[rec.promotion];
    var agency = advId ? (advertiserMap[advId] || advId) : '';
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
        generatedCount: 0,
        generatedGross: 0,
        confirmedCount: 0,
        confirmedGross: 0
      };
    }
    summary3[key3].generatedCount++;
    summary3[key3].generatedGross += grossReward;

    var keyAd = ad + '\u0000' + grossUnit + '\u0000' + netUnit;
    if (!summaryByAd[keyAd]) {
      summaryByAd[keyAd] = {
        ad: ad,
        grossUnit: grossUnit,
        netUnit: netUnit,
        generatedCount: 0,
        generatedGross: 0,
        confirmedCount: 0,
        confirmedGross: 0
      };
    }
    summaryByAd[keyAd].generatedCount++;
    summaryByAd[keyAd].generatedGross += grossReward;
  });

  confirmedRecords.forEach(function(rec) {
    var advId = (rec.advertiser || rec.advertiser === 0) ? rec.advertiser : promotionAdvertiserMap[rec.promotion];
    var agency = advId ? (advertiserMap[advId] || advId) : '';
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
        generatedCount: 0,
        generatedGross: 0,
        confirmedCount: 0,
        confirmedGross: 0
      };
    }
    summary3[key3].confirmedCount++;
    summary3[key3].confirmedGross += grossReward;

    var keyAd = ad + '\u0000' + grossUnit + '\u0000' + netUnit;
    if (!summaryByAd[keyAd]) {
      summaryByAd[keyAd] = {
        ad: ad,
        grossUnit: grossUnit,
        netUnit: netUnit,
        generatedCount: 0,
        generatedGross: 0,
        confirmedCount: 0,
        confirmedGross: 0
      };
    }
    summaryByAd[keyAd].confirmedCount++;
    summaryByAd[keyAd].confirmedGross += grossReward;

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
  setProgress_(80, '集計表作成完了', 6, TOTAL_STEPS);

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
    summarySheet.getRange(1, 23, 1, 5).setValues([[
      '広告',
      '発生成果数[件]',
      '発生成果額（グロス）[円]',
      '確定成果数[件]',
      '確定成果額（グロス）[円]'
    ]]);

    var rowsLeft = [];
    for (var k3 in summary3) {
      var s3 = summary3[k3];
      rowsLeft.push([
        s3.affiliate,
        s3.subject,
        s3.advertiser,
        s3.grossReward,
        s3.netReward
      ]);
    }

    var rowsRight = [];
    for (var kAd in summaryByAd) {
      var sa = summaryByAd[kAd];
      rowsRight.push([
        sa.ad,
        sa.generatedCount,
        sa.generatedGross,
        sa.confirmedCount,
        sa.confirmedGross
      ]);
    }

    if (rowsLeft.length > 0) {
      summarySheet.getRange(2, 15, rowsLeft.length, 5).setValues(rowsLeft);
    }
    if (rowsRight.length > 0) {
      summarySheet.getRange(2, 23, rowsRight.length, 5).setValues(rowsRight);
    }
    Logger.log('summarizeApprovedResultsByAgency: wrote ' + rowsLeft.length + ' row(s) (left) and ' + rowsRight.length + ' row(s) (right) to ' + summarySheet.getName());
  }
  setProgress_(100, '処理完了', TOTAL_STEPS, TOTAL_STEPS);
  Logger.log('summarizeApprovedResultsByAgency: complete');
  } catch (e) {
    Logger.log('summarizeApprovedResultsByAgency: error ' + e + (e.stack ? '\n' + e.stack : ''));
    alertUi_('エラーが発生しました: ' + e);
    setProgress_(100, 'エラー: ' + e, 0, TOTAL_STEPS);
  }
}

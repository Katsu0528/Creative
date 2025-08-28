// Enable strict mode to surface syntax issues early and fail fast when
// this script is executed outside the Google Apps Script runtime.
'use strict';

// Spreadsheet for outputting summarized data
var TARGET_SPREADSHEET_ID = '1qkae2jGCUlykwL-uTf0_eaBGzon20RCC-wBVijyvm8s';
// Spreadsheet that holds the date range used for the summary
var DATE_SPREADSHEET_ID = '13zQMfgfYlec1BOo0LwWZUerQD9Fm0Fkzav8Z20d5eDE';
var DATE_SHEET_NAME = '日付';
// Track last shown progress percentage to avoid excessive updates
var lastProgressPercent_ = -1;
// Cell used for in-sheet progress updates
var PROGRESS_CELL_ = 'A1';

function alertUi_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log('alertUi_: ' + message);
  }
}

function initProgress_() {
  // Reset progress tracking
  lastProgressPercent_ = -1;
  showProgress_(0, 1);
}

function clearProgress_() {
  try {
    SpreadsheetApp.getActiveSheet().getRange(PROGRESS_CELL_).clearContent();
  } catch (e) {}
}

function showProgress_(current, total) {
  if (total <= 0) return;
  var percent = Math.floor((current / total) * 100);
  if (percent === lastProgressPercent_) return;
  lastProgressPercent_ = percent;
  var barLength = 20;
  var filled = Math.round(barLength * current / total);
  var bar = '[' + '■'.repeat(filled) + '□'.repeat(barLength - filled) + '] ' +
            percent + '% (' + current + '/' + total + ')';
  try {
    SpreadsheetApp.getActiveSheet().getRange(PROGRESS_CELL_).setValue(bar);
  } catch (e) {
    Logger.log(bar);
  }
}

// Convert half-width spaces to full-width spaces for consistent name matching.
function toFullWidthSpace_(str) {
  return typeof str === 'string' ? str.replace(/ /g, '　') : str;
}


// Normalize advertiser IDs by removing whitespace, converting full-width digits
// to half-width, and stripping leading zeros so that logically identical IDs
// match even if formatted differently in sheets.
function normalizeAdvId_(id) {
  if (id === 0 || id) {
    var str = String(id);
    str = str.replace(/[０-９]/g, function(c) {
      return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
    });
    str = str.replace(/\s+/g, '');
    if (/^\d+$/.test(str)) {
      // Remove leading zeros without using parseInt to avoid precision loss
      str = str.replace(/^0+/, '');
      if (str === '') str = '0';
    }
    return str;
  }
  return '';
}

function summarizeApprovedResultsByAgency(targetSheetName) {
  Logger.log('summarizeApprovedResultsByAgency: start' + (targetSheetName ? ' target=' + targetSheetName : ''));
  try {
    lastProgressPercent_ = -1;
    var counts = { confirmed: 0, generated: 0, adListRows: 0, summaryLeftRows: 0, summaryRightRows: 0, summarySheetName: '' };
  var targetSs = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  var dateSs = SpreadsheetApp.openById(DATE_SPREADSHEET_ID);
  var dateSheet = dateSs.getSheetByName(DATE_SHEET_NAME);
  var start = dateSheet.getRange('B2').getValue();
  var end = dateSheet.getRange('C2').getValue();
  if (!(start instanceof Date) || !(end instanceof Date)) {
    alertUi_('B2/C2 に日付が入力されていません。');
    Logger.log('summarizeApprovedResultsByAgency: invalid date range');
    throw new Error('日付が正しく入力されていません');
  }
  start.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);
  Logger.log('summarizeApprovedResultsByAgency: fetching records from ' + start + ' to ' + end);

  alertUi_('対象期間: ' +
    Utilities.formatDate(start, 'Asia/Tokyo', 'yyyy-MM-dd') + ' ～ ' +
    Utilities.formatDate(end, 'Asia/Tokyo', 'yyyy-MM-dd'));

  var baseUrl = 'https://otonari-asp.com/api/v1/m';
  var accessKey = 'agqnoournapf';
  var secretKey = '1kvu9dyv1alckgocc848socw';
  baseUrl = baseUrl.replace(/\/+$/, '');
    var headers = { 'X-Auth-Token': accessKey + ':' + secretKey };
    function fetchRecords(dateField, states) {
    Logger.log('fetchRecords: ' + dateField + ' を between_date で ' + start + ' ～ ' + end +
      (states && states.length ? '、state=' + states.join(',') : '、state 指定なし') + ' の条件で検索');
    var params = [
      dateField + '=between_date',
      dateField + '_A_Y=' + start.getFullYear(),
      dateField + '_A_M=' + (start.getMonth() + 1),
      dateField + '_A_D=' + start.getDate(),
      dateField + '_B_Y=' + end.getFullYear(),
      dateField + '_B_M=' + (end.getMonth() + 1),
      dateField + '_B_D=' + end.getDate(),
      'limit=500'
    ];
    if (states && states.length) {
      states.forEach(function(s) {
        params.push('state[]=' + s);
        params.push('state=' + s);
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
    function getRecordDate_(rec, unixField, dateField) {
      var d = null;
      var unixVal = rec[unixField];
      if (unixVal !== undefined && unixVal !== null && unixVal !== '') {
      d = new Date(Number(unixVal) * 1000);
    } else if (rec[dateField]) {
      var str = String(rec[dateField]).replace(' ', 'T');
      d = new Date(str);
      if (isNaN(d.getTime())) {
        d = new Date(str.replace(/-/g, '/'));
      }
    }
    return d;
  }

    var generatedRecords = fetchRecords('regist_unix', [1]);
    if (generatedRecords === null) {
        alertUi_('発生成果の取得に失敗しました');
        throw new Error('発生成果の取得に失敗しました');
    }

    var confirmedRecords = fetchRecords('apply_unix', [1]);
    if (confirmedRecords === null) {
        alertUi_('確定成果の取得に失敗しました');
        throw new Error('確定成果の取得に失敗しました');
    }

  counts.generated = generatedRecords.length;
  counts.confirmed = confirmedRecords.length;
  Logger.log('fetchGeneratedRecords: 取得した件数=' + counts.generated + '件');
  Logger.log('fetchConfirmedRecords: 取得した件数=' + counts.confirmed + '件');
  if (confirmedRecords.length > 0) {
    Logger.log('例: 確定成果の一部: ' + JSON.stringify(confirmedRecords[0]));
  }
  alertUi_('発生件数: ' + counts.generated + ' 件');
  alertUi_('確定件数: ' + counts.confirmed + ' 件');

  Logger.log('summarizeApprovedResultsByAgency: fetched ' + counts.generated + ' generated record(s) and ' + counts.confirmed + ' confirmed record(s)');

  // Combine generated and confirmed records for subsequent lookups
  var allRecords = generatedRecords.concat(confirmedRecords);

  var advertiserMap = {};
  var advertiserInfoMap = {};
  var userMap = {};
  var promotionMap = {};
  var promotionAdvertiserMap = {};
  var adToAdvIdMap = {};
  var mediaMap = {};
  var mediaInfoMap = {};

  var advertiserSet = {};
  var userSet = {};
  var promotionSet = {};
  var mediaSet = {};

  allRecords.forEach(function(rec) {
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

  fetchNames(Object.keys(advertiserSet), 'advertiser', advertiserInfoMap, function(rec) {
    if (!rec) return { company: '', name: '' };
    return { company: rec.company || '', name: rec.name || '' };
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

  // 会社名と担当者名を結合した広告主名のマップを作成
  Object.keys(advertiserInfoMap).forEach(function(id) {
    var info = advertiserInfoMap[id];
    var company = info.company || '';
    var person = info.name || '';
    advertiserMap[id] = toFullWidthSpace_(company && person ? company + ' ' + person : (company || person));
  });

  // 会社名と担当者名を結合したアフィリエイター名のマップを作成
  Object.keys(mediaInfoMap).forEach(function(id) {
    var info = mediaInfoMap[id];
    var person = info.user ? (userMap[info.user] || '') : '';
    mediaMap[id] = info.company && person ? info.company + ' ' + person : (info.company || person);
  });

  var resultSheet = dateSs.getSheetByName('シート4') || dateSs.getSheetByName('Sheet4');
  if (!resultSheet) {
    resultSheet = dateSs.insertSheet('シート4');
  }
  resultSheet.clearContents();
  resultSheet.getRange(1, 1, 1, 5).setValues([[
    '確定/発生日時',
    '承認状態',
    '広告主名',
    '広告名',
    'アフィリエイター名'
  ]]);
  var resultRows = allRecords.map(function(rec) {
    var advId = (rec.advertiser || rec.advertiser === 0) ? rec.advertiser : promotionAdvertiserMap[rec.promotion];
    var advertiserName = advId ? (advertiserMap[advId] || advId) : '';
    var adName = rec.promotion ? (promotionMap[rec.promotion] || rec.promotion) : '';
    var affiliateName = rec.media ? (mediaMap[rec.media] || rec.media) : '';
    var date = getRecordDate_(rec, 'apply_unix', 'apply_at') ||
               getRecordDate_(rec, 'regist_unix', 'regist_at');
    return [
      date,
      rec.state,
      advertiserName,
      adName,
      affiliateName
    ];
  });
  if (resultRows.length > 0) {
    resultSheet.getRange(2, 1, resultRows.length, 5).setValues(resultRows);
  }

  var adListSheet = targetSs.getSheetByName('【毎月更新】広告一覧');
  if (!adListSheet) {
    adListSheet = targetSs.insertSheet('【毎月更新】広告一覧');
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
    var advertiserName = (advId || advId === 0) ? (advertiserMap[advId] || '') : '';
    adRows.push([adName, advertiserName]);
    adToAdvIdMap[adName] = advId;
  });
  if (adRows.length > 0) {
    adListSheet.getRange(2, 1, adRows.length, 2).setValues(adRows);
  }
  counts.adListRows = adRows.length;
  Logger.log('summarizeApprovedResultsByAgency: wrote ' + adRows.length + ' row(s) to 【毎月更新】広告一覧');

  var rowsLeft = [];
  var summaryByAd = {};
  var totalRecords = generatedRecords.length + confirmedRecords.length;
  var processedRecords = 0;
  initProgress_();
  generatedRecords.forEach(function(rec) {
    processedRecords++;
    showProgress_(processedRecords, totalRecords);
    var ad = rec.promotion ? (promotionMap[rec.promotion] || rec.promotion) : '';
    var grossUnit = Number(rec.gross_action_cost || 0);
    var netUnit = Number(rec.net_action_cost || 0);
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
    summaryByAd[keyAd].generatedGross += Number(rec.gross_reward || 0);
  });

  confirmedRecords.forEach(function(rec) {
    processedRecords++;
    showProgress_(processedRecords, totalRecords);
    var advId = (rec.advertiser || rec.advertiser === 0) ? rec.advertiser : promotionAdvertiserMap[rec.promotion];
    var agency = advId ? (advertiserMap[advId] || advId) : '';
    var ad = rec.promotion ? (promotionMap[rec.promotion] || rec.promotion) : '';
    var affiliate = rec.media ? (mediaMap[rec.media] || rec.media) : '';
    var grossUnit = Number(rec.gross_action_cost || 0);
    var netUnit = Number(rec.net_action_cost || 0);
    var subject = rec.subject || '';
    var grossReward = Number(rec.gross_reward || 0);
    var netReward = Number(rec.net_reward || 0);
    rowsLeft.push([affiliate, subject, agency, grossReward, netReward]);

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
  });

  var summarySheet = null;
  if (targetSheetName) {
    summarySheet = targetSs.getSheetByName(targetSheetName);
    Logger.log('summarizeApprovedResultsByAgency: using target sheet ' + targetSheetName);
    if (!summarySheet) {
      Logger.log('summarizeApprovedResultsByAgency: target sheet not found');
    }
  } else {
    var latestDate = null;
    var pattern = /^(\d{4})年(\d{1,2})月対応_データ格納$/;
    targetSs.getSheets().forEach(function(sheet) {
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
      summarySheet = targetSs.getSheetByName(fallbackName);
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
      var advIds = rowsRight.map(function(row) {
        var adName = row[0];
        return adToAdvIdMap[adName] || '';
      });
      summarySheet.getRange(2, 28, advIds.length, 1).setValues(advIds.map(function(id){ return [id]; }));
    }
    counts.summaryLeftRows = rowsLeft.length;
    counts.summaryRightRows = rowsRight.length;
    counts.summarySheetName = summarySheet.getName();
    Logger.log('summarizeApprovedResultsByAgency: wrote ' + rowsLeft.length + ' row(s) (left) and ' + rowsRight.length + ' row(s) (right) to ' + summarySheet.getName());
    classifyResultsByClientSheet(summarySheet);
  }

  clearProgress_();

  var msg = '処理が完了しました。' +
            '\n確定成果 ' + counts.confirmed + ' 件' +
            '\n発生成果 ' + counts.generated + ' 件' +
            '\n【毎月更新】広告一覧 ' + counts.adListRows + ' 行';
  if (counts.summarySheetName) {
    msg += '\n' + counts.summarySheetName + ' 左 ' + counts.summaryLeftRows + ' 行 右 ' + counts.summaryRightRows + ' 行';
  }
  Logger.log('summarizeApprovedResultsByAgency: complete');
  Logger.log(msg);
  return msg;
  } catch (e) {
    clearProgress_();
    Logger.log('summarizeApprovedResultsByAgency: error ' + e + (e.stack ? '\n' + e.stack : ''));
    throw e;
  }
}

function classifyResultsByClientSheet_old(records, startDate, endDate) {
  var validRange =
    startDate instanceof Date && !isNaN(startDate) &&
    endDate instanceof Date && !isNaN(endDate);
  if (!validRange) {
    Logger.log('classifyResultsByClientSheet: invalid date range');
    return { generated: 0, confirmed: 0 };
  }
  startDate.setHours(0, 0, 0, 0);
  endDate.setHours(0, 0, 0, 0);

  var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  if (!Array.isArray(records)) records = [];

  var combinedSummary = {};
  var confirmedTotal = 0;

  records.forEach(function(rec) {
    if (!rec) return;
    var advName = rec.advertiser_name || rec.advertiserName || rec.advertiser || '';
    var ad = rec.ad || rec.ad_name || rec.adName || '';
    var unit = Number(rec.gross_action_cost || 0);
    var d = rec.apply_unix ? new Date(Number(rec.apply_unix) * 1000)
                           : (rec.apply ? new Date(String(rec.apply).replace(' ', 'T')) : null);
    if (!d || d < startDate || d > endDate) return;
    var key = advName + '\u0000' + ad + '\u0000' + unit;
    var entry = combinedSummary[key] || (combinedSummary[key] = {advertiser: advName, ad: ad, unit: unit, count: 0, amount: 0});
    entry.count++;
    entry.amount += unit;
    confirmedTotal++;
  });

  var headers = ['広告主名', '広告名', '単価', '件数', '金額'];
  var sheet = ss.getSheetByName('代理店集計') || ss.insertSheet('代理店集計');
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var rows = Object.keys(combinedSummary).map(function(k) {
    var s = combinedSummary[k];
    return [s.advertiser, s.ad, s.unit, s.count, s.amount];
  }).sort(function(a, b) {
    if (a[0] < b[0]) return -1;
    if (a[0] > b[0]) return 1;
    if (a[1] < b[1]) return -1;
    if (a[1] > b[1]) return 1;
    if (a[2] < b[2]) return -1;
    if (a[2] > b[2]) return 1;
    return 0;
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  return { generated: 0, confirmed: confirmedTotal };
}


function summarizeAgencyAds(targetSheetName) {
  Logger.log('処理を開始します');
  try {
    var msg = summarizeApprovedResultsByAgency(targetSheetName);
    alertUi_(msg);
  } catch (e) {
    alertUi_('エラーが発生しました: ' + e);
    throw e;
  }
}

function classifyResultsByClientSheet(summarySheet) {
  if (!summarySheet) return;
  var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);

  var clientSheet = ss.getSheetByName('クライアント情報');
  var typeMap = {};
  if (clientSheet) {
    var lastRow = clientSheet.getLastRow();
    if (lastRow >= 2) {
      var types = clientSheet.getRange(2, 14, lastRow - 1, 1).getValues();
      var advIds = clientSheet.getRange(2, 15, lastRow - 1, 1).getValues();
      advIds.forEach(function(row, idx) {
        var id = normalizeAdvId_(row[0] || '');
        if (id) {
          typeMap[id] = (types[idx][0] || '').toString().trim();
        }
      });
    }
  }

  var invoiceSheet = ss.getSheetByName('発行') || ss.insertSheet('発行');
  invoiceSheet.clearContents();
  invoiceSheet.getRange(1, 1, 1, 5).setValues([[
    '広告主ID', '広告主', '広告', '件数', '金額'
  ]]);

  var unmatchedSheet = ss.getSheetByName('該当なし') || ss.insertSheet('該当なし');
  unmatchedSheet.clearContents();
  unmatchedSheet.getRange(1, 1, 1, 7).setValues([[
    '広告主ID', '広告主', '広告', '発生成果数', '発生成果額', '確定成果数', '確定成果額'
  ]]);

  var last = summarySheet.getLastRow();
  if (last < 2) return;
  var data = summarySheet.getRange(2, 22, last - 1, 7).getValues();
  var invoiceRows = [];
  var unmatchedRows = [];

  data.forEach(function(row) {
    var advertiser = row[0];
    var ad = row[1];
    var genCount = Number(row[2] || 0);
    var genAmount = Number(row[3] || 0);
    var confCount = Number(row[4] || 0);
    var confAmount = Number(row[5] || 0);
    var advId = normalizeAdvId_(row[6] || '');
    if (!(genCount || genAmount || confCount || confAmount)) return;
    var t = typeMap[advId];
    if (t === '発生') {
      invoiceRows.push([advId, advertiser, ad, genCount, genAmount]);
    } else if (t === '確定') {
      invoiceRows.push([advId, advertiser, ad, confCount, confAmount]);
    } else {
      unmatchedRows.push([advId, advertiser, ad, genCount, genAmount, confCount, confAmount]);
    }
  });

  if (invoiceRows.length > 0) {
    // Merge rows with the same advertiser ID, ad name, and unit price into one line
    var merged = {};
    invoiceRows.forEach(function(r) {
      var advId = r[0];
      var advertiser = r[1];
      var ad = r[2];
      var count = Number(r[3] || 0);
      var amount = Number(r[4] || 0);
      var unit = count ? (amount / count).toString() : '0';
      var key = advId + '\u0000' + ad + '\u0000' + unit;
      if (!merged[key]) {
        merged[key] = [advId, advertiser, ad, 0, 0];
      }
      merged[key][3] += count;
      merged[key][4] += amount;
    });
    invoiceRows = Object.keys(merged).map(function(k) { return merged[k]; });
    invoiceSheet.getRange(2, 1, invoiceRows.length, 5).setValues(invoiceRows);
  }
  if (unmatchedRows.length > 0) {
    unmatchedSheet.getRange(2, 1, unmatchedRows.length, 7).setValues(unmatchedRows);
  }
}

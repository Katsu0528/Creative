'use strict';

function normalizeName_(str) {
  return typeof str === 'string' ? str.replace(/[\s\u3000]/g, '') : '';
}

function normalizeAdvId_(id) {
  if (id === 0 || id) {
    var str = String(id);
    str = str.replace(/[０-９]/g, function(c) {
      return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
    });
    str = str.replace(/\s+/g, '');
    if (/^\d+$/.test(str)) {
      str = str.replace(/^0+/, '');
      if (str === '') str = '0';
    }
    return str;
  }
  return '';
}

function createDateWithDayClamp_(year, monthIndex, day) {
  var lastDay = new Date(year, monthIndex + 1, 0).getDate();
  var safeDay = Math.min(Math.max(1, day), lastDay);
  return new Date(year, monthIndex, safeDay);
}

function resolveClientPeriod_(closingValue, today) {
  var now = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  var text = closingValue == null ? '' : String(closingValue).trim();
  if (!text) return null;

  if (/月末/.test(text)) {
    return {
      start: new Date(now.getFullYear(), now.getMonth() - 1, 1),
      end: new Date(now.getFullYear(), now.getMonth() + 1, 0)
    };
  }

  var day = Number(text);
  if (!isFinite(day) || day < 1 || day > 31) return null;

  var start;
  var endBase;
  if (now.getDate() >= 20) {
    start = createDateWithDayClamp_(now.getFullYear(), now.getMonth() - 1, day);
    endBase = createDateWithDayClamp_(now.getFullYear(), now.getMonth(), day);
  } else {
    start = createDateWithDayClamp_(now.getFullYear(), now.getMonth() - 2, day);
    endBase = createDateWithDayClamp_(now.getFullYear(), now.getMonth() - 1, day);
  }
  var end = new Date(endBase.getFullYear(), endBase.getMonth(), endBase.getDate() - 1);
  return { start: start, end: end };
}

function toDateAtMidnight_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  var str = String(value).trim();
  if (!str) return null;
  var normalized = str.replace(/\./g, '/').replace(/-/g, '/');
  var date = new Date(normalized);
  if (isNaN(date)) return null;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function getMonthRange_(baseDate) {
  return {
    start: new Date(baseDate.getFullYear(), baseDate.getMonth(), 1),
    end: new Date(baseDate.getFullYear(), baseDate.getMonth() + 1, 0)
  };
}

function getTargetPeriod_() {
  var props = PropertiesService.getScriptProperties();
  var start = toDateAtMidnight_(props.getProperty('RECEIPT_START_DATE'));
  var end = toDateAtMidnight_(props.getProperty('RECEIPT_END_DATE'));
  if (start && end && start.getTime() <= end.getTime()) {
    return { start: start, end: end };
  }
  return getMonthRange_(new Date());
}

function summarizeConfirmedResultsByAffiliate() {
  var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  function reportProgress_(message) {
    Logger.log(message);
    if (activeSpreadsheet) {
      activeSpreadsheet.toast(message, '受領処理の進行状況', 5);
    }
  }

  reportProgress_('処理を開始します');
  var clientSheet = ss.getSheetByName('クライアント情報');
  if (!clientSheet) {
    throw new Error('クライアント情報シートが見つかりません');
  }

  var baseUrl = 'https://otonari-asp.com/api/v1/m'.replace(/\/+$/, '');
  var accessKey = 'agqnoournapf';
  var secretKey = '5j39q2hzsmsccck0ccgo4w0o';
  var headers = { 'X-Auth-Token': accessKey + ':' + secretKey };

  function fetchRecords_(advertiserId, dateField, start, end, states) {
    var params = [
      'advertiser=' + encodeURIComponent(advertiserId),
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
        response = UrlFetchApp.fetch(url, {
          method: 'get',
          headers: headers,
          muteHttpExceptions: true
        });

        var statusCode = response.getResponseCode();
        if (statusCode < 200 || statusCode >= 300) {
          throw new Error(
            'status=' + statusCode +
            ' body=' + response.getContentText()
          );
        }
        break;
      } catch (e) {
        if (attempt === 2) {
          throw new Error('API取得に失敗しました advertiser=' + advertiserId + ': ' + e);
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
          if (j.records && j.records.length) result = result.concat(j.records);
        } catch (e) {}
      });
    }
    return result;
  }

  var lastRow = clientSheet.getLastRow();
  if (lastRow < 2) {
    reportProgress_('クライアント情報がないため処理を終了します');
    return;
  }
  var clientValues = clientSheet.getRange(2, 1, lastRow - 1, 16).getValues(); // A:P
  reportProgress_('クライアント情報を取得しました: ' + clientValues.length + '件');

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var basePeriod = getTargetPeriod_();
  var monthRange = getMonthRange_(today);

  var clients = [];
  var fetchPlans = {};

  clientValues.forEach(function(row, index) {
    var projectName = (row[0] || '').toString().trim(); // A
    var advertiserName = row[1]; // B
    var type = (row[13] || '').toString().trim(); // N
    var advertiserId = normalizeAdvId_(row[14]); // O
    var closing = row[15]; // P

    if (!advertiserId || (type !== '発生' && type !== '確定')) return;

    var dateField = type === '発生' ? 'regist_unix' : 'apply_auto_unix';
    var baseFetchKey = [advertiserId, dateField, basePeriod.start.getTime(), basePeriod.end.getTime()].join('|');
    if (!fetchPlans[baseFetchKey]) {
      fetchPlans[baseFetchKey] = {
        advertiserId: advertiserId,
        dateField: dateField,
        start: basePeriod.start,
        end: basePeriod.end,
        clients: []
      };
    }

    var closingPeriod = resolveClientPeriod_(closing, today);
    var closingFetchKey = '';
    if (closingPeriod && (closingPeriod.start.getTime() !== basePeriod.start.getTime() || closingPeriod.end.getTime() !== basePeriod.end.getTime())) {
      closingFetchKey = [advertiserId, dateField, closingPeriod.start.getTime(), closingPeriod.end.getTime()].join('|');
      if (!fetchPlans[closingFetchKey]) {
        fetchPlans[closingFetchKey] = {
          advertiserId: advertiserId,
          dateField: dateField,
          start: closingPeriod.start,
          end: closingPeriod.end,
          clients: []
        };
      }
    }

    var clientInfo = {
      rowNo: index + 2,
      projectName: projectName,
      projectNorm: normalizeName_(projectName),
      advertiserName: advertiserName || '',
      advertiserId: advertiserId,
      type: type,
      dateField: dateField,
      baseFetchKey: baseFetchKey,
      closingFetchKey: closingFetchKey
    };
    fetchPlans[baseFetchKey].clients.push(clientInfo);
    if (closingFetchKey) fetchPlans[closingFetchKey].clients.push(clientInfo);
    clients.push(clientInfo);
  });

  if (clients.length === 0) {
    reportProgress_('対象クライアントがないため処理を終了します');
    return;
  }

  var allRecords = [];
  var promotionSet = {};

  var fetchKeys = Object.keys(fetchPlans);
  Object.keys(fetchPlans).forEach(function(fetchKey, idx) {
    var plan = fetchPlans[fetchKey];
    var records = fetchRecords_(plan.advertiserId, plan.dateField, plan.start, plan.end, [1]);
    reportProgress_((idx + 1) + '/' + fetchKeys.length + '件目: advertiser=' + plan.advertiserId + ' 期間 ' +
      Utilities.formatDate(plan.start, 'Asia/Tokyo', 'yyyy/MM/dd') + ' - ' + Utilities.formatDate(plan.end, 'Asia/Tokyo', 'yyyy/MM/dd') +
      ' の成果件数 ' + records.length + '件を取得');

    records.forEach(function(rec) {
      rec._fetchKey = fetchKey;
      allRecords.push(rec);
      if (rec.promotion || rec.promotion === 0) promotionSet[rec.promotion] = true;
    });
  });

  var promotionMap = {};
  var promotionIds = Object.keys(promotionSet);
  reportProgress_('案件名の取得を開始します: ' + promotionIds.length + '件');
  for (var i = 0; i < promotionIds.length; i += 100) {
    var batch = promotionIds.slice(i, i + 100);
    var requests = batch.map(function(id) {
      return { url: baseUrl + '/promotion/search?id=' + encodeURIComponent(id), method: 'get', headers: headers };
    });
    var responses = UrlFetchApp.fetchAll(requests);
    responses.forEach(function(res, idx) {
      var id = batch[idx];
      try {
        var json = JSON.parse(res.getContentText());
        var rec = Array.isArray(json.records) ? json.records[0] : json.records;
        promotionMap[id] = rec && rec.name ? rec.name : String(id);
      } catch (e) {
        promotionMap[id] = String(id);
      }
    });
    reportProgress_('案件名の取得進捗: ' + Math.min(i + 100, promotionIds.length) + '/' + promotionIds.length + '件');
  }

  var recordsByFetchKey = {};
  allRecords.forEach(function(rec) {
    if (!recordsByFetchKey[rec._fetchKey]) recordsByFetchKey[rec._fetchKey] = [];
    recordsByFetchKey[rec._fetchKey].push(rec);
  });

  var summary = {};
  var matchedRecordKeySet = {};

  function addSummary_(resultType, adName, advertiserName, unit, count, amount) {
    var key = [resultType, adName, advertiserName, unit].join('\t');
    if (!summary[key]) {
      summary[key] = {
        type: resultType,
        ad: adName,
        advertiser: advertiserName,
        unit: unit,
        count: 0,
        amount: 0
      };
    }
    summary[key].count += count;
    summary[key].amount += amount;
  }

  function getRecordMeta_(rec) {
    var promotionId = rec.promotion || rec.promotion === 0 ? String(rec.promotion) : '';
    var adName = promotionMap[promotionId] || rec.promotion_name || promotionId || '';
    var unit = Number(rec.gross_action_cost || 0);
    var recKey = [String(rec.id || ''), String(rec.advertiser || ''), promotionId, unit, String(rec.regist_unix || ''), String(rec.apply_auto_unix || '')].join('|');
    return { adName: adName, adNorm: normalizeName_(adName), unit: unit, recKey: recKey };
  }

  clients.forEach(function(client) {
    var targetFetchKey = client.closingFetchKey || client.baseFetchKey;
    var records = recordsByFetchKey[targetFetchKey] || [];
    records.forEach(function(rec) {
      var meta = getRecordMeta_(rec);
      if (client.projectNorm && meta.adNorm !== client.projectNorm) return;
      addSummary_(client.type, meta.adName, client.advertiserName, meta.unit, 1, meta.unit);
      matchedRecordKeySet[meta.recKey] = true;
    });
  });

  clients.forEach(function(client) {
    if (client.type !== '発生') return;
    var monthFetchKey = [client.advertiserId, client.dateField, monthRange.start.getTime(), monthRange.end.getTime()].join('|');
    if (!fetchPlans[monthFetchKey]) {
      fetchPlans[monthFetchKey] = {
        advertiserId: client.advertiserId,
        dateField: client.dateField,
        start: monthRange.start,
        end: monthRange.end,
        clients: []
      };
      var monthRecords = fetchRecords_(client.advertiserId, client.dateField, monthRange.start, monthRange.end, [1]);
      recordsByFetchKey[monthFetchKey] = monthRecords;
    }
  });

  Object.keys(recordsByFetchKey).forEach(function(fetchKey) {
    var plan = fetchPlans[fetchKey];
    if (!plan || plan.dateField !== 'regist_unix') return;
    var isMonthPlan = plan.start.getTime() === monthRange.start.getTime() && plan.end.getTime() === monthRange.end.getTime();
    if (!isMonthPlan) return;
    var records = recordsByFetchKey[fetchKey] || [];
    records.forEach(function(rec) {
      var meta = getRecordMeta_(rec);
      if (matchedRecordKeySet[meta.recKey]) return;
      addSummary_('発生', meta.adName, rec.advertiser_name || String(rec.advertiser || ''), meta.unit, 1, meta.unit);
    });
  });

  var targetSheet = ss.getActiveSheet();
  targetSheet.clearContents();
  var outputHeaders = ['成果種別', '成果内容', '広告主', '単価', '成果数[件]', '成果額（グロス）[円]'];
  targetSheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);

  var rows = Object.keys(summary).map(function(key) {
    var item = summary[key];
    return [item.type, item.ad, item.advertiser, item.unit, item.count, item.amount];
  }).sort(function(a, b) {
    if (a[0] < b[0]) return -1;
    if (a[0] > b[0]) return 1;
    if (a[2] < b[2]) return -1;
    if (a[2] > b[2]) return 1;
    if (a[1] < b[1]) return -1;
    if (a[1] > b[1]) return 1;
    if (a[3] < b[3]) return -1;
    if (a[3] > b[3]) return 1;
    return 0;
  });

  if (rows.length > 0) {
    targetSheet.getRange(2, 1, rows.length, outputHeaders.length).setValues(rows);
  }
  reportProgress_('処理が完了しました: 出力 ' + rows.length + '行');
}

function 受領() {
  summarizeConfirmedResultsByAffiliate();
}

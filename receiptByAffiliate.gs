'use strict';

var RECEIPT_OUTPUT_SPREADSHEET_ID = '13zQMfgfYlec1BOo0LwWZUerQD9Fm0Fkzav8Z20d5eDE';
var RECEIPT_OUTPUT_SHEET_NAME = '結果';

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

function getDefaultTargetPeriod_(today) {
  var baseDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  if (baseDate.getDate() <= 15) {
    baseDate = new Date(baseDate.getFullYear(), baseDate.getMonth() - 1, 1);
  }
  return getMonthRange_(baseDate);
}

function getTargetPeriod_() {
  var props = PropertiesService.getScriptProperties();
  var start = toDateAtMidnight_(props.getProperty('RECEIPT_START_DATE'));
  var end = toDateAtMidnight_(props.getProperty('RECEIPT_END_DATE'));
  if (start && end && start.getTime() <= end.getTime()) {
    return { start: start, end: end };
  }
  return getDefaultTargetPeriod_(new Date());
}

function summarizeConfirmedResultsByAffiliate() {
  var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  var outputSs = SpreadsheetApp.openById(RECEIPT_OUTPUT_SPREADSHEET_ID);
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

  var props = PropertiesService.getScriptProperties();
  var baseUrl = (props.getProperty('OTONARI_BASE_URL') || 'https://otonari-asp.com/api/v1/m').trim().replace(/\/+$/, '');
  var accessKey = (props.getProperty('OTONARI_ACCESS_KEY') || 'agqnoournapf').trim();
  var secretKey = (props.getProperty('OTONARI_SECRET_KEY') || '5j39q2hzsmsccck0ccgo4w0o').trim();
  if (!accessKey || !secretKey) {
    throw new Error('APIのアクセスキーまたはシークレットキーが設定されていません');
  }
  var headers = { 'X-Auth-Token': accessKey + ':' + secretKey };

  function fetchAdvertiserNamesByIds_(ids) {
    var result = {};
    var uniqueIds = Array.from(new Set((ids || []).map(normalizeAdvId_).filter(function(id) {
      return !!id;
    })));
    uniqueIds.forEach(function(id) {
      try {
        var url = baseUrl + '/advertiser/search?id=' + encodeURIComponent(id) + '&limit=1';
        var response = UrlFetchApp.fetch(url, {
          method: 'get',
          headers: headers,
          muteHttpExceptions: true
        });
        var statusCode = response.getResponseCode();
        if (statusCode < 200 || statusCode >= 300) {
          Logger.log('広告主名取得に失敗 id=' + id + ' status=' + statusCode);
          return;
        }
        var json = JSON.parse(response.getContentText());
        var records = Array.isArray(json.records) ? json.records : [];
        if (records.length === 0) return;
        var advertiser = records[0] || {};
        var name = (advertiser.name || advertiser.company || advertiser.company_name || '').toString().trim();
        if (!name) {
          var company = (advertiser.company || advertiser.company_name || '').toString().trim();
          var person = (advertiser.person || advertiser.name || '').toString().trim();
          name = company && person ? (company + ' ' + person) : (company || person);
        }
        if (name) result[id] = name;
      } catch (e) {
        Logger.log('広告主名取得で例外 id=' + id + ' error=' + e);
      }
    });
    return result;
  }

  function fetchPromotionInfoByIds_(ids) {
    var result = {};
    var uniqueIds = Array.from(new Set((ids || []).map(function(id) {
      return (id || '').toString().trim();
    }).filter(function(id) {
      return !!id;
    })));

    uniqueIds.forEach(function(id) {
      try {
        var url = baseUrl + '/promotion/search?id=' + encodeURIComponent(id) + '&limit=1';
        var response = UrlFetchApp.fetch(url, {
          method: 'get',
          headers: headers,
          muteHttpExceptions: true
        });
        var statusCode = response.getResponseCode();
        if (statusCode < 200 || statusCode >= 300) {
          Logger.log('広告情報取得に失敗 id=' + id + ' status=' + statusCode);
          return;
        }

        var json = JSON.parse(response.getContentText());
        var records = Array.isArray(json.records) ? json.records : [];
        if (records.length === 0) return;

        var promotion = records[0] || {};
        var name = (promotion.name || promotion.promotion_name || promotion.ad_name || '').toString().trim();
        var company = (promotion.company || promotion.advertiser_company || '').toString().trim();
        var advertiserName = (promotion.advertiser_name || promotion.name_adv || '').toString().trim();
        var advertiserDisplay = (company + ' ' + advertiserName).trim();

        if (!advertiserDisplay) {
          var advertiserId = normalizeAdvId_(promotion.advertiser);
          advertiserDisplay = advertiserNameMap[advertiserId] || '';
        }

        result[id] = {
          name: name,
          advertiserDisplay: advertiserDisplay
        };
      } catch (e) {
        Logger.log('広告情報取得で例外 id=' + id + ' error=' + e);
      }
    });

    return result;
  }

  function fetchRecordsByDateField_(dateField, start, end, states) {
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
    var result = [];
    var offset = 0;
    var totalCount = null;

    while (totalCount === null || offset < totalCount) {
      var url = baseUrl + '/action_log_raw/search?' + baseParams + '&offset=' + offset;
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
            throw new Error('status=' + statusCode + ' body=' + response.getContentText());
          }
          break;
        } catch (e) {
          if (attempt === 2) {
            throw new Error('API取得に失敗しました dateField=' + dateField + ' offset=' + offset + ': ' + e);
          }
          Utilities.sleep(1000 * Math.pow(2, attempt));
        }
      }

      var json = JSON.parse(response.getContentText());
      var chunk = Array.isArray(json.records) ? json.records : [];
      if (totalCount === null) {
        totalCount = json.header && json.header.count ? Number(json.header.count) : chunk.length;
      }
      result = result.concat(chunk);
      offset += 500;
      if (chunk.length === 0) break;
    }

    return result;
  }

  var targetPeriod = getTargetPeriod_();
  var allStart = targetPeriod.start;
  var today = targetPeriod.end;

  reportProgress_('発生成果を取得します: ' + Utilities.formatDate(allStart, 'Asia/Tokyo', 'yyyy/MM/dd') + ' - ' + Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd'));
  var generatedRecords = fetchRecordsByDateField_('regist_unix', allStart, today, [1]);
  reportProgress_('発生成果を取得しました: ' + generatedRecords.length + '件');

  reportProgress_('確定成果を取得します: ' + Utilities.formatDate(allStart, 'Asia/Tokyo', 'yyyy/MM/dd') + ' - ' + Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd'));
  var confirmedRecords = fetchRecordsByDateField_('apply_auto_unix', allStart, today, [1]);
  reportProgress_('確定成果を取得しました: ' + confirmedRecords.length + '件');

  var lastRow = clientSheet.getLastRow();
  var clientValues = lastRow >= 2 ? clientSheet.getRange(2, 1, lastRow - 1, 16).getValues() : [];
  reportProgress_('クライアント情報を取得しました: ' + clientValues.length + '件');

  var clientByAdvertiserId = {};
  clientValues.forEach(function(row) {
    var advertiserId = normalizeAdvId_(row[14]); // O
    if (!advertiserId) return;
    if (!clientByAdvertiserId[advertiserId]) {
      clientByAdvertiserId[advertiserId] = {
        advertiserName: (row[1] || '').toString().trim(), // B
        resultType: (row[13] || '').toString().trim(), // N
        closing: (row[15] || '').toString().trim() // P
      };
    }
  });

  var byAdUnit = {};

  function getAdName_(rec) {
    return (rec.promotion_name || rec.ad_name || rec.name || '').toString().trim();
  }

  function getPromotionId_(rec) {
    if (rec.promotion === 0 || rec.promotion) {
      return String(rec.promotion).trim();
    }
    return '';
  }

  function ensureAdUnitSummary_(rec) {
    var advertiserId = normalizeAdvId_(rec.advertiser);
    if (!advertiserId) return null;
    var unitPrice = Number(rec.gross_action_cost || 0);
    var adName = getAdName_(rec);
    var promotionId = getPromotionId_(rec);
    var key = advertiserId + '\u0000' + promotionId + '\u0000' + adName + '\u0000' + unitPrice;
    if (!byAdUnit[key]) {
      byAdUnit[key] = {
        advertiserId: advertiserId,
        advertiserName: (rec.advertiser_name || '').toString().trim(),
        promotionId: promotionId,
        adName: adName,
        unitPrice: unitPrice,
        generatedCount: 0,
        generatedAmount: 0,
        confirmedCount: 0,
        confirmedAmount: 0
      };
    }
    return byAdUnit[key];
  }

  generatedRecords.forEach(function(rec) {
    var row = ensureAdUnitSummary_(rec);
    if (!row) return;
    row.generatedCount += 1;
    row.generatedAmount += Number(rec.gross_reward || rec.gross_action_cost || 0);
  });

  confirmedRecords.forEach(function(rec) {
    var row = ensureAdUnitSummary_(rec);
    if (!row) return;
    row.confirmedCount += 1;
    row.confirmedAmount += Number(rec.gross_reward || rec.gross_action_cost || 0);
  });

  var periodStartText = Utilities.formatDate(allStart, 'Asia/Tokyo', 'yyyy/MM/dd');
  var periodEndText = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy/MM/dd');

  var advertiserNameMap = fetchAdvertiserNamesByIds_(Object.keys(byAdUnit).map(function(key) {
    return byAdUnit[key].advertiserId;
  }));
  var promotionInfoMap = fetchPromotionInfoByIds_(Object.keys(byAdUnit).map(function(key) {
    return byAdUnit[key].promotionId;
  }));

  var outputRows = Object.keys(byAdUnit).reduce(function(rows, key) {
    var item = byAdUnit[key];
    var client = clientByAdvertiserId[item.advertiserId];
    var closing = client && client.closing ? client.closing : 'ヒットなし';
    var advertiserName = item.advertiserName || (client ? client.advertiserName : '') || advertiserNameMap[item.advertiserId] || '';
    var promotionInfo = promotionInfoMap[item.promotionId] || {};
    var adColumnValue = item.promotionId;
    var adNameValue = item.adName || promotionInfo.name || '';

    if (!client) {
      adColumnValue = promotionInfo.advertiserDisplay || advertiserName;
      adNameValue = promotionInfo.name || adNameValue;
    }

    var preferredType = client && client.resultType === '確定' ? '確定' : '発生';
    if (!client || preferredType === '発生') {
      rows.push([
        periodStartText,
        periodEndText,
        '発生',
        closing,
        advertiserName,
        adColumnValue,
        adNameValue,
        item.unitPrice,
        item.generatedCount,
        item.generatedAmount
      ]);
    }
    if (client && preferredType === '確定') {
      rows.push([
        periodStartText,
        periodEndText,
        '確定',
        closing,
        advertiserName,
        adColumnValue,
        adNameValue,
        item.unitPrice,
        item.confirmedCount,
        item.confirmedAmount
      ]);
    }
    return rows;
  }, []).sort(function(a, b) {
    if (a[0] < b[0]) return -1;
    if (a[0] > b[0]) return 1;
    if (a[1] < b[1]) return -1;
    if (a[1] > b[1]) return 1;
    if (a[3] < b[3]) return -1;
    if (a[3] > b[3]) return 1;
    if (a[4] < b[4]) return -1;
    if (a[4] > b[4]) return 1;
    if (a[5] < b[5]) return -1;
    if (a[5] > b[5]) return 1;
    if (a[6] < b[6]) return -1;
    if (a[6] > b[6]) return 1;
    if (a[7] < b[7]) return -1;
    if (a[7] > b[7]) return 1;
    if (a[2] < b[2]) return -1;
    if (a[2] > b[2]) return 1;
    return 0;
  });

  var targetSheet = outputSs.getSheetByName(RECEIPT_OUTPUT_SHEET_NAME);
  if (!targetSheet) {
    targetSheet = outputSs.insertSheet(RECEIPT_OUTPUT_SHEET_NAME);
  }
  targetSheet.clearContents();
  var outputHeaders = ['集計期間開始日', '集計期間終了日', '成果区分', '締め日', '広告主', '広告', '広告名', '単価[円]', '成果数[件]', '成果額（グロス）[円]'];
  targetSheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);

  if (outputRows.length > 0) {
    targetSheet.getRange(2, 1, outputRows.length, outputHeaders.length).setValues(outputRows);
  }

  reportProgress_('処理が完了しました: 出力 ' + outputRows.length + '行');
}

function 受領() {
  summarizeConfirmedResultsByAffiliate();
}

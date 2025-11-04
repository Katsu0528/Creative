/**
 * Automation utilities for working with affiliates and media IDs.
 *
 * Step 1: Populate affiliate IDs (column B) for each display name in column A.
 * Step 2: Create per-affiliate sheets that list all related media IDs.
 * Step 3: Submit partnership applications for every media ID to a promotion.
 */

var AFFILIATE_MEDIA_SHEET_PREFIX = 'Media-';
var AFFILIATE_MEDIA_HEADER = ['Media ID', 'Media Name', 'Media URL'];
const PROMOTION_ID = 'pi6s45pgyy50';

function populateAffiliateIdsAndMediaSheets(promotionId) {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (!sheet) {
    throw new Error('アクティブなシートを取得できませんでした。');
  }

  var affiliates = populateAffiliateIdsFromColumnA(sheet);
  createAffiliateMediaSheets(affiliates);

  if (promotionId) {
    applyAllAffiliateMediaToPromotion(promotionId, affiliates);
  }
}

function populateAffiliateIdsFromColumnA(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('処理対象の行が存在しません。');
    return [];
  }

  if (!sheet.getRange(1, 2).getValue()) {
    sheet.getRange(1, 2).setValue('アフィリエイターID');
  }

  var names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var results = [];
  var affiliates = [];
  var seenIds = {};
  var cache = {};

  for (var i = 0; i < names.length; i++) {
    var displayName = acsSanitizeString(names[i][0]);
    if (!displayName) {
      results.push(['']);
      continue;
    }

    if (!cache.hasOwnProperty(displayName)) {
      try {
        cache[displayName] = resolveAffiliateId(displayName);
      } catch (error) {
        Logger.log('Failed to resolve affiliate id for "%s": %s', displayName, error);
        cache[displayName] = '';
      }
    }

    var affiliateId = cache[displayName] || '';
    results.push([affiliateId]);

    if (affiliateId && !seenIds[affiliateId]) {
      affiliates.push({
        id: affiliateId,
        displayName: displayName
      });
      seenIds[affiliateId] = true;
    }
  }

  sheet.getRange(2, 2, results.length, 1).setValues(results);
  return affiliates;
}

function createAffiliateMediaSheets(affiliates) {
  if (!affiliates || !affiliates.length) {
    return;
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var existingSheets = spreadsheet.getSheets();
  var reusableSheets = {};
  for (var i = 0; i < existingSheets.length; i++) {
    var sheet = existingSheets[i];
    var name = sheet.getName();
    if (name.indexOf(AFFILIATE_MEDIA_SHEET_PREFIX) === 0) {
      reusableSheets[name] = sheet;
    }
  }

  for (var j = 0; j < affiliates.length; j++) {
    var affiliate = affiliates[j];
    var sheetName = AFFILIATE_MEDIA_SHEET_PREFIX + affiliate.id;
    var sheet = reusableSheets[sheetName];
    if (!sheet) {
      sheet = spreadsheet.getSheetByName(sheetName);
    }
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
    }

    sheet.clearContents();
    sheet.getRange(1, 1, 1, AFFILIATE_MEDIA_HEADER.length).setValues([AFFILIATE_MEDIA_HEADER]);
    sheet.getRange(1, AFFILIATE_MEDIA_HEADER.length + 1).setValue(affiliate.displayName);

    var mediaRecords = fetchMediaRecordsByAffiliate(affiliate.id);
    if (!mediaRecords.length) {
      continue;
    }

    var values = [];
    for (var k = 0; k < mediaRecords.length; k++) {
      var record = mediaRecords[k];
      values.push([
        acsSanitizeString(record.id),
        acsSanitizeString(record.name),
        acsSanitizeString(record.url)
      ]);
    }

    sheet.getRange(2, 1, values.length, AFFILIATE_MEDIA_HEADER.length).setValues(values);
  }
}

function applyAllAffiliateMediaToPromotion(promotionId, affiliates) {
  promotionId = promotionId || PROMOTION_ID;
  if (!promotionId) {
    throw new Error('提携申請を行う広告IDが指定されていません。');
  }

  var affiliateList = affiliates;
  if (!affiliateList || !affiliateList.length) {
    affiliateList = populateAffiliateIdsFromColumnA(SpreadsheetApp.getActiveSheet());
  }

  var totalMedia = 0;
  var appliedCount = 0;
  var duplicateCount = 0;
  var errorCount = 0;
  var errorMessages = [];

  for (var i = 0; i < affiliateList.length; i++) {
    var affiliate = affiliateList[i];
    var mediaRecords = fetchMediaRecordsByAffiliate(affiliate.id);
    for (var j = 0; j < mediaRecords.length; j++) {
      var media = mediaRecords[j];
      var mediaId = acsSanitizeString(media.id);
      if (!mediaId) {
        continue;
      }
      totalMedia++;
      var result = ensurePromotionApplication(mediaId, promotionId);
      if (result.status === 'success') {
        appliedCount++;
      } else if (result.status === 'duplicate') {
        duplicateCount++;
      } else if (result.status === 'skipped') {
        continue;
      } else {
        errorCount++;
        errorMessages.push(mediaId + ': ' + result.message);
      }
    }
  }

  var summary = [
    '提携申請処理が完了しました。',
    '対象メディア件数: ' + totalMedia,
    '新規申請件数: ' + appliedCount,
    '既存申請件数: ' + duplicateCount,
    'エラー件数: ' + errorCount
  ];

  if (errorMessages.length) {
    summary.push('エラー詳細:\n' + errorMessages.join('\n'));
  }

  SpreadsheetApp.getUi().alert(summary.join('\n'));
}

function resolveAffiliateId(displayName) {
  var parsed = acsParseAffiliateIdentifier(displayName);
  var searchPatterns = [];
  if (parsed.company && parsed.name) {
    searchPatterns.push({ company: parsed.company, name: parsed.name });
  }
  if (parsed.company) {
    searchPatterns.push({ company: parsed.company });
  }
  if (parsed.name) {
    searchPatterns.push({ name: parsed.name });
  }
  searchPatterns.push({ company: displayName });
  searchPatterns.push({ name: displayName });

  var seen = {};
  for (var i = 0; i < searchPatterns.length; i++) {
    var params = searchPatterns[i];
    var key = JSON.stringify(params);
    if (seen[key]) {
      continue;
    }
    seen[key] = true;

    var record = acsFindAffiliateRecord(params);
    if (record && record.id) {
      return record.id;
    }
  }
  return '';
}

function acsFindAffiliateRecord(params) {
  var records = fetchAllPages('/user/search', params);
  if (!records.length) {
    return null;
  }
  if (records.length === 1) {
    return records[0];
  }
  var exact = acsFindExactAffiliateRecord(records, params);
  return exact || records[0];
}

function acsFindExactAffiliateRecord(records, params) {
  var company = acsSanitizeString(params.company);
  var name = acsSanitizeString(params.name);

  if (company && name) {
    for (var i = 0; i < records.length; i++) {
      var record = records[i];
      if (acsSanitizeString(record.company) === company && acsSanitizeString(record.name) === name) {
        return record;
      }
    }
  }
  if (company) {
    for (var j = 0; j < records.length; j++) {
      var recordByCompany = records[j];
      if (acsSanitizeString(recordByCompany.company) === company) {
        return recordByCompany;
      }
    }
  }
  if (name) {
    for (var k = 0; k < records.length; k++) {
      var recordByName = records[k];
      if (acsSanitizeString(recordByName.name) === name) {
        return recordByName;
      }
    }
  }
  return null;
}

function acsParseAffiliateIdentifier(identifier) {
  var value = acsSanitizeString(identifier);
  if (!value) {
    return { company: '', name: '' };
  }

  var delimiters = ['＋', '+', '/', '／', '|', '｜', '>', '→'];
  for (var i = 0; i < delimiters.length; i++) {
    var delimiter = delimiters[i];
    if (value.indexOf(delimiter) !== -1) {
      var parts = value.split(delimiter).map(function(part) {
        return acsSanitizeString(part);
      }).filter(function(part) {
        return !!part;
      });
      if (parts.length >= 2) {
        return {
          company: parts[0],
          name: parts.slice(1).join(' ')
        };
      }
    }
  }

  var whitespaceParts = value.split(/[ \u3000]+/).filter(function(part) {
    return !!part;
  });
  if (whitespaceParts.length >= 2) {
    return {
      company: whitespaceParts.slice(0, whitespaceParts.length - 1).join(' '),
      name: whitespaceParts[whitespaceParts.length - 1]
    };
  }

  return {
    company: value,
    name: ''
  };
}

function fetchMediaRecordsByAffiliate(affiliateId) {
  if (!affiliateId) {
    return [];
  }
  return fetchAllPages('/media/search', { user: affiliateId });
}

function ensurePromotionApplication(mediaId, promotionId) {
  var existing = fetchPromotionApplication(mediaId, promotionId);
  if (existing) {
    return { status: 'duplicate', record: existing };
  }

  var payload = {
    media: mediaId,
    promotion: promotionId,
    state: 0
  };

  var response = callApi('/promotion_apply/regist', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  });

  if (response.status >= 200 && response.status < 300) {
    return { status: 'success', record: response.json && response.json.record };
  }

  return {
    status: 'error',
    message: acsExtractApiErrorMessage(response.json) || acsSanitizeString(response.text) || ('HTTPステータス: ' + response.status)
  };
}

function fetchPromotionApplication(mediaId, promotionId) {
  var records = fetchAllPages('/promotion_apply/search', {
    media: mediaId,
    promotion: promotionId,
    limit: 1
  });
  return records.length ? records[0] : null;
}

function fetchAllPages(path, params) {
  var config = getAcsApiConfig();
  var limit = params && params.limit ? params.limit : 200;
  var offset = 0;
  var results = [];
  var baseParams = params || {};

  while (true) {
    var queryParams = {};
    for (var key in baseParams) {
      if (!baseParams.hasOwnProperty(key)) {
        continue;
      }
      if (key === 'limit' || key === 'offset') {
        continue;
      }
      queryParams[key] = baseParams[key];
    }
    queryParams.limit = limit;
    queryParams.offset = offset;

    var response = callApi(path + '?' + acsBuildQueryString(queryParams));
    if (response.status !== 200) {
      throw new Error('APIリクエストに失敗しました。status=' + response.status + ' body=' + response.text);
    }

    var records = acsNormalizeRecords(response.json && response.json.records);
    if (!records.length) {
      break;
    }

    results = results.concat(records);
    if (records.length < limit) {
      break;
    }
    offset += records.length;
  }

  return results;
}

function callApi(path, options) {
  var config = getAcsApiConfig();
  var url = config.baseUrl + path;
  var requestOptions = {
    method: (options && options.method) || 'get',
    muteHttpExceptions: true,
    headers: {
      'X-Auth-Token': config.token
    }
  };

  if (options && options.contentType) {
    requestOptions.contentType = options.contentType;
  }
  if (options && options.payload !== undefined) {
    requestOptions.payload = options.payload;
  }
  if (options && options.headers) {
    for (var key in options.headers) {
      if (options.headers.hasOwnProperty(key)) {
        requestOptions.headers[key] = options.headers[key];
      }
    }
  }

  var response = UrlFetchApp.fetch(url, requestOptions);
  return {
    status: response.getResponseCode(),
    text: response.getContentText(),
    json: acsParseJsonSafe(response.getContentText())
  };
}

function getAcsApiConfig() {
  var props = PropertiesService.getScriptProperties();
  var baseUrl = acsSanitizeString(props.getProperty('OTONARI_BASE_URL')) || 'https://otonari-asp.com/api/v1/m';
  baseUrl = baseUrl.replace(/\/+$/, '');
  var accessKey = acsSanitizeString(props.getProperty('OTONARI_ACCESS_KEY')) || 'agqnoournapf';
  var secretKey = acsSanitizeString(props.getProperty('OTONARI_SECRET_KEY')) || '5j39q2hzsmsccck0ccgo4w0o';
  if (!accessKey || !secretKey) {
    throw new Error('APIキーが設定されていません。');
  }
  return {
    baseUrl: baseUrl,
    token: accessKey + ':' + secretKey
  };
}

function acsBuildQueryString(params) {
  var parts = [];
  for (var key in params) {
    if (!params.hasOwnProperty(key)) {
      continue;
    }
    var value = params[key];
    if (value === null || value === undefined || value === '') {
      continue;
    }
    if (Array.isArray(value)) {
      for (var i = 0; i < value.length; i++) {
        parts.push(encodeURIComponent(key) + '=' + encodeURIComponent(value[i]));
      }
    } else {
      parts.push(encodeURIComponent(key) + '=' + encodeURIComponent(value));
    }
  }
  return parts.join('&');
}

function acsNormalizeRecords(records) {
  if (!records) {
    return [];
  }
  if (Array.isArray(records)) {
    return records;
  }
  if (records.id) {
    return [records];
  }
  var normalized = [];
  for (var key in records) {
    if (records.hasOwnProperty(key) && records[key]) {
      normalized.push(records[key]);
    }
  }
  return normalized;
}

function acsExtractApiErrorMessage(responseJson) {
  if (!responseJson) {
    return '';
  }
  if (typeof responseJson.message === 'string') {
    return acsSanitizeString(responseJson.message);
  }
  var error = responseJson.error;
  if (!error) {
    return '';
  }
  if (typeof error === 'string') {
    return acsSanitizeString(error);
  }
  var messages = [];
  if (error.message) {
    messages.push(acsSanitizeString(error.message));
  }
  if (error.field_error) {
    var fieldErrors = Array.isArray(error.field_error) ? error.field_error : [error.field_error];
    for (var i = 0; i < fieldErrors.length; i++) {
      messages.push(acsSanitizeString(fieldErrors[i]));
    }
  }
  if (error.global_error) {
    var globalErrors = Array.isArray(error.global_error) ? error.global_error : [error.global_error];
    for (var j = 0; j < globalErrors.length; j++) {
      messages.push(acsSanitizeString(globalErrors[j]));
    }
  }
  return messages.filter(function(message) {
    return !!message;
  }).join(' / ');
}

function acsParseJsonSafe(text) {
  if (!text) {
    return null;
  }
  try {
    return JSON.parse(text);
  } catch (error) {
    Logger.log('JSON parse error: ' + error + ' body=' + text);
    return null;
  }
}

function acsSanitizeString(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

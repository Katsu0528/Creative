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

    sheet.getRange(1, 1, values.length, AFFILIATE_MEDIA_HEADER.length).setValues(values);
  }
}

function applyAllAffiliateMediaToPromotion(promotionId, affiliates) {
  promotionId = promotionId || PROMOTION_ID;
  if (!promotionId) {
    throw new Error('提携申請を行う広告IDが指定されていません。');
  }

  var mediaEntries = collectMediaEntriesFromSheets(affiliates);
  if (!mediaEntries.length) {
    SpreadsheetApp.getUi().alert('提携申請の対象となるメディアIDが見つかりませんでした。');
    return;
  }

  var totalMedia = mediaEntries.length;
  var appliedCount = 0;
  var duplicateCount = 0;
  var errorCount = 0;
  var errorMessages = [];

  Logger.log('提携申請処理を開始します。広告ID: %s, 対象メディア数: %s', promotionId, totalMedia);

  for (var i = 0; i < mediaEntries.length; i++) {
    var entry = mediaEntries[i];
    var result = ensurePromotionApplication(entry.id, promotionId);
    if (result.status === 'success') {
      appliedCount++;
      Logger.log('[%s/%s] %s の提携申請を新規で登録しました。', i + 1, totalMedia, entry.id);
    } else if (result.status === 'duplicate') {
      duplicateCount++;
      Logger.log('[%s/%s] %s は既に申請済みでした。', i + 1, totalMedia, entry.id);
    } else if (result.status === 'skipped') {
      Logger.log('[%s/%s] %s の提携申請はスキップされました。', i + 1, totalMedia, entry.id);
      continue;
    } else {
      errorCount++;
      var message = entry.id + ' (' + entry.sheetName + '!' + entry.rowNumber + ')';
      if (result.message) {
        message += ': ' + result.message;
      }
      errorMessages.push(message);
      Logger.log('[%s/%s] %s の提携申請でエラーが発生しました: %s', i + 1, totalMedia, entry.id, result.message || '理由不明');
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

  var summaryText = summary.join('\n');
  Logger.log('提携申請処理のサマリー:\n%s', summaryText);
  SpreadsheetApp.getUi().alert(summaryText);
}

function collectMediaEntriesFromSheets(affiliates) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var entries = [];
  var targetSheetNames = null;

  if (affiliates && affiliates.length) {
    targetSheetNames = {};
    for (var i = 0; i < affiliates.length; i++) {
      var affiliate = affiliates[i];
      var sheetName = AFFILIATE_MEDIA_SHEET_PREFIX + acsSanitizeString(affiliate.id);
      targetSheetNames[sheetName] = true;
    }
  }

  for (var j = 0; j < sheets.length; j++) {
    var sheet = sheets[j];
    var sheetName = sheet.getName();
    if (sheetName.indexOf(AFFILIATE_MEDIA_SHEET_PREFIX) !== 0) {
      continue;
    }
    if (targetSheetNames && !targetSheetNames[sheetName]) {
      continue;
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 1) {
      continue;
    }

    var values = sheet.getRange(1, 1, lastRow, 1).getValues();
    for (var rowIndex = 0; rowIndex < values.length; rowIndex++) {
      var mediaId = acsSanitizeString(values[rowIndex][0]);
      if (!mediaId) {
        continue;
      }
      entries.push({
        id: mediaId,
        sheetName: sheetName,
        rowNumber: rowIndex + 1
      });
    }
  }

  return entries;
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
  var existingApplication = findExistingPromotionApplication(mediaId, promotionId);
  if (existingApplication) {
    return { status: 'duplicate', record: existingApplication };
  }

  var payload = {
    media: mediaId,
    promotion: promotionId,
    state: 1
  };

  var response = callApi('/promotion_apply/regist', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  });

  if (response.status >= 200 && response.status < 300) {
    return { status: 'success', record: response.json && response.json.record };
  }

  if (acsIsDuplicateApplicationResponse(response)) {
    return {
      status: 'duplicate',
      message: acsExtractApiErrorMessage(response.json) || acsSanitizeString(response.text) || ''
    };
  }

  return {
    status: 'error',
    message: acsExtractApiErrorMessage(response.json) || acsSanitizeString(response.text) || ('HTTPステータス: ' + response.status)
  };
}

function findExistingPromotionApplication(mediaId, promotionId) {
  try {
    var query = acsBuildQueryString({
      promotion: promotionId,
      media: mediaId
    });
    var response = callApi('/promotion_apply/search' + (query ? '?' + query : ''));
    if (response.status !== 200) {
      Logger.log('提携申請検索に失敗しました (media=%s, promotion=%s, status=%s, body=%s)', mediaId, promotionId, response.status, response.text);
      return null;
    }

    var records = acsNormalizeRecords(response.json && response.json.records);
    if (records.length > 0) {
      return records[0];
    }
  } catch (error) {
    Logger.log('提携申請検索エラー: media=%s, promotion=%s, error=%s', mediaId, promotionId, error);
  }
  return null;
}

function acsIsDuplicateApplicationResponse(response) {
  if (!response) {
    return false;
  }

  if (response.status === 409) {
    return true;
  }

  var message = acsExtractApiErrorMessage(response.json);
  if (acsContainsDuplicatePhrase(message)) {
    return true;
  }

  if (acsContainsDuplicatePhrase(acsSanitizeString(response.text))) {
    return true;
  }

  var error = response.json && response.json.error;
  if (!error) {
    return false;
  }

  if (acsContainsDuplicatePhrase(acsSanitizeString(error.message))) {
    return true;
  }

  var fieldErrors = error.field_error;
  if (fieldErrors) {
    if (!Array.isArray(fieldErrors)) {
      fieldErrors = [fieldErrors];
    }
    for (var i = 0; i < fieldErrors.length; i++) {
      if (acsContainsDuplicatePhrase(acsSanitizeString(fieldErrors[i]))) {
        return true;
      }
    }
  }

  var globalErrors = error.global_error;
  if (globalErrors) {
    if (!Array.isArray(globalErrors)) {
      globalErrors = [globalErrors];
    }
    for (var j = 0; j < globalErrors.length; j++) {
      if (acsContainsDuplicatePhrase(acsSanitizeString(globalErrors[j]))) {
        return true;
      }
    }
  }

  return false;
}

function acsContainsDuplicatePhrase(text) {
  if (!text) {
    return false;
  }

  var lower = text.toLowerCase();
  return text.indexOf('既に') !== -1 || text.indexOf('すでに') !== -1 || lower.indexOf('already') !== -1;
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

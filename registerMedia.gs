var MEDIA_BASE_API_URL = 'https://otonari-asp.com/api/v1/m'.replace(/\/+$/, '');
var MEDIA_ACCESS_KEY = 'agqnoournapf';
var MEDIA_SECRET_KEY = '1kvu9dyv1alckgocc848socw';
var MEDIA_AUTH_TOKEN = MEDIA_ACCESS_KEY + ':' + MEDIA_SECRET_KEY;
var MEDIA_SHEET_NAME = 'メディア登録';
var MEDIA_DATA_START_ROW = 2;
var DEFAULT_MEDIA_CATEGORY_ID = '';
var DEFAULT_MEDIA_TYPE_ID = '';
var DEFAULT_MEDIA_PV_ROW = 0;
var DEFAULT_MEDIA_UU_ROW = 0;
var MEDIA_CATEGORY_LOOKUP_CACHE = null;
var MEDIA_TYPE_LOOKUP_CACHE = null;

var MEDIA_COLUMNS = {
  AFFILIATE_IDENTIFIER: 0,
  MEDIA_NAME: 1,
  MEDIA_URL: 2,
  MEDIA_CATEGORY_ID: 3,
  MEDIA_TYPE_ID: 4,
  MEDIA_COMMENT: 5,
  RESULT_MEDIA_ID: 6,
  RESULT_MESSAGE: 7
};

var MEDIA_COLUMN_COUNT = MEDIA_COLUMNS.RESULT_MESSAGE + 1;


function registerMediaFromSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(MEDIA_SHEET_NAME) || spreadsheet.getActiveSheet();
  if (!sheet) {
    throw new Error('処理対象のシートを取得できませんでした。');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < MEDIA_DATA_START_ROW) {
    SpreadsheetApp.getUi().alert('データ行が存在しません。');
    return;
  }

  var rowCount = lastRow - MEDIA_DATA_START_ROW + 1;
  var values = sheet.getRange(MEDIA_DATA_START_ROW, 1, rowCount, MEDIA_COLUMN_COUNT).getValues();
  var processingResult = processMediaRegistrationRows(values, {
    rowNumberOffset: MEDIA_DATA_START_ROW - 1,
    respectExistingResults: true,
  });

  if (processingResult.sheetValues.length) {
    sheet
      .getRange(
        MEDIA_DATA_START_ROW,
        MEDIA_COLUMNS.RESULT_MEDIA_ID + 1,
        processingResult.sheetValues.length,
        2
      )
      .setValues(processingResult.sheetValues);
  }
}

function registerMediaFromWeb(rowEntries) {
  if (!Array.isArray(rowEntries)) {
    throw new Error('送信されたデータの形式が正しくありません。');
  }

  var normalizedRows = [];
  var customRowNumbers = [];

  rowEntries.forEach(function(entry, index) {
    var data = entry || {};
    var rowNumber = Number(data.rowNumber);
    if (!rowNumber || rowNumber < 1) {
      rowNumber = index + 1;
    }

    var rowValues = [
      sanitizeString(data.affiliateIdentifier || data.affiliate || ''),
      sanitizeString(data.mediaName),
      sanitizeString(data.mediaUrl),
      sanitizeString(data.mediaCategory),
      sanitizeString(data.mediaType),
      sanitizeString(data.mediaComment),
      '',
      '',
    ];

    var hasValue = false;
    for (var i = 0; i <= MEDIA_COLUMNS.MEDIA_COMMENT; i++) {
      if (rowValues[i]) {
        hasValue = true;
        break;
      }
    }

    if (!hasValue) {
      return;
    }

    normalizedRows.push(rowValues);
    customRowNumbers.push(rowNumber);
  });

  var processingResult = processMediaRegistrationRows(normalizedRows, {
    rowNumberOffset: 0,
    respectExistingResults: false,
    rowNumbers: customRowNumbers,
  });

  return {
    summary: processingResult.summary,
    results: processingResult.details,
  };
}

function processMediaRegistrationRows(rows, options) {
  if (!Array.isArray(rows) || !rows.length) {
    return {
      sheetValues: [],
      details: [],
      summary: { total: 0, success: 0, skipped: 0, errors: 0 },
    };
  }

  var opts = options || {};
  var offset = Number(opts.rowNumberOffset || 0);
  if (isNaN(offset)) {
    offset = 0;
  }
  var respectExistingResults = opts.respectExistingResults !== false;
  var customRowNumbers = Array.isArray(opts.rowNumbers) ? opts.rowNumbers : null;

  var sheetValues = [];
  var detailedResults = [];
  var summary = {
    total: rows.length,
    success: 0,
    skipped: 0,
    errors: 0,
  };

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i] || [];
    var rowNumber = customRowNumbers && customRowNumbers[i]
      ? customRowNumbers[i]
      : offset + i + 1;
    var affiliateIdentifier = sanitizeString(row[MEDIA_COLUMNS.AFFILIATE_IDENTIFIER]);
    var mediaName = sanitizeString(row[MEDIA_COLUMNS.MEDIA_NAME]);
    var mediaUrl = sanitizeString(row[MEDIA_COLUMNS.MEDIA_URL]);
    var mediaCategoryInput = sanitizeString(row[MEDIA_COLUMNS.MEDIA_CATEGORY_ID]);
    var mediaTypeInput = sanitizeString(row[MEDIA_COLUMNS.MEDIA_TYPE_ID]);
    var mediaCategoryId = resolveMediaCategoryId(mediaCategoryInput || DEFAULT_MEDIA_CATEGORY_ID);
    var mediaTypeId = resolveMediaTypeId(mediaTypeInput || DEFAULT_MEDIA_TYPE_ID);
    var mediaComment = sanitizeString(row[MEDIA_COLUMNS.MEDIA_COMMENT]);
    var existingMediaId = respectExistingResults
      ? sanitizeString(row[MEDIA_COLUMNS.RESULT_MEDIA_ID])
      : '';
    var existingMessage = respectExistingResults
      ? sanitizeString(row[MEDIA_COLUMNS.RESULT_MESSAGE])
      : '';

    Logger.log(
      'Processing media registration row ' +
        rowNumber +
        ' affiliateIdentifier=' +
        affiliateIdentifier +
        ' mediaName=' +
        mediaName +
        ' mediaUrl=' +
        mediaUrl
    );

    var resultStatus = 'skipped';
    var resultMessage = '';
    var resultMediaId = existingMediaId;

    if (!affiliateIdentifier || !mediaName || !mediaUrl) {
      resultMessage = existingMessage || '必須項目が空欄のため処理をスキップしました。';
      Logger.log('Row ' + rowNumber + ' skipped: ' + resultMessage);
      appendResult();
      continue;
    }

    if (!mediaCategoryId) {
      resultMessage = mediaCategoryInput
        ? 'メディアカテゴリー名からIDを取得できませんでした。'
        : 'メディアカテゴリーIDが空欄のため処理をスキップしました。';
      Logger.log('Row ' + rowNumber + ' skipped: ' + resultMessage);
      appendResult();
      continue;
    }

    if (!mediaTypeId) {
      resultMessage = mediaTypeInput
        ? 'メディアタイプ名からIDを取得できませんでした。'
        : 'メディアタイプIDが空欄のため処理をスキップしました。';
      Logger.log('Row ' + rowNumber + ' skipped: ' + resultMessage);
      appendResult();
      continue;
    }

    if (!mediaComment) {
      resultMessage = 'メディア説明が空欄のため処理をスキップしました。';
      Logger.log('Row ' + rowNumber + ' skipped: ' + resultMessage);
      appendResult();
      continue;
    }

    if (respectExistingResults && existingMediaId) {
      resultMessage =
        existingMessage || '登録済みです。再登録する場合G列とH列を空にしてください。';
      Logger.log('Row ' + rowNumber + ' skipped: ' + resultMessage);
      appendResult();
      continue;
    }

    var userId = resolveAffiliateUserId(affiliateIdentifier);
    if (!userId) {
      resultStatus = 'error';
      resultMediaId = '';
      resultMessage = 'アフィリエイターIDを取得できませんでした。';
      Logger.log('Row ' + rowNumber + ' skipped: ' + resultMessage);
      appendResult();
      continue;
    }

    Logger.log('Row ' + rowNumber + ' resolved userId=' + userId);

    var registrationResult = registerMedia(
      userId,
      mediaName,
      mediaUrl,
      mediaCategoryId,
      mediaTypeId,
      mediaComment
    );
    if (registrationResult.success) {
      resultStatus = 'success';
      resultMediaId = registrationResult.id;
      resultMessage = '登録に成功しました。';
      Logger.log('Row ' + rowNumber + ' media registration succeeded. mediaId=' + resultMediaId);
    } else {
      resultStatus = 'error';
      resultMediaId = '';
      resultMessage = registrationResult.message;
      Logger.log('Row ' + rowNumber + ' media registration failed: ' + resultMessage);
    }

    appendResult();

    function appendResult() {
      sheetValues.push([resultMediaId || '', resultMessage || '']);
      detailedResults.push({
        rowIndex: i,
        rowNumber: rowNumber,
        mediaId: resultMediaId || '',
        message: resultMessage || '',
        status: resultStatus,
      });

      if (resultStatus === 'success') {
        summary.success += 1;
      } else if (resultStatus === 'error') {
        summary.errors += 1;
      } else {
        summary.skipped += 1;
      }
    }
  }

  return {
    sheetValues: sheetValues,
    details: detailedResults,
    summary: summary,
  };
}

function registerPromotionApply(promotionId) {
  var targetPromotionId = sanitizeString(promotionId);
  if (!targetPromotionId) {
    throw new Error('提携申請する広告IDが指定されていません。');
  }

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(MEDIA_SHEET_NAME) || spreadsheet.getActiveSheet();
  if (!sheet) {
    throw new Error('処理対象のシートを取得できませんでした。');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < MEDIA_DATA_START_ROW) {
    SpreadsheetApp.getUi().alert('データ行が存在しません。');
    return;
  }

  var rowCount = lastRow - MEDIA_DATA_START_ROW + 1;
  var values = sheet.getRange(MEDIA_DATA_START_ROW, 1, rowCount, MEDIA_COLUMN_COUNT).getValues();
  var results = [];

  var summary = {
    total: values.length,
    successApplications: 0,
    duplicateApplications: 0,
    skippedRows: 0,
    errorRows: 0
  };

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowNumber = MEDIA_DATA_START_ROW + i;
    var affiliateIdentifier = sanitizeString(row[MEDIA_COLUMNS.AFFILIATE_IDENTIFIER]);
    var existingMediaId = sanitizeString(row[MEDIA_COLUMNS.RESULT_MEDIA_ID]);

    Logger.log('Processing promotion application for row ' + rowNumber + ' affiliateIdentifier=' + affiliateIdentifier);

    if (!affiliateIdentifier) {
      var missingAffiliateMessage = 'アフィリエイター識別子が空欄のため提携申請をスキップしました。';
      Logger.log('Row ' + rowNumber + ' skipped: ' + missingAffiliateMessage);
      results.push([existingMediaId, missingAffiliateMessage]);
      summary.skippedRows++;
      continue;
    }

    var userId = resolveAffiliateUserId(affiliateIdentifier);
    if (!userId) {
      var userMissingMessage = 'アフィリエイターIDを取得できなかったため提携申請をスキップしました。';
      Logger.log('Row ' + rowNumber + ' skipped: ' + userMissingMessage);
      results.push([existingMediaId, userMissingMessage]);
      summary.errorRows++;
      continue;
    }

    var mediaList = listActiveMediaByAffiliate(userId);
    if (!mediaList.length) {
      var noMediaMessage = '提携申請対象のメディアが見つかりませんでした。';
      Logger.log('Row ' + rowNumber + ' skipped: ' + noMediaMessage);
      results.push([existingMediaId, noMediaMessage]);
      summary.skippedRows++;
      continue;
    }

    var appliedCount = 0;
    var duplicateCount = 0;
    var errorMessages = [];

    for (var j = 0; j < mediaList.length; j++) {
      var mediaRecord = mediaList[j];
      var mediaId = sanitizeString(mediaRecord && mediaRecord.id);
      if (!mediaId) {
        continue;
      }

      var applicationResult = submitPromotionApplication(mediaId, targetPromotionId);
      if (applicationResult.success) {
        if (applicationResult.duplicate) {
          duplicateCount++;
        } else {
          appliedCount++;
        }
      } else {
        errorMessages.push(applicationResult.message || '提携申請に失敗しました。');
      }
    }

    summary.successApplications += appliedCount;
    summary.duplicateApplications += duplicateCount;
    if (errorMessages.length) {
      summary.errorRows++;
    }

    var messageParts = [];
    if (appliedCount > 0) {
      messageParts.push('提携申請を送信しました（' + appliedCount + '件）。');
    }
    if (duplicateCount > 0) {
      messageParts.push('既に提携済みのメディアが' + duplicateCount + '件ありました。');
    }
    if (errorMessages.length) {
      messageParts.push('エラー: ' + uniqueMessages(errorMessages).join(' / '));
    }
    if (!messageParts.length) {
      messageParts.push('提携申請の送信対象がありませんでした。');
    }

    results.push([existingMediaId, messageParts.join(' ')]);
  }

  sheet.getRange(MEDIA_DATA_START_ROW, MEDIA_COLUMNS.RESULT_MEDIA_ID + 1, results.length, 2).setValues(results);

  try {
    SpreadsheetApp.getUi().alert([
      '提携申請処理が完了しました。',
      '申請成功件数: ' + summary.successApplications,
      '既存提携件数: ' + summary.duplicateApplications,
      'スキップ件数: ' + summary.skippedRows,
      'エラー件数: ' + summary.errorRows
    ].join('\n'));
  } catch (uiError) {
    Logger.log('Unable to display summary alert: ' + uiError);
  }
}

function registerMedia(userId, mediaName, mediaUrl, mediaCategoryId, mediaTypeId, mediaComment) {
  var payload = {
    user: userId,
    name: mediaName,
    url: mediaUrl,
    media_category: mediaCategoryId,
    media_type: mediaTypeId,
    pv_row: DEFAULT_MEDIA_PV_ROW,
    uu_row: DEFAULT_MEDIA_UU_ROW,
    comment: mediaComment
  };

  Logger.log('Submitting media registration request payload: ' + JSON.stringify(payload));

  var options = {
    method: 'post',
    headers: createAuthHeaders({ 'Content-Type': 'application/json' }),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(MEDIA_BASE_API_URL + '/media/regist', options);
    var statusCode = response.getResponseCode();
    var contentText = response.getContentText();

    if (statusCode >= 200 && statusCode < 300) {
      var json = parseJson(contentText);
      var record = json && json.record;
      if (record && record.id) {
        Logger.log('Media registration API success. userId=' + userId + ' mediaId=' + record.id);
        return { success: true, id: record.id, message: '' };
      }
      Logger.log('Media registration API success response missing id. userId=' + userId + ' response=' + contentText);
      return { success: false, message: 'API応答からメディアIDを取得できませんでした。' };
    }

    Logger.log('Media registration failed. Status: ' + statusCode + ' Body: ' + contentText);
    return { success: false, message: 'メディア登録に失敗しました。レスポンスコード: ' + statusCode };
  } catch (error) {
    Logger.log('Media registration error: ' + error);
    return { success: false, message: 'メディア登録中にエラーが発生しました。' };
  }
}

function listActiveMediaByAffiliate(userId) {
  if (!userId) {
    return [];
  }

  var records = callAllPagesAPI(MEDIA_BASE_API_URL + '/media/search', MEDIA_AUTH_TOKEN, { user: userId });
  var active = [];

  for (var i = 0; i < records.length; i++) {
    var record = records[i];
    if (!record) {
      continue;
    }

    var state = record.state;
    var isActive = state === undefined || state === null || state === '' || state === 1 || state === '1';
    if (isActive) {
      active.push(record);
    }
  }

  Logger.log('listActiveMediaByAffiliate: userId=' + userId + ' activeCount=' + active.length + ' totalRecords=' + records.length);
  return active;
}

function submitPromotionApplication(mediaId, promotionId, options) {
  if (!mediaId || !promotionId) {
    return { success: false, message: '提携申請に必要な情報が不足しています。' };
  }

  var existingRecord = findExistingPromotionApplication(mediaId, promotionId);
  if (existingRecord) {
    Logger.log('Promotion application already exists: mediaId=' + mediaId + ' promotionId=' + promotionId);
    return {
      success: true,
      duplicate: true,
      message: '既に提携申請済みです。',
      record: existingRecord
    };
  }

  var desiredState = 1;
  if (options && options.state !== undefined) {
    desiredState = Number(options.state);
    if (isNaN(desiredState)) {
      desiredState = 1;
    }
  }

  var payload = {
    media: mediaId,
    promotion: promotionId,
    state: desiredState
  };

  var requestOptions = {
    method: 'post',
    headers: createAuthHeaders({ 'Content-Type': 'application/json' }),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(MEDIA_BASE_API_URL + '/promotion_apply/regist', requestOptions);
    var statusCode = response.getResponseCode();
    var contentText = response.getContentText();

    if (statusCode >= 200 && statusCode < 300) {
      Logger.log('Promotion application succeeded: mediaId=' + mediaId + ' promotionId=' + promotionId);
      var json = parseJson(contentText) || {};
      return { success: true, duplicate: false, message: '', record: json.record || null };
    }

    var jsonBody = parseJson(contentText);
    var errorMessage = extractApiErrorMessage(jsonBody) || sanitizeString(contentText) || '提携申請に失敗しました。レスポンスコード: ' + statusCode;

    if (isDuplicateApplicationMessage(errorMessage) || isDuplicateApplicationFromResponse(jsonBody)) {
      Logger.log('Promotion application duplicate detected by API response: mediaId=' + mediaId + ' promotionId=' + promotionId + ' message=' + errorMessage);
      return {
        success: true,
        duplicate: true,
        message: errorMessage || '既に提携済みです。',
        record: jsonBody && jsonBody.record
      };
    }

    Logger.log('Promotion application failed: mediaId=' + mediaId + ' promotionId=' + promotionId + ' status=' + statusCode + ' body=' + contentText);
    return {
      success: false,
      message: errorMessage || '提携申請に失敗しました。レスポンスコード: ' + statusCode
    };
  } catch (error) {
    Logger.log('Promotion application error: mediaId=' + mediaId + ' promotionId=' + promotionId + ' error=' + error);
    return { success: false, message: '提携申請中にエラーが発生しました。' };
  }
}

function findExistingPromotionApplication(mediaId, promotionId) {
  if (!mediaId || !promotionId) {
    return null;
  }

  var query = buildQueryString({ promotion: promotionId, media: mediaId });
  var url = MEDIA_BASE_API_URL + '/promotion_apply/search' + (query ? '?' + query : '');
  var options = {
    method: 'get',
    headers: createAuthHeaders(),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var statusCode = response.getResponseCode();
    var contentText = response.getContentText();
    if (statusCode >= 200 && statusCode < 300) {
      var json = parseJson(contentText);
      var records = json && json.records;
      if (Array.isArray(records) && records.length) {
        return records[0];
      }
      return null;
    }

    Logger.log('Promotion application search failed: mediaId=' + mediaId + ' promotionId=' + promotionId + ' status=' + statusCode + ' body=' + contentText);
  } catch (error) {
    Logger.log('Promotion application search error: mediaId=' + mediaId + ' promotionId=' + promotionId + ' error=' + error);
  }

  return null;
}

function extractApiErrorMessage(responseJson) {
  if (!responseJson) {
    return '';
  }

  if (typeof responseJson.message === 'string') {
    return sanitizeString(responseJson.message);
  }

  var error = responseJson.error;
  if (!error) {
    return '';
  }

  if (typeof error === 'string') {
    return sanitizeString(error);
  }

  var messages = [];

  if (error.message) {
    messages.push(sanitizeString(error.message));
  }

  if (error.field_error) {
    var fieldErrors = error.field_error;
    if (!Array.isArray(fieldErrors)) {
      fieldErrors = [fieldErrors];
    }
    for (var i = 0; i < fieldErrors.length; i++) {
      messages.push(sanitizeString(fieldErrors[i]));
    }
  }

  if (error.global_error) {
    var globalErrors = error.global_error;
    if (!Array.isArray(globalErrors)) {
      globalErrors = [globalErrors];
    }
    for (var j = 0; j < globalErrors.length; j++) {
      messages.push(sanitizeString(globalErrors[j]));
    }
  }

  return messages.filter(function(message) {
    return !!message;
  }).join(' / ');
}

function isDuplicateApplicationFromResponse(responseJson) {
  if (!responseJson || !responseJson.error) {
    return false;
  }

  var error = responseJson.error;
  if (typeof error === 'string') {
    return isDuplicateApplicationMessage(error);
  }

  if (error.message && isDuplicateApplicationMessage(error.message)) {
    return true;
  }

  var fields = [].concat(error.field_error || [], error.global_error || []);
  for (var i = 0; i < fields.length; i++) {
    if (isDuplicateApplicationMessage(fields[i])) {
      return true;
    }
  }

  return false;
}

function isDuplicateApplicationMessage(message) {
  var text = sanitizeString(message);
  if (!text) {
    return false;
  }
  var lower = text.toLowerCase();
  return text.indexOf('既に') !== -1 || text.indexOf('すでに') !== -1 || lower.indexOf('already') !== -1;
}

function uniqueMessages(messages) {
  var seen = {};
  var unique = [];
  for (var i = 0; i < messages.length; i++) {
    var message = sanitizeString(messages[i]);
    if (!message || seen[message]) {
      continue;
    }
    seen[message] = true;
    unique.push(message);
  }
  return unique;
}

function resolveAffiliateUserId(identifier) {
  var parsed = parseAffiliateIdentifier(identifier);
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

  searchPatterns.push({ company: identifier });
  searchPatterns.push({ name: identifier });

  var uniqueKeys = {};
  for (var i = 0; i < searchPatterns.length; i++) {
    var params = searchPatterns[i];
    var key = JSON.stringify(params);
    if (uniqueKeys[key]) {
      continue;
    }
    uniqueKeys[key] = true;

    var record = searchAffiliate(params);
    if (record && record.id) {
      return record.id;
    }
  }
  return '';
}

function searchAffiliate(params) {
  var records = callAllPagesAPI(MEDIA_BASE_API_URL + '/user/search', MEDIA_AUTH_TOKEN, params);

  if (records.length === 1) {
    return records[0];
  }

  if (records.length > 1) {
    var exact = findExactAffiliateRecord(records, params);
    if (exact) {
      return exact;
    }
  }

  return null;
}

function findExactAffiliateRecord(records, params) {
  var company = params.company ? params.company.trim() : '';
  var name = params.name ? params.name.trim() : '';

  if (company && name) {
    for (var i = 0; i < records.length; i++) {
      var record = records[i];
      if (record.company === company && record.name === name) {
        return record;
      }
    }
  }

  if (company) {
    for (var j = 0; j < records.length; j++) {
      var recordByCompany = records[j];
      if (recordByCompany.company === company) {
        return recordByCompany;
      }
    }
  }

  if (name) {
    for (var k = 0; k < records.length; k++) {
      var recordByName = records[k];
      if (recordByName.name === name) {
        return recordByName;
      }
    }
  }

  return null;
}

function parseAffiliateIdentifier(identifier) {
  var value = (identifier || '').toString();
  var normalized = value.replace(/\r?\n/g, ' ').trim();
  var delimiters = ['＋', '+', '/', '／', '|', '｜', '>', '→'];

  for (var i = 0; i < delimiters.length; i++) {
    var delimiter = delimiters[i];
    if (normalized.indexOf(delimiter) !== -1) {
      var parts = normalized.split(delimiter).map(function(part) {
        return part.trim();
      }).filter(function(part) {
        return part;
      });
      if (parts.length >= 2) {
        return {
          company: parts[0],
          name: parts.slice(1).join(' ')
        };
      }
    }
  }

  var whitespaceParts = normalized.split(/[ \u3000]+/).filter(function(part) {
    return part;
  });

  if (whitespaceParts.length >= 2) {
    return {
      company: whitespaceParts.slice(0, whitespaceParts.length - 1).join(' '),
      name: whitespaceParts[whitespaceParts.length - 1]
    };
  }

  return {
    company: normalized,
    name: ''
  };
}

function normalizeRecords(records) {
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

function buildQueryString(params) {
  var queryParts = [];
  for (var key in params) {
    if (!params.hasOwnProperty(key)) {
      continue;
    }
    var value = params[key];
    if (value !== null && value !== undefined && value !== '') {
      queryParts.push(encodeURIComponent(key) + '=' + encodeURIComponent(value));
    }
  }
  return queryParts.join('&');
}

function createAuthHeaders(additionalHeaders) {
  var headers = {
    'X-Auth-Token': MEDIA_AUTH_TOKEN
  };

  if (additionalHeaders) {
    for (var key in additionalHeaders) {
      if (additionalHeaders.hasOwnProperty(key)) {
        headers[key] = additionalHeaders[key];
      }
    }
  }
  return headers;
}

function callAllPagesAPI(baseUrl, token, params) {
  var allRecords = [];
  var limit = 500;
  var offset = 0;
  var query = '';
  if (typeof params === 'string') {
    query = params;
  } else {
    query = buildQueryString(params || {});
  }
  var headers = {
    'X-Auth-Token': token
  };

  while (true) {
    var queryParts = [];
    if (query) {
      queryParts.push(query);
    }
    queryParts.push('limit=' + limit);
    queryParts.push('offset=' + offset);
    var url = baseUrl + (queryParts.length ? '?' + queryParts.join('&') : '');

    try {
      var response = UrlFetchApp.fetch(url, {
        method: 'get',
        headers: headers,
        muteHttpExceptions: true
      });
      var statusCode = response.getResponseCode();
      var contentText = response.getContentText();

      if (statusCode < 200 || statusCode >= 300) {
        Logger.log('callAllPagesAPI failed. Status: ' + statusCode + ' Body: ' + contentText);
        break;
      }

      var json = parseJson(contentText);
      var records = normalizeRecords(json ? json.records : null);
      if (!records.length) {
        break;
      }

      allRecords = allRecords.concat(records);
      if (records.length < limit) {
        break;
      }
      offset += records.length;
    } catch (error) {
      Logger.log('callAllPagesAPI error: ' + error);
      break;
    }
  }

  return allRecords;
}

function parseJson(contentText) {
  if (!contentText) {
    return null;
  }
  try {
    return JSON.parse(contentText);
  } catch (error) {
    Logger.log('JSON parse error: ' + error + ' content: ' + contentText);
    return null;
  }
}

function sanitizeString(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function resolveMediaCategoryId(value) {
  var resolved = resolveMediaLookupValue(value, getMediaCategoryLookup());
  if (!resolved && value) {
    Logger.log('Unable to resolve media category identifier: ' + value);
  }
  return resolved;
}

function resolveMediaTypeId(value) {
  var resolved = resolveMediaLookupValue(value, getMediaTypeLookup());
  if (!resolved && value) {
    Logger.log('Unable to resolve media type identifier: ' + value);
  }
  return resolved;
}

function resolveMediaLookupValue(value, lookup) {
  var key = sanitizeString(value);
  if (!key) {
    return '';
  }

  if (lookup.byId[key]) {
    return lookup.byId[key];
  }

  if (lookup.byName[key]) {
    return lookup.byName[key];
  }

  if (looksLikeId(key)) {
    return key;
  }

  return '';
}

function getMediaCategoryLookup() {
  if (MEDIA_CATEGORY_LOOKUP_CACHE) {
    return MEDIA_CATEGORY_LOOKUP_CACHE;
  }

  MEDIA_CATEGORY_LOOKUP_CACHE = buildMediaLookup(MEDIA_BASE_API_URL + '/media_category/search');
  return MEDIA_CATEGORY_LOOKUP_CACHE;
}

function getMediaTypeLookup() {
  if (MEDIA_TYPE_LOOKUP_CACHE) {
    return MEDIA_TYPE_LOOKUP_CACHE;
  }

  MEDIA_TYPE_LOOKUP_CACHE = buildMediaLookup(MEDIA_BASE_API_URL + '/media_type/search');
  return MEDIA_TYPE_LOOKUP_CACHE;
}

function buildMediaLookup(url) {
  var lookup = { byId: {}, byName: {} };
  var records = callAllPagesAPI(url, MEDIA_AUTH_TOKEN, {});

  for (var i = 0; i < records.length; i++) {
    var record = records[i];
    if (!record) {
      continue;
    }

    var id = sanitizeString(record.id);
    var name = sanitizeString(record.name);

    if (id) {
      lookup.byId[id] = id;
    }

    if (name && id) {
      lookup.byName[name] = id;
    }
  }

  return lookup;
}

function looksLikeId(value) {
  return /^(?:[0-9a-f]{32}|[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})$/i.test(value);
}


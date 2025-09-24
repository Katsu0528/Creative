var MEDIA_BASE_API_URL = 'https://otonari-asp.com/api/v1/m'.replace(/\/+$/, '');
var MEDIA_ACCESS_KEY = 'agqnoournapf';
var MEDIA_SECRET_KEY = '1kvu9dyv1alckgocc848socw';
var MEDIA_SHEET_NAME = 'メディア登録';
var MEDIA_DATA_START_ROW = 2;
var DEFAULT_MEDIA_CATEGORY_ID = '';
var DEFAULT_MEDIA_TYPE_ID = '';
var DEFAULT_MEDIA_PV_ROW = 0;
var DEFAULT_MEDIA_UU_ROW = 0;

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
  var results = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var rowNumber = MEDIA_DATA_START_ROW + i;
    var affiliateIdentifier = sanitizeString(row[MEDIA_COLUMNS.AFFILIATE_IDENTIFIER]);
    var mediaName = sanitizeString(row[MEDIA_COLUMNS.MEDIA_NAME]);
    var mediaUrl = sanitizeString(row[MEDIA_COLUMNS.MEDIA_URL]);
    var mediaCategoryId = sanitizeString(row[MEDIA_COLUMNS.MEDIA_CATEGORY_ID]) || DEFAULT_MEDIA_CATEGORY_ID;
    var mediaTypeId = sanitizeString(row[MEDIA_COLUMNS.MEDIA_TYPE_ID]) || DEFAULT_MEDIA_TYPE_ID;
    var mediaComment = sanitizeString(row[MEDIA_COLUMNS.MEDIA_COMMENT]);
    var existingMediaId = sanitizeString(row[MEDIA_COLUMNS.RESULT_MEDIA_ID]);
    var existingMessage = sanitizeString(row[MEDIA_COLUMNS.RESULT_MESSAGE]);

    Logger.log(
      'Processing row ' + rowNumber +
      ' affiliateIdentifier=' + affiliateIdentifier +
      ' mediaName=' + mediaName +
      ' mediaUrl=' + mediaUrl
    );

    if (!affiliateIdentifier || !mediaName || !mediaUrl) {
      var missingMessage = existingMessage || '必須項目が空欄のため処理をスキップしました。';
      Logger.log('Row ' + rowNumber + ' skipped: ' + missingMessage);
      results.push([existingMediaId, missingMessage]);
      continue;
    }

    if (!mediaCategoryId) {
      Logger.log('Row ' + rowNumber + ' skipped: メディアカテゴリーIDが空欄のため処理をスキップしました。');
      results.push([existingMediaId, 'メディアカテゴリーIDが空欄のため処理をスキップしました。']);
      continue;
    }

    if (!mediaTypeId) {
      Logger.log('Row ' + rowNumber + ' skipped: メディアタイプIDが空欄のため処理をスキップしました。');
      results.push([existingMediaId, 'メディアタイプIDが空欄のため処理をスキップしました。']);
      continue;
    }

    if (!mediaComment) {
      Logger.log('Row ' + rowNumber + ' skipped: メディア説明が空欄のため処理をスキップしました。');
      results.push([existingMediaId, 'メディア説明が空欄のため処理をスキップしました。']);
      continue;
    }

    if (existingMediaId) {
      var skippedMessage = existingMessage || '登録済みです。再登録する場合はG列とH列を空にしてください。';
      Logger.log('Row ' + rowNumber + ' skipped: ' + skippedMessage);
      results.push([existingMediaId, skippedMessage]);
      continue;
    }

    var userId = resolveAffiliateUserId(affiliateIdentifier);
    if (!userId) {
      Logger.log('Row ' + rowNumber + ' skipped: アフィリエイターIDを取得できませんでした。');
      results.push(['', 'アフィリエイターIDを取得できませんでした。']);
      continue;
    }

    Logger.log('Row ' + rowNumber + ' resolved userId=' + userId);

    var registrationResult = registerMedia(userId, mediaName, mediaUrl, mediaCategoryId, mediaTypeId, mediaComment);
    if (registrationResult.success) {
      Logger.log('Row ' + rowNumber + ' media registration succeeded. mediaId=' + registrationResult.id);
      results.push([registrationResult.id, '登録に成功しました。']);
    } else {
      Logger.log('Row ' + rowNumber + ' media registration failed: ' + registrationResult.message);
      results.push(['', registrationResult.message]);
    }
  }

  sheet.getRange(MEDIA_DATA_START_ROW, MEDIA_COLUMNS.RESULT_MEDIA_ID + 1, results.length, 2).setValues(results);
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
  var query = buildQueryString(params);
  var url = MEDIA_BASE_API_URL + '/user/search' + (query ? '?' + query : '');
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
      var records = normalizeRecords(json ? json.records : null);

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

    Logger.log('Affiliate search failed. Status: ' + statusCode + ' Body: ' + contentText);
    return null;
  } catch (error) {
    Logger.log('Affiliate search error: ' + error);
    return null;
  }
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
    'X-Auth-Token': MEDIA_ACCESS_KEY + ':' + MEDIA_SECRET_KEY
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


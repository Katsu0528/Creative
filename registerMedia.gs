var MEDIA_BASE_API_URL = 'https://otonari-asp.com/api/v1/m'.replace(/\/+$/, '');
var MEDIA_ACCESS_KEY = 'agqnoournapf';
var MEDIA_SECRET_KEY = '1kvu9dyv1alckgocc848socw';
var MEDIA_SHEET_NAME = 'メディア登録';
var MEDIA_DATA_START_ROW = 2;
var DEFAULT_MEDIA_CATEGORY_ID = '';
var DEFAULT_MEDIA_TYPE_ID = '';

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
  var values = sheet.getRange(MEDIA_DATA_START_ROW, 1, rowCount, 6).getValues();
  var results = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var affiliateIdentifier = sanitizeString(row[0]);
    var mediaName = sanitizeString(row[1]);
    var mediaUrl = sanitizeString(row[2]);
    var mediaCategory = sanitizeString(row[3]);
    var existingMediaId = sanitizeString(row[4]);
    var existingMessage = sanitizeString(row[5]);

    if (!affiliateIdentifier || !mediaName || !mediaUrl || !mediaCategory) {
      var missingMessage = existingMessage || '必須項目が空欄のため処理をスキップしました。';
      results.push([existingMediaId, missingMessage]);
      continue;
    }

    if (existingMediaId) {
      var skippedMessage = existingMessage || '登録済みです。再登録する場合はE列とF列を空にしてください。';
      results.push([existingMediaId, skippedMessage]);
      continue;
    }

    var userId = resolveAffiliateUserId(affiliateIdentifier);
    if (!userId) {
      results.push(['', 'アフィリエイターIDを取得できませんでした。']);
      continue;
    }

    var registrationResult = registerMedia(userId, mediaName, mediaUrl, mediaCategory);
    if (registrationResult.success) {
      results.push([registrationResult.id, '登録に成功しました。']);
    } else {
      results.push(['', registrationResult.message]);
    }
  }

  sheet.getRange(MEDIA_DATA_START_ROW, 5, results.length, 2).setValues(results);
}

function registerMedia(userId, mediaName, mediaUrl, mediaCategory) {
  var payload = {
    user: userId,
    name: mediaName,
    url: mediaUrl
  };

  if (mediaCategory) {
    payload.media_category = mediaCategory;
  } else if (DEFAULT_MEDIA_CATEGORY_ID) {
    payload.media_category = DEFAULT_MEDIA_CATEGORY_ID;
  }
  if (DEFAULT_MEDIA_TYPE_ID) {
    payload.media_type = DEFAULT_MEDIA_TYPE_ID;
  }

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
        return { success: true, id: record.id, message: '' };
      }
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


// Script to apply affiliate media to promotions listed in the active sheet.
// Columns:
// A: Advertiser name
// B: Promotion (advertisement) name
// C: Affiliate name
// D: (output) Result of the application attempt
//
// For each data row, the script searches the API for matching affiliates,
// advertisers, and promotions, then registers partnership applications for
// every media that belongs to the affiliate.

const API_BASE_URL = 'https://otonari-asp.com/api/v1/m';
const API_ACCESS_KEY = 'agqnoournapf';
const API_SECRET_KEY = '1kvu9dyv1alckgocc848socw';

// --- Public entry point ----------------------------------------------------

function applyAffiliatePartnerships() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('データ行が存在しません。');
    return;
  }

  if (!data[0][3]) {
    sheet.getRange(1, 4).setValue('申請結果');
  }

  const statusMessages = [];
  let successCount = 0;
  let skipCount = 0;
  let errorCount = 0;

  for (let i = 1; i < data.length; i++) {
    const rowNumber = i + 1; // 1-indexed for users
    const advertiserName = String(data[i][0] || '').trim();
    const promotionName = String(data[i][1] || '').trim();
    const affiliateName = String(data[i][2] || '').trim();

    Logger.log('Row %s: Start processing advertiser="%s" promotion="%s" affiliate="%s"', rowNumber, advertiserName, promotionName, affiliateName);

    if (!affiliateName || !promotionName) {
      Logger.log('Row %s: Skipped due to missing affiliate or promotion name.', rowNumber);
      statusMessages.push('アフィリエイター名または広告名が未入力のためスキップしました。');
      skipCount++;
      continue;
    }

    try {
      const affiliate = findAffiliateByName(affiliateName);
      if (!affiliate) {
        Logger.log('Row %s: Affiliate not found.', rowNumber);
        statusMessages.push('アフィリエイターが見つかりません。');
        skipCount++;
        continue;
      }

      const mediaList = listMediaByAffiliate(affiliate.id);
      if (mediaList.length === 0) {
        Logger.log('Row %s: No media found for affiliate id=%s.', rowNumber, affiliate.id);
        statusMessages.push('アフィリエイターに紐づくメディアが見つかりません。');
        skipCount++;
        continue;
      }

      const promotions = findPromotions(advertiserName, promotionName);
      if (promotions.length === 0) {
        Logger.log('Row %s: Promotion not found for advertiser="%s" promotion="%s".', rowNumber, advertiserName, promotionName);
        statusMessages.push('広告が見つかりません。');
        skipCount++;
        continue;
      }

      let applied = 0;
      promotions.forEach(function(promotion) {
        mediaList.forEach(function(media) {
          if (registerPromotionApplication(media.id, promotion.id)) {
            applied++;
            Logger.log('Row %s: Applied promotion id=%s to media id=%s.', rowNumber, promotion.id, media.id);
          }
        });
      });

      if (applied > 0) {
        Logger.log('Row %s: Completed with %s applications.', rowNumber, applied);
        statusMessages.push('提携申請を送信しました（' + applied + '件）。');
        successCount += applied;
      } else {
        Logger.log('Row %s: No new applications submitted (possibly already partnered).', rowNumber);
        statusMessages.push('既に提携済み、または新規申請はありませんでした。');
        skipCount++;
      }
    } catch (error) {
      statusMessages.push('エラー: ' + error.message);
      Logger.log('Row %s: %s', rowNumber, error.stack || error);
      errorCount++;
    }
  }

  sheet.getRange(2, 4, statusMessages.length, 1).setValues(statusMessages.map(function(message) {
    return [message];
  }));

  SpreadsheetApp.getUi().alert([
    '提携申請処理が完了しました。',
    '申請成功件数: ' + successCount,
    'スキップ件数: ' + skipCount,
    'エラー件数: ' + errorCount
  ].join('\n'));
}

// --- API helpers -----------------------------------------------------------

var affiliateCacheByName = {};
var advertiserCacheByName = {};
var promotionCacheByKey = {};
var mediaCacheByAffiliate = {};
var applicationCache = {};

function normalizeSearchText(value) {
  return String(value || '')
    .replace(/[＋+]/g, ' ')
    .replace(/[\u3000\s]+/g, ' ')
    .trim()
    .toLowerCase();
}

function buildAffiliateDisplayName(record) {
  if (!record) {
    return '';
  }
  var company = record.company_name || record.company || record.corporate_name || record.corporation || '';
  var person = record.name || record.user_name || record.contact_name || '';
  var pieces = [company, person].map(function(value) {
    return String(value || '').trim();
  }).filter(function(value) {
    return value.length > 0;
  });
  return pieces.join(' ');
}

function findAffiliateByName(name) {
  if (!name) {
    return null;
  }

  var normalized = normalizeSearchText(name);
  if (Object.prototype.hasOwnProperty.call(affiliateCacheByName, normalized)) {
    return affiliateCacheByName[normalized];
  }

  var parsed = parseAffiliateIdentifierParts(name);
  var searchParamsList = buildAffiliateSearchParams(parsed, name);
  var aggregatedRecords = [];

  for (var i = 0; i < searchParamsList.length; i++) {
    var params = searchParamsList[i];
    var query = buildAffiliateQueryString(params);
    var response = callApi('/user/search' + (query ? '?' + query : ''));
    var records = extractRecords(response.records);
    if (!records.length) {
      continue;
    }

    aggregatedRecords = aggregatedRecords.concat(records);

    var exact = selectAffiliateCandidate(records, parsed, normalized);
    if (exact) {
      affiliateCacheByName[normalized] = exact;
      return exact;
    }
  }

  if (aggregatedRecords.length) {
    var fallback = selectAffiliateCandidate(aggregatedRecords, parsed, normalized) || aggregatedRecords[0];
    affiliateCacheByName[normalized] = fallback || null;
    return fallback;
  }

  affiliateCacheByName[normalized] = null;
  return null;
}

function selectAffiliateCandidate(records, parsed, normalizedName) {
  if (!records || !records.length) {
    return null;
  }

  var combinedMatches = records.filter(function(record) {
    return normalizeSearchText(buildAffiliateDisplayName(record)) === normalizedName;
  });
  if (combinedMatches.length) {
    return combinedMatches[0];
  }

  var companyNormalized = normalizeSearchText(parsed.company);
  var personNormalized = normalizeSearchText(parsed.name);

  if (companyNormalized && personNormalized) {
    for (var i = 0; i < records.length; i++) {
      var record = records[i];
      if (normalizeSearchText(record.company || record.company_name || record.corporate_name || '') === companyNormalized &&
          normalizeSearchText(record.name || record.user_name || record.contact_name || '') === personNormalized) {
        return record;
      }
    }
  }

  if (companyNormalized) {
    for (var j = 0; j < records.length; j++) {
      var companyRecord = records[j];
      if (normalizeSearchText(companyRecord.company || companyRecord.company_name || companyRecord.corporate_name || '') === companyNormalized) {
        return companyRecord;
      }
    }
  }

  if (personNormalized) {
    for (var k = 0; k < records.length; k++) {
      var nameRecord = records[k];
      if (normalizeSearchText(nameRecord.name || nameRecord.user_name || nameRecord.contact_name || '') === personNormalized) {
        return nameRecord;
      }
    }
  }

  return null;
}

function buildAffiliateSearchParams(parsed, originalName) {
  var paramsList = [];

  if (parsed.company && parsed.name) {
    paramsList.push({ company: parsed.company, name: parsed.name });
  }

  if (parsed.company) {
    paramsList.push({ company: parsed.company });
  }

  if (parsed.name) {
    paramsList.push({ name: parsed.name });
  }

  if (originalName) {
    paramsList.push({ name: originalName });
  }

  var unique = [];
  var seen = {};
  for (var i = 0; i < paramsList.length; i++) {
    var params = paramsList[i];
    var key = JSON.stringify(params);
    if (seen[key]) {
      continue;
    }
    seen[key] = true;
    unique.push(params);
  }
  return unique;
}

function parseAffiliateIdentifierParts(identifier) {
  var value = String(identifier || '');
  var normalized = value.replace(/\r?\n/g, ' ').trim();
  if (!normalized) {
    return { company: '', name: '' };
  }

  var delimiters = ['＋', '+', '/', '／', '|', '｜', '>', '→'];
  for (var i = 0; i < delimiters.length; i++) {
    var delimiter = delimiters[i];
    if (normalized.indexOf(delimiter) !== -1) {
      var parts = normalized.split(delimiter).map(function(part) {
        return part.trim();
      }).filter(function(part) {
        return part.length > 0;
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
    return part.length > 0;
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

function buildAffiliateQueryString(params) {
  if (!params) {
    return '';
  }
  var parts = [];
  for (var key in params) {
    if (!params.hasOwnProperty(key)) {
      continue;
    }
    var value = params[key];
    if (value === null || value === undefined || value === '') {
      continue;
    }
    parts.push(encodeURIComponent(key) + '=' + encodeURIComponent(value));
  }
  return parts.join('&');
}

function findAdvertiserByName(name) {
  if (!name) {
    return null;
  }
  const normalized = name.toLowerCase();
  if (Object.prototype.hasOwnProperty.call(advertiserCacheByName, normalized)) {
    return advertiserCacheByName[normalized];
  }

  const response = callApi('/advertiser/search?name=' + encodeURIComponent(name));
  const all = extractRecords(response.records);
  const exact = all.filter(function(record) {
    return record && String(record.name || '').toLowerCase() === normalized;
  });
  const advertiser = exact.length > 0 ? exact[0] : all[0] || null;
  advertiserCacheByName[normalized] = advertiser || null;
  return advertiser;
}

function findPromotions(advertiserName, promotionName) {
  const key = [advertiserName || '', promotionName].join('::').toLowerCase();
  if (Object.prototype.hasOwnProperty.call(promotionCacheByKey, key)) {
    return promotionCacheByKey[key];
  }

  var query = '/promotion/search?name=' + encodeURIComponent(promotionName);
  var advertiser = null;
  if (advertiserName) {
    advertiser = findAdvertiserByName(advertiserName);
    if (advertiser && advertiser.id) {
      query += '&advertiser=' + encodeURIComponent(advertiser.id);
    }
  }

  const response = callApi(query);
  const all = extractRecords(response.records);
  const normalizedPromotion = promotionName.toLowerCase();
  const filtered = all.filter(function(record) {
    return record && String(record.name || '').toLowerCase() === normalizedPromotion;
  });

  const promotions = filtered.length > 0 ? filtered : all;
  promotionCacheByKey[key] = promotions;
  return promotions;
}

function listMediaByAffiliate(affiliateId) {
  if (!affiliateId) {
    return [];
  }
  if (Object.prototype.hasOwnProperty.call(mediaCacheByAffiliate, affiliateId)) {
    return mediaCacheByAffiliate[affiliateId];
  }

  const response = callApi('/media/search?user=' + encodeURIComponent(affiliateId));
  const allMedia = extractRecords(response.records) || [];
  const mediaList = [];
  var excludedStates = [];

  allMedia.forEach(function(record) {
    if (!record) {
      return;
    }
    const state = record.state;
    const isActive = (!state && state !== 0) || state === 1 || state === '1';
    if (isActive) {
      mediaList.push(record);
    } else {
      excludedStates.push(state);
    }
  });

  const stringStateOneIncludedCount = mediaList.filter(function(record) {
    return record && record.state === '1';
  }).length;
  const stringStateOneTotalCount = allMedia.filter(function(record) {
    return record && record.state === '1';
  }).length;

  Logger.log(
    'listMediaByAffiliate: affiliateId=%s totalRecords=%s activeReturned=%s stringStateOneIncluded=%s/%s excludedStates=%s',
    affiliateId,
    allMedia.length,
    mediaList.length,
    stringStateOneIncludedCount,
    stringStateOneTotalCount,
    excludedStates.map(function(state) {
      return state === undefined ? 'undefined' : state;
    }).join(',')
  );

  mediaCacheByAffiliate[affiliateId] = mediaList;
  return mediaList;
}

function registerPromotionApplication(mediaId, promotionId) {
  if (!mediaId || !promotionId) {
    Logger.log('registerPromotionApplication skipped: missing mediaId or promotionId (mediaId=%s, promotionId=%s)', mediaId, promotionId);
    return false;
  }
  const cacheKey = mediaId + '::' + promotionId;
  if (Object.prototype.hasOwnProperty.call(applicationCache, cacheKey)) {
    Logger.log('registerPromotionApplication skipped: already attempted cacheKey=%s', cacheKey);
    return false;
  }

  const payload = {
    media: mediaId,
    promotion: promotionId,
    state: 1
  };

  try {
    callApi('/promotion_apply/regist', {
      method: 'post',
      payload: payload
    });
    Logger.log('registerPromotionApplication success: mediaId=%s promotionId=%s', mediaId, promotionId);
    applicationCache[cacheKey] = true;
    return true;
  } catch (error) {
    // Ignore duplicate registration errors but keep the cache entry to avoid retries
    applicationCache[cacheKey] = true;
    if (isDuplicateApplicationError(error)) {
      Logger.log('registerPromotionApplication duplicate detected: mediaId=%s promotionId=%s', mediaId, promotionId);
      return false;
    }
    throw error;
  }
}

function isDuplicateApplicationError(error) {
  if (!error) {
    return false;
  }

  var messageText = error && error.message ? String(error.message) : '';
  if (messageText.indexOf('既に') !== -1 ||
      messageText.indexOf('すでに') !== -1 ||
      messageText.toLowerCase().indexOf('already') !== -1) {
    return true;
  }

  var errorResponse = error && error.response ? error.response : null;
  if (!errorResponse || !errorResponse.error) {
    return false;
  }

  var fieldErrors = errorResponse.error.field_error || [];
  if (!Array.isArray(fieldErrors)) {
    fieldErrors = [fieldErrors];
  }

  for (var i = 0; i < fieldErrors.length; i++) {
    var text = String(fieldErrors[i] || '');
    if (text.indexOf('既に') !== -1 || text.indexOf('すでに') !== -1 || text.toLowerCase().indexOf('already') !== -1) {
      return true;
    }
  }

  return false;
}

function callApi(path, options) {
  options = options || {};
  const url = API_BASE_URL + path;
  const headers = Object.assign({}, {
    'X-Auth-Token': API_ACCESS_KEY + ':' + API_SECRET_KEY
  }, options.headers || {});

  const params = {
    method: options.method || 'get',
    headers: headers,
    muteHttpExceptions: true
  };

  if (options.payload !== undefined) {
    params.payload = typeof options.payload === 'string'
      ? options.payload
      : JSON.stringify(options.payload);
    params.contentType = 'application/json';
  }

  Logger.log('API Request: %s %s payload=%s', params.method.toUpperCase(), url, params.payload || '');
  const response = UrlFetchApp.fetch(url, params);
  const status = response.getResponseCode();
  const text = response.getContentText();
  Logger.log('API Response: %s status=%s body=%s', url, status, text);

  if (status >= 200 && status < 300) {
    var json;
    if (text) {
      try {
        json = JSON.parse(text);
      } catch (parseError) {
        Logger.log('API Response Parse Error: %s error=%s', url, parseError);
        throw parseError;
      }
    } else {
      json = {};
    }
    Logger.log('API Response Parsed: %s data=%s', url, JSON.stringify(json));
    return json;
  }

  var errorMessage = 'APIリクエストに失敗しました。HTTP ' + status;
  var parsedBody = null;
  if (text) {
    try {
      parsedBody = JSON.parse(text);
    } catch (e) {
      parsedBody = null;
    }

    if (parsedBody) {
      var errorDetails = [];
      if (parsedBody.message) {
        errorDetails.push(String(parsedBody.message));
      }
      if (parsedBody.error) {
        if (parsedBody.error.message) {
          errorDetails.push(String(parsedBody.error.message));
        }
        if (parsedBody.error.field_error) {
          var fieldErrorList = parsedBody.error.field_error;
          if (!Array.isArray(fieldErrorList)) {
            fieldErrorList = [fieldErrorList];
          }
          fieldErrorList.forEach(function(entry) {
            errorDetails.push(String(entry));
          });
        }
      }
      if (errorDetails.length) {
        errorMessage += ' - ' + errorDetails.join(' / ');
      } else {
        errorMessage += ' - ' + text;
      }
    } else {
      errorMessage += ' - ' + text;
    }
  }

  var error = new Error(errorMessage);
  if (parsedBody) {
    error.response = parsedBody;
  }
  throw error;
}

function extractRecords(records) {
  if (!records) {
    return [];
  }
  if (Array.isArray(records)) {
    return records;
  }
  return [records];
}

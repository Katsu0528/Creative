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

    if (!affiliateName || !promotionName) {
      statusMessages.push('アフィリエイター名または広告名が未入力のためスキップしました。');
      skipCount++;
      continue;
    }

    try {
      const affiliate = findAffiliateByName(affiliateName);
      if (!affiliate) {
        statusMessages.push('アフィリエイターが見つかりません。');
        skipCount++;
        continue;
      }

      const mediaList = listMediaByAffiliate(affiliate.id);
      if (mediaList.length === 0) {
        statusMessages.push('アフィリエイターに紐づくメディアが見つかりません。');
        skipCount++;
        continue;
      }

      const promotions = findPromotions(advertiserName, promotionName);
      if (promotions.length === 0) {
        statusMessages.push('広告が見つかりません。');
        skipCount++;
        continue;
      }

      let applied = 0;
      promotions.forEach(function(promotion) {
        mediaList.forEach(function(media) {
          if (registerPromotionApplication(media.id, promotion.id)) {
            applied++;
          }
        });
      });

      if (applied > 0) {
        statusMessages.push('提携申請を送信しました（' + applied + '件）。');
        successCount += applied;
      } else {
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

function findAffiliateByName(name) {
  if (!name) {
    return null;
  }
  const normalized = name.toLowerCase();
  if (Object.prototype.hasOwnProperty.call(affiliateCacheByName, normalized)) {
    return affiliateCacheByName[normalized];
  }

  const response = callApi('/user/search?name=' + encodeURIComponent(name));
  const candidates = extractRecords(response.records).filter(function(record) {
    return record && String(record.name || '').toLowerCase() === normalized;
  });

  const affiliate = candidates.length > 0 ? candidates[0] : extractRecords(response.records)[0] || null;
  affiliateCacheByName[normalized] = affiliate || null;
  return affiliate;
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
  const mediaList = extractRecords(response.records).filter(function(record) {
    return record && (!record.state || record.state === 1);
  });
  mediaCacheByAffiliate[affiliateId] = mediaList;
  return mediaList;
}

function registerPromotionApplication(mediaId, promotionId) {
  if (!mediaId || !promotionId) {
    return false;
  }
  const cacheKey = mediaId + '::' + promotionId;
  if (Object.prototype.hasOwnProperty.call(applicationCache, cacheKey)) {
    return false;
  }

  const payload = {
    media: mediaId,
    promotion: promotionId,
    state: 0
  };

  try {
    callApi('/promotion_apply/regist', {
      method: 'post',
      payload: payload
    });
    applicationCache[cacheKey] = true;
    return true;
  } catch (error) {
    // Ignore duplicate registration errors but keep the cache entry to avoid retries
    applicationCache[cacheKey] = true;
    if (error && error.message) {
      const messageText = String(error.message);
      if (messageText.indexOf('既に') !== -1 || messageText.toLowerCase().indexOf('already') !== -1) {
        return false;
      }
    }
    throw error;
  }
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

  const response = UrlFetchApp.fetch(url, params);
  const status = response.getResponseCode();
  const text = response.getContentText();

  if (status >= 200 && status < 300) {
    return text ? JSON.parse(text) : {};
  }

  var errorMessage = 'APIリクエストに失敗しました。HTTP ' + status;
  if (text) {
    try {
      const parsed = JSON.parse(text);
      if (parsed && parsed.message) {
        errorMessage += ' - ' + parsed.message;
      } else {
        errorMessage += ' - ' + text;
      }
    } catch (e) {
      errorMessage += ' - ' + text;
    }
  }
  throw new Error(errorMessage);
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

var PROMOTION_LOOKUP_CACHE_BY_INPUT = {};
var PROMOTION_LOOKUP_CACHE_BY_ID = {};
var MEDIA_LOOKUP_CACHE_BY_INPUT = {};
var MEDIA_LOOKUP_CACHE_BY_ID = {};

function registerPromotionApplicationsFromWeb(rowEntries) {
  if (!Array.isArray(rowEntries)) {
    throw new Error('送信されたデータの形式が正しくありません。');
  }

  var normalizedRows = [];
  var logs = [];

  logs.push({ level: 'info', message: '受信したリクエスト行数: ' + rowEntries.length });

  rowEntries.forEach(function(entry, index) {
    var data = entry || {};
    var rowNumber = Number(data.rowNumber);
    if (!rowNumber || rowNumber < 1) {
      rowNumber = index + 1;
    }

    var promotionIdentifier = sanitizeString(
      data.promotionIdentifier || data.promotionName || data.promotion || data.advertisement || ''
    );
    var mediaIdentifier = sanitizeString(
      data.mediaIdentifier || data.mediaName || data.media || ''
    );

    if (!promotionIdentifier && !mediaIdentifier) {
      return;
    }

    normalizedRows.push({
      rowNumber: rowNumber,
      promotionIdentifier: promotionIdentifier,
      mediaIdentifier: mediaIdentifier,
    });
  });

  if (!normalizedRows.length) {
    return {
      summary: { total: 0, success: 0, skipped: 0, errors: 0 },
      results: [],
      logs: logs.concat({ level: 'warning', message: '処理対象の行がありませんでした。' }),
    };
  }

  var results = [];
  var summary = { total: normalizedRows.length, success: 0, skipped: 0, errors: 0 };

  normalizedRows.forEach(function(row) {
    logs.push({
      level: 'info',
      message: formatRowLabel(row.rowNumber) + ' 入力内容を確認しました。',
      detail: '広告: ' + (row.promotionIdentifier || '(未入力)') + ' / メディア: ' + (row.mediaIdentifier || '(未入力)'),
    });

    if (!row.promotionIdentifier || !row.mediaIdentifier) {
      results.push({
        rowNumber: row.rowNumber,
        status: 'skipped',
        promotionId: '',
        mediaId: '',
        message: '広告名/IDまたはメディア名/IDが未入力のためスキップしました。',
      });
      summary.skipped++;
      logs.push({
        level: 'warning',
        message: formatRowLabel(row.rowNumber) + ' 広告またはメディアの指定が無いためスキップしました。',
      });
      return;
    }

    var promotionRecord = findPromotionRecordByIdentifier(row.promotionIdentifier);
    if (!promotionRecord) {
      results.push({
        rowNumber: row.rowNumber,
        status: 'error',
        promotionId: '',
        mediaId: '',
        message: '広告名からIDを取得できませんでした。',
      });
      summary.errors++;
      logs.push({
        level: 'error',
        message: formatRowLabel(row.rowNumber) + ' 広告IDを特定できませんでした。',
        detail: '入力: ' + row.promotionIdentifier,
      });
      return;
    }

    var mediaRecord = findMediaRecordByIdentifier(row.mediaIdentifier);
    if (!mediaRecord) {
      results.push({
        rowNumber: row.rowNumber,
        status: 'error',
        promotionId: promotionRecord.id || '',
        mediaId: '',
        message: 'メディア名からIDを取得できませんでした。',
      });
      summary.errors++;
      logs.push({
        level: 'error',
        message: formatRowLabel(row.rowNumber) + ' メディアIDを特定できませんでした。',
        detail: '入力: ' + row.mediaIdentifier,
      });
      return;
    }

    var applicationResult = submitPromotionApplication(mediaRecord.id, promotionRecord.id);
    var message = applicationResult.message || '';
    if (applicationResult.success) {
      summary.success++;
      results.push({
        rowNumber: row.rowNumber,
        status: 'success',
        promotionId: promotionRecord.id || '',
        mediaId: mediaRecord.id || '',
        message: message || (applicationResult.duplicate ? '既に提携済みです。' : '提携申請を送信しました。'),
      });
      logs.push({
        level: 'success',
        message: formatRowLabel(row.rowNumber) + ' 提携申請を送信しました。',
        detail: '広告ID: ' + promotionRecord.id + ' / メディアID: ' + mediaRecord.id,
      });
    } else {
      summary.errors++;
      results.push({
        rowNumber: row.rowNumber,
        status: 'error',
        promotionId: promotionRecord.id || '',
        mediaId: mediaRecord.id || '',
        message: message || '提携申請に失敗しました。',
      });
      logs.push({
        level: 'error',
        message: formatRowLabel(row.rowNumber) + ' 提携申請に失敗しました。',
        detail: message || '詳細は処理結果を確認してください。',
      });
    }
  });

  return {
    summary: summary,
    results: results,
    logs: logs,
  };
}

function formatRowLabel(rowNumber) {
  return '行' + (rowNumber || '-');
}

function findPromotionRecordByIdentifier(identifier) {
  var key = normalizeLookupKey(identifier);
  if (Object.prototype.hasOwnProperty.call(PROMOTION_LOOKUP_CACHE_BY_INPUT, key)) {
    return PROMOTION_LOOKUP_CACHE_BY_INPUT[key];
  }

  var record = null;
  if (looksLikePromotionIdentifier(identifier)) {
    record = fetchPromotionRecordById(identifier);
  }
  if (!record) {
    record = searchPromotionRecordByName(identifier);
  }

  PROMOTION_LOOKUP_CACHE_BY_INPUT[key] = record || null;
  return record;
}

function findMediaRecordByIdentifier(identifier) {
  var key = normalizeLookupKey(identifier);
  if (Object.prototype.hasOwnProperty.call(MEDIA_LOOKUP_CACHE_BY_INPUT, key)) {
    return MEDIA_LOOKUP_CACHE_BY_INPUT[key];
  }

  var record = null;
  if (looksLikeMediaIdentifier(identifier)) {
    record = fetchMediaRecordById(identifier);
  }
  if (!record) {
    record = searchMediaRecordByName(identifier);
  }

  MEDIA_LOOKUP_CACHE_BY_INPUT[key] = record || null;
  return record;
}

function fetchPromotionRecordById(id) {
  var normalizedId = sanitizeString(id);
  if (!normalizedId) {
    return null;
  }
  if (Object.prototype.hasOwnProperty.call(PROMOTION_LOOKUP_CACHE_BY_ID, normalizedId)) {
    return PROMOTION_LOOKUP_CACHE_BY_ID[normalizedId];
  }
  var records = callAllPagesAPI(
    MEDIA_BASE_API_URL + '/promotion/search',
    MEDIA_AUTH_TOKEN,
    { id: normalizedId }
  );
  var record = records && records.length ? records[0] : null;
  PROMOTION_LOOKUP_CACHE_BY_ID[normalizedId] = record || null;
  return record;
}

function fetchMediaRecordById(id) {
  var normalizedId = sanitizeString(id);
  if (!normalizedId) {
    return null;
  }
  if (Object.prototype.hasOwnProperty.call(MEDIA_LOOKUP_CACHE_BY_ID, normalizedId)) {
    return MEDIA_LOOKUP_CACHE_BY_ID[normalizedId];
  }
  var records = callAllPagesAPI(
    MEDIA_BASE_API_URL + '/media/search',
    MEDIA_AUTH_TOKEN,
    { id: normalizedId }
  );
  var record = records && records.length ? records[0] : null;
  MEDIA_LOOKUP_CACHE_BY_ID[normalizedId] = record || null;
  return record;
}

function searchPromotionRecordByName(name) {
  var query = sanitizeString(name);
  if (!query) {
    return null;
  }
  var records = callAllPagesAPI(
    MEDIA_BASE_API_URL + '/promotion/search',
    MEDIA_AUTH_TOKEN,
    { name: query }
  );
  return selectBestMatchRecord(records, query);
}

function searchMediaRecordByName(name) {
  var query = sanitizeString(name);
  if (!query) {
    return null;
  }
  var records = callAllPagesAPI(
    MEDIA_BASE_API_URL + '/media/search',
    MEDIA_AUTH_TOKEN,
    { name: query }
  );
  return selectBestMatchRecord(records, query);
}

function selectBestMatchRecord(records, query) {
  var list = Array.isArray(records) ? records : [];
  if (!list.length) {
    return null;
  }
  var normalizedQuery = normalizeLookupKey(query);
  for (var i = 0; i < list.length; i++) {
    var record = list[i];
    if (normalizeLookupKey(record && record.name) === normalizedQuery) {
      return record;
    }
  }
  return list[0];
}

function normalizeLookupKey(value) {
  return sanitizeString(value).replace(/[\s\u3000]+/g, ' ').toLowerCase();
}

function looksLikePromotionIdentifier(value) {
  var text = sanitizeString(value);
  return looksLikeId(text) || /^[0-9]+$/.test(text);
}

function looksLikeMediaIdentifier(value) {
  var text = sanitizeString(value);
  return looksLikeId(text) || /^[0-9]+$/.test(text);
}

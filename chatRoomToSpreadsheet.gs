'use strict';

/**
 * Google ChatのスペースURLからスペースIDを抜き出してメッセージを取得し、
 * 代理店No. / 案件名 / 単価をシートに出力する。
 *
 * 必要スコープ例:
 * - https://www.googleapis.com/auth/chat.messages.readonly
 * - https://www.googleapis.com/auth/spreadsheets
 */
var CHAT_ROOM_URL = 'https://chat.google.com/room/AAAAecfI5fs?cls=7';
var AGENCY_MASTER_SHEET_NAME = '代理店マスタ';
var OUTPUT_SHEET_NAME = 'チャット案件一覧';

function summarizeChatRoomToSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName(AGENCY_MASTER_SHEET_NAME);
  if (!masterSheet) {
    throw new Error('代理店マスタシートが見つかりません: ' + AGENCY_MASTER_SHEET_NAME);
  }

  var agencyMap = loadAgencyMaster_(masterSheet);
  var spaceId = parseSpaceIdFromRoomUrl_(CHAT_ROOM_URL);
  var messages = listChatMessages_(spaceId);

  if (!messages.length) {
    throw new Error('対象スペースにメッセージが見つかりませんでした。');
  }

  messages.sort(function(a, b) {
    return new Date(a.createTime || 0).getTime() - new Date(b.createTime || 0).getTime();
  });

  var rows = [];
  messages.forEach(function(message) {
    var body = message.text || message.formattedText || '';
    if (!body) {
      return;
    }

    var agencyKeyword = extractAgencyKeywordAfterMention_(body);
    var agencyNo = agencyKeyword ? (agencyMap[normalizeKey_(agencyKeyword)] || '') : '';
    var projectName = extractProjectName_(body);
    var unitPrice = extractUnitPrice_(body);

    if (!agencyNo && !projectName && !unitPrice) {
      return;
    }

    rows.push([agencyNo, projectName, unitPrice]);
  });

  writeSummaryRows_(ss, rows);
}

function parseSpaceIdFromRoomUrl_(roomUrl) {
  var match = String(roomUrl || '').match(/\/room\/([^?/#]+)/);
  if (!match) {
    throw new Error('ルームURLからスペースIDを取得できませんでした: ' + roomUrl);
  }
  return match[1];
}

function listChatMessages_(spaceId) {
  var token = ScriptApp.getOAuthToken();
  var baseUrl = 'https://chat.googleapis.com/v1/spaces/' + encodeURIComponent(spaceId) + '/messages';
  var pageToken = '';
  var all = [];

  while (true) {
    var url = baseUrl + '?pageSize=1000';
    if (pageToken) {
      url += '&pageToken=' + encodeURIComponent(pageToken);
    }

    var res = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + token }
    });

    var code = res.getResponseCode();
    if (code < 200 || code >= 300) {
      throw new Error('Google Chat APIエラー: HTTP ' + code + ' ' + res.getContentText());
    }

    var json = JSON.parse(res.getContentText() || '{}');
    if (json.messages && json.messages.length) {
      all = all.concat(json.messages);
    }

    if (!json.nextPageToken) {
      break;
    }
    pageToken = json.nextPageToken;
  }

  return all;
}

function loadAgencyMaster_(sheet) {
  var values = sheet.getDataRange().getValues();
  var map = {};

  for (var i = 1; i < values.length; i++) {
    var agencyNo = values[i][0];
    var agencyName = values[i][1];
    if (!agencyNo || !agencyName) {
      continue;
    }
    map[normalizeKey_(agencyName)] = agencyNo;
  }

  return map;
}

function extractAgencyKeywordAfterMention_(text) {
  var normalized = sanitizeText_(text);
  var lines = normalized.split(/\r?\n/);

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;

    var match = line.match(/^(@[^\s　]+|<[^>]+>|\[[^\]]+\])\s+([^\s　]+)/);
    if (match) {
      return match[2];
    }
  }

  return '';
}

function extractProjectName_(text) {
  var normalized = sanitizeText_(text);
  var pattern = /(?:【?\s*(?:商材名|商品名|案件名)\s*】?)\s*[：:\-]\s*([^\n\r]+)/i;
  var match = normalized.match(pattern);
  return match ? cleanValue_(match[1]) : '';
}

function extractUnitPrice_(text) {
  var normalized = sanitizeText_(text);
  var pattern = /(?:【?\s*(?:成果単価\s*\(貴社卸・税抜\)|成果単価|単価|成果報酬|広告単価)\s*】?)\s*[：:\-]?\s*([0-9０-９,，]+(?:\.[0-9０-９]+)?)/i;
  var match = normalized.match(pattern);
  if (!match) {
    return '';
  }

  var num = toHalfWidthNumber_(match[1]).replace(/[，,]/g, '');
  return num;
}

function sanitizeText_(text) {
  return String(text || '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\u00A0/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .trim();
}

function cleanValue_(value) {
  return String(value || '').replace(/^\s+|\s+$/g, '');
}

function normalizeKey_(value) {
  return String(value || '')
    .replace(/[！-～]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    })
    .replace(/\s+/g, '')
    .toLowerCase();
}

function toHalfWidthNumber_(value) {
  return String(value || '').replace(/[０-９．]/g, function(s) {
    if (s === '．') return '.';
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
}

function writeSummaryRows_(ss, rows) {
  var sheet = ss.getSheetByName(OUTPUT_SHEET_NAME) || ss.insertSheet(OUTPUT_SHEET_NAME);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 3).setValues([['代理店No.', '案件名', '単価']]);

  if (!rows.length) {
    return;
  }

  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
}

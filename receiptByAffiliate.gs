'use strict';

// Summary of confirmed results by affiliate and output to "受領" sheet.
function summarizeConfirmedResultsByAffiliate() {
  var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  var dateSs = SpreadsheetApp.openById(DATE_SPREADSHEET_ID);
  var dateSheet = dateSs.getSheetByName(DATE_SHEET_NAME);
  var start = dateSheet.getRange('B2').getValue();
  var end = dateSheet.getRange('C2').getValue();
  if (!(start instanceof Date) || !(end instanceof Date)) {
    alertUi_('B2/C2 に日付が入力されていません。');
    throw new Error('日付が正しく入力されていません');
  }
  start.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);

  var baseUrl = 'https://otonari-asp.com/api/v1/m'.replace(/\/+$/, '');
  var headers = { 'X-Auth-Token': 'agqnoournapf:1kvu9dyv1alckgocc848socw' };

  function fetchRecords(dateField, states) {
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
    var url = baseUrl + '/action_log_raw/search?' + baseParams + '&offset=0';
    var response;
    for (var attempt = 0; attempt < 3; attempt++) {
      try {
        response = UrlFetchApp.fetch(url, { method: 'get', headers: headers });
        break;
      } catch (e) {
        if (attempt === 2) {
          alertUi_('API取得に失敗しました: ' + e);
          return null;
        }
        Utilities.sleep(1000 * Math.pow(2, attempt));
      }
    }
    var json = JSON.parse(response.getContentText());
    var result = json.records && json.records.length ? json.records : [];
    var count = json.header && json.header.count ? json.header.count : result.length;
    var fetched = result.length;
    if (fetched < count) {
      var requests = [];
      for (var offset = fetched; offset < count; offset += 500) {
        requests.push({
          url: baseUrl + '/action_log_raw/search?' + baseParams + '&offset=' + offset,
          method: 'get',
          headers: headers
        });
      }
      var responses = UrlFetchApp.fetchAll(requests);
      responses.forEach(function(res) {
        try {
          var j = JSON.parse(res.getContentText());
          if (j.records && j.records.length) {
            result = result.concat(j.records);
          }
        } catch (e) {}
      });
    }
    return result;
  }

  var records = fetchRecords('apply_unix', [1]);
  if (records === null) {
    throw new Error('確定成果の取得に失敗しました');
  }

  var advertiserSet = {}, promotionSet = {}, mediaSet = {}, userSet = {};
  records.forEach(function(rec) {
    if (rec.advertiser || rec.advertiser === 0) advertiserSet[rec.advertiser] = true;
    if (rec.promotion) promotionSet[rec.promotion] = true;
    if (rec.media) mediaSet[rec.media] = true;
    if (rec.user) userSet[rec.user] = true;
  });

  var advertiserInfoMap = {}, promotionMap = {}, promotionAdvertiserMap = {}, mediaInfoMap = {}, userMap = {};

  function fetchNames(ids, endpoint, map, nameResolver) {
    for (var i = 0; i < ids.length; i += 100) {
      var batch = ids.slice(i, i + 100);
      var requests = batch.map(function(id) {
        return { url: baseUrl + '/' + endpoint + '/search?id=' + encodeURIComponent(id), method: 'get', headers: headers };
      });
      var responses = UrlFetchApp.fetchAll(requests);
      responses.forEach(function(res, idx) {
        var id = batch[idx];
        try {
          var json = JSON.parse(res.getContentText());
          var rec = Array.isArray(json.records) ? json.records[0] : json.records;
          map[id] = nameResolver(rec) || id;
        } catch (e) {
          map[id] = id;
        }
      });
    }
  }

  function fetchPromotions(ids) {
    for (var i = 0; i < ids.length; i += 100) {
      var batch = ids.slice(i, i + 100);
      var requests = batch.map(function(id) {
        return { url: baseUrl + '/promotion/search?id=' + encodeURIComponent(id), method: 'get', headers: headers };
      });
      var responses = UrlFetchApp.fetchAll(requests);
      responses.forEach(function(res, idx) {
        var id = batch[idx];
        try {
          var json = JSON.parse(res.getContentText());
          var rec = Array.isArray(json.records) ? json.records[0] : json.records;
          promotionMap[id] = rec && rec.name;
          if (rec && (rec.advertiser || rec.advertiser === 0)) {
            promotionAdvertiserMap[id] = rec.advertiser;
            advertiserSet[rec.advertiser] = true;
          }
        } catch (e) {
          promotionMap[id] = id;
        }
      });
    }
  }

  fetchPromotions(Object.keys(promotionSet));
  fetchNames(Object.keys(advertiserSet), 'advertiser', advertiserInfoMap, function(rec) {
    if (!rec) return { company: '', name: '' };
    return { company: rec.company || '', name: rec.name || '' };
  });
  fetchNames(Object.keys(mediaSet), 'media', mediaInfoMap, function(rec) {
    if (!rec) return { company: '', user: '' };
    if (rec.user) userSet[rec.user] = true;
    return { company: rec.name || '', user: rec.user || '' };
  });
  fetchNames(Object.keys(userSet), 'user', userMap, function(rec) {
    return rec && rec.name;
  });

  var advertiserMap = {}, mediaMap = {};
  Object.keys(advertiserInfoMap).forEach(function(id) {
    var info = advertiserInfoMap[id];
    var company = info.company || '';
    var person = info.name || '';
    advertiserMap[id] = toFullWidthSpace_(company && person ? company + ' ' + person : (company || person));
  });
  Object.keys(mediaInfoMap).forEach(function(id) {
    var info = mediaInfoMap[id];
    var person = info.user ? (userMap[info.user] || '') : '';
    mediaMap[id] = info.company && person ? info.company + ' ' + person : (info.company || person);
  });

  var summary = {};
  records.forEach(function(rec) {
    var advId = (rec.advertiser || rec.advertiser === 0) ? rec.advertiser : promotionAdvertiserMap[rec.promotion];
    var advertiser = advId ? (advertiserMap[advId] || advId) : '';
    var ad = rec.promotion ? (promotionMap[rec.promotion] || rec.promotion) : '';
    var affiliate = rec.media ? (mediaMap[rec.media] || rec.media) : '';
    var unit = Number(rec.gross_action_cost || 0);
    var key = advertiser + '\u0000' + ad + '\u0000' + affiliate + '\u0000' + unit;
    var entry = summary[key] || (summary[key] = {advertiser: advertiser, ad: ad, affiliate: affiliate, unit: unit, count: 0, amount: 0});
    entry.count++;
    entry.amount += unit;
  });

  var sheet = ss.getSheetByName('受領') || ss.insertSheet('受領');
  sheet.clearContents();
  var headers = ['広告主', '広告', 'アフィリエイター', '単価', '件数', '金額'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var rows = Object.keys(summary).map(function(k) {
    var s = summary[k];
    return [s.advertiser, s.ad, s.affiliate, s.unit, s.count, s.amount];
  }).sort(function(a, b) {
    if (a[0] < b[0]) return -1;
    if (a[0] > b[0]) return 1;
    if (a[1] < b[1]) return -1;
    if (a[1] > b[1]) return 1;
    if (a[2] < b[2]) return -1;
    if (a[2] > b[2]) return 1;
    if (a[3] < b[3]) return -1;
    if (a[3] > b[3]) return 1;
    return 0;
  });
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

// Entry point named 受領 as requested.
function 受領() {
  summarizeConfirmedResultsByAffiliate();
}


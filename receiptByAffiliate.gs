'use strict';

// Normalize names by removing both half-width and full-width spaces for
// consistent aggregation regardless of spacing differences.
function normalizeName_(str) {
  return typeof str === 'string' ? str.replace(/[\s\u3000]/g, '') : '';
}

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

  function getId_(val) {
    return (typeof val === 'object' && val !== null && 'id' in val) ? val.id : val;
  }

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
    var advId = getId_(rec.advertiser);
    if (advId || advId === 0) advertiserSet[advId] = true;
    var promotionId = getId_(rec.promotion);
    if (promotionId) promotionSet[promotionId] = true;
    var mediaId = getId_(rec.media);
    if (mediaId || mediaId === 0) mediaSet[mediaId] = true;
    var userId = getId_(rec.user);
    if (userId || userId === 0) userSet[userId] = true;
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
            var adv = getId_(rec.advertiser);
            promotionAdvertiserMap[id] = adv;
            if (adv || adv === 0) advertiserSet[adv] = true;
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
    var userId = getId_(rec.user);
    if (userId || userId === 0) userSet[userId] = true;
    return { company: rec.name || '', user: userId };
  });
  fetchNames(Object.keys(userSet), 'user', userMap, function(rec) {
    if (!rec) return { company: '', name: '' };
    return { company: rec.company || '', name: rec.name || '' };
  });

  var advertiserMap = {}, mediaMap = {};
  Object.keys(advertiserInfoMap).forEach(function(id) {
    var info = advertiserInfoMap[id];
    advertiserMap[id] = {
      company: toFullWidthSpace_(info.company || ''),
      person: toFullWidthSpace_(info.name || '')
    };
  });
  Object.keys(mediaInfoMap).forEach(function(id) {
    var info = mediaInfoMap[id];
    var userInfo = info.user ? userMap[info.user] : null;
    mediaMap[id] = {
      company: toFullWidthSpace_((userInfo && userInfo.company) || info.company || ''),
      person: toFullWidthSpace_((userInfo && userInfo.name) || '')
    };
  });

  var excludedNames = {};
  var genreSheet = ss.getSheetByName('【毎月更新】ジャンル');
  if (genreSheet) {
    var genreValues = genreSheet.getRange(1, 1, genreSheet.getLastRow(), 1).getValues();
    genreValues.forEach(function(row) {
      var name = row[0];
      if (name) excludedNames[normalizeName_(String(name))] = true;
    });
  }

  var summary = {};
  records.forEach(function(rec) {
    var promotionId = getId_(rec.promotion);
    var advId = (rec.advertiser || rec.advertiser === 0) ? getId_(rec.advertiser) : promotionAdvertiserMap[promotionId];
    var advertiserInfo = advId ? (advertiserMap[advId] || { company: toFullWidthSpace_(String(advId)), person: '' }) : { company: '', person: '' };
    var ad = promotionId ? (promotionMap[promotionId] || promotionId) : '';
    var mediaId = getId_(rec.media);
    var affiliateInfo = (mediaId || mediaId === 0) ? (mediaMap[mediaId] || { company: toFullWidthSpace_(String(mediaId)), person: '' }) : { company: '', person: '' };

    var excluded = false;
    if (affiliateInfo.company && affiliateInfo.person) {
      var combinedKey = normalizeName_(affiliateInfo.company + affiliateInfo.person);
      excluded = excludedNames[combinedKey];
    } else {
      var personKey = normalizeName_(affiliateInfo.person);
      excluded = personKey && excludedNames[personKey];
    }
    if (excluded) return; // Skip excluded affiliates

    // Use net unit price for receipts
    var unit = Number(rec.net_action_cost || 0);
    var key = [
      normalizeName_(advertiserInfo.company),
      normalizeName_(advertiserInfo.person),
      ad,
      normalizeName_(affiliateInfo.company),
      normalizeName_(affiliateInfo.person),
      unit
    ].join('\u0000');
    var entry = summary[key] || (summary[key] = {
      advertiserCompany: advertiserInfo.company,
      advertiserPerson: advertiserInfo.person,
      ad: ad,
      affiliateCompany: affiliateInfo.company,
      affiliatePerson: affiliateInfo.person,
      unit: unit,
      count: 0,
      amount: 0
    });
    entry.count++;
    entry.amount += unit;
  });

  var sheet = ss.getSheetByName('受領') || ss.insertSheet('受領');
  sheet.clearContents();
  var headers = ['広告主会社', '広告主氏名', '広告', 'アフィリエイター会社', 'アフィリエイター氏名', '単価', '件数', '金額'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var rows = Object.keys(summary).map(function(k) {
    var s = summary[k];
    return [s.advertiserCompany, s.advertiserPerson, s.ad, s.affiliateCompany, s.affiliatePerson, s.unit, s.count, s.amount];
  }).sort(function(a, b) {
    if (a[3] < b[3]) return -1; // sort by affiliate company first
    if (a[3] > b[3]) return 1;
    if (a[4] < b[4]) return -1;
    if (a[4] > b[4]) return 1;
    if (a[0] < b[0]) return -1;
    if (a[0] > b[0]) return 1;
    if (a[1] < b[1]) return -1;
    if (a[1] > b[1]) return 1;
    if (a[2] < b[2]) return -1;
    if (a[2] > b[2]) return 1;
    if (a[5] < b[5]) return -1;
    if (a[5] > b[5]) return 1;
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


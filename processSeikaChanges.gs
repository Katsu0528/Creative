function processSeikaChanges() {
  var originalSs = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = originalSs.getSheetByName('成果変更用');
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert('成果変更用 sheet not found.');
    Logger.log('成果変更用 sheet not found');
    return;
  }
  Logger.log('処理開始');


  // Step 1: Build lookup maps from source sheet
  // Read rows starting from A3:E3 until a blank row is encountered.
  // Track approve/cancel keys separately for cid, args and combined pairs
  var approvePairs = {};
  var cancelPairs = {};
  var approveCids = {};
  var approveArgsMap = {};
  var cancelCids = {};
  var cancelArgsMap = {};
  var srcData = [];
  var row = 3;
  while (true) {
    var values = sourceSheet.getRange(row, 1, 1, 5).getValues()[0];
    Logger.log('Row ' + row + ' values: ' + values.join(', '));
    var approveCid = values[0];
    var approveArgs = values[1];
    var cancelCid = values[3];
    var cancelArgs = values[4];
    if (!approveCid && !approveArgs && !cancelCid && !cancelArgs) break;
    srcData.push(values);
    if (approveCid && approveArgs) {
      approvePairs[approveCid + '\u0000' + approveArgs] = true;
    } else {
      if (approveCid) approveCids[approveCid] = true;
      if (approveArgs) approveArgsMap[approveArgs] = true;
    }
    if (cancelCid && cancelArgs) {
      cancelPairs[cancelCid + '\u0000' + cancelArgs] = true;
    } else {
      if (cancelCid) cancelCids[cancelCid] = true;
      if (cancelArgs) cancelArgsMap[cancelArgs] = true;
    }
    row++;
  }
  Logger.log('Lookup maps: approvePairs=' + Object.keys(approvePairs).length +
              ', approveCids=' + Object.keys(approveCids).length +
              ', approveArgs=' + Object.keys(approveArgsMap).length +
              ', cancelPairs=' + Object.keys(cancelPairs).length +
              ', cancelCids=' + Object.keys(cancelCids).length +
              ', cancelArgs=' + Object.keys(cancelArgsMap).length);

  // Step 2: fetch only records listed in the spreadsheet from the API
  Logger.log('Fetching results for listed records from API');

  var lookupKeys = [];
  srcData.forEach(function(row) {
    var aCid = row[0];
    var aArgs = row[1];
    var cCid = row[3];
    var cArgs = row[4];

    if (aCid && aArgs) {
      lookupKeys.push({ cid: aCid, args: aArgs });
    } else {
      if (aCid) lookupKeys.push({ cid: aCid });
      if (aArgs) lookupKeys.push({ args: aArgs });
    }

    if (cCid && cArgs) {
      lookupKeys.push({ cid: cCid, args: cArgs });
    } else {
      if (cCid) lookupKeys.push({ cid: cCid });
      if (cArgs) lookupKeys.push({ args: cArgs });
    }
  });
  if (lookupKeys.length === 0) {
    Logger.log('No lookup keys found - check source sheet data');
  }
  lookupKeys.forEach(function(k) {
    Logger.log('Lookup key - cid: ' + k.cid + ', args: ' + k.args);
  });

  var records = fetchResultsByKeys(lookupKeys);
  Logger.log('API search complete, records found: ' + (records ? records.length : 0));
  if (!records || records.length === 0) {
    Logger.log('No data returned from API');
    return;
  }
  Logger.log(records.length + ' record(s) fetched');

  var matched = [];
  Logger.log('Matching API records against lookup maps');
  for (var i = 0; i < records.length; i++) {
    var rec = records[i];
    var pairKey = rec.cid + '\u0000' + rec.args;

    var approveHit = approvePairs[pairKey] || approveCids[rec.cid] || approveArgsMap[rec.args];
    var cancelHit = cancelPairs[pairKey] || cancelCids[rec.cid] || cancelArgsMap[rec.args];

    if (approveHit) {
      rec.state = '承認';
      matched.push(rec);
    } else if (cancelHit) {
      rec.state = 'キャンセル';
      matched.push(rec);
    }
  }

  if (matched.length === 0) {
    Logger.log('No matching records');
    return;
  }
  Logger.log(matched.length + ' matching records found');

  // Step 3: create DL sheet with matched records
  var dlSheet = originalSs.getSheetByName('DL');
  if (dlSheet) originalSs.deleteSheet(dlSheet);
  dlSheet = originalSs.insertSheet('DL');

  var keys = Object.keys(matched[0]);
  dlSheet.getRange(1, 1, 1, keys.length).setValues([keys]);
  var rows = matched.map(function(rec) {
    return keys.map(function(k) { return rec[k]; });
  });
  dlSheet.getRange(2, 1, rows.length, keys.length).setValues(rows);
  Logger.log('DL sheet updated - ' + rows.length + ' row(s) written');

  // Step 4: download DL sheet as Shift_JIS
  Logger.log('Exporting DL CSV');
  downloadCsvDlShiftJis();
  Logger.log('DL CSV exported');

  // Step 5: delete DL sheet
  Logger.log('Deleting DL sheet');
  originalSs.deleteSheet(dlSheet);
  Logger.log('DL sheet deleted');
}



// Export the DL sheet as a Shift_JIS encoded CSV.
// This mirrors the standalone implementation in downloadCsvDl.gs so that
// processSeikaChanges.gs can operate independently.
function downloadCsvDlShiftJis() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('DL');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('DL sheet not found.');
    return;
  }
  Logger.log('Exporting DL sheet');

  var data = sheet.getDataRange().getValues();
  var csvContent = '';
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    for (var j = 0; j < row.length; j++) {
      csvContent += row[j];
      if (j < row.length - 1) {
        csvContent += ',';
      }
    }
    csvContent += '\n';
  }

  Logger.log('CSV rows prepared: ' + data.length);

  var blob = Utilities.newBlob(csvContent, 'text/csv', 'DL.csv')
    .setContentTypeFromExtension();
  var sjisBlob = convertToShiftJis(blob);
  Logger.log('Converted CSV to Shift_JIS');
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(
      '<a href="' + sjisBlob.getBlob().getDataUrl() + '" target="_blank">Download</a>'
    ),
    'Download DL CSV (Shift_JIS)'
  );
}

// Convert a UTF-8 blob to Shift_JIS encoding.
function convertToShiftJis(blob) {
  var uint8Array = new Uint8Array(blob.getBytes());
  var sjisArray = Encoding.convert(uint8Array, { to: 'SJIS', from: 'UTF8' });
  var sjisBlob = Utilities.newBlob(sjisArray, 'text/csv', blob.getName());
  return sjisBlob;
}

// Fetch all results that occurred last month via the API.
function fetchLastMonthResults() {
  var baseUrl = 'https://otonari-asp.com/api/v1/m';
  var accessKey = 'agqnoournapf';
  var secretKey = '1kvu9dyv1alckgocc848socw';

  baseUrl = baseUrl.replace(/\/+$/, '');

  Logger.log('Fetching results from API');
  Logger.log('Base URL: ' + baseUrl);
  var now = new Date();
  var start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  var end = new Date(now.getFullYear(), now.getMonth(), 0);

  var params = [
    'apply_unix=between_date',
    'apply_unix_A_Y=' + start.getFullYear(),
    'apply_unix_A_M=' + (start.getMonth() + 1),
    'apply_unix_A_D=' + start.getDate(),
    'apply_unix_B_Y=' + end.getFullYear(),
    'apply_unix_B_M=' + (end.getMonth() + 1),
    'apply_unix_B_D=' + end.getDate(),
    'limit=500',
    'offset=0'
  ];

  var headers = { 'X-Auth-Token': accessKey + ':' + secretKey };

  var records = [];
  var offset = 0;
  while (true) {
    params[params.length - 1] = 'offset=' + offset;
    var url = baseUrl + '/action_log_raw/search?' + params.join('&');
    Logger.log('Requesting: ' + url);
    var response;
    try {
      response = UrlFetchApp.fetch(url, { method: 'get', headers: headers });
    } catch (e) {
      SpreadsheetApp.getUi().alert('Failed to fetch ' + url + ': ' + e);
      Logger.log('Fetch failed: ' + e);
      break;
    }
    var json = JSON.parse(response.getContentText());
    if (json.records && json.records.length) {
      records = records.concat(json.records);
      Logger.log('Received ' + json.records.length + ' records (total ' + records.length + ')');
    }
    var count = json.header && json.header.count ? json.header.count : 0;
    if (records.length >= count) {
      Logger.log('Reached end of records');
      break;
    }
    offset += json.records.length;
  }
  Logger.log('Fetched ' + records.length + ' total record(s)');

  return records;
}

// Fetch records for cid/args combinations listed in the spreadsheet.
// Each key may contain either or both properties: { cid: 'xxx', args: 'yyy' }
function fetchResultsByKeys(keys) {
  var baseUrl = 'https://otonari-asp.com/api/v1/m';
  var accessKey = 'agqnoournapf';
  var secretKey = '1kvu9dyv1alckgocc848socw';

  baseUrl = baseUrl.replace(/\/+$/, '');

  var headers = { 'X-Auth-Token': accessKey + ':' + secretKey };

  var records = [];
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    Logger.log('Searching API for cid=' + k.cid + ', args=' + k.args);
    var params = [];
    if (k.cid) params.push('cid=' + encodeURIComponent(k.cid));
    if (k.args) params.push('args=' + encodeURIComponent(k.args));
    if (params.length === 0) {
      Logger.log('Skipping empty key');
      continue;
    }
    params.push('limit=1');
    var url = baseUrl + '/action_log_raw/search?' + params.join('&');
    Logger.log('Requesting: ' + url);
    try {
      var response = UrlFetchApp.fetch(url, { method: 'get', headers: headers });
      var json = JSON.parse(response.getContentText());
        if (json.records && json.records.length) {
          Logger.log('Found ' + json.records.length + ' record(s) for cid=' + k.cid + ', args=' + k.args);
          records = records.concat(json.records);
        } else {
          Logger.log('No record found for cid=' + k.cid + ', args=' + k.args);
        }
    } catch (e) {
      SpreadsheetApp.getUi().alert('Failed to fetch ' + url + ': ' + e);
      Logger.log('Fetch failed: ' + e);
    }
  }
  Logger.log('Fetched ' + records.length + ' total record(s)');
  return records;
}

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
  var srcData = sourceSheet.getRange(3, 1, sourceSheet.getLastRow() - 2, 4).getValues();
  var approve = {};
  var cancel = {};
  srcData.forEach(function(row) {
    if (row[0] && row[1]) approve[row[0] + '\u0000' + row[1]] = true;
    if (row[2] && row[3]) cancel[row[2] + '\u0000' + row[3]] = true;
  });
  Logger.log('Lookup maps: approve=' + Object.keys(approve).length + ', cancel=' + Object.keys(cancel).length);

  // Step 2: fetch last month's results from the API
  Logger.log('Fetching last month results from API');
  var records = fetchLastMonthResults();
  if (!records || records.length === 0) {
    Logger.log('No data returned from API');
    return;
  }
  Logger.log(records.length + ' record(s) fetched');

  var matched = [];
  for (var i = 0; i < records.length; i++) {
    var rec = records[i];
    var key = rec.cid + '\u0000' + rec.args;
    if (approve[key]) {
      rec.state = '承認';
      matched.push(rec);
    } else if (cancel[key]) {
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
  Logger.log(rows.length + ' row(s) written to DL');

  // Step 4: download DL sheet as Shift_JIS
  downloadCsvDlShiftJis();
  Logger.log('DL CSV exported');

  // Step 5: delete DL sheet
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

  var blob = Utilities.newBlob(csvContent, 'text/csv', 'DL.csv')
    .setContentTypeFromExtension();
  var sjisBlob = convertToShiftJis(blob);
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
  var props = PropertiesService.getScriptProperties();
  var baseUrl = props.getProperty('https://otonari-asp.com/api/v1/m');
  var accessKey = props.getProperty('agqnoournapf');
  var secretKey = props.getProperty('1kvu9dyv1alckgocc848socw');
  if (!baseUrl || !accessKey || !secretKey) {
    SpreadsheetApp.getUi().alert('API credentials are not set.');
    return [];
  }
  if (baseUrl.indexOf('your-api-host') !== -1) {
    SpreadsheetApp.getUi().alert(
      'API_BASE_URL is still set to the placeholder "your-api-host". ' +
      'Please update the script properties with the actual API endpoint.'
    );
    return [];
  }
  baseUrl = baseUrl.replace(/\/+$/, '');

  Logger.log('Fetching results from API');
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
    var response;
    try {
      response = UrlFetchApp.fetch(url, { method: 'get', headers: headers });
    } catch (e) {
      SpreadsheetApp.getUi().alert('Failed to fetch ' + url + ': ' + e);
      break;
    }
    var json = JSON.parse(response.getContentText());
    if (json.records && json.records.length) {
      records = records.concat(json.records);
    }
    var count = json.header && json.header.count ? json.header.count : 0;
    if (records.length >= count) {
      break;
    }
    offset += json.records.length;
  }
  Logger.log('Fetched ' + records.length + ' total record(s)');

  return records;
}

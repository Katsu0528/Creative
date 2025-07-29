function downloadLastMonthApprovedCsvShiftJis() {
  var props = PropertiesService.getScriptProperties();
  var baseUrl = props.getProperty('API_BASE_URL');
  var accessKey = props.getProperty('API_ACCESS_KEY');
  var secretKey = props.getProperty('API_SECRET_KEY');
  if (!baseUrl || !accessKey || !secretKey) {
    SpreadsheetApp.getUi().alert('API credentials are not set.');
    return;
  }

  // calculate previous month range
  var now = new Date();
  var start = new Date(now.getFullYear(), now.getMonth() - 1, 1); // first day of previous month
  var end = new Date(now.getFullYear(), now.getMonth(), 0); // last day of previous month

  var params = [
    'apply_unix=between_date',
    'apply_unix_A_Y=' + start.getFullYear(),
    'apply_unix_A_M=' + (start.getMonth() + 1),
    'apply_unix_A_D=' + start.getDate(),
    'apply_unix_B_Y=' + end.getFullYear(),
    'apply_unix_B_M=' + (end.getMonth() + 1),
    'apply_unix_B_D=' + end.getDate(),
    'state=1',
    'limit=500',
    'offset=0'
  ];

  var headers = {
    'X-Auth-Token': accessKey + ':' + secretKey
  };

  var records = [];
  var offset = 0;
  while (true) {
    params[params.length - 1] = 'offset=' + offset;
    var url = baseUrl + '/action_log_raw/search?' + params.join('&');
    var response = UrlFetchApp.fetch(url, { 'method': 'get', 'headers': headers });
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

  if (records.length === 0) {
    SpreadsheetApp.getUi().alert('No data found for last month.');
    return;
  }

  var keys = Object.keys(records[0]);
  var csvRows = [keys.join(',')];
  records.forEach(function(rec) {
    var row = [];
    keys.forEach(function(k) {
      var val = rec[k];
      if (val === null || val === undefined) val = '';
      row.push('"' + String(val).replace(/"/g, '""') + '"');
    });
    csvRows.push(row.join(','));
  });
  var csvContent = csvRows.join('\n');

  var blob = Utilities.newBlob(csvContent, 'text/csv', 'approved_results.csv').setContentTypeFromExtension();
  var sjisBlob = convertToShiftJis(blob);
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput('<a href="' + sjisBlob.getBlob().getDataUrl() + '" target="_blank">Download</a>'),
    'Download Approved Results CSV (Shift_JIS)');
}

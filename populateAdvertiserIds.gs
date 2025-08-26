var TARGET_SPREADSHEET_ID = '1qkae2jGCUlykwL-uTf0_eaBGzon20RCC-wBVijyvm8s';

function updateAdvertiserIds() {
  var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  var sheet = ss.getSheetByName('クライアント情報');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('クライアント情報シートが見つかりません');
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var names = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  var results = [];
  for (var i = 0; i < names.length; i++) {
    var name = names[i][0];
    if (!name) {
      results.push(['']);
      continue;
    }
    var id = fetchAdvertiserId(name);
    results.push([id || '']);
  }
  sheet.getRange(2, 15, results.length, 1).setValues(results);
}

function fetchAdvertiserId(name) {
  // TODO: Replace with actual API endpoint and parameters.
  var url = 'https://api.example.com/advertisers?name=' + encodeURIComponent(name);
  try {
    var response = UrlFetchApp.fetch(url);
    var data = JSON.parse(response.getContentText());
    return data && data.id ? data.id : '';
  } catch (e) {
    Logger.log('fetchAdvertiserId: ' + e);
    return '';
  }
}

var CONTACT_SPREADSHEET_ID = '1qkae2jGCUlykwL-uTf0_eaBGzon20RCC-wBVijyvm8s';
var CONTACT_SHEET_NAME = 'お問い合わせ';
var CONTACT_HEADERS = ['タイムスタンプ', 'メールアドレス', '内容'];

/**
 * 問い合わせ内容をスプレッドシートに保存します。
 * @param {string} message - ユーザーからのご意見・ご要望
 * @return {string}
 */
function submitContactFeedback(message) {
  var sanitizedMessage = (message || '').trim();
  if (!sanitizedMessage) {
    throw new Error('ご意見・ご要望が未入力です。');
  }

  var activeUser = Session.getActiveUser();
  var email = (activeUser && activeUser.getEmail()) || '';
  var spreadsheet = SpreadsheetApp.openById(CONTACT_SPREADSHEET_ID);
  var sheet = spreadsheet.getSheetByName(CONTACT_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONTACT_SHEET_NAME);
  }

  ensureContactHeader(sheet);
  sheet.appendRow([new Date(), email || 'メールアドレス未取得', sanitizedMessage]);

  return '送信が完了しました。貴重なご意見をありがとうございます。';
}

function ensureContactHeader(sheet) {
  if (!sheet) {
    return;
  }

  var lastRow = sheet.getLastRow();
  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, CONTACT_HEADERS.length).setValues([CONTACT_HEADERS]);
    return;
  }

  if (lastRow === 1) {
    var firstRow = sheet.getRange(1, 1, 1, CONTACT_HEADERS.length).getValues()[0];
    var hasHeader = firstRow.some(function(value) {
      return String(value).trim() !== '';
    });
    if (!hasHeader) {
      sheet.getRange(1, 1, 1, CONTACT_HEADERS.length).setValues([CONTACT_HEADERS]);
    }
  }
}

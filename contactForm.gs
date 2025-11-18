var CONTACT_SPREADSHEET_ID = '1L5yR0paTmOX3Hk8wN2cuKGm_T3028TtgQXo8TU8RAJE';
var CONTACT_SHEET_NAME = 'お問い合わせ';

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

  var nextRow = sheet.getLastRow() + 1;
  var rowValues = [new Date(), email, sanitizedMessage];
  sheet.getRange(nextRow, 1, 1, rowValues.length).setValues([rowValues]);

  return '送信が完了しました。貴重なご意見をありがとうございます。';
}

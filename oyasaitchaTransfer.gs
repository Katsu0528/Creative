const OYASAITCHA_SOURCE_URL = "https://docs.google.com/spreadsheets/d/1V1Nu25By-BKY5YxN7BYodJJikIYoNfQl90GrflfrcCo/edit";
const OYASAITCHA_TARGET_URL = "https://docs.google.com/spreadsheets/d/10FyUPwfw6BYaZRaITIM6cwhfU76AcDzNlpA7XvatD1Y/edit";
const OYASAITCHA_SOURCE_SHEET_ID = 0;
const OYASAITCHA_TARGET_SHEET_ID = 124820548;

function transferOyasaitchaSubmissions() {
  Logger.log("[transferOyasaitchaSubmissions] start");
  const sourceSs = SpreadsheetApp.openByUrl(OYASAITCHA_SOURCE_URL);
  const targetSs = SpreadsheetApp.openByUrl(OYASAITCHA_TARGET_URL);
  const sourceSheet = getSheetById_(sourceSs, OYASAITCHA_SOURCE_SHEET_ID) || sourceSs.getSheets()[0];
  const targetSheet = getSheetById_(targetSs, OYASAITCHA_TARGET_SHEET_ID);

  if (!targetSheet) {
    Logger.log("[transferOyasaitchaSubmissions] target sheet not found: %s", OYASAITCHA_TARGET_SHEET_ID);
    SpreadsheetApp.getUi().alert("転記先シートが見つかりませんでした。");
    return;
  }

  const lastRow = sourceSheet.getLastRow();
  Logger.log("[transferOyasaitchaSubmissions] source sheet: %s (%s), lastRow=%s", sourceSheet.getName(), sourceSheet.getSheetId(), lastRow);
  if (lastRow === 0) {
    Logger.log("[transferOyasaitchaSubmissions] no rows in source sheet");
    return;
  }

  const startRow = 7;
  if (lastRow < startRow) {
    Logger.log("[transferOyasaitchaSubmissions] no rows to process after row %s", startRow);
    return;
  }

  const data = sourceSheet.getRange(startRow, 1, lastRow - startRow + 1, 5).getValues();
  const today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
  Logger.log("[transferOyasaitchaSubmissions] fetched rows=%s (startRow=%s), today=%s", data.length, startRow, today);

  for (let i = 0; i < data.length; i++) {
    const [adv, colB, colC, , folderUrl] = data[i];

    if (!adv) {
      Logger.log("[transferOyasaitchaSubmissions] skip row %s (adv empty)", startRow + i);
      continue;
    }
    if (String(adv).trim() !== "おやさいっちゃ") {
      Logger.log("[transferOyasaitchaSubmissions] skip row %s (adv=%s)", startRow + i, adv);
      continue;
    }

    const targetRow = targetSheet.getLastRow() + 1;
    targetSheet.getRange(targetRow, 2).setValue(colB);
    targetSheet.getRange(targetRow, 3).setValue(colC);
    targetSheet.getRange(targetRow, 4).setValue(today);
    targetSheet.getRange(targetRow, 6).setValue(folderUrl);
    Logger.log("[transferOyasaitchaSubmissions] wrote base data to targetRow=%s", targetRow);

    const folderId = extractDriveId_(folderUrl);
    if (!folderId) {
      Logger.log("[transferOyasaitchaSubmissions] folder id not found for row %s url=%s", i + 1, folderUrl);
      continue;
    }

    const files = listFolderFiles_(folderId);
    Logger.log("[transferOyasaitchaSubmissions] folder files=%s for row %s", files.length, i + 1);
    let targetCol = 7;

    files.forEach(file => {
      const mimeType = file.getMimeType();
      Logger.log("[transferOyasaitchaSubmissions] file name=%s mimeType=%s", file.getName(), mimeType);
      if (mimeType.startsWith("image/")) {
        targetSheet.getRange(targetRow, targetCol).setValue(createCellImage_(file, file.getName()));
        targetCol += 2;
        return;
      }
      if (mimeType.startsWith("video/")) {
        const fileUrl = `https://drive.google.com/file/d/${file.getId()}/view`;
        targetSheet.getRange(targetRow, targetCol).setValue(fileUrl);
        targetCol += 2;
      }
    });
  }
}

function getSheetById_(spreadsheet, sheetId) {
  const sheets = spreadsheet.getSheets();
  for (const sheet of sheets) {
    if (sheet.getSheetId() === sheetId) return sheet;
  }
  return null;
}

function extractDriveId_(url) {
  if (!url) return "";
  const match = String(url).match(/[-\w]{25,}/);
  return match ? match[0] : "";
}

function listFolderFiles_(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const files = [];
  const iter = folder.getFiles();
  while (iter.hasNext()) {
    files.push(iter.next());
  }
  return files.sort((a, b) => a.getName().localeCompare(b.getName(), "ja"));
}

function createCellImage_(file, name) {
  const blob = file.getBlob();
  return SpreadsheetApp.newCellImage()
    .setSourceUrl("data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes()))
    .setAltTextTitle(name)
    .build();
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (sheetName !== "業務履歴" || col !== 2 || row < 2) {
    return;
  }

  const targetCell = sheet.getRange(row, 3);
  targetCell.clearDataValidations();

  const keyword = e.range.getValue();
  if (!keyword) {
    targetCell.clearContent();
    return;
  }

  const masterSheet = e.source.getSheetByName("マスタ");
  if (!masterSheet) {
    return;
  }

  const lastRow = masterSheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  const masterValues = masterSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const matched = Array.from(new Set(
    masterValues
      .filter(([code, name]) => name === keyword && code !== "")
      .map(([code]) => code)
  ));

  if (matched.length === 0) {
    targetCell.clearContent();
    return;
  }

  const dvRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(matched, true)
    .setAllowInvalid(false)
    .build();
  targetCell.setDataValidation(dvRule);
}

function setStatusToCompletedForUrlRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // 一括取得 → 配列編集 → 一括書込（高速化のベストプラクティス）
  const numRows = lastRow - 1;
  const urlVals = sheet.getRange(2, COL_URL, numRows, 1).getValues();
  const wRange  = sheet.getRange(2, COL_W,   numRows, 1);
  const adRange = sheet.getRange(2, COL_AD,  numRows, 1);
  const ajRange = sheet.getRange(2, COL_AJ,  numRows, 1);

  const wVals  = wRange.getValues();
  const adVals = adRange.getValues();
  const ajVals = ajRange.getValues();

  const wVali  = wRange.getDataValidations();
  const adVali = adRange.getDataValidations();
  const ajVali = ajRange.getDataValidations();

  let updated = false;
  for (let i = 0; i < numRows; i++) {
    const url = urlVals[i][0];
    if (typeof url === 'string' && url.trim() !== '') {
      if (wVali[i][0])  wVals[i][0]  = STATUS_VALUE;
      if (adVali[i][0]) adVals[i][0] = STATUS_VALUE;
      if (ajVali[i][0]) ajVals[i][0] = STATUS_VALUE;
      updated = true;
    }
  }
  if (updated) {
    wRange.setValues(wVals);
    adRange.setValues(adVals);
    ajRange.setValues(ajVals);
  }
}

function onEdit(e) {
  const range = e.range;
  if (!range) return;

  const sheet = range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;
  if (range.getColumn() !== COL_URL)    return;
  if (range.getRow() < 2)               return;

  const value = range.getValue();
  if (typeof value !== 'string' || value.trim() === '') return;

  updateRowStatuses_(sheet, range.getRow());
}

function updateRowStatuses_(sheet, row) {
  const wCell  = sheet.getRange(row, COL_W);
  const adCell = sheet.getRange(row, COL_AD);
  const ajCell = sheet.getRange(row, COL_AJ);

  if (wCell.getDataValidation())  wCell.setValue(STATUS_VALUE);
  if (adCell.getDataValidation()) adCell.setValue(STATUS_VALUE);
  if (ajCell.getDataValidation()) ajCell.setValue(STATUS_VALUE);
}

function createInstallableOnEdit_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers
    .filter(t => t.getHandlerFunction() === 'onEdit' &&
                 t.getEventType()      === ScriptApp.EventType.ON_EDIT)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('onEdit')
           .forSpreadsheet(SpreadsheetApp.getActive())
           .onEdit()
           .create();
}


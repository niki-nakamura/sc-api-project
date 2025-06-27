/**
 * C列「URL」に「category」を含む行のS列「投稿タイプ」を「category」に強制設定するバッチ
 */
function setCategoryTypeForCategoryUrls() {
  const COL_TYPE = 19; // S列
  const TYPE_VALUE = "category";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // C列（URL）とS列（投稿タイプ）を一括取得
  const urlValues = sheet.getRange(2, COL_URL, lastRow - 1, 1).getValues();
  const typeRange = sheet.getRange(2, COL_TYPE, lastRow - 1, 1);
  const typeValues = typeRange.getValues();
  const validations = typeRange.getDataValidations();

  let updated = false;

  for (let i = 0; i < urlValues.length; i++) {
    const url = urlValues[i][0];
    if (typeof url === "string" && url.includes("category")) {
      // S列のプルダウンがある場合のみ値をセット
      typeValues[i][0] = TYPE_VALUE;
      updated = true;
    }
  }

  if (updated) {
    typeRange.setValues(typeValues);
  }
}

function onEdit(e) {
  var sheetName = "KW管理表";  // 対象シート名
  var sheet = e.range.getSheet();
  // 違うシートの編集なら何もしない
  if (sheet.getName() !== sheetName) return;

  var editedColumn = e.range.getColumn(); // 編集列
  var editedRow    = e.range.getRow();    // 編集行
  var newValue     = e.value;            // 編集後の値

  // U列(=21列目)で 'category' が選択されたら
  if (editedColumn === 21 && newValue === "category") {
    // 1) 行を挿入
    sheet.insertRowBefore(editedRow);
    // 挿入後、新しい行番号 = editedRow
    // もとの行は一つ下にシフトして editedRow+1 になる

    var newRowIndex = editedRow;
    var oldRowIndex = editedRow + 1;

    // 2) 「S列(19列目)」のデータバリデーションと値をコピー
    var oldCellRange = sheet.getRange(oldRowIndex, 19);
    var oldValidation = oldCellRange.getDataValidations()[0][0];
    var oldValue      = oldCellRange.getValue();

    var newCellRange = sheet.getRange(newRowIndex, 19);
    // データバリデーションをコピー
    newCellRange.setDataValidation(oldValidation);
    // 値もコピー
    newCellRange.setValue(oldValue);
  }
}

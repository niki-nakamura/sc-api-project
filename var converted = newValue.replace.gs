function removeFullWidthSpacesInColumnB() {
  // 対象のシートを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // シート上の最終行を取得
  var lastRow = sheet.getLastRow();
  
  // 2行目以降がある場合のみ処理
  if (lastRow > 1) {
    // B列の2行目から最終行までを対象
    var range = sheet.getRange(2, 2, lastRow - 1, 1);
    var values = range.getValues();
    
    for (var i = 0; i < values.length; i++) {
      var val = values[i][0];
      // 値が文字列の場合のみ全角スペースを半角スペースに置き換え
      if (typeof val === 'string') {
        var converted = val.replace(/　/g, ' ');
        if (val !== converted) {
          values[i][0] = converted;
        }
      }
    }
    // 変換後の値をまとめて反映
    range.setValues(values);
  }
}

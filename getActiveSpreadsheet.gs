function onEdit(e) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var source   = ss.getSheetByName('KW管理表');
  var target   = ss.getSheetByName('松浦 嵩');
  
  // 1) 列幅をコピー
  var lastCol  = source.getLastColumn();  
  for (var c = 1; c <= lastCol; c++) {
    var width = source.getColumnWidth(c);
    target.setColumnWidth(c, width);
  }
  
  // 2) 「松浦 嵩」シートにFILTER関数を入れる(すでに手入力してあれば不要)
  //    例: target.getRange("A2").setValue("=FILTER('KW管理表'!A2:AR, 'KW管理表'!O2:O = \"松浦\")");
  //    あるいは QUERY 関数などご自身の好みの方法で
}

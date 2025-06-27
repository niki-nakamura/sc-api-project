// update24hMetrics.gs
/**
 * B 列（query）・C 列（page URL）の組み合わせを 2〜MAX_ROW 行目まで走査し、
 * 直近 1 日の平均掲載順位を AG 列へ出力
 */
function updateDailyPositions() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error(`シート '${CONFIG.SHEET_NAME}' が見つかりません`);

  // 直近 24 時間分（日付範囲：昨日〜今日）
  const tz        = Session.getScriptTimeZone();
  const fmt       = 'yyyy-MM-dd';
  const today     = new Date();
  const yesterday = new Date(today.getTime() - 24 * 60 * 60 * 1000);
  const startDate = Utilities.formatDate(yesterday, tz, fmt);
  const endDate   = Utilities.formatDate(today,     tz, fmt);

  for (let row = 2; row <= CONFIG.MAX_ROW; row++) {
    const query = (sheet.getRange(row, 2).getValue() || '').toString().trim();
    const page  = (sheet.getRange(row, 3).getValue() || '').toString().trim();

    if (!query || !page) {                            // 未入力はクリア
      sheet.getRange(row, CONFIG.OUTPUT_COL).setValue('');
      continue;
    }

    try {
      const pos = fetchAvgPosition_(
        CONFIG.SITE_URL, page, query, startDate, endDate
      );
      sheet.getRange(row, CONFIG.OUTPUT_COL).setValue(pos);
    } catch (err) {
      sheet.getRange(row, CONFIG.OUTPUT_COL).setValue(`ERR: ${err.message}`);
    }
  }
}

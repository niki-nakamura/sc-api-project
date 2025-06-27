/**
 * KW 管理表 C 列（URL）書式リセット（バッチ版）
 */
const URL_FMT_CFG = {        // ← ★旧 CFG を URL_FMT_CFG に変更
  SHEET: 'KW管理表',
  COL  : 3,
  START: 2,
  BATCH: 1000
};

function formatUrlColumnNow() { run_(); }
function formatUrlColumnHour() { run_(); }

function addHourlyTriggerKw() {
  if (!ScriptApp.getProjectTriggers()
        .some(t => t.getHandlerFunction() === 'formatUrlColumnHour')) {
    ScriptApp.newTrigger('formatUrlColumnHour')
             .timeBased().everyHours(1).create();
  }
}

function run_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet()
              .getSheetByName(URL_FMT_CFG.SHEET);
  if (!sh) return;

  const last = sh.getLastRow();
  if (last < URL_FMT_CFG.START) return;

  for (let row = URL_FMT_CFG.START; row <= last; row += URL_FMT_CFG.BATCH) {
    const rows = Math.min(CFG.BATCH, last - row + 1);
    const rng  = sh.getRange(row, CFG.COL, rows, 1);
    const rich = rng.getRichTextValues();

    let dirty = false;
    for (let i = 0; i < rich.length; i++) {
      const cell = rich[i][0];
      if (cell.getLinkUrl()) {
        dirty = true;
        // plain text だけ残す
        rich[i][0] = SpreadsheetApp.newRichTextValue()
                      .setText(cell.getText())
                      .build();
      }
    }
    if (dirty) rng.setRichTextValues(rich);

    // 見た目のフォーマット
    rng.setFontFamily('Noto Sans JP')
       .setFontSize(10)
       .setFontWeight('normal')
       .setFontLine('none')
       .setFontColor('#000000')
       .setHorizontalAlignment('right');

    SpreadsheetApp.flush();           // バッチ毎に確定
  }
}

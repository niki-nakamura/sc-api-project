/*******************  KW管理表 – 前回クロール日時  *******************/
const CRAWL_CFG = {          // ← ★旧 CFG を CRAWL_CFG に変更
  SHEET   : 'KW管理表',
  COL_URL : 3,
  COL_OUT : 43,
  SITE_URL: 'https://digi-mado.jp/',
  TZ      : 'Asia/Tokyo',
  ENDPOINT: 'https://searchconsole.googleapis.com/v1/urlInspection/index:inspect',
  QPS     : 10,
  BATCH   : 10,
  GAP     : 1100
};

/** メイン ― 途中再開にも対応 */
function updateLastCrawlDates() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CRAWL_CFG.SHEET);
  const lastRow = sh.getLastRow();
  const urls = sh.getRange(2, CRAWL_CFG.COL_URL, lastRow - 1).getValues().flat();

  // 進捗（次に処理すべきインデックス）を取得
  const startIdx = Number(PropertiesService.getScriptProperties().getProperty('next') || 0);

  const out = [];                    // 結果をバッファ
  const t0  = Date.now();            // 開始時刻

  for (let i = startIdx; i < urls.length; i += CRAWL_CFG.BATCH) {

    // ===== タイムアウト 5.5 分で一時停止 =====
    if (Date.now() - t0 > 330000) {        // 5.5 分    // 5.5分
      saveState(i);                     // 進捗保存
      scheduleRerun();                  // 1分後に自動再開
      writeBack(sh, out);               // ここまでの結果を反映
      return;                           // 強制終了
    }

    // ===== fetchAll で並列取得 =====
    const reqs  = urls.slice(i, i + CFG.BATCH).map(makeReq);
    const resps = UrlFetchApp.fetchAll(reqs);             // :contentReference[oaicite:0]{index=0}

    resps.forEach((r, idx) => out[i + idx] = [parseLastCrawl(r)]);
    Utilities.sleep(CFG.GAP);                             // QPS 制御 :contentReference[oaicite:1]{index=1}
  }

  // ===== 全件終了 =====
  writeBack(sh, out);
  PropertiesService.getScriptProperties().deleteProperty('next'); // 進捗クリア
  deleteOldTriggers();                                            // ごみ掃除
}

/** URL Inspection API リクエストを生成 */
function makeReq(url) {
  if (!url) {                   // 空セルはダミー
    return { url: 'about:blank', muteHttpExceptions: true };
  }
  const payload = { inspectionUrl: url, siteUrl: CFG.SITE_URL };
  return {
    url            : CFG.ENDPOINT,
    method         : 'post',
    contentType    : 'application/json',
    payload        : JSON.stringify(payload),
    headers        : { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
}

/** API レスポンスから最終クロール日時を抽出 */
function parseLastCrawl(resp) {
  try {
    const code = resp.getResponseCode();
    const data = JSON.parse(resp.getContentText());
    if (code === 200 && data.inspectionResult) {
      const iso = (data.inspectionResult &&
              data.inspectionResult.indexStatusResult &&
              data.inspectionResult.indexStatusResult.lastCrawlTime) || '';
      return iso
        ? Utilities.formatDate(new Date(iso), CFG.TZ, 'yyyy/MM/dd HH:mm:ss')   // :contentReference[oaicite:2]{index=2}
        : 'NO DATA';
    }
    return `HTTP ${code}`;
  } catch (e) {
    return 'ERROR';
  }
}

/** 途中再開用のインデックスを保存 */
function saveState(nextIdx) {
  PropertiesService.getScriptProperties().setProperty('next', nextIdx);       // :contentReference[oaicite:3]{index=3}
}

/** 1分後に自分を再実行するトリガーをセット */
function scheduleRerun() {
  ScriptApp.newTrigger('updateLastCrawlDates')
           .timeBased().after(60 * 1000).create();                            // :contentReference[oaicite:4]{index=4}
}

/** 既存の同名トリガーを削除（重複防止） */
function deleteOldTriggers() {
  ScriptApp.getProjectTriggers()                                              // :contentReference[oaicite:5]{index=5}
           .filter(t => t.getHandlerFunction() === 'updateLastCrawlDates')
           .forEach(t => ScriptApp.deleteTrigger(t));
}

/** シートへの書き戻し（まとめ書きで高速化） */
function writeBack(sheet, values) {
  if (values.length) {
    sheet.getRange(2, CFG.COL_OUT, values.length, 1).setValues(values);       // :contentReference[oaicite:6]{index=6}
  }
}

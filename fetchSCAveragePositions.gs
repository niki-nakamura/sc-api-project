/******************************************************
 * KW管理表 – Search Console 平均順位バッチ 2025‑06‑26 v2.3
 *   ● C列終端まで処理（空行を許容）
 *   ● LockService で排他制御
 *   ● 残時間に応じた動的バッチサイズ
 ******************************************************/

/* ===== 0. 基本設定 =============================== */
const SPREADSHEET_ID   = '1LqMts2qQr-T6r0nH3XxnM0r20LuQJPqaDuUg9R8Li28';
const PROPERTY_URL     = 'https://digi-mado.jp/';
const SHEET_NAME       = 'KW管理表';
const BATCH_SIZE       = 300;           // 初期値（動的に縮小）
const PROBE_BACK_DAYS  = 30;
const MAX_EXEC_MS      = 330000;        // 5.5 分
const EXEC_SAFETY_MS   = 25000;         // 25 秒安全マージン

/* ===== 1. エントリ =============================== */
function batchFetchSCAveragePositions() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    console.info('⚠️ データ行が 0 行のため終了');
    return;
  }
  PropertiesService.getScriptProperties()
    .setProperty(`NEXT_START_ROW_${SHEET_NAME}`, '2');
  batchProcessWorker();
}

/* ===== 2. バッチ本体 ============================== */
function batchProcessWorker() {
  /* ---- 排他ロック ---- */
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // 30 秒待機
  } catch (e) {
    console.error('⚠️ 他プロセスが実行中のためスキップ');
    return;
  }

  cleanupEphemeralTriggers();

  const t0 = Date.now();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = findLastDataRow(sheet, 3); // ★C列終端
  const propKey = `NEXT_START_ROW_${SHEET_NAME}`;
  const scriptPs = PropertiesService.getScriptProperties();
  let startRow = Number(scriptPs.getProperty(propKey)) || 2;
  if (startRow > lastRow) {
    console.info(`🏁 全処理完了（startRow=${startRow} > lastRow=${lastRow}）`);
    scriptPs.deleteProperty(propKey);
    lock.releaseLock();
    return;
  }

  /* ---- 最新データ日 ---- */
  const { latestDate, lag } = probeLatestDateWithData(PROBE_BACK_DAYS);
  if (!latestDate) throw new Error(`GSC に過去 ${PROBE_BACK_DAYS} 日分のデータがありません。`);
  console.info(`📅 最新データ日: ${latestDate}（遅延 ${lag} 日）`);

  /* ---- 日付ウインドウ ---- */
  const newEnd   = latestDate;
  const newStart = getDateStrGap(latestDate, -6);
  const oldEnd   = latestDate;
  const oldStart = getDateStrGap(latestDate, -29);
  const start30  = oldStart;
  const end30    = oldEnd;

  /* ---- ループ ---- */
  let globalWritten = 0;
  let batchSize     = BATCH_SIZE;
  while (startRow <= lastRow) {
    /* 残り時間チェック */
    const elapsed = Date.now() - t0;
    const remaining = MAX_EXEC_MS - elapsed;
    if (remaining < EXEC_SAFETY_MS) break;
    if (remaining < 120000) batchSize = 100; // 残り <2 分なら縮小

    const endRow = Math.min(startRow + batchSize - 1, lastRow);
    const range  = sheet.getRange(startRow, 1, endRow - startRow + 1, 10);
    const data   = range.getValues();

    const avgCache = {};
    const metCache = {};
    let   written  = 0;

    data.forEach(row => {
      const kw  = row[1];
      const url = row[2];
      if (!kw || !url) {
        row[3]=row[4]=row[5]=row[7]=row[8]=row[9]='';
        return;
      }
      /* ---- 平均順位 ---- */
      const oKey = `${kw}|${url}|old`;
      const oldPos = oKey in avgCache ? avgCache[oKey] :
                     (avgCache[oKey] = fetchAveragePosition(kw,url,oldStart,oldEnd));
      const nKey = `${kw}|${url}|new`;
      const newPos = nKey in avgCache ? avgCache[nKey] :
                     (avgCache[nKey] = fetchAveragePosition(kw,url,newStart,newEnd));

      row[3] = oldPos !== null ? oldPos.toFixed(1) : '-';
      row[4] = newPos !== null ? newPos.toFixed(1) : '-';
      row[5] = (oldPos !== null && newPos !== null)
                 ? ((oldPos-newPos>0?'+':'')+(oldPos-newPos).toFixed(1)) : '-';

      /* ---- クリック & インプレ ---- */
      const mKey = `${kw}|${url}`;
      const met  = mKey in metCache ? metCache[mKey] :
                   (metCache[mKey] = fetchPageMetricsForKW(kw,url,start30,end30));
      if (met) {
        const {clicks, impressions} = met;
        row[7] = clicks;
        row[8] = impressions;
        row[9] = impressions ? (clicks/impressions*100).toFixed(1)+'%' : '0%';
      } else {
        row[7]=row[8]=row[9]='-';
      }
      if (row[3] !== '-' || row[4] !== '-' || row[7] !== '-') written++;
    });

    range.setValues(data);
    SpreadsheetApp.flush();
    console.info(`✅ 範囲 ${startRow}-${endRow} 行: ${written}/${data.length} 行反映`);

    globalWritten += written;
    startRow = endRow + 1;
    scriptPs.setProperty(propKey, String(startRow));
  }

  /* ---- 次回トリガ ---- */
  if (startRow <= lastRow) {
    console.info(`⏭ 残 ${lastRow - startRow + 1} 行。1 分後に再開`);
    ScriptApp.newTrigger('batchProcessWorker').timeBased().after(60 * 1000).create();
  } else {
    scriptPs.deleteProperty(propKey);
    console.info(`🎉 全処理完了。総書き込み行: ${globalWritten}`);
  }
  lock.releaseLock();
}

/* ===== 3. 遅延検出 =============================== */
function probeLatestDateWithData(maxDays) {
  for (let i = 0; i <= maxDays; i++) {
    const d = getDateStrNDaysAgo(i);
    const res = callSearchAnalyticsRaw(
      {startDate:d,endDate:d,searchType:'web',dimensions:['date'],rowLimit:1},
      'probe');
    if (res?.rows?.length) return { latestDate:d, lag:i };
  }
  return { latestDate:null, lag:maxDays };
}

/* ===== 4. Search Console API ==================== */
function fetchAveragePosition(kw,url,s,e){
  const res = callSearchAnalyticsRaw({
    startDate:s,endDate:e,searchType:'web',
    dimensions:['query','page'],
    dimensionFilterGroups:[{filters:[
      {dimension:'query',operator:'equals',expression:kw},
      {dimension:'page', operator:'equals',expression:url}
    ]}],
    rowLimit:1
  },'avg');
  return res?.rows?.[0]?.position ?? null;
}

function fetchPageMetricsForKW(kw,url,s,e){
  const res = callSearchAnalyticsRaw({
    startDate:s,endDate:e,searchType:'web',
    dimensions:['query','page'],
    dimensionFilterGroups:[{filters:[
      {dimension:'query',operator:'equals',expression:kw},
      {dimension:'page', operator:'equals',expression:url}
    ]}],
    rowLimit:25000
  },'met');
  if (!res?.rows?.length) return null;
  let clicks=0, impr=0;
  res.rows.forEach(r=>{ clicks+=r.clicks||0; impr+=r.impressions||0; });
  return {clicks, impressions:impr};
}

function callSearchAnalyticsRaw(payload, tag){
  const url = `https://www.googleapis.com/webmasters/v3/sites/${encodeURIComponent(PROPERTY_URL)}/searchAnalytics/query`;
  try{
    const res = UrlFetchApp.fetch(url,{
      method:'post',
      contentType:'application/json',
      payload:JSON.stringify(payload),
      muteHttpExceptions:true,
      headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()}
    });
    if(res.getResponseCode()>=400){
      console.error(`❌ HTTP${res.getResponseCode()} (${tag}) : ${res.getContentText().slice(0,150)}`);
      return null;
    }
    return JSON.parse(res.getContentText());
  }catch(e){
    console.error(`🔥 ${tag} error: ${e}`);
    return null;
  }
}

/* ===== 5. 日付ヘルパ ============================= */
function getDateStrNDaysAgo(n){
  return Utilities.formatDate(
    new Date(Date.now()-n*86400000),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd');
}
function getDateStrGap(base, diff){
  return Utilities.formatDate(
    new Date(new Date(base).getTime()+diff*86400000),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd');
}

/* ===== 6. 終端行検索 & トリガ掃除 ================ */
function findLastDataRow(sheet, col){
  const maxRow = sheet.getLastRow();
  const values = sheet.getRange(2, col, maxRow-1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0]) return i + 2;
  }
  return 1;
}
function cleanupEphemeralTriggers(){
  ScriptApp.getProjectTriggers()
    .filter(t=>t.getHandlerFunction()==='batchProcessWorker')
    .forEach(t=>ScriptApp.deleteTrigger(t));
}

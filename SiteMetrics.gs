/*************************************************
 * サイト全体データ抽出スクリプト（例：SiteMetrics.gs）
 * 対象シート: カテゴリKWデイリー順位
 *
 * 2行目に以下のデータを抽出します：
 *   E列：対象URL（例："https://digi-mado.jp/"）
 *   F列：30日間の平均掲載順位（※最終的にF2にはF列（F3以降）の平均を設定）
 *   G列：7日間の平均掲載順位（※最終的にG2にはG列（G3以降）の平均を設定）
 *   H列：両期間の比較（30日平均 - 7日平均）
 *   I列：30日間クリック数
 *   J列：30日間合計表示回数
 *   K列：30日間平均CTR（クリック数/表示回数×100%）
 *************************************************/
function fetchSiteData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "カテゴリKWデイリー順位";
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("シートが見つかりません: " + sheetName);
    return;
  }
  
  // 2行目を対象とする（サイト全体データ用）
  const row = 2;
  
  // E列のURLプレフィックスを取得（例："https://digi-mado.jp/"）
  const siteUrl = sheet.getRange(row, 5).getValue();
  if (!siteUrl) {
    Logger.log("E2 に URL が設定されていません。");
    return;
  }
  
  // 30日間および7日間の集計期間を設定
  const start30 = getDateStrNDaysAgo(30);
  const end30   = getDateStrNDaysAgo(0);
  const start7  = getDateStrNDaysAgo(7);
  const end7    = getDateStrNDaysAgo(0);
  
  // ※以下、F2,G2にはSearch Consoleからの平均掲載順位を直接書かず、
  //    後でシート内F列、G列の平均値（F2以外、G2以外）を算出するか、
  //    もしくはF2,G2は空欄のままとします。
  // ここではとりあえず一旦空欄にしておきます。
  sheet.getRange(row, 6).clearContent(); // F2
  sheet.getRange(row, 7).clearContent(); // G2
  
  // H列：両期間の比較（※ここではSearch Consoleから取得したデータを使って比較）
  // ※なお、平均掲載順位はF2,G2に設定しないので、ここは参考として残すか不要なら空欄にできます。
  const pos30 = fetchAveragePositionForSite(siteUrl, start30, end30);
  const pos7  = fetchAveragePositionForSite(siteUrl, start7, end7);
  const hCell = sheet.getRange(row, 8);
  let diffText = "-";
  if (pos30 !== null && pos7 !== null) {
    const diff = pos30 - pos7;
    if (diff > 0) {
      diffText = "+" + diff.toFixed(1);
    } else if (diff < 0) {
      diffText = diff.toFixed(1);
    } else {
      diffText = "±0";
    }
  }
  hCell.setValue(diffText);
  
  // I～K列は従来通りサイト全体（URLに siteUrl を含む）の30日間のクリック数と表示回数を取得
  const metrics = fetchSiteMetrics(siteUrl, start30, end30);
  const iCell = sheet.getRange(row, 9);  // I列：30日間クリック数
  const jCell = sheet.getRange(row, 10); // J列：30日間合計表示回数
  const kCell = sheet.getRange(row, 11); // K列：30日間平均CTR
  if (metrics === null) {
    iCell.setValue("-");
    jCell.setValue("-");
    kCell.setValue("-");
  } else {
    const clicks = metrics.clicks;
    const impressions = metrics.impressions;
    const ctr = impressions > 0 ? (clicks / impressions) : 0;
    iCell.setValue(clicks);
    jCell.setValue(impressions);
    kCell.setValue((ctr * 100).toFixed(2) + "%");
  }
  
  // --- 以下は任意の処理 ---
  // F2およびG2に、シート内のF列、G列（2行目以外）の数値の平均を算出して設定する例です。
  // もしスプレッドシート側でAVERAGE関数などを使って計算する場合は、この部分は不要です。
  computeAveragesForFandG(sheet);
}


/*************************************************
 * [サブ関数]
 * n日前の日付 (YYYY-MM-DD) を返す
 *************************************************/
function getDateStrNDaysAgo(n) {
  const d = new Date();
  d.setDate(d.getDate() - n);
  const yyyy = d.getFullYear();
  const mm = ("0" + (d.getMonth() + 1)).slice(-2);
  const dd = ("0" + d.getDate()).slice(-2);
  return `${yyyy}-${mm}-${dd}`;
}


/*************************************************
 * [サブ関数]
 * 平均掲載順位を取得（サイト全体対象）
 * Search Console API のクエリでは、dimensions を ["page"] とし、
 * フィルタ条件に "page" が siteUrl を含む（operator: "contains"）条件を設定しています。
 *************************************************/
function fetchAveragePositionForSite(siteUrl, startDate, endDate) {
  const payload = {
    startDate: startDate,
    endDate: endDate,
    dimensions: ["page"],
    dimensionFilterGroups: [{
      filters: [
        { dimension: "page", operator: "contains", expression: siteUrl }
      ]
    }],
    rowLimit: 100
  };
  
  const propertyUrlEncoded = encodeURIComponent("https://digi-mado.jp/");
  const apiUrl = `https://www.googleapis.com/webmasters/v3/sites/${propertyUrlEncoded}/searchAnalytics/query`;
  
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  };
  
  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const json = JSON.parse(response.getContentText());
    if (!json.rows || json.rows.length === 0) {
      return null;
    }
    // Search Console が返す weighted average position を利用
    return json.rows[0].position;
  } catch (e) {
    return null;
  }
}


/*************************************************
 * [サブ関数]
 * 30日間のクリック数と表示回数を取得（サイト全体対象）
 * Search Console API のクエリでは、dimensions を ["page"] とし、
 * フィルタ条件に "page" が siteUrl を含む（operator: "contains"）条件を設定しています。
 *************************************************/
function fetchSiteMetrics(siteUrl, startDate, endDate) {
  const payload = {
    startDate: startDate,
    endDate: endDate,
    dimensions: ["page"],
    dimensionFilterGroups: [{
      filters: [
        { dimension: "page", operator: "contains", expression: siteUrl }
      ]
    }],
    rowLimit: 100
  };
  
  const propertyUrlEncoded = encodeURIComponent("https://digi-mado.jp/");
  const apiUrl = `https://www.googleapis.com/webmasters/v3/sites/${propertyUrlEncoded}/searchAnalytics/query`;
  
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  };
  
  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const json = JSON.parse(response.getContentText());
    if (!json.rows || json.rows.length === 0) {
      return null;
    }
    let sumClicks = 0;
    let sumImpressions = 0;
    json.rows.forEach(row => {
      sumClicks += row.clicks;
      sumImpressions += row.impressions;
    });
    return { clicks: sumClicks, impressions: sumImpressions };
  } catch (e) {
    return null;
  }
}


/*************************************************
 * [サブ関数]
 * シート内のF列とG列（2行目以外）の数値の平均を算出して、
 * F2とG2にそれぞれ設定する
 * ※単純な平均計算です。必要なければ、この関数呼び出しを削除してください。
 *************************************************/
function computeAveragesForFandG(sheet) {
  const lastRow = sheet.getLastRow();
  // F列の平均（F2以外＝3行目～最終行）
  if (lastRow >= 3) {
    const fValues = sheet.getRange(3, 6, lastRow - 2, 1).getValues();
    const fNumbers = fValues.flat().filter(v => typeof v === "number" && !isNaN(v));
    if (fNumbers.length > 0) {
      const fAvg = fNumbers.reduce((sum, val) => sum + val, 0) / fNumbers.length;
      sheet.getRange(2, 6).setValue(fAvg.toFixed(1));
    }
  }
  // G列の平均（G2以外＝3行目～最終行）
  if (lastRow >= 3) {
    const gValues = sheet.getRange(3, 7, lastRow - 2, 1).getValues();
    const gNumbers = gValues.flat().filter(v => typeof v === "number" && !isNaN(v));
    if (gNumbers.length > 0) {
      const gAvg = gNumbers.reduce((sum, val) => sum + val, 0) / gNumbers.length;
      sheet.getRange(2, 7).setValue(gAvg.toFixed(1));
    }
  }
}

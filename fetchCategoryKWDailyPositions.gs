/*************************************************
 * カテゴリKWデイリー順位シート専用の処理
 *************************************************/
function fetchCategoryKWDailyPositions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "カテゴリKWデイリー順位"; // 対象シート
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log("シートが見つかりません: " + sheetName);
    return;
  }

  // シートの最終行を取得
  const lastRow = sheet.getLastRow();
  
  // 30日分・7日分の集計期間
  const start30 = getDateStrNDaysAgo(30);  // 30日前
  const end30   = getDateStrNDaysAgo(0);   // 今日
  const start7  = getDateStrNDaysAgo(7);   // 7日前
  const end7    = getDateStrNDaysAgo(0);   // 今日

  // A2 から最終行までをループ（見出し行が1行目の場合）
  for (let row = 2; row <= lastRow; row++) {
    const query   = sheet.getRange(row, 3).getValue(); // C列: キーワード
    const pageUrl = sheet.getRange(row, 5).getValue(); // E列: URL

    // キーワードまたはURLが未入力ならスキップ
    if (!query || !pageUrl) {
      continue;
    }

    // 30日間平均順位 (F列用)
    const pos30 = fetchAveragePosition(query, pageUrl, start30, end30);
    
    // 7日間平均順位 (G列用)
    const pos7 = fetchAveragePosition(query, pageUrl, start7, end7);

    // F列: 30日平均順位
    const fCell = sheet.getRange(row, 6); 
    if (pos30 === null) {
      fCell.setValue("-");
    } else {
      fCell.setValue(pos30.toFixed(1));
    }

    // G列: 7日平均順位
    const gCell = sheet.getRange(row, 7); 
    if (pos7 === null) {
      gCell.setValue("-");
    } else {
      gCell.setValue(pos7.toFixed(1));
    }

    // H列: 30日平均(F列) - 7日平均(G列) の差分
    const hCell = sheet.getRange(row, 8);
    let diffText = "-";
    if (pos30 !== null && pos7 !== null) {
      const diff = pos30 - pos7;  // 30日 - 7日
      if (diff > 0) {
        diffText = "+" + diff.toFixed(1);
      } else if (diff < 0) {
        diffText = diff.toFixed(1);
      } else {
        diffText = "±0";
      }
    }
    hCell.setValue(diffText);

    // I列, J列, K列:
    // I列: 30日間クリック合計
    // J列: 30日間合計表示回数
    // K列: 30日間平均CTR（クリック数/表示回数×100%）
    // ※抽出時は「C列のKW」と「E列のURL」の2条件フィルターを適用
    const metrics = fetchPageMetricsForKW(query, pageUrl, start30, end30);
    const iCell = sheet.getRange(row, 9);
    const jCell = sheet.getRange(row, 10);
    const kCell = sheet.getRange(row, 11);
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
      kCell.setValue((ctr * 100).toFixed(1) + "%");
    }
  }

  // シート全体の I2 にサイト全体(https://digi-mado.jp/)の 30日合計クリック数を表示
  const totalSiteSessions30 = fetchSitewideClicksSum(start30, end30);
  const siteSessionCell = sheet.getRange("I2");
  if (totalSiteSessions30 === null) {
    siteSessionCell.setValue("-");
  } else {
    siteSessionCell.setValue(totalSiteSessions30);
  }
}


/*************************************************
 * [サブ関数] n日前の日付(YYYY-MM-DD)を返す
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
 * Search Console API で (query × pageUrl) の期間平均順位を返す
 *************************************************/
function fetchAveragePosition(query, pageUrl, startDate, endDate) {
  const payload = {
    startDate: startDate,
    endDate: endDate,
    dimensions: ["query", "page"],
    dimensionFilterGroups: [{
      filters: [
        { dimension: "query", operator: "equals", expression: query },
        { dimension: "page", operator: "equals", expression: pageUrl }
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
    return json.rows[0].position;
  } catch (e) {
    return null;
  }
}


/*************************************************
 * [サブ関数]
 * キーワード (query) と URL (pageUrl) の2条件で
 * クリック数と表示回数の合計(期間内)を返す
 *************************************************/
function fetchPageMetricsForKW(query, pageUrl, startDate, endDate) {
  const payload = {
    startDate: startDate,
    endDate: endDate,
    dimensions: ["query", "page"],
    dimensionFilterGroups: [{
      filters: [
        { dimension: "query", operator: "equals", expression: query },
        { dimension: "page", operator: "equals", expression: pageUrl }
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
 * ドメイン全体でのクリック数合計(期間内)を返す
 *************************************************/
function fetchSitewideClicksSum(startDate, endDate) {
  const payload = {
    startDate: startDate,
    endDate: endDate,
    dimensions: ["page"], // フィルターなしで全ページ合算
    rowLimit: 25000
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
    json.rows.forEach(row => {
      sumClicks += row.clicks;
    });
    return sumClicks;
  } catch (e) {
    return null;
  }
}

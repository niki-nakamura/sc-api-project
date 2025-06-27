function updateModifiedDatesByUrl() {
  // -------------------------------
  // 1) スプレッドシート設定
  // -------------------------------
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // シート名を実際のものに合わせてください
  const sheet = ss.getSheetByName("KW管理表"); 
  const lastRow = sheet.getLastRow();

  // データがない場合は処理終了
  if (lastRow < 2) {
    Logger.log("データが存在しないため処理を終了します。");
    return;
  }

  // -------------------------------
  // 2) C列(列番号3)のURLを取得 (2行目～最終行)
  // -------------------------------
  const urlRange = sheet.getRange(2, 3, lastRow - 1, 1);
  const urlValues = urlRange.getValues(); // 2次元配列

  // -------------------------------
  // 3) WordPress認証情報
  // -------------------------------
  const username = "nakamura_niki";
  const password = "1koCprcDrGX@8mcVVKUJPAKM";
  const encodedAuth = Utilities.base64Encode(username + ":" + password);

  // -------------------------------
  // 4) ループでURL→最終更新日を取得 → AQ列に反映
  // -------------------------------
  const output = [];  // AQ列書き込み用配列

  for (let i = 0; i < urlValues.length; i++) {
    const urlStr = urlValues[i][0]; // C列のURL
    let modifiedDate = ""; // 取得できなければ空欄

    // URLが文字列で、'digi-mado.jp' を含むもののみ対象
    if (typeof urlStr === "string" && urlStr.includes("digi-mado.jp")) {
      const postId = extractPostId(urlStr);
      // デバッグ用
      Logger.log(`Row ${i+2} => URL: ${urlStr}, postId: ${postId}`);

      if (postId) {
        // 単一投稿エンドポイントを叩いて modified を取る
        modifiedDate = fetchModifiedByPostId(postId, encodedAuth);
      }
    } else {
      // デバッグ用
      Logger.log(`Row ${i+2} => URLが空、または digi-mado.jp を含まないためスキップ: ${urlStr}`);
    }

    // 書き込み配列に1行分を追加
    output.push([modifiedDate]);
  }

  // AQ列(列番号42) にまとめて書き込み
  sheet.getRange(2, 42, lastRow - 1, 1).setValues(output);

  Logger.log("AQ列への更新日書き込みが完了しました。(404は無視)");
}

/**
 * URL中のパターンから投稿IDを抜き出す。
 * 例:
 *  - https://digi-mado.jp/article/12345/
 *  - https://digi-mado.jp/article/12345?foo=bar
 *  - 上記URLの12345を返す。
 * 
 * 見つからない場合は空文字 "" を返す。
 */
function extractPostId(urlStr) {
  // 連続する末尾のスラッシュを取り除き、クエリパラメータもあれば削除
  // （ただしクエリが含まれていても下記の正規表現で問題ない場合は不要です）
  urlStr = urlStr.trim(); // 前後スペース除去
  urlStr = urlStr.replace(/\?[^/]+$/, "");    // ?以降を削除（必要であれば）
  urlStr = urlStr.replace(/\/+$/, "");        // 末尾 / を削除

  // '/article/12345' というパターンを想定
  let match = urlStr.match(/\/article\/(\d+)$/);
  if (!match) {
    // もし別パターン(たとえば /archives/12345)があるならこちらも試す
    match = urlStr.match(/\/archives\/(\d+)$/);
  }

  if (match && match[1]) {
    return match[1]; // 数字IDを返す
  }
  return "";
}

/**
 * 投稿IDを指定し、 /wp-json/wp/v2/posts/<ID> から modified を取得
 * 成功時は "2025-03-31T06:11:03" のような日時文字列を返す。
 * 取得失敗・404などの場合は空文字。
 */
function fetchModifiedByPostId(postId, encodedAuth) {
  try {
    const url = `https://digi-mado.jp/wp-json/wp/v2/posts/${postId}`;
    const options = {
      method: "get",
      headers: {
        "Authorization": "Basic " + encodedAuth
      },
      muteHttpExceptions: true // 200以外でもレスポンス取得
    };

    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      const postData = JSON.parse(response.getContentText());
      if (postData && postData.modified) {
        return postData.modified;
      }
    }
    // ステータス200以外は無視 (空欄)
  } catch (err) {
    // 例外は無視
    Logger.log(`Error fetching modified date for postId:${postId} => ${err}`);
  }
  return "";
}

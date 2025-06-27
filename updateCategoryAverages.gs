/**
 * S列の「category」を区切りに、直下のデータ行〜
 * 次の “非連続” の category 行の直前までの D列数値平均を、
 * 連続ブロック先頭の G 列に出力する。
 *
 * ・対象行 : 2〜5000 行（1 行目は見出しで触らない）
 * ・列番号 : D = 4, G = 7, S = 19
 *
 * 例）S 列が
 *   5: category
 *   6: category
 *   7: —（データ）
 *   8: category      ← ここでブロック切替
 * ⇒ G5 に D7 までの平均を書き込み、G6 は空欄のまま
 */
function updateCategoryAverages() {
  /* ---- 0. 設定 ---- */
  const SHEET_NAME  = 'KW管理表';   // ★シート名
  const MAX_ROW     = 5000;         // 処理上限
  const KW_CATEGORY = 'category';   // プルダウン値（小文字比較）
  const COL_D = 4, COL_G = 7, COL_S = 19;

  /* ---- 1. シート取得 ---- */
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`シート『${SHEET_NAME}』が見つかりません`);

  /* ---- 2. 実データ行数を把握（最大 5000 行まで） ---- */
  const lastRow     = Math.min(sh.getLastRow(), MAX_ROW);
  if (lastRow < 2) return;                          // データ無し
  const numRows = lastRow - 1;                      // 2 行目以降

  /* ---- 3. 必要列を一括取得 ---- */
  const sVals = sh.getRange(2, COL_S, numRows).getValues(); // S2:S
  const dVals = sh.getRange(2, COL_D, numRows).getValues(); // D2:D

  /* ---- 4. G 列クリア ---- */
  sh.getRange(2, COL_G, numRows).clearContent();

  /* ---- 5. 「category」ブロックの先頭インデックスを抽出 ---- */
  const topIdxList = [];
  for (let i = 0; i < numRows; i++) {
    const cur = (sVals[i][0] || '').toString().trim().toLowerCase();
    const prev = i > 0
      ? (sVals[i - 1][0] || '').toString().trim().toLowerCase()
      : '';
    if (cur === KW_CATEGORY && prev !== KW_CATEGORY) { // 連続の先頭
      topIdxList.push(i);
    }
  }
  if (!topIdxList.length) return;                    // category 無し

  /* ---- 6. 各ブロックごとに平均計算 ---- */
  topIdxList.forEach((topIdx, idx) => {
    /* 6-1. 連続 category の最後（blockEnd）を求める */
    let blockEnd = topIdx;
    while (
      blockEnd + 1 < numRows &&
      (sVals[blockEnd + 1][0] || '').toString().trim().toLowerCase() === KW_CATEGORY
    ) {
      blockEnd++;
    }

    /* 6-2. 次ブロックの先頭インデックス（無ければデータ末尾） */
    const nextTop = (idx + 1 < topIdxList.length) ? topIdxList[idx + 1] : numRows;

    /* 6-3. 平均対象範囲: blockEnd の次行 〜 nextTop-1 行 */
    const startIdx = blockEnd + 1;   // 0-based
    const endIdxEx = nextTop;        // exclusive

    let sum = 0, cnt = 0;
    for (let r = startIdx; r < endIdxEx; r++) {
      const v = dVals[r][0];
      if (typeof v === 'number' && !isNaN(v)) {
        sum += v; cnt++;
      }
    }

    /* 6-4. 平均をブロック先頭行の G 列へ出力 */
    if (cnt) {
      sh.getRange(topIdx + 2, COL_G).setValue(sum / cnt); // +2: 0-based→実行番号
    }
  });
}

/* ---- オプション：D列 or S列が編集されたら自動再計算 ---- */
function onEdit(e) {
  const col = e.range.getColumn();
  if (col === 4 || col === 19) updateCategoryAverages();
}

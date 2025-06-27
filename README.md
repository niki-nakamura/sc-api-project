# page-load-speed
Good!Apps Page Load Speed &amp; Performance Optimisation – Summary of Results

# sc-api-project

このプロジェクトは、Google Apps Script (GAS) を使用してスプレッドシートを操作し、データの自動処理や外部APIとの連携を行うスクリプト群を提供します。以下は各スクリプトの概要です。

## スクリプト一覧

### 1. removeFullWidthSpacesInColumnB.gs
- **概要**: スプレッドシートのB列に含まれる全角スペースを半角スペースに置き換えます。
- **用途**: データの整形。

---

### 2. updateModifiedDatesByPostId.gs
- **概要**: WordPressの投稿IDを取得し、投稿の最終更新日をスプレッドシートのAQ列に記録します。
- **用途**: WordPressの投稿データ管理。

---

### 3. updateLastCrawlDates.gs
- **概要**: Google Search Console APIを使用して、URLの最終クロール日時を取得し、スプレッドシートに記録します。
- **用途**: クロール状況の追跡。

---

### 4. updateCategoryAverages.gs
- **概要**: S列の「category」ごとにD列の数値の平均を計算し、G列に出力します。
- **用途**: カテゴリごとのデータ集計。

---

### 5. update24hMetrics.gs
- **概要**: B列（query）とC列（page URL）の組み合わせに基づき、直近24時間の平均掲載順位をAG列に出力します。
- **用途**: 検索順位の短期的な変動を追跡。

---

### 6. SiteMetrics.gs
- **概要**: サイト全体のデータを抽出し、30日間および7日間の平均掲載順位やクリック数、CTRなどを計算します。
- **用途**: サイト全体のパフォーマンス分析。

---

### 7. getActiveSpreadsheet.gs
- **概要**: KW管理表の列幅を別シートにコピーし、FILTER関数を設定します。
- **用途**: シート間のデータ同期。

---

### 8. formatUrlColumn.gs
- **概要**: KW管理表のC列（URL列）の書式をリセットし、リンクを削除してテキスト形式に整形します。
- **用途**: URL列のフォーマット統一。

---

### 9. fetchSCAveragePositions.gs
- **概要**: Google Search Console APIを使用して、キーワードとURLごとの平均掲載順位やクリック数、CTRを取得します。
- **用途**: 検索パフォーマンスの詳細分析。

---

### 10. fetchCategoryKWDailyPositions.gs
- **概要**: カテゴリKWデイリー順位シート専用のスクリプトで、キーワードとURLごとの7日間および30日間の平均順位やクリック数を取得します。
- **用途**: カテゴリごとの検索パフォーマンス分析。

---

### 11. セルに「category」と入力されたときに、入力された行の直上に新しい行を作成.gs
- **概要**: U列に「category」と入力された場合、その行の直上に新しい行を挿入し、S列のデータバリデーションと値をコピーします。
- **用途**: カテゴリデータの管理。

---

### 12. lighthouserc.json
- **概要**: プロジェクトの設定ファイルで、タイムゾーンやOAuthスコープなどを定義しています。
- **用途**: プロジェクト全体の設定管理。

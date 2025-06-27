# sc-api-project Google Apps Script 一覧

このプロジェクトは Google スプレッドシート「KW管理表」や「カテゴリKWデイリー順位」などを対象に、SEO・検索順位・WordPress・Search Console 連携などの自動化を行う Google Apps Script (GAS) 群を管理しています。  
各スクリプトの役割・機能を以下にまとめます。

---

## スクリプト一覧と概要

### 1. fetchSCAveragePositions.gs
- **KW管理表**のC列（URL）終端までを対象に、Search Console API からキーワード×URLごとの平均掲載順位・クリック数・インプレッション・CTRをバッチで取得し、D～K列に反映します。
- LockServiceによる排他制御、動的バッチサイズ、途中再開・自動トリガー再実行に対応。

### 2. fetchCategoryKWDailyPositions.gs
- **カテゴリKWデイリー順位**シート専用。
- C列（KW）×E列（URL）ごとに、30日間・7日間の平均掲載順位やクリック数・インプレッション・CTRを取得し、F～K列に反映します。
- シート全体の30日クリック合計もI2セルに出力。

### 3. SiteMetrics.gs
- **カテゴリKWデイリー順位**シートの2行目（サイト全体用）に、URL・平均掲載順位（F/G列）・30日クリック数・インプレッション・CTRなどを集計して出力します。
- F2/G2にはシート内のF列/G列の平均値も自動計算。

### 4. update24hMetrics.gs
- **KW管理表**のB列（query）・C列（page URL）ごとに、直近24時間の平均掲載順位をAG列に出力します。

### 5. updateCategoryAverages.gs (onEdit)
- **KW管理表**のS列「category」行ごとに、直下のデータ行～次のcategory行直前までのD列数値平均を、ブロック先頭のG列に出力します。
- D列またはS列編集時に自動再計算。

### 6. updateLastCrawlDates.gs
- **KW管理表**のC列（URL）ごとに、Google Search Console Index APIを使い、前回クロール日時をAQ列に出力します。
- タイムアウト・途中再開・自動トリガー再実行に対応。

### 7. updateModifiedDatesByPostId.gs
- **KW管理表**のC列（URL）から WordPress 投稿IDを抽出し、REST API経由で最終更新日（modified）を取得、AQ列に反映します。

### 8. formatUrlColumn.gs
- **KW管理表**のC列（URL）の書式（リッチテキスト/リンク/フォント等）を一括リセットします。
- バッチ処理・毎時自動トリガー対応。

### 9. getActiveSpreadsheet.gs (onEdit)
- **KW管理表**→「松浦 嵩」シートへの列幅コピーや、FILTER関数の自動挿入などを行います。

### 10. セルに「category」と入力されたときに、入力された行の直上に新しい行を作成.gs (onEdit)
- **KW管理表**のU列で「category」と入力された際、直上に新規行を挿入し、S列のデータバリデーションと値をコピーします。

### 11. var converted = newValue.replace.gs
- **KW管理表**のB列（2行目以降）で、全角スペースを半角スペースに一括変換します。

---

## 注意事項
- 各スクリプトはシート名や列番号などを前提にしているため、運用時は該当シート構成・カラム構成に注意してください。
- Search Console APIやWordPress REST APIの認証情報・権限設定が必要なものがあります。
- バッチ処理系はGoogle Apps Scriptの実行時間制限（6分）に配慮し、途中再開や自動トリガー再実行に対応しています。

---

## ファイル構成
- fetchSCAveragePositions.gs
- fetchCategoryKWDailyPositions.gs
- SiteMetrics.gs
- update24hMetrics.gs
- updateCategoryAverages.gs
- updateLastCrawlDates.gs
- updateModifiedDatesByPostId.gs
- formatUrlColumn.gs
- getActiveSpreadsheet.gs
- セルに「category」と入力されたときに、入力された行の直上に新しい行を作成.gs
- var converted = newValue.replace.gs

---

## 連絡先
- スクリプトの仕様・運用に関する質問は管理者までご連絡ください。
- **用途**: プロジェクト全体の設定管理。

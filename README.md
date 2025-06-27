# sc-api-project Google Apps Script 一覧

このプロジェクトは Google スプレッドシート「KW管理表」や「カテゴリKWデイリー順位」などを対象に、SEO・検索順位・WordPress・Search Console 連携などの自動化を行う Google Apps Script (GAS) 群を管理しています。  
各スクリプトの役割・機能を以下にまとめます。

---

## 列構成（KW管理表）

| 列 | 項目名 |
|----|----------------------|
| A  | KW                   |
| B  | URL                  |
| C  | 30日間の平均掲載順位 |
| D  | 7日間                |
| E  | 比較                 |
| F  | 記事全体平均順位     |
| G  | クリック数           |
| H  | 合計表示回数         |
| I  | 平均CTR              |
| J  | 進捗（%）            |
| K  | --                   |
| L  | 担当                 |
| M  | KW選定者             |
| N  | KW選定シート         |
| O  | volume               |
| P  | 大カテゴリ（部門）   |
| Q  | 製品カテゴリ         |
| R  | 投稿タイプ           |
| S  | 想定文字数（WP）     |
| T  | 骨格作成者           |
| U  | 骨格期限日           |
| V  | 骨格ステータス       |
| W  | 骨格完了日           |
| X  | 骨格URL              |
| Y  | 骨格修正ステータス   |
| Z  | 骨格修正理由         |
| AA | 記事執筆者           |
| AB | 執筆期限日           |
| AC | 執筆ステータス       |
| AD | 執筆完了日           |
| AE | WordPress URL        |
| AF | 記事修正ステータス   |
| AG | 記事修正理由         |
| AH | 編集・校閲担当者     |
| AI | 編集・校閲ステータス |
| AJ | 編集・校閲完了日     |
| AK | 最終文字数(WP)      |
| AL | リライト担当者       |
| AM | リライトステータス   |
| AN | リライト日           |
| AO | modified             |
| AQ | crawl                |

---

## スクリプト一覧と概要

### 1. fetchSCAveragePositions.gs
- KW管理表のC列「URL」終端までを対象に、Search Console API からA列「KW」×B列「URL」ごとのC列「30日間の平均掲載順位」、D列「7日間」、E列「比較」、G列「クリック数」、H列「合計表示回数」、I列「平均CTR」などをバッチで取得し、該当列に反映します。
- LockServiceによる排他制御、動的バッチサイズ、途中再開・自動トリガー再実行に対応。

### 2. fetchCategoryKWDailyPositions.gs
- カテゴリKWデイリー順位シート専用。
- C列「KW」×E列「URL」ごとに、F列「30日間の平均掲載順位」、G列「7日間の平均掲載順位」、H列「比較」、I列「クリック数」、J列「合計表示回数」、K列「平均CTR」などを取得し、該当列に反映します。
- シート全体のI2セルに30日クリック合計も出力。

### 3. SiteMetrics.gs
- カテゴリKWデイリー順位シートの2行目（サイト全体用）に、E列「URL」、F列「30日間の平均掲載順位」、G列「7日間の平均掲載順位」、H列「比較」、I列「クリック数」、J列「合計表示回数」、K列「平均CTR」などを集計して出力します。
- F2/G2にはシート内のF列/G列の平均値も自動計算。

### 4. update24hMetrics.gs
- KW管理表のA列「KW」・B列「URL」の組み合わせごとに、直近24時間の平均掲載順位をAG列「記事修正理由」へ出力します。

### 5. updateCategoryAverages.gs (onEdit)
- KW管理表のS列「投稿タイプ」「category」行ごとに、直下のデータ行～次のcategory行直前までのD列「7日間」数値平均を、ブロック先頭のG列「記事全体平均順位」に出力します。
- D列「7日間」またはS列「投稿タイプ」編集時に自動再計算。

### 6. updateLastCrawlDates.gs
- KW管理表のB列「URL」ごとに、Google Search Console Index APIを使い、AQ列「crawl」に前回クロール日時を出力します。
- タイムアウト・途中再開・自動トリガー再実行に対応。

### 7. updateModifiedDatesByPostId.gs
- KW管理表のB列「URL」から WordPress 投稿IDを抽出し、REST API経由でAO列「modified」に最終更新日を反映します。

### 8. formatUrlColumn.gs
- KW管理表のB列「URL」の書式（リッチテキスト/リンク/フォント等）を一括リセットします。
- バッチ処理・毎時自動トリガー対応。

### 9. getActiveSpreadsheet.gs (onEdit)
- KW管理表→「松浦 嵩」シートへの列幅コピーや、FILTER関数の自動挿入などを行います。

### 10. セルに「category」と入力されたときに、入力された行の直上に新しい行を作成.gs (onEdit)
- KW管理表のU列「骨格期限日」で「category」と入力された際、直上に新規行を挿入し、S列「投稿タイプ」のデータバリデーションと値をコピーします。

### 11. var converted = newValue.replace.gs
- KW管理表のA列「KW」またはB列「URL」（2行目以降）で、全角スペースを半角スペースに一括変換します。

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

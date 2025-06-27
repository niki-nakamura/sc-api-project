# sc-api-project Google Apps Script 一覧

このプロジェクトは Google スプレッドシート「KW管理表」や「カテゴリKWデイリー順位」などを対象に、SEO・検索順位・WordPress・Search Console 連携などの自動化を行う Google Apps Script (GAS) 群を管理しています。

---

## スクリプト一覧と概要

### 1. fetchSCAveragePositions.gs
- **KW管理表のB列「KW」・C列「URL」ごとに、D列「30日間の平均掲載順位」、E列「7日間」、F列「比較」、H列「クリック数」、I列「合計表示回数」、J列「平均CTR」などをSearch Console APIからバッチ取得し、該当列に反映します。**
- LockServiceによる排他制御、動的バッチサイズ、途中再開・自動トリガー再実行に対応。
- **トリガー:** 定期（午前3時～4時, batchFetchSCAveragePositions）

### 2. fetchCategoryKWDailyPositions.gs
- **カテゴリKWデイリー順位シートのC列「KW」・E列「URL」ごとに、F列「30日間の平均掲載順位」、G列「7日間の平均掲載順位」、H列「比較」、I列「クリック数」、J列「合計表示回数」、K列「平均CTR」などを取得し、該当列に反映します。**
- シート全体のI2セルに30日クリック合計も出力。
- **トリガー:** 定期（午前0時～1時, fetchCategoryKWDailyPositions）

### 3. SiteMetrics.gs
- **カテゴリKWデイリー順位シートの2行目（サイト全体用）に、E列「URL」、F列「30日間の平均掲載順位」、G列「7日間の平均掲載順位」、H列「比較」、I列「クリック数」、J列「合計表示回数」、K列「平均CTR」などを集計して出力します。**
- F2/G2にはシート内のF列/G列の平均値も自動計算。
- **トリガー:** 定期（午前1時～2時, fetchSiteData）

### 4. updateCategoryAverages.gs (onEdit)
- **KW管理表のS列「投稿タイプ」で「category」行ごとに、直下のデータ行～次のcategory行直前までのD列「30日間の平均掲載順位」数値平均を、ブロック先頭のG列「記事全体平均順位」に出力します。**
- D列またはS列編集時に自動再計算。
- **トリガー:** onEdit, 定期（午前0時～1時, updateCategoryAverages）

### 5. updateLastCrawlDates.gs
- **KW管理表のC列「URL」ごとに、Google Search Console Index APIを使い、AQ列「crawl」に前回クロール日時を出力します。**
- タイムアウト・途中再開・自動トリガー再実行に対応。
- **トリガー:** 定期（午前6時～7時, updateLastCrawlDatesFirst50）

### 6. updateModifiedDatesByPostId.gs
- **KW管理表のC列「URL」から WordPress 投稿IDを抽出し、REST API経由でAP列「modified」に最終更新日を反映します。**
- **トリガー:** 定期（午前2時～3時, updateModifiedDatesByUrl）

### 7. formatUrlColumn.gs
- **KW管理表のC列「URL」の書式（リッチテキスト/リンク/フォント等）を一括リセットします。**
- バッチ処理・毎時自動トリガー対応。
- **トリガー:** 定期（1時間おき, formatUrlColumnHour）

### 8. var converted = newValue.replace.gs
- **KW管理表のB列「KW」またはC列「URL」（2行目以降）で、全角スペースを半角スペースに一括変換します。**
- **トリガー:** 定期（1時間おき, removeFullWidthSpacesInColumnB）

### 9. setStatusToCompletedForUrlRows.gs
- **KW管理表のC列「URL」に値がある行のW列（骨格ステータス）、AD列（執筆ステータス）、AJ列（編集・校閲ステータス）を「完了」に一括更新します。**
- **トリガー:** 手動実行、onEdit対応

### 10. setCategoryTypeForCategoryUrls.gs
- **KW管理表のC列「URL」に「category」を含む行のS列「投稿タイプ」を「category」に一括設定します。**
- **トリガー:** 手動実行

---

## トリガー設定のないGAS一覧

- **setStatusToCompletedForUrlRows.gs**（手動実行・onEdit対応だが定期トリガーなし）
- **setCategoryTypeForCategoryUrls.gs**（手動実行のみ）

※ `update24hMetrics.gs` は削除済みです。

---

## ファイル構成

- fetchSCAveragePositions.gs
- fetchCategoryKWDailyPositions.gs
- SiteMetrics.gs
- updateCategoryAverages.gs
- updateLastCrawlDates.gs
- updateModifiedDatesByPostId.gs
- formatUrlColumn.gs
- var converted = newValue.replace.gs
- setStatusToCompletedForUrlRows.gs
- setCategoryTypeForCategoryUrls.gs

---

## Scheduled Triggers（定期実行トリガー一覧）

| トリガー名                        | 実行時間帯         | 備考                      |
|-----------------------------------|--------------------|---------------------------|
| batchFetchSCAveragePositions      | 午前3時～4時       | Search Console平均順位バッチ |
| updateLastCrawlDatesFirst50       | 午前6時～7時       | 前回クロール日時（最初の50件）|
| formatUrlColumnHour               | 1時間おき          | URL列書式リセット           |
| updateCategoryAverages            | 午前0時～1時       | カテゴリ平均自動計算         |
| fetchCategoryKWDailyPositions     | 午前0時～1時       | カテゴリKWデイリー順位抽出    |
| fetchSiteData                     | 午前1時～2時       | サイト全体データ抽出         |
| updateModifiedDatesByUrl          | 午前2時～3時       | WordPress最終更新日取得      |
| removeFullWidthSpacesInColumnB    | 1時間おき          | B列全角スペース除去          |

---

## 注意事項

- 各スクリプトはシート名や列番号などを前提にしているため、運用時は該当シート構成・カラム構成に注意してください。
- Search Console APIやWordPress REST APIの認証情報・権限設定が必要なものがあります。
- バッチ処理系はGoogle Apps Scriptの実行時間制限（6分）に配慮し、途中再開や自動トリガー再実行に対応しています。

---

## 列構成（KW管理表）

| 列 | 項目名                |
|----|-----------------------|
| B  | KW                    |
| C  | URL                   |
| D  | 30日間の平均掲載順位  |
| E  | 7日間                 |
| F  | 比較                  |
| G  | 記事全体平均順位      |
| H  | クリック数            |
| I  | 合計表示回数          |
| J  | 平均CTR               |
| K  | 進捗（%）             |
| L  | （空欄 ※予備列）      |
| M  | 担当                  |
| N  | KW選定者              |
| O  | KW選定シート          |
| P  | volume                |
| Q  | 大カテゴリ（部門）    |
| R  | 製品カテゴリ          |
| S  | 投稿タイプ            |
| T  | 想定文字数（WP）      |
| U  | 骨格作成者            |
| V  | 骨格期限日            |
| W  | 骨格ステータス        |
| X  | 骨格完了日            |
| Y  | 骨格URL               |
| Z  | 骨格修正ステータス    |
| AA | 骨格修正理由          |
| AB | 記事執筆者            |
| AC | 執筆期限日            |
| AD | 執筆ステータス        |
| AE | 執筆完了日            |
| AF | WordPress URL         |
| AG | 記事修正ステータス    |
| AH | 記事修正理由          |
| AI | 編集・校閲担当者      |
| AJ | 編集・校閲ステータス  |
| AK | 編集・校閲完了日      |
| AL | 最終文字数 (WP)       |
| AM | リライト担当者        |
| AN | リライトステータス    |
| AO | リライト日            |
| AP | modified              |
| AQ | crawl                 |

---

## 連絡先

- スクリプトの仕様・運用に関する質問は管理者までご連絡ください。
- **用途**: プロジェクト全体の設定管理。

---

## トリガー設定の最適化について

### 最適な実行順・時間帯の考え方

- **依存関係があるものは順番を守る**（例：順位データ取得→集計→書式リセット）
- **API制限や負荷分散を考慮し、深夜～早朝に分散実行**
- **onEdit系は手動編集時のみ発火なので定期トリガー不要**
- **書式リセットやスペース除去など軽い処理は他バッチと重複してもOK**

---

### 推奨トリガー順・時間帯

| 実行順 | トリガー名                        | 実行時間帯         | 備考・理由                                  |
|--------|-----------------------------------|--------------------|---------------------------------------------|
| 1      | batchFetchSCAveragePositions      | 午前3時～4時       | 検索順位・クリック等の生データ取得（最初に） |
| 2      | fetchCategoryKWDailyPositions     | 午前4時～5時       | カテゴリKWデイリー順位抽出                   |
| 3      | fetchSiteData                     | 午前5時～6時       | サイト全体集計（カテゴリKWデータ依存）        |
| 4      | updateCategoryAverages            | 午前6時～7時       | カテゴリ平均自動計算（順位データ依存）        |
| 5      | updateLastCrawlDatesFirst50       | 午前7時～8時       | 前回クロール日時（順位データ後でOK）          |
| 6      | updateModifiedDatesByUrl          | 午前8時～9時       | WordPress最終更新日取得                      |
| 7      | formatUrlColumnHour               | 1時間おき          | URL列書式リセット（いつでもOK）              |
| 8      | removeFullWidthSpacesInColumnB    | 1時間おき          | B列全角スペース除去（いつでもOK）            |

- **onEdit系（updateCategoryAverages, getActiveSpreadsheet, セルに「category」と入力...）は定期トリガー不要**
- **update24hMetrics.gs** は手動実行推奨（API負荷・用途限定のため）

---

### 補足

- **順位データ取得（batchFetchSCAveragePositions）→カテゴリ集計（fetchCategoryKWDailyPositions, fetchSiteData, updateCategoryAverages）→クロール・WP更新日系**の順がデータ整合性・運用効率の観点で最適です。
- **書式リセット・スペース除去**は他バッチと重複しても問題ありません。
- **API制限やシート負荷を避けるため、1時間ごとに分散実行**するのがベストです。
- **API制限やシート負荷を避けるため、1時間ごとに分散実行**するのがベストです。
- **API制限やシート負荷を避けるため、1時間ごとに分散実行**するのがベストです。
- **API制限やシート負荷を避けるため、1時間ごとに分散実行**するのがベストです。
- **書式リセット・スペース除去**は他バッチと重複しても問題ありません。
- **API制限やシート負荷を避けるため、1時間ごとに分散実行**するのがベストです。
- **API制限やシート負荷を避けるため、1時間ごとに分散実行**するのがベストです。
- **API制限やシート負荷を避けるため、1時間ごとに分散実行**するのがベストです。
- **API制限やシート負荷を避けるため、1時間ごとに分散実行**するのがベストです。
- **API制限やシート負荷を避けるため、1時間ごとに分散実行**するのがベストです。
- **API制限やシート負荷を避けるため、1時間ごとに分散実行**するのがベストです。
- **API制限やシート負荷を避けるため、1時間ごとに分散実行**するのがベストです。
- **API制限やシート負荷を避けるため、1時間ごとに分散実行**するのがベストです。

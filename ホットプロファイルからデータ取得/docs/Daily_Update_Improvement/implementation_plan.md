# 実装計画: 日次更新の堅牢化とAPIリトライ

## ゴール
1. **日次更新のタイムアウト対策**: `updateCompanyListDailyDiff` をバッチ分割（トリガー実行）に対応させ、大量更新時もタイムアウトせず確実に処理できるようにする。
2. **APIリトライ処理の導入**: 通信エラー発生時に自動リトライを行う共通関数を導入し、データ取得の確実性を高める。

## ユーザーレビューが必要な事項
- 特になし（内部ロジックの改善のみ）

## 提案例

### `共通.js`

#### [MODIFY] [共通.js](file:///c:/Users/ten01/OneDrive/ドキュメント/VS-Code/VS-Code/ホットプロファイルからデータ取得/共通.js)
- **`fetchWithRetry` 関数の追加**: `UrlFetchApp.fetch` をラップし、ステータスコードが異常または例外発生時に、指定回数（デフォルト3回）のリトライを行う関数を追加します。
    - **ログ仕様**: リトライ発生時には `Logger.log` に「APIリトライ実行中 (1/3): エラー内容...」のような警告ログを出力します。最終的に失敗した場合はエラーをスローし、呼び出し元でエラーハンドリング（取得ログへの MISS 記録など）を行います。

### `企業情報.js`

#### [MODIFY] [企業情報.js](file:///c:/Users/ten01/OneDrive/ドキュメント/VS-Code/VS-Code/ホットプロファイルからデータ取得/企業情報.js)
- **`fetchCompanyById`, `fetchCompanyPage`, `fetchAllClientsByUpdatedRange` の修正**:
    - `UrlFetchApp.fetch` の代わりに `fetchWithRetry` を使用するように置換します。
    - `fetchAllClientsByUpdatedRange` はページネーションで全件取得する方式から、バッチ処理で部分取得する方式へ変更するための準備（あるいはロジック変更）を行います。
- **日次更新のバッチ化**:
    - `startDailyDiffStep` 関数を新設（週次と同様の初期化処理）。
    - `updateCompanyListDailyDiff` を `updateCompanyListDailyDiffStep`（トリガー呼び出し対応）にリファクタリング。
        - 実行時間監視（タイムアウト防止）。
        - カーソル管理（プロパティ使用）。
        - 処理完了後に `reapplyFormulasFromSettings` と `startLastTouchStep` を呼び出す。

### `ラストタッチ確認.js`

#### [MODIFY] [ラストタッチ確認.js](file:///c:/Users/ten01/OneDrive/ドキュメント/VS-Code/VS-Code/ホットプロファイルからデータ取得/ラストタッチ確認.js)
- **`fetchLatestVisitForClient`, `updateLastTouchStep` の修正**:
    - API呼び出し箇所を `fetchWithRetry` に置き換え可能であれば適用します（ただし `UrlFetchApp.fetchAll` を使用している箇所は、個別のリトライ制御が複雑になるため、今回は単発取得の `fetchLatestVisitForClient` を優先して適用検討します。並列取得のリトライは複雑なため、リスク回避で現状維持または簡易的な再試行のみ検討）。

## 検証計画

### 自動テスト
- 現在自動テスト環境はないため、GAS上での実行確認となります。

### 手動検証
1. **APIリトライ確認**:
    - 意図的に無効なURLやパラメータでエラーを起こし、リトライログが出力されるか確認（既存への影響を極小化するため、デバッグ関数で確認）。
2. **日次更新バッチ動作確認**:
    - `startDailyDiffStep` を実行し、プロパティにカーソルが保存され、トリガーが設定されるか確認。
    - トリガー経由で `updateCompanyListDailyDiffStep` が実行され、ログに処理件数と継続/完了ステータスが記録されるか確認。
    - 完了後にラストタッチ更新処理へ遷移するか確認。

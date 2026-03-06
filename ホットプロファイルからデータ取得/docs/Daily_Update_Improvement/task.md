# タスク: 日次更新の堅牢化とAPIリトライ

## チェックリスト

### 1. APIリトライ処理の実装
- [x] `共通.js` に `fetchWithRetry` 関数を追加
- [x] `企業情報.js` の `fetchCompanyById` を `fetchWithRetry` に置換
- [x] `企業情報.js` の `fetchCompanyPage` を `fetchWithRetry` に置換
- [x] `企業情報.js` の `fetchAllClientsByUpdatedRange` を `fetchWithRetry` に置換

### 2. 日次更新のバッチ化
- [x] `startDailyDiffStep` 関数を新設
- [x] `updateCompanyListDailyDiffStep` 関数を新設（トリガー対応版）
- [x] `finishDailyDiffStep` 関数を新設（終了処理）
- [x] 既存の `updateCompanyListDailyDiff` は互換用として維持

### 3. 検証
- [x] コード構文チェック（実装完了）
- [ ] 手動テスト手順の確認（ユーザー実施）

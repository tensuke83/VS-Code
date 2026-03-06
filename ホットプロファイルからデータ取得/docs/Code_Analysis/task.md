# タスクリスト: ホットプロファイル連携の改善

- [x] **現状のコード分析** <!-- id: 0 -->
    - [x] 既存ファイルのレビュー (`ラストタッチ確認.js`, `企業情報.js`, `共通.js`, `不足分の再取得.js`)
    - [x] 未定義関数の特定 (`fetchLatestVisitForClient`, `fetchCompanyPage`, `fetchCompanyById`)
    - [x] パフォーマンスボトルネックの特定 (ループ内での同期的なAPI呼び出し)
- [x] **未定義関数の実装** <!-- id: 1 -->
    - [x] `fetchLatestVisitForClient` の実装 (`ラストタッチ確認.js` または `共通.js`)
    - [x] `fetchCompanyPage` の実装 (`企業情報.js` または `共通.js`)
    - [x] `fetchCompanyById` の実装 (`企業情報.js` または `共通.js`)
- [x] **パフォーマンスの最適化** <!-- id: 2 -->
    - [x] `upsertClients` から `fetchLatestVisitForClient` の同期呼び出しを削除 (または条件付き/バッチ最適化)
    - [x] 既存の `startLastTouchStep` (並列処理) を活用して一括更新を行うようにする
- [ ] **リファクタリングとクリーンアップ** <!-- id: 3 -->
    - [ ] API呼び出しロジックの共通化
    - [ ] エラーハンドリングとログ出力の統一
    - [x] **取得ログの項目整理** (不要な列の削除)
    - [x] **通知精度の向上** (複数トリガーまたぎでの集計)

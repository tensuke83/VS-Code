# Teams通知レイアウト改善 - Walkthrough

## 変更内容

### `共通.js` の変更

#### 1. `sendTeamsAdaptiveCard` の拡張
- `payload.body`（配列）が渡された場合、Adaptive Card の body としてそのまま使用
- 従来の `payload.title` + `payload.markdown` 形式も互換維持
- 変数名 `body` → `resBody` に変更（`body` 配列との競合回避）

#### 2. `safeNotifySummary` のリッチレイアウト化
- **タイトル**: エラー時は赤色（`attention`）で強調表示
- **サブタイトル**: モード名を薄い文字で表示
- **対象別サマリ**: `■ 対象名` + `FactSet` で数値を整列表示
- **時刻情報**: `FactSet` で開始/終了/所要時間を整列表示
- **区切り線**: 罫線文字でセクション区切り

#### 3. 全幅表示（前回対応済み）
- `msteams: { width: "Full" }` による全幅表示は維持

## 後方互換性
- `safeNotify({ title, markdown, level })` は従来通り動作（`ラストタッチ確認.js` から呼ばれている）
- `sendTeamsAdaptiveCard` は `title`+`markdown` 形式でも `body` 配列形式でも受け付ける

## 検証
- コードレビューにより、既存の呼び出し元（`企業情報.js`、`ラストタッチ確認.js`）との互換性を確認済み
- 実際のTeams通知の表示はスクリプト実行時に確認が必要

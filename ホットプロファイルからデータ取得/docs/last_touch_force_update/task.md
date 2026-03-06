# タスクリスト: 企業情報更新時のラストタッチ再取得

- [x] `ラストタッチ確認.js` に `forceUpdateLastTouch` 関数を作成する
- [x] `企業情報.js` の `upsertClients` を修正する（更新行追跡と戻り値の追加）
- [x] `企業情報.js` の `updateCompanyListDailyDiffStep` を修正する（`forceUpdateLastTouch` 呼び出し）
- [x] 週次更新 (`updateCompanyListStepSingleIdMode`) は影響なし確認済み（週次は最後に全件ラストタッチ走査が走るため変更不要）

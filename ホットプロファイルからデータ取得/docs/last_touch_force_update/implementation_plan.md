# 実装計画: 企業情報更新時のラストタッチ再取得

## 目的
「日次の企業情報更新」(`updateCompanyListDailyDiffStep`) が実行される際、既存の行が更新されると「ラストタッチ」(`ラストタッチ`) 列が上書きされて消えてしまう問題があります。
企業情報が更新されたレコードについては、即座に「ラストタッチ」情報を再取得して更新するようにします。

## 変更内容

### [ラストタッチ確認.js](file:///c:/Users/ten01/OneDrive/%E3%83%89%E3%82%AD%E3%83%A5%E3%83%A1%E3%83%B3%E3%83%88/VS-Code/VS-Code/%E3%83%9B%E3%83%83%E3%83%88%E3%83%97%E3%83%AD%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E3%81%8B%E3%82%89%E3%83%87%E3%83%BC%E3%82%BF%E5%8F%96%E5%BE%97/%E3%83%A9%E3%82%B9%E3%83%88%E3%82%BF%E3%83%83%E3%83%81%E7%A2%BA%E8%AA%8D.js)

#### [NEW] `forceUpdateLastTouch` 関数の追加
- **パラメータ**: `targets` - `{ rowIndex, companyName, companyId }` の配列
- **ロジック**:
    1.  `targets` をループ処理します。
    2.  各ターゲットについて、HotProfile API (`daily_reports/get_entry_list`) を `client_name` で検索します（`order: { key: "visit_on", type: "desc" }`, `limit: 1`）。
    3.  レポートが見つかった場合、その `rowIndex` の「ラストタッチ」日時とURLをシートに書き込みます。
    4.  既存の `updateLastTouchStep` のロジック（`getAllClientIdsFromReport` など）を可能な限り再利用します。

### [企業情報.js](file:///c:/Users/ten01/OneDrive/%E3%83%89%E3%82%AD%E3%83%A5%E3%83%A1%E3%83%B3%E3%83%88/VS-Code/VS-Code/%E3%83%9B%E3%83%83%E3%83%88%E3%83%97%E3%83%AD%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E3%81%8B%E3%82%89%E3%83%87%E3%83%BC%E3%82%BF%E5%8F%96%E5%BE%97/%E4%BC%81%E6%A5%AD%E6%83%85%E5%A0%B1.js)

#### [MODIFY] `upsertClients`
- 戻り値または内部ロジックを変更し、更新された行を収集できるようにします。
- 現在のシグネチャ: `upsertClients(clients, tenants, sheet, existingDataMap, processedIds, logEntries, context)`
- 戻り値: `updatedRows` (更新された行を特定するオブジェクトの配列) を追加します。
- ループ内で `UPDATE` が発生した際 (820行目付近):
    - `{ rowIndex, name: client.name, id: responseId }` を `updatedRows` に追加します。
- 最後に `updatedRows` を返します。

#### [MODIFY] `updateCompanyListDailyDiffStep`
- `upsertClients` から返された `updatedRows` を受け取ります。
- このリストを `forceUpdateLastTouch` に渡して実行します。

## 検証計画

### マニュアル検証
1.  **モックテスト**: `updateCompanyListDailyDiffStep` を模倣し、特定の既知の会社ID（シートに存在するもの）のみを処理するテスト関数を作成します。
2.  **事前準備**: シート上の対象企業の「ラストタッチ」値を手動で削除し、「更新日時」を古くして更新対象になるようにします。
3.  **実行**: テスト関数を実行します。
4.  **事後確認**: 以下を確認します。
    -   企業情報が更新されていること（「更新日時」や「取得日時」で確認）。
    -   「ラストタッチ」列に値が入っていること（空でないこと）。

### コード検証
- ログを確認し、「Force Update Last Touch」等のエントリが出力されていることを確認します。

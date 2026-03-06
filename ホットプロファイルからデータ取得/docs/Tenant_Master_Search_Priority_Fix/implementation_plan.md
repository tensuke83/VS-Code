# テナントマスタ検索優先順位の修正

現状、企業情報の更新処理において、テナントマスタから情報を取得する際に「リクエスト時の会社ID」を優先して検索しています。
しかし、ホットプロファイル側で会社IDが更新（旧IDから新IDへ変更）されている場合、テナントマスタに新旧両方のIDが紐づいているケースがあり、旧IDに紐づく不要な情報（テナント_顧客ID等）を参照してしまう問題が発生しています。

本修正では、APIレスポンスから得られる最新のID（新会社ID）を最優先でテナントマスタから検索するように変更します。

## 提案される変更点

### 企業情報の取得・更新ロジック

#### [MODIFY] [企業情報.js](file:///c:/Users/ten01/OneDrive/%E3%83%89%E3%82%AD%E3%83%A5%E3%83%A1%E3%83%B3%E3%83%88/VS-Code/VS-Code/%E3%83%9B%E3%83%83%E3%83%88%E3%83%97%E3%83%AD%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E3%81%8B%E3%82%89%E3%83%87%E3%83%BC%E3%82%BF%E5%8F%96%E5%BE%97/%E4%BC%81%E6%A5%AD%E6%83%85%E5%A0%B1.js)

- `upsertClients` 関数内の `tenantInfo` 取得ロジックを変更します。
  - 現在: `primaryCompanyId` (リクエストID優先) -> `responseId` (APIレスポンスID) の順で検索
  - 修正後: `responseId` (新会社ID) -> `primaryCompanyId` (リクエストID) の順で検索

```javascript
// 修正前のロジック
const tenantInfo = tenants.find(t => Number(t.companyId) === Number(primaryCompanyId))
  || tenants.find(t => Number(t.companyId) === Number(responseId));

// 修正後のロジック
const tenantInfo = tenants.find(t => Number(t.companyId) === Number(responseId))
  || tenants.find(t => Number(t.companyId) === Number(primaryCompanyId));
```

## 検証計画

### 手動確認
1. テナントマスタに以下のテストデータを準備します。
   - レコードA: ホットプロファイル会社ID=1001, テナント_顧客ID=OLD_TENANT
   - レコードB: ホットプロファイル会社ID=2001, テナント_顧客ID=NEW_TENANT
2. ホットプロファイル側でID 1001 が 2001 に更新されているケースを想定します。
3. スクリプトを実行し、ID 2001（新会社ID）がレスポンスに含まれる場合、出力シートの「テナント_顧客ID」が `NEW_TENANT` になっていることを確認します。
   - 現状のロジックでは `OLD_TENANT` が優先されてしまいます。

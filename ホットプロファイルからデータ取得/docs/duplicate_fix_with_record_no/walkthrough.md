# ウォークスルー: レコード番号による重複排除精度の向上

## 課題
同一会社IDで「テナント_顧客ID」が重複（または空欄で重複）している場合、従来の複合キー（会社ID + テナント顧客ID）では、以下の問題が発生していました：
1. `processedIds` でキー衝突により2件目以降がスキップされる
2. `existingDataMap` でキー衝突により片方のレコードのみが更新される

## 解決策
テナントマスタ固有の「id」列（= `recordNo`）をユニークキーとして利用することで、これらの問題を解決しました。

## 変更ファイル

### `企業情報.js`

#### 1. `buildExistingDataMapByHeader` の修正 (行 504-564)

**レコード番号列の読み取り追加:**
```javascript
const colRecordNo = headers.indexOf("レコード番号") + 1;
const recordNos = (colRecordNo > 0) 
  ? sheet.getRange(2, colRecordNo, numRows, 1).getValues() 
  : null;
```

**キー生成ロジックの変更:**
- **レコード番号あり**: `REC:{recordNo}` をキーとして登録
- **レコード番号なし**: 従来通り複合キー（会社ID + テナント顧客ID）で登録（過去データ救済）

```javascript
if (recordNo !== "") {
  const recordKey = `REC:${recordNo}`;
  map.set(recordKey, rowData);
} else {
  // フォールバック: 複合キー
}
```

#### 2. `upsertClients` の修正

**A. 重複チェックキーの変更 (行 677-681):**
```javascript
const dedupeKey = tenantInfo.recordNo
  ? `PROC:${responseId}_${tenantInfo.recordNo}`
  : generateCompositeKey(responseId, tenantInfo.tenantCustomerId);
```

**B. 既存行検索の優先順位変更 (行 761-773):**
1. まずレコード番号で検索
2. 見つからなければ複合キーで検索（フォールバック）

**C. マップ更新ロジックの変更 (行 776-787, 819-830):**
- レコード番号がある場合は `REC:{recordNo}` で登録
- ない場合は従来の複合キーで登録

## 効果
- テナント顧客IDが空欄または重複していても、`recordNo` により一意に識別・更新可能
- 既存データ（レコード番号がない行）も複合キーのフォールバックで正常動作

## 動作イメージ

```
【テナントマスタ】
| id  | ホットプロファイル会社ID | テナント_顧客ID |
|-----|--------------------------|-----------------|
| 123 | 12345                    | (空欄)          |
| 456 | 12345                    | (空欄)          |

【企業情報シート（更新後）】
| レコード番号 | 会社ID | 新会社ID | テナント_顧客ID | ... |
|--------------|--------|----------|-----------------|-----|
| 123          | 12345  | 12345    | (空欄)          | ... |
| 456          | 12345  | 12345    | (空欄)          | ... |
```

両方のレコードが正しく登録・更新されます。

# ウォークスルー: 契約ステータスによるフィルタリング

## 概要
テナントマスタシートの「契約ステータス」列を確認し、「解約済み」となっているレコードを企業情報の取得対象から除外するよう改修しました。

## 変更ファイル

### `企業情報.js` の `getTenantMasterData` 関数

#### 1. ヘッダエイリアスの追加 (行 333-334)
```javascript
// 契約ステータス（「解約済み」を除外するため）
contractStatus: ["契約ステータス", "ステータス", "契約状態"]
```

#### 2. 列番号の解決 (行 350)
```javascript
const colContractStatus = resolveColumnIndex(headerAliases.contractStatus);    // 任意（解約済み除外用）
```
- 必須項目ではないため、列が見つからない場合でもエラーにならない
- 列がなければ全件対象

#### 3. データ読み取り (行 373)
```javascript
const contractStatuses = (colContractStatus > 0) 
  ? sheet.getRange(2, colContractStatus, numRows, 1).getValues() 
  : null;
```

#### 4. フィルタリングロジック (行 380-385)
```javascript
// 契約ステータスが「解約済み」の場合はスキップ
if (contractStatuses) {
  const statusVal = String(contractStatuses[i][0] || "").trim();
  if (statusVal === "解約済み") {
    continue;
  }
}
```

## 動作仕様
| 契約ステータス | 処理 |
|----------------|------|
| 本契約         | 対象 |
| 解約済み       | 除外 |
| (空欄)         | 対象 |
| その他の値     | 対象 |

## 検証方法
1. テナントマスタに「契約ステータス」列を追加し、様々な値を設定
2. 週次または日次更新を実行
3. 「解約済み」のレコードのみが企業情報シートに出力されていないことを確認

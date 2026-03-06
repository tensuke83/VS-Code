# ウォークスルー: 企業情報出力時の列名エイリアス対応

## 課題
「テナント_顧客ID」などの列に値が入らないケースや、「顧客ID」との混同を防ぐため、企業情報シートへの出力ロジックをより堅牢にする必要がありました。

## 変更ファイル

### `企業情報.js` の `upsertClients`

#### 1. エイリアス検索ヘルパーの導入 (行 604-610)
```javascript
const findColIdx = (candidates) => {
  for (const c of candidates) {
    if (h2i[c] !== undefined) return h2i[c];
  }
  return undefined;
};
```

#### 2. 主要列のエイリアス定義 (行 612-624)
- **テナント_顧客ID**: `["テナント_顧客ID", "テナント顧客ID", "顧客ID（テナント）", "TenantCustomerId"]`
  - ※ "顧客ID" を除外して混同を防止
- **顧客ID**: `["顧客ID", "会社コード", "メンテナンス用(会社コード)"]`
- **レコード番号**: `["レコード番号", "レコードID", "id"]`

#### 3. 書き込みロジックの修正 (行 755-776)
`switch` 文内でヘッダ文字列による判定にもエイリアスを適用し、インデックスが一致する場合に値を書き込むよう変更しました。

```javascript
// テナント_顧客ID（テナントマスタ由来）
case "テナント_顧客ID":
case "テナント顧客ID":
case "顧客ID（テナント）":
case "TenantCustomerId":
  if (i === idxTenantCustomerId) rowData[i] = ...;
  break;
```

## 効果
- シートのヘッダ名が「テナント顧客ID」や「TenantCustomerId」であっても正しく値が出力されます。
- HotProfile由来の「顧客ID」とテナントマスタ由来の「テナント_顧客ID」が混同されず、それぞれ正しい列に出力されます。

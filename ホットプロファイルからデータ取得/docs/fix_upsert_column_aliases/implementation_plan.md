# 企業情報シート出力時の列特定ロジック修正

## 課題
「テナント_顧客ID」が空欄になる問題が発生している。これはヘッダ名の揺らぎ（エイリアス）が原因と考えられる。
また、「顧客ID」列（HotProfile由来）と「テナント_顧客ID」列が混同されるリスクがあるため、厳密なエイリアス定義が必要。

## 解決策
`upsertClients` 関数内で列特定を行う際、エイリアス対応を行いつつ、「テナント_顧客ID」と「顧客ID」を厳密に区別する。

### 変更点: `企業情報.js` / `upsertClients`

1. `headerAliases` の定義を追加。
   - **テナント_顧客ID**: `["テナント_顧客ID", "テナント顧客ID", "顧客ID（テナント）", "TenantCustomerId"]`
     - ※ "顧客ID" は含めない（HotProfile由来の列と区別するため）
   - **顧客ID**: `["顧客ID", "会社コード", "メンテナンス用(会社コード)"]`
   - **レコード番号**: `["レコード番号", "レコードID", "id"]`

2. ヘッダ列の探索ロジックをエイリアス検索に変更する。

## 実装イメージ

```javascript
  // ヘッダ名正規化マップ (trimのみ)
  const h2i = {};
  headers.forEach((h, i) => { h2i[String(h || "").trim()] = i; });

  // 柔軟な検索ヘルパー
  const findColIdx = (candidates) => {
    for (const c of candidates) {
      if (h2i[c] !== undefined) return h2i[c];
    }
    return undefined;
  };

  // テナント_顧客ID（"顧客ID"単体は含めない）
  const idxTenantCustomerId = findColIdx(["テナント_顧客ID", "テナント顧客ID", "顧客ID（テナント）", "TenantCustomerId"]);
  
  // 顧客ID（HotProfile由来）
  const idxCustomerId = findColIdx(["顧客ID", "会社コード"]);
  
  const idxRecordNo = findColIdx(["レコード番号", "レコードID", "id"]);
```

## 検証
変更後、再実行して以下を確認する：
1. 「テナント_顧客ID」列に正しい値が入っていること。
2. 「顧客ID」列（存在する場合）に HotProfile の会社コードが入っていること。
3. 両者が混同されていないこと。

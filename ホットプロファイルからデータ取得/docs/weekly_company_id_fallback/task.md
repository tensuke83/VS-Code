# 週次リセット時の会社ID変更対応

- [x] テナントマスタから「ご契約者_法人名」を取得する処理の追加 (`getTenantMasterData`)
- [x] 会社名でHotProfileの企業情報を検索するAPI関数の実装 (`fetchCompanyByName`)
- [x] `updateCompanyListStepSingleIdMode` にて、取得失敗時に会社名で再検索するフォールバック処理の組み込み

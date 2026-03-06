# 週次リセット時の会社ID変更への対応

## Goal
週次のリセットによる全件取得（`updateCompanyListStepSingleIdMode`）において、会社IDが変更されてしまい取得が失敗しているケースに対応する。
具体的には、取得に失敗した場合に、テナントマスタの「ご契約者_法人名」列から会社名を取得し、その会社名をキーにして再度HotProfileから会社情報を検索・取得して処理を継続する。

## Proposed Changes

### 企業情報.js

#### `getTenantMasterData` の修正
- ヘッダエイリアスに `companyName: ["ご契約者_法人名", "法人名", "会社名"]` などを追加し、テナントマスタから会社名を取得する。
- 戻り値のオブジェクトに `companyName` を含める。

#### `fetchCompanyByName` の追加 (新規)
- 会社名を引数として受け取り、HotProfile API (`clients/get_entry_list`) を用いて会社名で検索を行う関数を追加する。
- 検索のpayloadは `search: { name: companyName }` (または適切な仕様) とし、最初に見つかった企業情報を返す。

#### `updateCompanyListStepSingleIdMode` の修正
- `fetchCompaniesByIds` で `clients` 配列を取得した後、`null` となっている要素について、テナントから `companyName` を取得し、`fetchCompanyByName` で再取得を試みる。
- 再取得に成功した場合は、その情報を元にUPSERTを行い、新会社IDも更新されるようにする。

## Verification Plan
- スクリプト上のロジックに不整合がないことを確認する。
- 実際の環境で、取得失敗時に会社名検索がフォールバックとして走り、正常に情報が更新されることを（可能であれば）テストする。

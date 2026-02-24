/**
 * 1IDモード専用 完全版スクリプト
 *
 * - 新会社ID 列を 2列目に追加
 * - テナントマスタは Spreadsheet ID をプロパティに保存し openById で参照
 * - 会社ID / 新会社ID の両方をキーに既存行を参照
 * - 範囲モード関連は削除済み
 */

/* ---------------------------
   定数
   --------------------------- */
const PAGE_SIZE = 100;

/**
 * 複合キーを生成する（会社ID + テナント顧客ID）
 * - 同じ会社IDでもテナント顧客IDが異なれば別レコードとして扱うためのキー
 * @param {number|string} companyId - 会社ID
 * @param {string} tenantCustomerId - テナント顧客ID（先頭の ' は正規化時に除去）
 * @returns {string} 複合キー（例: "12345_ABC"）
 */
function generateCompositeKey(companyId, tenantCustomerId) {
  const normalizedCompanyId = String(companyId || "").trim();
  // テナント顧客IDの先頭 ' を除去して正規化
  let normalizedTenantId = String(tenantCustomerId || "").trim();
  if (normalizedTenantId.startsWith("'")) {
    normalizedTenantId = normalizedTenantId.substring(1);
  }
  return `${normalizedCompanyId}_${normalizedTenantId}`;
}

/**
 * 単一の会社IDで企業情報を取得する
 * @param {number} companyId - 取得対象の会社ID
 * @returns {Object|null} - 会社情報オブジェクト、または見つからない場合null
 */
function fetchCompanyById(companyId) {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error("fetchCompanyById: API_KEY が未設定です");

  const url = "https://hammock.hot-profile.com/rest_api/v1/clients/get_entry_list";
  const payload = {
    api_key: apiKey,
    search: { from_id: companyId, to_id: companyId },
    page: { display_number: 1, number: 1 }
  };

  try {
    const res = fetchWithRetry(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload)
    });

    const status = res.getResponseCode();
    const body = res.getContentText();

    if (status >= 200 && status < 300) {
      const data = JSON.parse(body);
      const clients = Array.isArray(data.clients) ? data.clients : [];
      return clients[0] || null;
    } else {
      Logger.log(`fetchCompanyById HTTP error (id=${companyId}): ${status}`);
      return null;
    }
  } catch (e) {
    Logger.log(`fetchCompanyById error (id=${companyId}): ${e}`);
    return null;
  }
}

/**
 * 会社名で企業情報を検索する（会社ID変更時のフォールバック用）
 * @param {string} companyName - 検索対象の会社名
 * @returns {Object|null} - 会社情報オブジェクト、または見つからない場合null
 */
function fetchCompanyByName(companyName) {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error("fetchCompanyByName: API_KEY が未設定です");
  if (!companyName || String(companyName).trim() === "") return null;

  const url = "https://hammock.hot-profile.com/rest_api/v1/clients/get_entry_list";
  const payload = {
    api_key: apiKey,
    search: { name: companyName },
    page: { display_number: 10, number: 1 }
  };

  try {
    const res = fetchWithRetry(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload)
    });

    const status = res.getResponseCode();
    const body = res.getContentText();

    if (status >= 200 && status < 300) {
      const data = JSON.parse(body);
      const clients = Array.isArray(data.clients) ? data.clients : [];
      if (clients.length === 0) {
        Logger.log(`fetchCompanyByName: 該当なし (name=${companyName})`);
        return null;
      }
      // 完全一致を優先、なければ先頭を返す
      const exact = clients.find(c => c && c.name === companyName);
      return exact || clients[0];
    } else {
      Logger.log(`fetchCompanyByName HTTP error (name=${companyName}): ${status}`);
      return null;
    }
  } catch (e) {
    Logger.log(`fetchCompanyByName error (name=${companyName}): ${e}`);
    return null;
  }
}

/**
 * 指定範囲・ページの企業一覧を取得する
 * @param {number} fromId - 開始会社ID
 * @param {number} toId - 終了会社ID
 * @param {number} page - ページ番号
 * @param {string} fromDate - 開始日時 (yyyy-MM-dd HH:mm:ss)
 * @param {string} toDate - 終了日時 (yyyy-MM-dd HH:mm:ss)
 * @returns {{clients: Array, count: number}} - 取得したクライアント配列と件数
 */
function fetchCompanyPage(fromId, toId, page, fromDate, toDate) {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error("fetchCompanyPage: API_KEY が未設定です");

  const url = "https://hammock.hot-profile.com/rest_api/v1/clients/get_entry_list";
  const payload = {
    api_key: apiKey,
    search: {
      from_id: fromId,
      to_id: toId,
      from_datetime_updated_on: fromDate,
      to_datetime_updated_on: toDate
    },
    page: { display_number: PAGE_SIZE, number: page }
  };

  try {
    const res = fetchWithRetry(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload)
    });

    const status = res.getResponseCode();
    const body = res.getContentText();

    if (status >= 200 && status < 300) {
      const data = JSON.parse(body);
      const clients = Array.isArray(data.clients) ? data.clients : [];
      return { clients: clients, count: clients.length };
    } else {
      Logger.log(`fetchCompanyPage HTTP error: ${status}`);
      return { clients: [], count: 0 };
    }
  } catch (e) {
    Logger.log(`fetchCompanyPage error: ${e}`);
    return { clients: [], count: 0 };
  }
}

/* ---------------------------
   プロパティ管理: APIキー / テナントマスタID
   --------------------------- */
function setApiKey() {
  PropertiesService.getScriptProperties().setProperty("API_KEY", "ここにAPIキーを入力");
}
function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty("API_KEY");
}

function setTenantMasterSheetId() {
  PropertiesService.getScriptProperties().setProperty("TENANT_MASTER_SHEET_ID", "ここにスプレッドシートIDを入力");
}
function getTenantMasterSheetId() {
  return PropertiesService.getScriptProperties().getProperty("TENANT_MASTER_SHEET_ID");
}

/* ---------------------------
   シート確保（ヘッダ強化版）
   --------------------------- */
function ensureOutputSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("企業情報");

  // ユーザー要望: ラストタッチは「テナント_顧客ID」の後、ラストタッチURLはその次
  const expectedHeader = [
    "会社ID", "新会社ID", "会社名", "都道府県", "住所",
    "業種_大分類", "業種_中分類", "業種_小分類", "業種_細分類",
    "従業員規模", "上場区分", "資本金", "法人番号", "顧客ID",
    "レコード番号", "テナント_顧客ID", "ラストタッチ", "ラストタッチURL", "取得日時", "更新日時"
  ];

  if (!sheet) {
    sheet = ss.insertSheet("企業情報");
    sheet.appendRow(expectedHeader);
    return sheet;
  }

  // シート存在時: ラストタッチ列とラストタッチURL列の位置調整
  const lastCol = sheet.getLastColumn();
  if (lastCol > 0) {
    const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    const h2i = {};
    currentHeaders.forEach((h, i) => { h2i[h] = i + 1; });

    const colTenant = h2i["テナント_顧客ID"];
    let colLastTouch = h2i["ラストタッチ"];
    let colLastTouchUrl = h2i["ラストタッチURL"];

    if (colTenant) {
      // ラストタッチ列がなければテナントIDの後に挿入
      if (!colLastTouch) {
        sheet.insertColumnAfter(colTenant);
        sheet.getRange(1, colTenant + 1).setValue("ラストタッチ");
        colLastTouch = colTenant + 1;
      } else if (colLastTouch !== colTenant + 1) {
        // 位置がテナントID + 1 でなければ移動
        const range = sheet.getRange(1, colLastTouch, sheet.getMaxRows(), 1);
        sheet.moveColumns(range, colTenant + 1);
        colLastTouch = colTenant + 1;
      }

      // ラストタッチURL列がなければラストタッチの後に挿入
      // h2i を再取得して最新の列位置を反映
      const updatedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
      const h2iUpdated = {};
      updatedHeaders.forEach((h, i) => { h2iUpdated[h] = i + 1; });
      colLastTouch = h2iUpdated["ラストタッチ"];
      colLastTouchUrl = h2iUpdated["ラストタッチURL"];

      if (!colLastTouchUrl && colLastTouch) {
        sheet.insertColumnAfter(colLastTouch);
        sheet.getRange(1, colLastTouch + 1).setValue("ラストタッチURL");
      } else if (colLastTouchUrl && colLastTouch && colLastTouchUrl !== colLastTouch + 1) {
        // 位置がラストタッチ + 1 でなければ移動
        const range = sheet.getRange(1, colLastTouchUrl, sheet.getMaxRows(), 1);
        sheet.moveColumns(range, colLastTouch + 1);
      }
    }
  }

  return sheet;
}

/**
 * ログシート「取得ログ」を確保する（存在しなければ作成）
 * @returns {Sheet} ログシート
 */
function ensureLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("取得ログ");
  if (!sheet) {
    sheet = ss.insertSheet("取得ログ");
    // 整理後のヘッダ構成
    sheet.appendRow([
      "日時", "操作", "会社ID", "会社名",
      "更新日時", "進捗状況"
    ]);
  }
  return sheet;
}

/**
 * 設定_列固定シートから対象シートのヘッダを取得して設定する
 * @param {string} sheetName - 出力対象シート名
 */
function ensureOutputSheetFromSettings(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const settingsSheet = ss.getSheetByName("設定_列固定");
  if (!settingsSheet) throw new Error("設定_列固定シートが存在しません");

  const settings = settingsSheet.getDataRange().getValues();
  let headerRow = null;
  for (let i = 1; i < settings.length; i++) { // 1行目はヘッダ
    if (settings[i][0] === sheetName) {
      headerRow = settings[i][1];
      break;
    }
  }
  if (!headerRow) throw new Error("設定_列固定シートにヘッダ定義がありません: " + sheetName);

  const headers = headerRow.split(",");
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  return sheet;
}

/* ---------------------------
   業種取得（配列は先頭要素のみ使用）
   --------------------------- */
function getIndustryPartsFromClient(client) {
  const raw = client.jsic_business_category_kbn;

  // 配列の場合 → 先頭の有効文字列だけを利用
  if (Array.isArray(raw)) {
    const first = raw.find(s => typeof s === "string" && s.trim().length > 0);
    return parseIndustry(first || "");
  }

  // 文字列の場合 → そのまま処理
  if (typeof raw === "string" && raw.trim().length > 0) {
    return parseIndustry(raw);
  }

  // それ以外 → 空を返す
  return { large: "", medium: "", small: "", detail: "" };
}

/* ---------------------------
   業種パース（複数候補保持版）
   --------------------------- */
function parseIndustry(industryRaw) {
  if (!industryRaw || typeof industryRaw !== "string") {
    return { large: "", medium: "", small: "", detail: "" };
  }

  // 区切り記号を全角に正規化
  const normalized = industryRaw.replace(/>/g, "＞");

  // ＞で階層分割（カンマや読点は保持）
  const parts = normalized.split("＞").map(p => p.trim());

  return {
    large: parts[0] || "",
    medium: parts[1] || "",
    small: parts[2] || "",
    detail: parts[3] || ""
  };
}

/**
 * テナントマスタ（シート名「テナントマスタ」）を見出し名で読み取る修正版。
 * - 列位置ではなく「見出し名」を基準に列を特定する
 * - 「ホットプロファイル会社ID」を必須キーとして companyId に格納
 * - 「テナント顧客ID」は先頭ゼロ保持のため "'" を付けて文字列化
 * - ★ テナントマスタの「id」列の値を recordNo として返す（そのまま出力シートの「レコード番号」列へ出力される）
 *
 * 返却：[{ recordNo:string, tenantCustomerId:string, companyId:number }, ...]
 *  - recordNo …… テナントマスタの「id」列の値（文字列のまま）
 *  - tenantCustomerId …… テナント顧客ID（先頭ゼロ保持のため "'" 付与）
 *  - companyId …… ホットプロファイル会社ID（数値）
 */
function getTenantMasterData() {
  // ---- 0) スプレッドシート取得 ----
  const sheetId = getTenantMasterSheetId();
  if (!sheetId) throw new Error("TENANT_MASTER_SHEET_ID が未設定です");
  const ss = SpreadsheetApp.openById(sheetId);

  // ---- 1) 対象シート ----
  const sheetName = "テナントマスタ";
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート名 '${sheetName}' が見つかりません`);

  // ---- 2) ヘッダ行 ----
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; // データなし
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());

  // ---- 3) 見出しの同義語（柔軟に対応） ----
  const headerAliases = {
    // ★ recordNo は「id」列を参照する
    recordNo: ["id", "ID", "レコードID", "テナントID"],

    // 文字列として保持（先頭ゼロ保持のため出力時に "'" を付ける）
    tenantCustomerId: ["テナント_顧客ID", "テナント顧客ID", "顧客ID（テナント）", "TenantCustomerId", "顧客ID"],

    // companyId は HotProfile の会社ID（必須キー）
    companyId: ["ホットプロファイル会社ID", "会社ID", "HPF会社ID", "HotProfile会社ID", "client.id"],

    // 契約ステータス（「解約済み」を除外するため）
    contractStatus: ["契約ステータス", "ステータス", "契約状態"],

    // 会社名（会社ID変更時のフォールバック検索用）
    companyName: ["ご契約者_法人名", "法人名", "会社名"]
  };

  // ---- 4) 見出し→列番号の解決 ----
  const resolveColumnIndex = (aliases) => {
    for (const label of aliases) {
      const idx = headerRow.indexOf(label);
      if (idx >= 0) return idx + 1; // 1始まり
    }
    return -1;
  };

  const colRecordNo = resolveColumnIndex(headerAliases.recordNo);                // ★ 必須（id列）
  const colTenantCustomerId = resolveColumnIndex(headerAliases.tenantCustomerId); // 任意
  const colCompanyId = resolveColumnIndex(headerAliases.companyId);              // ★ 必須
  const colContractStatus = resolveColumnIndex(headerAliases.contractStatus);    // 任意（解約済み除外用）
  const colCompanyName = resolveColumnIndex(headerAliases.companyName);            // 任意（フォールバック用）

  if (colRecordNo < 1) {
    const msg = [
      "テナントマスタの見出しに『id』列が見つかりません。",
      "見出し名候補：",
      `- ${headerAliases.recordNo.join(" / ")}`
    ].join("\n");
    throw new Error(msg);
  }

  if (colCompanyId < 1) {
    const msg = [
      "テナントマスタの見出しに『ホットプロファイル会社ID』が見つかりません。",
      "見出し名候補：",
      `- ${headerAliases.companyId.join(" / ")}`
    ].join("\n");
    throw new Error(msg);
  }

  // ---- 5) データ読み取り ----
  const numRows = lastRow - 1;
  const recordNos = sheet.getRange(2, colRecordNo, numRows, 1).getValues();                  // ★ id列（必須）
  const tenantIds = (colTenantCustomerId > 0) ? sheet.getRange(2, colTenantCustomerId, numRows, 1).getValues() : null;
  const companyIds = sheet.getRange(2, colCompanyId, numRows, 1).getValues();                 // 必須
  const contractStatuses = (colContractStatus > 0) ? sheet.getRange(2, colContractStatus, numRows, 1).getValues() : null; // 契約ステータス
  const companyNames = (colCompanyName > 0) ? sheet.getRange(2, colCompanyName, numRows, 1).getValues() : null; // 会社名（フォールバック用）

  const out = [];
  for (let i = 0; i < numRows; i++) {
    const recordNoVal = recordNos[i][0];         // ★ 文字列のまま保持
    const tenantIdVal = tenantIds ? tenantIds[i][0] : undefined;
    const companyIdVal = companyIds[i][0];

    // 契約ステータスが「解約済み」の場合はスキップ
    if (contractStatuses) {
      const statusVal = String(contractStatuses[i][0] || "").trim();
      if (statusVal === "解約済み") {
        continue;
      }
    }

    // companyId を数値に
    const nCompanyId = Number(companyIdVal);
    if (isNaN(nCompanyId) || nCompanyId <= 0) {
      // 無効な会社IDはスキップ
      continue;
    }

    // テナント顧客IDは先頭ゼロ維持のため "'" を付けて文字列化
    const tenantCustomerIdStr =
      (tenantIdVal === null || tenantIdVal === undefined || String(tenantIdVal).trim() === "")
        ? ""
        : "'" + String(tenantIdVal);

    // ★ recordNo はテナントマスタの id 列そのまま（空なら空文字）
    const recordNoStr =
      (recordNoVal === null || recordNoVal === undefined) ? "" : String(recordNoVal);

    // 会社名（フォールバック検索用）
    const companyNameStr = companyNames
      ? String(companyNames[i][0] || "").trim()
      : "";

    out.push({
      recordNo: recordNoStr,             // ← upsertClients() が「レコード番号」列に書き込む
      tenantCustomerId: tenantCustomerIdStr,
      companyId: nCompanyId,
      companyName: companyNameStr          // ★ 会社ID変更時フォールバック用
    });
  }

  return out;
}

/**
 * HotProfileの会社ID配列を受け取り、各IDの会社情報（clients[0]）を並列で取得して返す修正版。
 * - UrlFetchApp.fetchAll() を使用して並列取得
 * - 同時接続数（CONCURRENCY）でチャンク分割し、サーバー負荷を抑制
 *
 * @param {number[]} ids - HotProfileの会社ID配列
 * @returns {(Object|null)[]} - 各IDに対応する client オブジェクト（見つからない/失敗時は null）
 */
function fetchCompaniesByIds(ids) {
  // ---- 0) 前提チェック -------------------------------------------------------
  if (!Array.isArray(ids)) throw new Error(`fetchCompaniesByIds: ids is not array`);
  if (ids.length === 0) return [];

  const apiKey = getApiKey();
  if (!apiKey) throw new Error("fetchCompaniesByIds: API_KEY missing");

  const url = "https://hammock.hot-profile.com/rest_api/v1/clients/get_entry_list";

  // ★設定：並列数と待機時間（APIレート制限: 5分間1000回 = 1分200回）
  const CONCURRENCY = 10;      // 同時接続数（レート制限内で最大スループット）
  const CHUNK_SLEEP_MS = 100;  // チャンク間の待機時間（100ms × 250チャンク = 25秒）

  const resultsMap = new Map(); // id -> client

  // ---- 1) IDリストをチャンクに分割して処理 -----------------------------------
  for (let i = 0; i < ids.length; i += CONCURRENCY) {
    const chunkIds = ids.slice(i, i + CONCURRENCY);

    // リクエスト作成
    const requests = chunkIds.map(id => {
      const payload = {
        api_key: apiKey,
        search: { from_id: id, to_id: id },
        page: { display_number: 1, number: 1 }
      };
      return {
        url: url,
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
    });

    try {
      // ★並列実行
      const responses = UrlFetchApp.fetchAll(requests);

      // レスポンス処理
      responses.forEach((res, idx) => {
        const id = chunkIds[idx];
        const status = res.getResponseCode();
        const body = res.getContentText();

        if (status >= 200 && status < 300) {
          try {
            const data = JSON.parse(body);
            const clients = Array.isArray(data.clients) ? data.clients : [];
            resultsMap.set(id, clients[0] || null);
          } catch (e) {
            Logger.log(`JSON parse error (id=${id}): ${e}`);
            resultsMap.set(id, null);
          }
        } else {
          Logger.log(`HTTP error (id=${id}) status=${status} body=${truncateString(body, 200)}`);
          resultsMap.set(id, null);
        }
      });

    } catch (e) {
      Logger.log(`fetchAll error at chunk start index ${i}: ${e}`);
      // 失敗したチャンクのIDはすべてnullにする
      chunkIds.forEach(id => resultsMap.set(id, null));
    }

    // チャンク間のスロットリング（最後のチャンク以外）
    if (i + CONCURRENCY < ids.length) {
      Utilities.sleep(CHUNK_SLEEP_MS);
    }
  }

  // ---- 2) 元の順序で配列にして返す --------------------------------------------
  return ids.map(id => resultsMap.get(id) || null);
}

/* ---------------------------
   既存データ読み込み（レコード番号優先、複合キーはフォールバック）
   - レコード番号がある行は REC:{recordNo} をキーとする
   - レコード番号がない行は従来の複合キー（会社ID + テナント顧客ID）を使用
   --------------------------- */
function buildExistingDataMapByHeader(sheet) {
  const map = new Map();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return map;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const colCompanyId = headers.indexOf("会社ID") + 1;
  const colNewCompanyId = headers.indexOf("新会社ID") + 1;
  const colUpdatedAt = headers.indexOf("更新日時") + 1;
  const colTenantCustomerId = headers.indexOf("テナント_顧客ID") + 1;
  const colRecordNo = headers.indexOf("レコード番号") + 1;

  if (colCompanyId < 1) throw new Error("企業情報シートに '会社ID' 列がありません");
  if (colNewCompanyId < 1) throw new Error("企業情報シートに '新会社ID' 列がありません");
  if (colUpdatedAt < 1) throw new Error("企業情報シートに '更新日時' 列がありません");

  const numRows = lastRow - 1;
  const companyIds = sheet.getRange(2, colCompanyId, numRows, 1).getValues();
  const newCompanyIds = sheet.getRange(2, colNewCompanyId, numRows, 1).getValues();
  const updatedAts = sheet.getRange(2, colUpdatedAt, numRows, 1).getValues();
  const tenantCustomerIds = (colTenantCustomerId > 0)
    ? sheet.getRange(2, colTenantCustomerId, numRows, 1).getValues()
    : null;
  const recordNos = (colRecordNo > 0)
    ? sheet.getRange(2, colRecordNo, numRows, 1).getValues()
    : null;

  for (let i = 0; i < numRows; i++) {
    const rowIndex = i + 2;
    const id1 = Number(companyIds[i][0]);
    const id2 = Number(newCompanyIds[i][0]);
    const updatedAtVal = updatedAts[i][0];
    const updatedAtMs = toEpochMs(updatedAtVal);
    const tenantId = tenantCustomerIds ? String(tenantCustomerIds[i][0] || "") : "";
    const recordNo = recordNos ? String(recordNos[i][0] || "").trim() : "";

    const rowData = { rowIndex, updatedAtMs, tenantCustomerId: tenantId, recordNo };

    // レコード番号がある場合は、それを優先キーとして登録
    if (recordNo !== "") {
      const recordKey = `REC:${recordNo}`;
      map.set(recordKey, rowData);
    } else {
      // レコード番号がない場合は、従来の複合キーで登録（過去データ救済）
      if (!isNaN(id1) && id1 > 0) {
        const compositeKey1 = generateCompositeKey(id1, tenantId);
        map.set(compositeKey1, rowData);
      }
      if (!isNaN(id2) && id2 > 0) {
        const compositeKey2 = generateCompositeKey(id2, tenantId);
        map.set(compositeKey2, rowData);
      }
    }
  }

  return map;
}

/**
 * テナントマスタに載っている会社だけを対象に、差分（clients）をシートへ反映する（UPSERT）。
 *
 * 前提：
 * - 企業情報シートは「設定_列固定」のヘッダ定義に従う（列順が変わってもOK）
 * - existingDataMap は以下の形式：
 *     Map<number, { rowIndex: number, updatedAtMs: number }>
 *   キーは「会社ID」または「新会社ID（= HotProfileのclient.id）」のどちらでも引けるようにしておく
 *
 * 更新判定：
 * - APIの client.updated_at（例: "2025/04/04 15:38:15"）を toEpochMs() で数値化し、
 *   既存行の updatedAtMs より新しい場合のみ UPDATE する
 *
 * @param {Array<Object>} clients HotProfile API の clients 配列
 * @param {Array<{recordNo:any, tenantCustomerId:string, companyId:number}>} tenants テナントマスタ配列
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 出力先シート（企業情報）
 * @param {Map<number, {rowIndex:number, updatedAtMs:number}>} existingDataMap 既存行マップ
 * @param {Set<number>} processedIds 同一実行内の重複処理防止用（responseIdベース）
 * @param {Array<Array<any>>} logEntries ログ追記用の行配列
 * @param {{from:any, to:any, page?:number, requestId?:number}} context 呼び出しコンテキスト（週次/日次で使い分け）
 */
function upsertClients(clients, tenants, sheet, existingDataMap, processedIds, logEntries, context) {
  // --- 1) 対象テナント（会社ID）の集合を作る ---
  const tenantIds = new Set(tenants.map(t => Number(t.companyId)).filter(n => !isNaN(n) && n > 0));

  // 取得日時（シートに書く用）
  const nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");

  // --- 2) ヘッダを取得し、列位置を前計算する（毎レコードで indexOf しない） ---
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // ヘッダ名 → 0始まりIndex へのマップ
  // ヘッダ名 → 0始まりIndex へのマップ
  const h2i = {};
  headers.forEach((h, i) => { h2i[String(h || "").trim()] = i; });

  // 柔軟な検索ヘルパー
  const findColIdx = (candidates) => {
    for (const c of candidates) {
      if (h2i[c] !== undefined) return h2i[c];
    }
    return undefined;
  };

  const idxCompanyId = findColIdx(["会社ID", "ホットプロファイル会社ID"]);
  const idxNewCompanyId = findColIdx(["新会社ID", "client.id"]);
  const idxUpdatedAt = findColIdx(["更新日時", "client.updated_at"]);
  const idxFetchedAt = findColIdx(["取得日時"]);
  const idxLastTouch = findColIdx(["ラストタッチ"]);

  // テナント_顧客ID（"顧客ID"単体は含めない ※HotProfile由来の列と区別）
  const idxTenantCustomerId = findColIdx(["テナント_顧客ID", "テナント顧客ID", "顧客ID（テナント）", "TenantCustomerId"]);

  // 顧客ID（HotProfile由来）
  const idxCustomerId = findColIdx(["顧客ID", "会社コード", "メンテナンス用(会社コード)"]);

  const idxRecordNo = findColIdx(["レコード番号", "レコードID", "id"]);

  // 必須列が無い場合は即エラー（データ破壊を防ぐ）
  if (idxCompanyId === undefined) throw new Error("企業情報ヘッダに '会社ID' がありません");
  if (idxNewCompanyId === undefined) throw new Error("企業情報ヘッダに '新会社ID' がありません");
  if (idxUpdatedAt === undefined) throw new Error("企業情報ヘッダに '更新日時' がありません");
  if (idxFetchedAt === undefined) throw new Error("企業情報ヘッダに '取得日時' がありません");

  // URL列がある場合だけ埋める（設定に無ければ無視）
  const idxHpfUrl = h2i["ホットプロファイルURL"]; // 任意列

  // まとめて追記するためのバッファ
  const rowsToAppend = [];
  const appendMeta = []; // { responseId, primaryCompanyId, updatedAtMs, tenantCustomerId } を保持

  // ★更新された行を追跡する（forceUpdateLastTouch 用）
  const updatedRows = [];

  // --- 3) clients を1件ずつ UPSERT ---
  clients.forEach(client => {
    // HotProfileの会社ID（レスポンス側）
    const responseId = Number(client && client.id);
    if (isNaN(responseId) || responseId <= 0) return;

    // 週次（1IDモード）で requestId が渡る場合がある。日次差分では undefined の想定。
    const requestId = context && context.requestId ? Number(context.requestId) : null;

    // --- 3-1) テナントマスタに載っている会社だけ処理する（運用方針どおり） ---
    // 日次差分：tenantIds.has(responseId) のみで判定される
    // 週次 1ID：requestId が tenantIds に居る場合があるので OR にしておく
    const allowed =
      tenantIds.has(responseId) ||
      (requestId && tenantIds.has(requestId));

    if (!allowed) return;

    // --- 3-2) primaryCompanyId を決める ---
    // 週次1IDモードでは requestId を優先（テナントマスタの companyId が requestId の想定）
    // 日次差分は requestId が無いので responseId になる
    const primaryCompanyId = (requestId && tenantIds.has(requestId)) ? requestId : responseId;

    // --- 3-3) 更新日時の取得と比較用ミリ秒化 ---
    const updatedAtStr = (client && client.updated_at) ? String(client.updated_at) : "";
    const updatedAtMs = toEpochMs(updatedAtStr); // ※ toEpochMs は別途定義（Date/文字列揺れに強い）

    // --- 3-4) 顧客ID（カスタム項目：メンテナンス用(会社コード)） ---
    // 先頭ゼロを守りたいので "'" を付けて文字列として扱う（スプレッドシートの自動数値化回避）
    let customerId = "";
    if (Array.isArray(client.items)) {
      const found = client.items.find(i => i && i.label === "メンテナンス用(会社コード)");
      if (found && found.value !== undefined && found.value !== null && String(found.value).trim() !== "") {
        customerId = "'" + String(found.value);
      }
    }

    // --- 3-5) 業種（階層パース） ---
    // getIndustryPartsFromClient は既存関数を利用（配列/文字列どちらでもOKの実装になっている前提）
    const industryParts = getIndustryPartsFromClient(client);

    // --- 3-6) テナント情報（レコード番号/テナント顧客ID）を取得 ---
    // ★修正: find → filter に変更し、同じ会社IDに紐づく「すべての」テナントレコードを取得
    const matchingTenants = tenants.filter(t =>
      Number(t.companyId) === Number(responseId) ||
      Number(t.companyId) === Number(primaryCompanyId)
    );

    // マッチするテナントがない場合はスキップ（通常はここには来ないはず）
    if (matchingTenants.length === 0) return;

    // --- 3-7) 各テナントレコードに対してループ処理（重複会社ID対応） ---
    matchingTenants.forEach(tenantInfo => {
      // ★重複処理チェック: レコード番号がある場合はそれを優先、ない場合は複合キー
      const dedupeKey = tenantInfo.recordNo
        ? `PROC:${responseId}_${tenantInfo.recordNo}`
        : generateCompositeKey(responseId, tenantInfo.tenantCustomerId);
      if (processedIds.has(dedupeKey)) return;
      processedIds.add(dedupeKey);

      // --- 行データを組み立てる ---
      const rowData = new Array(headers.length).fill("");

      // よく使うフィールドは先に埋めておく（indexで代入する方が速い）
      rowData[idxCompanyId] = primaryCompanyId;
      rowData[idxNewCompanyId] = responseId;
      rowData[idxFetchedAt] = nowStr;
      rowData[idxUpdatedAt] = updatedAtStr; // 表示はAPIの文字列のまま（比較は updatedAtMs で行う）

      // テナント情報を直接設定
      if (idxTenantCustomerId !== undefined) {
        rowData[idxTenantCustomerId] = tenantInfo.tenantCustomerId !== undefined ? tenantInfo.tenantCustomerId : "";
      }
      if (idxRecordNo !== undefined) {
        rowData[idxRecordNo] = tenantInfo.recordNo !== undefined ? tenantInfo.recordNo : "";
      }

      // ★ラストタッチ更新は startLastTouchStep (並列処理) に委譲
      // パフォーマンス最適化のため、ここでの同期API呼び出しは削除
      // idxLastTouch は空のままにし、後続の updateLastTouchStep で更新される

      // 任意列：URL（設定にある時だけ）
      if (idxHpfUrl !== undefined) {
        rowData[idxHpfUrl] = "https://005108.hammock.hot-profile.com/clients/" + primaryCompanyId;
      }

      // それ以外の列はヘッダ名で埋める（列追加にも対応しやすい）
      headers.forEach((header, i) => {
        if (rowData[i] !== "") return; // すでに埋めた列はスキップ

        switch (String(header).trim()) {
          case "会社名":
            rowData[i] = (client && client.name) ? client.name : "";
            break;
          case "都道府県":
            rowData[i] = (client && client.pref) ? client.pref : "";
            break;
          case "住所":
            rowData[i] = (client && client.address) ? client.address : "";
            break;

          case "業種_大分類":
            rowData[i] = industryParts.large || "";
            break;
          case "業種_中分類":
            rowData[i] = industryParts.medium || "";
            break;
          case "業種_小分類":
            rowData[i] = industryParts.small || "";
            break;
          case "業種_細分類":
            rowData[i] = industryParts.detail || "";
            break;

          case "従業員規模":
            rowData[i] = (client && client.worker_number_kbn) ? client.worker_number_kbn : "";
            break;
          case "上場区分":
            rowData[i] = (client && client.ipo_kbn) ? client.ipo_kbn : "";
            break;
          case "資本金":
            rowData[i] = (client && client.capital_kbn) ? client.capital_kbn : "";
            break;
          case "法人番号":
            rowData[i] = (client && client.corporate_number) ? client.corporate_number : "";
            break;

          case "法人番号":
            rowData[i] = (client && client.corporate_number) ? client.corporate_number : "";
            break;

          // 顧客ID（HotProfile由来）
          case "顧客ID":
          case "会社コード":
          case "メンテナンス用(会社コード)":
            // すでに idxCustomerId で特定されている場合は rowData[idxCustomerId] に直接代入してもよいが
            // ここではヘッダ文字列マッチで汎用的に入るようにしておく
            if (i === idxCustomerId) rowData[i] = customerId;
            break;

          // レコード番号（テナントマスタ由来）
          case "レコード番号":
          case "レコードID":
          case "id":
            if (i === idxRecordNo) rowData[i] = tenantInfo && tenantInfo.recordNo !== undefined ? tenantInfo.recordNo : "";
            break;

          // テナント_顧客ID（テナントマスタ由来）
          case "テナント_顧客ID":
          case "テナント顧客ID":
          case "顧客ID（テナント）":
          case "TenantCustomerId":
            if (i === idxTenantCustomerId) rowData[i] = tenantInfo && tenantInfo.tenantCustomerId !== undefined ? tenantInfo.tenantCustomerId : "";
            break;

          // レコード番号とテナント_顧客ID は既に埋めたのでスキップ
          // 取得日時/更新日時/会社ID/新会社ID/URL も既に埋めたのでここでは何もしない
          default:
            // 未定義列は空のまま（設定シート側で増やしても壊れない）
            break;
        }
      });

      // --- 既存行の探索（レコード番号優先、複合キーはフォールバック） ---
      let existing = null;
      if (tenantInfo.recordNo) {
        // レコード番号がある場合は、それで検索
        const recordKey = `REC:${tenantInfo.recordNo}`;
        existing = existingDataMap.get(recordKey);
      }
      if (!existing) {
        // レコード番号がないか、見つからない場合は複合キーで検索（フォールバック）
        const compositeKeyForLookup = generateCompositeKey(responseId, tenantInfo.tenantCustomerId);
        const compositeKeyForLookup2 = generateCompositeKey(primaryCompanyId, tenantInfo.tenantCustomerId);
        existing = existingDataMap.get(compositeKeyForLookup) || existingDataMap.get(compositeKeyForLookup2);
      }

      // --- UPDATE か ADD かを判定 ---
      if (existing) {
        // 更新日時が新しい場合だけ上書き（同じ/古い場合は何もしない）
        const oldMs = existing.updatedAtMs || 0;
        if (updatedAtMs > oldMs) {
          // 単一行更新（差分件数が多い場合はここがボトルネックになり得るが、日次差分なら現実的）
          sheet.getRange(existing.rowIndex, 1, 1, rowData.length).setValues([rowData]);

          // ★更新された行を記録（ラストタッチ再取得用）
          updatedRows.push({
            rowIndex: existing.rowIndex,
            companyName: (client && client.name) ? client.name : "",
            companyId: responseId,
            oldCompanyId: primaryCompanyId !== responseId ? primaryCompanyId : undefined
          });

          // マップを更新（レコード番号優先で登録）
          const newEntry = { rowIndex: existing.rowIndex, updatedAtMs, tenantCustomerId: tenantInfo.tenantCustomerId, recordNo: tenantInfo.recordNo };
          if (tenantInfo.recordNo) {
            const recordKey = `REC:${tenantInfo.recordNo}`;
            existingDataMap.set(recordKey, newEntry);
          } else {
            // レコード番号がない場合は複合キーで登録
            const compositeKeyForLookup = generateCompositeKey(responseId, tenantInfo.tenantCustomerId);
            const compositeKeyForLookup2 = generateCompositeKey(primaryCompanyId, tenantInfo.tenantCustomerId);
            existingDataMap.set(compositeKeyForLookup, newEntry);
            existingDataMap.set(compositeKeyForLookup2, newEntry);
          }

          // ログ (新構成: ["日時", "操作", "会社ID", "会社名", "更新日時", "進捗状況"])
          logEntries.push([
            nowStr,
            "UPDATE",
            primaryCompanyId,
            (client && client.name) ? client.name : "",
            updatedAtStr,
            "", // 進捗状況はバッチ単位でappendRunLogが処理
          ]);
        } else {
          // 何もしない（ログに残したいなら "SKIP" を追加してもよい）
        }

      } else {
        // 新規行として追記（まとめて setValues するため配列に貯める）
        rowsToAppend.push(rowData);
        appendMeta.push({ responseId, primaryCompanyId, updatedAtMs, tenantCustomerId: tenantInfo.tenantCustomerId, recordNo: tenantInfo.recordNo });

        // ログ (新構成: ["日時", "操作", "会社ID", "会社名", "更新日時", "進捗状況"])
        logEntries.push([
          nowStr,
          "ADD",
          primaryCompanyId,
          (client && client.name) ? client.name : "",
          updatedAtStr,
          "", // 進捗状況はバッチ単位でappendRunLogが処理
        ]);
      }
    }); // matchingTenants.forEach
  });

  // --- 4) 追記行をまとめて書く（I/O削減） ---
  if (rowsToAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);

    // 追記した行の rowIndex と updatedAtMs を既存マップに登録（複合キーで）
    for (let i = 0; i < appendMeta.length; i++) {
      const rowIndex = startRow + i;
      const meta = appendMeta[i];

      const newEntry = { rowIndex, updatedAtMs: meta.updatedAtMs, tenantCustomerId: meta.tenantCustomerId, recordNo: meta.recordNo };
      if (meta.recordNo) {
        // レコード番号がある場合はそれで登録
        const recordKey = `REC:${meta.recordNo}`;
        existingDataMap.set(recordKey, newEntry);
      } else {
        // レコード番号がない場合は複合キーで登録
        const compositeKey1 = generateCompositeKey(meta.primaryCompanyId, meta.tenantCustomerId);
        const compositeKey2 = generateCompositeKey(meta.responseId, meta.tenantCustomerId);
        existingDataMap.set(compositeKey1, newEntry);
        existingDataMap.set(compositeKey2, newEntry);
      }
    }
  }

  // ★更新された行のリストを返す（呼び出し元でラストタッチ再取得に使用）
  return updatedRows;
}

/**
 * ログエントリを追記する関数
 * @param {Array} logEntries - [["日時","操作","会社ID","会社名","レコード番号","顧客ID","from","to","page","更新日時","備考"], ...] の形式
 */
function appendLogEntries(logEntries) {
  if (!logEntries || logEntries.length === 0) return;
  const sheet = ensureLogSheet();

  // 受け取った logEntries をそのまま行として追加
  sheet.getRange(sheet.getLastRow() + 1, 1, logEntries.length, logEntries[0].length).setValues(logEntries);
}


/**
 * 週次ステップ（1IDモード）実行の本体：スクリプトプロパティ上のカーソルから
 * テナントマスタの会社IDをバッチ分割して順次取得・UPSERTする。
 *
 * 帯域超過（Bandwidth quota exceeded）対策：
 *  - fetchCompaniesByIds(ids) を直列＋指数バックオフ版に差し替え前提
 *  - バッチサイズをプロパティ SINGLE_BATCH（既定50）で制御し、過負荷を回避
 *  - nullレスポンス（失敗/未取得）を "MISS" としてログに記録し処理継続
 *
 * 通知＆ログ：
 *  - 取得ログ（ensureLogSheet）に ADD/UPDATE/MISS を追記
 *  - ステップ完了時に safeNotifySummary() で Teams 通知
 *  - 異常時はトリガー停止＋エラー通知（スタックを truncateString で安全送信）
 *
 * 依存関数（既存のものを利用）：
 *  - ensureOutputSheetFromSettings, ensureLogSheet, getApiKey, getTenantMasterData
 *  - buildExistingDataMapByHeader, upsertClients, appendLogEntries
 *  - safeNotifySummary, stopStepTrigger, reapplyFormulasFromSettings
 *  - appendRunLog, truncateString
 */
function updateCompanyListStepSingleIdMode() {
  // ---- 0) 単一実行制御（並行実行の競合防止） --------------------------------
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { // 30秒で取得できない場合はスキップ
    Logger.log("別実行中のため中断（single id mode）");
    appendRunLog("SKIP", "Lock 競合（single）");
    return;
  }

  // 時刻などの文脈（通知用）
  const startTime = new Date();
  const props = PropertiesService.getScriptProperties();

  try {
    // ---- 1) 前提チェック：APIキー＆テナントマスタ ----------------------------
    const apiKey = getApiKey();
    if (!apiKey) {
      // ここでトリガーを止める（無限エラー防止）
      stopStepTrigger("updateCompanyListStepSingleIdMode");
      safeNotifySummary({
        summaries: ["API_KEY が未設定です"],
        startTime,
        endTime: new Date(),
        durationSec: 0,
        isError: true,
        mode: 'weekly'
      });
      appendRunLog("ERROR", "API_KEY 未設定（single）");
      return;
    }

    let tenants = [];
    try {
      tenants = getTenantMasterData(); // [{recordNo, tenantCustomerId, companyId}, ...]
    } catch (e) {
      stopStepTrigger("updateCompanyListStepSingleIdMode");
      safeNotifySummary({
        summaries: ["テナントマスタ読取に失敗: " + e],
        startTime,
        endTime: new Date(),
        durationSec: 0,
        isError: true,
        mode: 'weekly'
      });
      appendRunLog("ERROR", "テナントマスタ読取例外（single）: " + e);
      return;
    }
    if (!Array.isArray(tenants) || tenants.length === 0) {
      stopStepTrigger("updateCompanyListStepSingleIdMode");
      safeNotifySummary({
        summaries: ["テナントマスタが空です"],
        startTime,
        endTime: new Date(),
        durationSec: 0,
        isError: true,
        mode: 'weekly'
      });
      appendRunLog("NO_DATA", "テナントマスタが空（single）");
      return;
    }

    // ---- 2) 出力シート＆ログシートの保証 -------------------------------------
    const sheet = ensureOutputSheetFromSettings("企業情報"); // 列定義は「設定_列固定」参照
    ensureLogSheet("取得ログ");

    // ---- 3) 既存データマップの準備 -------------------------------------------
    const existing = buildExistingDataMapByHeader(sheet); // Map<会社ID or 新会社ID, {rowIndex, updatedAtMs}>
    const processedIds = new Set();                      // 実行内の重複処理防止
    const logEntries = [];                               // まとめてログ投入

    // ---- 4) スクリプトプロパティからカーソル／バッチを取得 -------------------
    let cursor = Number(props.getProperty("ID_CURSOR")) || 0;
    // バッチサイズはプロパティ優先。未設定なら安全側の既定50に。
    let BATCH = Number(props.getProperty("SINGLE_BATCH"));
    if (isNaN(BATCH) || BATCH <= 0) BATCH = 50; // 帯域配慮の既定値（従来200→50）
    // 任意：取得側スロットリングのヒント（ログに残す程度）
    const sleepHint = Number(props.getProperty("BASE_SLEEP_MS_HINT")) || 500;

    // ログに進捗状況を出力（備考ではなく進捗状況列に詳細を出力）
    appendRunLog("RUN_SINGLE", "", {
      progress: `cursor=${cursor} batch=${BATCH} sleepHint=${sleepHint}`
    });

    // ---- 5) 今回処理対象のID群を抽出 -----------------------------------------
    // tenants の並びに従って slice（カーソル位置から BATCH 件）
    const batchIds = tenants.slice(cursor, cursor + BATCH).map(t => t.companyId);

    // ---- 6) 会社情報の取得（直列＋バックオフ版を利用） -----------------------
    // ここは fetchCompaniesByIds の修正版が前提（UrlFetchApp.fetchAll 廃止）
    const clients = fetchCompaniesByIds(batchIds); // [client or null] 同順

    // ---- 7) レスポンスごとに UPSERT ＋ ログ -----------------------------------
    // ※ null は「未取得（レート制限含む）／未存在」を区別できないため MISS として扱う
    // ★ テナントの companyId → companyName マップを作成（フォールバック検索用）
    const tenantNameMap = new Map();
    tenants.forEach(t => {
      if (t.companyName && String(t.companyName).trim() !== "") {
        tenantNameMap.set(Number(t.companyId), t.companyName);
      }
    });

    batchIds.forEach((requestId, idx) => {
      let client = clients[idx];

      // ★ 会社IDでの取得に失敗した場合、会社名でフォールバック検索を行う
      if (!client) {
        const fallbackName = tenantNameMap.get(Number(requestId));
        if (fallbackName) {
          Logger.log(`会社ID=${requestId} の取得に失敗。会社名「${fallbackName}」で再検索します`);
          client = fetchCompanyByName(fallbackName);
          if (client) {
            Logger.log(`会社名「${fallbackName}」で再取得成功: 新会社ID=${client.id}`);
          } else {
            Logger.log(`会社名「${fallbackName}」でも取得できませんでした`);
          }
        }
      }

      if (!client) {
        // MISS でログ (新構成: ["日時", "操作", "会社ID", "会社名", "更新日時", "進捗状況"])
        logEntries.push([
          Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss"),
          "MISS",
          requestId,
          tenantNameMap.get(Number(requestId)) || "",  // 会社名（テナントマスタから）
          "",           // 更新日時（不明）
          "id=" + requestId  // 進捗状況列に移動
        ]);
        return;
      }

      // 取得できたものは通常通り UPSERT（ADD/UPDATE は upsertClients が判定）
      upsertClients(
        [client],
        tenants,
        sheet,
        existing,
        processedIds,
        logEntries,
        {
          from: requestId,
          to: requestId,
          page: 1,
          requestId: requestId,
          enableLastTouch: false, // ★週次はラストタッチを取らない
        }
      );
    });

    // ---- 8) ログ投入（まとめて書く） ------------------------------------------
    appendLogEntries(logEntries);

    // ---- 8.5) 累積統計を更新 --------------------------------------------------
    const batchAdded = logEntries.filter(e => e[1] === "ADD").length;
    const batchUpdated = logEntries.filter(e => e[1] === "UPDATE").length;
    const batchMiss = logEntries.filter(e => e[1] === "MISS").length;

    const prevAdded = Number(props.getProperty("STATS_ADDED")) || 0;
    const prevUpdated = Number(props.getProperty("STATS_UPDATED")) || 0;
    const prevMiss = Number(props.getProperty("STATS_MISS")) || 0;

    props.setProperty("STATS_ADDED", String(prevAdded + batchAdded));
    props.setProperty("STATS_UPDATED", String(prevUpdated + batchUpdated));
    props.setProperty("STATS_MISS", String(prevMiss + batchMiss));

    // ---- 9) カーソル更新／完了判定 --------------------------------------------
    cursor += BATCH;

    if (cursor >= tenants.length) {
      // 全件処理完了：プロパティとトリガーをクリーンアップ
      props.deleteProperty("ID_CURSOR");
      props.deleteProperty("STEP_MODE");
      stopStepTrigger("updateCompanyListStepSingleIdMode");

      // 累積統計から最終値を取得
      const endTime = new Date();
      const statsStartTimeStr = props.getProperty("STATS_START_TIME");
      const statsStartTime = statsStartTimeStr ? new Date(statsStartTimeStr) : startTime;
      const durationSec = Math.floor((endTime - statsStartTime) / 1000);

      const totalAdded = Number(props.getProperty("STATS_ADDED")) || 0;
      const totalUpdated = Number(props.getProperty("STATS_UPDATED")) || 0;
      const totalMiss = Number(props.getProperty("STATS_MISS")) || 0;

      safeNotifySummary({
        summaries: [{ target: '企業情報', added: totalAdded, updated: totalUpdated, notFound: totalMiss }],
        startTime: statsStartTime,
        endTime,
        durationSec,
        isError: false,
        mode: 'weekly'
      });

      // 累積統計プロパティも削除
      props.deleteProperty("STATS_START_TIME");
      props.deleteProperty("STATS_ADDED");
      props.deleteProperty("STATS_UPDATED");
      props.deleteProperty("STATS_MISS");

      // 列計算の式を再投入（運用要望対応）
      reapplyFormulasFromSettings();

      // ★ラストタッチ更新を連鎖起動 (batch=200, interval=5min)
      startLastTouchStep(500, 1);

      // 進捗状況付きでログ出力
      appendRunLog("DONE_SINGLE", `全 ${tenants.length} 件処理完了（ADD=${totalAdded}, UPDATE=${totalUpdated}, MISS=${totalMiss}）→ LastTouch開始`, {
        cursor: tenants.length,
        total: tenants.length,
        progress: `全${tenants.length}件完了`
      });
    } else {
      // まだ残件あり：カーソル保持して次のトリガーで続行
      props.setProperty("ID_CURSOR", String(cursor));
      props.setProperty("STEP_MODE", "single"); // モード明示

      // 進捗状況付きでログ出力
      appendRunLog("PROGRESS", "", {
        cursor: cursor,
        total: tenants.length
      });
    }

    // Logger.log削除（不要なログ出力）

  } catch (e) {
    // ---- 10) 異常時：トリガー停止＋通知 ---------------------------------------
    stopStepTrigger("updateCompanyListStepSingleIdMode");

    safeNotifySummary({
      summaries: [
        `メッセージ: \`${(e && e.message) || e}\``,
        'スタック:',
        '```',
        truncateString((e && e.stack) || '(no stack)', 1500),
        '```'
      ],
      startTime,
      endTime: new Date(),
      durationSec: 0,
      isError: true,
      mode: 'weekly'
    });

    appendRunLog("ERROR", `例外（single）: ${(e && e.message) || e}`);
    throw e; // ログのため再throw（必要に応じて握りつぶし可）

  } finally {
    // ---- 11) ロック解放 -------------------------------------------------------
    lock.releaseLock();
  }
}

/**
 * 週次の「1IDモード」ステップ実行を開始するための初期化関数（修正版）。
 * - 出力シート／ログシートの存在保証
 * - スクリプトプロパティの初期化（カーソル、バッチサイズ、モード、スロットリング用ヒント）
 * - 既存トリガーを安全に停止してから、新トリガーを作成
 * - トリガー間隔やバッチサイズを引数で上書き可能（レート制限に合わせて調整）
 *
 * 【推奨運用値】
 *   batchSize: 50（従来200は帯域超過になりやすい）
 *   intervalMinutes: 5（従来1分間隔は過密）
 *   baseSleepMsHint: 500（fetchCompaniesByIdsのスロットリング目安）
 *
 * @param {number} [batchSize=50]        - 1回のステップで処理する会社ID件数（SINGLE_BATCH）
 * @param {number} [intervalMinutes=5]   - ステップトリガーの実行間隔（分）
 * @param {number} [baseSleepMsHint=500] - 取得関数側のスロットリング目安（ミリ秒）。プロパティに保存して参照用に使う
 */
function startSingleIdStep(batchSize, intervalMinutes, baseSleepMsHint) {
  // ---- 0) デフォルト値＆入力バリデーション ---------------------------------
  // APIレート制限: 5分間1000回 → 1分あたり200回が上限
  // 最適設定: バッチ500件 × 1分間隔 → 5分で2500件処理可能
  const BATCH_DEFAULT = 500;   // 1ステップの処理件数（レート制限内で最大化）
  const INTERVAL_DEFAULT = 1;  // トリガー間隔（分）: 1分間隔で高速処理
  const SLEEP_HINT_DEFAULT = 100;

  // 数値化＆下限の安全値
  let batch = Number(batchSize);
  if (isNaN(batch) || batch <= 0) batch = BATCH_DEFAULT;
  // 過度な大きさは抑止（レート制限考慮: 1分200回が上限だが、余裕を持って500に設定）
  if (batch > 500) batch = 500;

  let interval = Number(intervalMinutes);
  if (isNaN(interval) || interval <= 0) interval = INTERVAL_DEFAULT;
  // 1分未満は設定不可／過密なので最低1分、推奨5分以上
  if (interval < 1) interval = 1;

  let sleepHint = Number(baseSleepMsHint);
  if (isNaN(sleepHint) || sleepHint < 0) sleepHint = SLEEP_HINT_DEFAULT;

  // ---- 1) 前提チェック：APIキー＆テナントマスタ --------------------------------
  const apiKey = getApiKey && getApiKey();
  if (!apiKey) {
    // トリガー作成前に失敗を明示（運用事故防止）
    throw new Error("startSingleIdStep: API_KEY が未設定です。setApiKey() またはスクリプトプロパティで設定してください。");
  }
  // テナントマスタの読み取り（空なら開始しても意味がない）
  let tenants = [];
  try {
    tenants = getTenantMasterData();
  } catch (e) {
    throw new Error("startSingleIdStep: テナントマスタ読み取りに失敗しました。原因: " + e);
  }
  if (!Array.isArray(tenants) || tenants.length === 0) {
    throw new Error("startSingleIdStep: テナントマスタが空です。テナントマスタを準備してから再実行してください。");
  }

  // ---- 2) 出力シート／ログシートの保証 ---------------------------------------
  // 列定義を設定シート「設定_列固定」から適用する前提の強化版（既存関数）
  const sheet = ensureOutputSheetFromSettings("企業情報");
  ensureLogSheet("取得ログ");

  // ---- 3) スクリプトプロパティ初期化 -----------------------------------------
  const props = PropertiesService.getScriptProperties();
  // 先頭から処理を開始
  props.setProperty("ID_CURSOR", "0");
  // 1ステップの処理量（安全側の値）
  props.setProperty("SINGLE_BATCH", String(batch));
  // 実行モード（週次 or single の識別に利用）
  props.setProperty("STEP_MODE", "single"); // 週次一括開始は別関数。ここは1IDモードステップ
  // スロットリングのヒント値（fetchCompaniesByIds側で参照してもよい）
  props.setProperty("BASE_SLEEP_MS_HINT", String(sleepHint));

  // 任意：開始ログ（運用観測用）
  appendRunLog("START_SINGLE", `cursor=0 batch=${batch} interval=${interval} sleepHint=${sleepHint}`);

  // ---- 4) 既存の該当トリガーを停止してから新規作成 ---------------------------
  // 実行関数名は既存の週次処理と共通（バッチを刻んで段階的に進める）
  const handlerName = "updateCompanyListStepSingleIdMode";

  // 安全に既存トリガー停止
  stopStepTrigger(handlerName);

  // 新規トリガー作成（指定分間隔）
  // 注意：Google Apps Script の timeBased().everyMinutes(n) は 1/5/10/15/30 などの実行間隔が推奨値
  ScriptApp.newTrigger(handlerName)
    .timeBased()
    .everyMinutes(interval)
    .create();

  // 任意：開始完了ログ
  Logger.log(`1IDモード開始: batch=${batch}, interval=${interval}min, sleepHint=${sleepHint}ms, tenants=${tenants.length}`);
}

/**
 * 指定した関数名のトリガーを停止する
 * @param {string} handlerName - 停止対象の関数名
 */
function stopStepTrigger(handlerName) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(t);
      Logger.log(`ステップトリガーを停止しました: ${handlerName}`);
    }
  });
}

/* ---------------------------
   NOT_FOUND一覧をシート全体ユニーク化して整理
   --------------------------- */
function outputNotFoundList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 出力先シートの用意（ヘッダは2列）
  let out = ss.getSheetByName("企業情報_NOT_FOUND");
  if (!out) {
    out = ss.insertSheet("企業情報_NOT_FOUND");
    out.appendRow(["ログ日時", "会社ID"]);
  }

  // 取得ログから抽出
  const logSheet = ensureLogSheet();
  const lastRow = logSheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log("取得ログにデータがありません");
    return;
  }

  // 必要な列数分だけ取得（現在の構成は6列）
  const values = logSheet.getRange(2, 1, lastRow - 1, 6).getValues();

  // NOT_FOUND のみ抽出し、会社IDを数値に正規化
  const allData = [];
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (row[1] !== "NOT_FOUND" && row[1] !== "MISS") continue; // MISSも含めるのが安全

    const ts = row[0]; // ログ日時
    let companyId = "";

    // 会社ID列(列2)を優先
    const colIdVal = row[2];
    if (colIdVal && (typeof colIdVal === "number" || /^\d+$/.test(colIdVal))) {
      companyId = String(colIdVal);
    }
    // 進捗状況(列5)に "id=..." がある場合のフォールバック（以前のログ形式等への対応）
    else {
      const progress = row[5] || "";
      const m = String(progress).match(/id\s*=\s*(\d+)/);
      if (m && m[1]) companyId = m[1];
    }

    if (companyId) {
      allData.push([ts, companyId]);
    }
  }

  if (allData.length === 0) {
    Logger.log("NOT_FOUNDログがありません");
    return;
  }

  // --- ユニーク化処理 ---
  const map = new Map();
  allData.forEach(row => {
    const ts = row[0];
    const id = row[1];
    // 既にある場合はログ日時を比較して新しい方を残す
    if (!map.has(id) || ts > map.get(id)[0]) {
      map.set(id, [ts, id]);
    }
  });

  const uniqueData = Array.from(map.values());

  // シートを初期化して再書き込み
  out.clear();
  out.appendRow(["ログ日時", "会社ID"]);
  out.getRange(2, 1, uniqueData.length, 2).setValues(uniqueData);

  Logger.log("企業情報_NOT_FOUND をユニーク化して整理しました: " + uniqueData.length + "件");
}


/* ---------------------------
   日次差分更新（バッチ処理版）
   - 長期休暇後など大量更新時もタイムアウトしない
   - トリガーで自動再開
   --------------------------- */

const DAILY_DIFF_BATCH_SIZE = 500;       // 1回のトリガーで処理する件数
const DAILY_DIFF_TIME_LIMIT_MS = 280000; // 4分40秒（トリガー上限は約6分）

/**
 * 日次差分更新の開始（バッチ処理版）
 * - プロパティを初期化し、トリガーを設定する
 * @param {number} [intervalMinutes=5] - トリガー間隔（分）
 */
function startDailyDiffStep(intervalMinutes) {
  const props = PropertiesService.getScriptProperties();
  const interval = Number(intervalMinutes) || 5;

  // 前回実行日時を取得（無ければ昨日の0時）
  let lastRun = props.getProperty("LAST_DIFF_RUN");
  if (!lastRun) {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    yesterday.setHours(0, 0, 0, 0);
    lastRun = Utilities.formatDate(yesterday, "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
  }

  const now = new Date();
  const nowStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");

  // プロパティ初期化
  props.setProperty("DAILY_DIFF_FROM", lastRun);
  props.setProperty("DAILY_DIFF_TO", nowStr);
  props.setProperty("DAILY_DIFF_PAGE", "1");
  props.setProperty("DAILY_DIFF_START_TIME", now.toISOString());
  props.setProperty("DAILY_DIFF_STATS_ADDED", "0");
  props.setProperty("DAILY_DIFF_STATS_UPDATED", "0");
  props.setProperty("DAILY_DIFF_STATS_NOT_FOUND", "0");

  // 既存トリガー停止＆新規作成
  stopStepTrigger("updateCompanyListDailyDiffStep");
  ScriptApp.newTrigger("updateCompanyListDailyDiffStep")
    .timeBased()
    .everyMinutes(interval)
    .create();

  appendRunLog("START_DAILY_DIFF", "", {
    progress: `対象期間: ${lastRun} ～ ${nowStr} (実行間隔: ${interval}分)`
  });
  Logger.log(`日次差分更新（バッチ版）開始: from=${lastRun} to=${nowStr}`);
}

/**
 * 日次差分更新のバッチ実行関数（トリガーから呼び出し）
 */
function updateCompanyListDailyDiffStep() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log("別実行中のため中断（daily diff step）");
    return;
  }

  const stepStartTime = new Date();
  const props = PropertiesService.getScriptProperties();

  try {
    const apiKey = getApiKey();
    if (!apiKey) {
      stopStepTrigger("updateCompanyListDailyDiffStep");
      throw new Error("API_KEY が未設定");
    }

    // プロパティから状態取得
    const fromDate = props.getProperty("DAILY_DIFF_FROM");
    const toDate = props.getProperty("DAILY_DIFF_TO");
    let page = Number(props.getProperty("DAILY_DIFF_PAGE")) || 1;

    if (!fromDate || !toDate) {
      stopStepTrigger("updateCompanyListDailyDiffStep");
      Logger.log("日次差分: プロパティが未設定のため終了");
      return;
    }

    // シートと既存データの準備
    const sheet = ensureOutputSheetFromSettings("企業情報");
    const tenants = getTenantMasterData();
    const existing = buildExistingDataMapByHeader(sheet);
    const processedIds = new Set();

    let totalAddedThisStep = 0;
    let totalUpdatedThisStep = 0;
    let totalNotFoundThisStep = 0;
    let hasMoreData = true;

    const url = "https://hammock.hot-profile.com/rest_api/v1/clients/get_entry_list";
    const pageSize = 200;

    // バッチ処理ループ
    while (hasMoreData) {
      // タイムアウトチェック
      const elapsed = new Date().getTime() - stepStartTime.getTime();
      if (elapsed > DAILY_DIFF_TIME_LIMIT_MS) {
        Logger.log(`時間制限超過のため中断 (経過: ${elapsed}ms, page=${page})`);
        break;
      }

      // API呼び出し
      const payload = {
        api_key: apiKey,
        search: { from_datetime_updated_on: fromDate, to_datetime_updated_on: toDate },
        page: { display_number: pageSize, number: page }
      };

      const res = fetchWithRetry(url, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload)
      });

      const data = JSON.parse(res.getContentText());
      const clients = Array.isArray(data.clients) ? data.clients : [];

      if (clients.length > 0) {
        const logEntries = [];
        const updatedRows = upsertClients(clients, tenants, sheet, existing, processedIds, logEntries, { from: fromDate, to: toDate, page: page });
        appendLogEntries(logEntries);

        totalAddedThisStep += logEntries.filter(e => e[1] === "ADD").length;
        totalUpdatedThisStep += logEntries.filter(e => e[1] === "UPDATE").length;
        totalNotFoundThisStep += logEntries.filter(e => e[1] === "NOT_FOUND").length;

        // ★企業情報が更新された行のラストタッチを即座に再取得
        if (updatedRows && updatedRows.length > 0) {
          forceUpdateLastTouch(updatedRows);
          Logger.log(`日次差分: ${updatedRows.length}件の更新行に対してラストタッチを再取得しました`);
        }
      }

      // 次ページがあるかチェック
      if (clients.length < pageSize) {
        hasMoreData = false;
      } else {
        page++;
      }
    }

    // 累積統計を更新
    const prevAdded = Number(props.getProperty("DAILY_DIFF_STATS_ADDED")) || 0;
    const prevUpdated = Number(props.getProperty("DAILY_DIFF_STATS_UPDATED")) || 0;
    const prevNotFound = Number(props.getProperty("DAILY_DIFF_STATS_NOT_FOUND")) || 0;

    props.setProperty("DAILY_DIFF_STATS_ADDED", String(prevAdded + totalAddedThisStep));
    props.setProperty("DAILY_DIFF_STATS_UPDATED", String(prevUpdated + totalUpdatedThisStep));
    props.setProperty("DAILY_DIFF_STATS_NOT_FOUND", String(prevNotFound + totalNotFoundThisStep));
    props.setProperty("DAILY_DIFF_PAGE", String(page));

    // 完了判定
    if (!hasMoreData) {
      finishDailyDiffStep(props);
    } else {
      // 進捗ログ
      appendRunLog("PROGRESS_DAILY_DIFF", "", {
        progress: `page=${page} 処理中、今回: 追加=${totalAddedThisStep} 更新=${totalUpdatedThisStep}`
      });
    }

  } catch (err) {
    stopStepTrigger("updateCompanyListDailyDiffStep");
    safeNotifySummary({
      summaries: [
        `メッセージ: \`${(err && err.message) || err}\``,
        'スタック:', '```', truncateString((err && err.stack) || '(no stack)', 1500), '```'
      ],
      startTime: stepStartTime, endTime: new Date(), durationSec: 0, isError: true, mode: 'daily'
    });
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 日次差分更新の終了処理
 */
function finishDailyDiffStep(props) {
  stopStepTrigger("updateCompanyListDailyDiffStep");

  // 統計情報を取得
  const startTimeStr = props.getProperty("DAILY_DIFF_START_TIME");
  const startTime = startTimeStr ? new Date(startTimeStr) : new Date();
  const endTime = new Date();
  const durationSec = Math.floor((endTime - startTime) / 1000);

  const totalAdded = Number(props.getProperty("DAILY_DIFF_STATS_ADDED")) || 0;
  const totalUpdated = Number(props.getProperty("DAILY_DIFF_STATS_UPDATED")) || 0;
  const totalNotFound = Number(props.getProperty("DAILY_DIFF_STATS_NOT_FOUND")) || 0;

  const toDate = props.getProperty("DAILY_DIFF_TO");

  // 通知
  safeNotifySummary({
    summaries: [{ target: '企業情報', added: totalAdded, updated: totalUpdated, notFound: totalNotFound }],
    startTime, endTime, durationSec, isError: false, mode: 'daily'
  });

  // 列計算式を再適用
  reapplyFormulasFromSettings();

  // ラストタッチ更新を連鎖起動 (日次差分モード=true)
  startLastTouchStep(500, 1, true);

  // 成功時のみ LAST_DIFF_RUN を更新
  if (toDate) {
    props.setProperty("LAST_DIFF_RUN", toDate);
  }

  // プロパティクリーンアップ
  props.deleteProperty("DAILY_DIFF_FROM");
  props.deleteProperty("DAILY_DIFF_TO");
  props.deleteProperty("DAILY_DIFF_PAGE");
  props.deleteProperty("DAILY_DIFF_START_TIME");
  props.deleteProperty("DAILY_DIFF_STATS_ADDED");
  props.deleteProperty("DAILY_DIFF_STATS_UPDATED");
  props.deleteProperty("DAILY_DIFF_STATS_NOT_FOUND");

  // ログ
  appendRunLog("DONE_DAILY_DIFF", `全処理完了 追加=${totalAdded} 更新=${totalUpdated} 見つからず=${totalNotFound} → LastTouch開始`);
}

/**
 * 週次全件更新開始（バッチサイズ指定可能）
 * - 出力シートを初期化（既存データを削除し、ヘッダを保証）
 * - スクリプトプロパティを初期化してステップモードを開始
 * - 以降はトリガーで updateCompanyListStepSingleIdMode がバッチ分割処理を実行
 */
function startWeeklyStep(batchSize, intervalMinutes) {
  // ✅ 出力シートを設定_列固定から保証
  const sheet = ensureOutputSheetFromSettings("企業情報");
  ensureLogSheet("取得ログ");

  // ✅ データ削除はヘッダ定義列数に限定
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
  }

  const batch = Number(batchSize) || 200;
  const interval = Number(intervalMinutes) || 5;

  // ✅ スクリプトプロパティを初期化
  const p = PropertiesService.getScriptProperties();
  p.setProperty("ID_CURSOR", "0"); // 先頭から処理開始
  p.setProperty("SINGLE_BATCH", String(batch));
  p.setProperty("STEP_MODE", "weekly"); // 実行モードを weekly に設定

  // ✅ 累積統計用プロパティを初期化（複数トリガーにまたがる正確な集計用）
  p.setProperty("STATS_START_TIME", new Date().toISOString());
  p.setProperty("STATS_ADDED", "0");
  p.setProperty("STATS_UPDATED", "0");
  p.setProperty("STATS_MISS", "0");

  // ✅ 既存トリガーを停止してから新しいトリガーを作成
  stopStepTrigger("updateCompanyListStepSingleIdMode");
  ScriptApp.newTrigger("updateCompanyListStepSingleIdMode")
    .timeBased()
    .everyMinutes(interval)
    .create();

  Logger.log(`週次開始 batch=${batch} interval=${interval}min`);
}

/* ---------------------------
   補助: ログ追記ユーティリティ
   --------------------------- */
function appendRunLog(operation, note, context) {
  const logSheet = ensureLogSheet();
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  // 進捗状況のフォーマット
  let progress = "";
  if (context?.cursor !== undefined && context?.total !== undefined) {
    const remaining = Math.max(0, context.total - context.cursor);
    progress = `${context.cursor}件完了 (残り: ${remaining}件)`;
  } else if (context?.progress) {
    progress = context.progress;
  }

  // note があれば progress に統合
  if (note) {
    if (progress) progress += " / " + note;
    else progress = note;
  }

  // 新構成: ["日時", "操作", "会社ID", "会社名", "更新日時", "進捗状況"] (備考なし)
  const row = [
    ts,
    operation || "",
    context?.companyId || "",
    context?.companyName || "",
    context?.updatedAt || "",
    progress
  ];
  logSheet.getRange(logSheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}

/* ---------------------------
   補助: 単発実行（必要時）
   --------------------------- */
function dailyUpdateOnce() {
  ensureOutputSheet();
  ensureLogSheet();

  const tenants = getTenantMasterData();
  const sheet = ensureOutputSheet();
  const existing = buildExistingDataMapByHeader(sheet);
  const logEntries = [];
  const processedIds = new Set();

  tenants.forEach(t => {
    const requestId = t.companyId;
    const client = fetchCompanyById(requestId);

    if (!client) {
      appendRunLog("NOT_FOUND", `id=${requestId}`);
      return;
    }

    upsertClients([client], tenants, sheet, existing, processedIds, logEntries, { from: requestId, to: requestId, page: 1, requestId: requestId });
  });

  appendLogEntries(logEntries);
  Logger.log("日次差分単発実行完了");
}
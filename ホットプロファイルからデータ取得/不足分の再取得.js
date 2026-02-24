/**
 * 指定した会社ID群をピンポイントで再取得して企業一覧へ反映する
 * - 期間は1970〜現在時刻で広めに設定
 * - テナント絞り込み（tenantIdsSet）を維持
 * - 取得ログに API/FILTER/ADD/UPDATE を記録
 */

function refetchCompaniesByIds(idList) {
  ensureOutputSheet();
  ensureLogSheet();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { appendRunLog("SKIP", "Lock 競合(refetch)"); return; }

  try {
    // 1) idList を検証
    if (!Array.isArray(idList)) {
      appendRunLog("ERROR", `idList が配列ではありません（typeof=${typeof idList}）`);
      return;
    }
    if (idList.length === 0) {
      appendRunLog("NO_DATA", "idList が空（refetch）");
      return;
    }

    // 2) 期間（1970〜現在時刻）
    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const fromDt = "1970-01-01 00:00:00";
    const toDt = Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss");

    // 3) テナントマスタ（必ず配列化）
    let tenants = [];
    try {
      const t = getTenantMasterData();
      tenants = Array.isArray(t) ? t : [];
    } catch (e) {
      appendRunLog("ERROR", "getTenantMasterData 例外(refetch): " + e);
      tenants = [];
    }

    const tenantCount = tenants.length;
    appendRunLog("READ", `tenants=${tenantCount} / idList=${idList.length}`);

    if (tenantCount === 0) {
      appendRunLog("NO_DATA", "テナントマスタが空/取得不可のため再取得不可");
      return;
    }

    const sheet = ensureOutputSheet();
    const existing = buildExistingDataMap(sheet);
    const processedIds = new Set();
    const logEntries = [];

    // 4) テナントの会社ID集合（undefined 回避）
    const tenantIdsSet = new Set((tenants || []).map(t => Number(t.companyId)));

    // 5) idList を数値化し、テナントに存在するものだけに限定
    const targetIds = idList
      .map(v => Number(v))
      .filter(v => !isNaN(v) && v > 0 && tenantIdsSet.has(v));

    appendRunLog("READ", `targetIds=${targetIds.length}`);
    if (targetIds.length === 0) {
      appendRunLog("NO_DATA", "テナント該当なし（idList→targetIds=0）");
      return;
    }

    // 6) 会社IDごとに from_id=to_id=ID で取得
    for (const cid of targetIds) {
      let page = 1;
      while (true) {
        const { clients, count } = fetchCompanyPage(cid, cid, page, fromDt, toDt);
        appendRunLog("API", `refetch cid=${cid} statusCount=${count}`, { from: cid, to: cid, page });

        if (clients.length > 0) {
          const filtered = clients.filter(c => tenantIdsSet.has(Number(c.id)));
          appendRunLog("FILTER", `cid=${cid} clients=${clients.length} filtered=${filtered.length}`, { from: cid, to: cid, page });
          upsertClients(filtered, tenants, sheet, existing, processedIds, logEntries, { from: cid, to: cid, page });
        }

        if (count < PAGE_SIZE) break;
        page++;
      }
    }

    appendLogEntries(logEntries);
    appendRunLog("DONE", `refetch 完了 target=${targetIds.length}`);

  } catch (e) {
    appendRunLog("ERROR", "refetchCompaniesByIds 例外: " + e);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 「再取得ID」シートから会社IDを読み込んで、数値の配列にして返却する
 * - ヘッダ1行目は自動スキップ（値が非数値なら除外）
 * - A列（1列目）を参照。必要なら列番号を変えて使う
 * - 空/重複/NaN は除外
 * - 常に配列を返す（例外時は []）
 */

function readCompanyIdsFromSheet(sheetName, colIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) { appendRunLog("ERROR", `IDシート不在: ${sheetName}`); return []; }

    const last = sheet.getLastRow();
    if (last < 2) { appendRunLog("READ", `IDシートが空: ${sheetName}`); return []; }

    const col = colIndex || 1;
    const values = sheet.getRange(2, col, last - 1, 1).getValues().map(r => r[0]);

    const ids = [];
    const seen = new Set();
    for (const v of values) {
      if (v === "" || v === null || v === undefined) continue;
      let s = String(v).trim();
      s = s.replace(/[^\d,]/g, "");          // 数字とカンマ以外を除去
      const parts = s.split(",").map(p => p.trim()).filter(p => p.length > 0);
      for (const part of parts) {
        const n = Number(part);
        if (!isNaN(n) && n > 0 && !seen.has(n)) { ids.push(n); seen.add(n); }
      }
    }
    appendRunLog("READ", `ID読み取り: ${ids.length}件`);
    return ids;
  } catch (e) {
    appendRunLog("ERROR", "readCompanyIdsFromSheet 例外: " + e);
    return [];
  }
}

/**
 * IDシート（例: 「再取得ID」A列）にある会社IDを読み込み、ピンポイント再取得を実行
 */

function refetchMissingFromSheet() {
  ensureOutputSheet();
  ensureLogSheet();

  const ids = readCompanyIdsFromSheet("再取得ID", 1); // ← シート名/列番号は運用に合わせて
  if (!Array.isArray(ids) || ids.length === 0) {
    appendRunLog("NO_DATA", "IDシートに有効な会社IDがありません");
    return;
  }
  refetchCompaniesByIds(ids);
}

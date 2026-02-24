/**
 * ラストタッチ確認.gs
 *
 * 概要:
 * 1. 「企業情報」シートから会社名を取得
 * 2. HotProfile API (daily_reports/get_entry_list) を会社名で検索
 *    - 最適化: API側で visit_on 降順ソート＋1件制限を行い、最新の日報のみを直接取得する
 * 3. 「企業情報」シートの「ラストタッチ」列にその日時を書き込む
 * 4. サーバー負荷低減のためバッチ分割（トリガー実行）＋並列取得（2並列）で行う
 */

// APIレート制限: 5分間1000回 = 1分あたり200回が上限
const LT_CONCURRENCY = 10;      // API同時接続数（レート制限内で最大スループット）
const LT_CHUNK_SLEEP_MS = 100;  // チャンク間待機(ms)（100ms × 250チャンク = 25秒）
const LT_BATCH_DEFAULT = 500;   // 1回のトリガーでの処理行数

/**
 * レポートオブジェクトから関連するすべてのクライアントIDを抽出する
 * - report.client_id
 * - report.client_histories 内の各履歴から client_id, current_client_id, histories 内の client_id
 * @param {Object} report - HotProfile APIから返されるレポートオブジェクト
 * @returns {Set<number>} - 関連するクライアントIDのセット
 */
function getAllClientIdsFromReport(report) {
    const ids = new Set();

    // トップレベルの client_id
    if (report.client_id) {
        ids.add(Number(report.client_id));
    }

    // client_histories の解析
    if (Array.isArray(report.client_histories)) {
        report.client_histories.forEach(history => {
            // history.client_id
            if (history.client_id) {
                ids.add(Number(history.client_id));
            }
            // history.current_client_id
            if (history.current_client_id) {
                ids.add(Number(history.current_client_id));
            }
            // history.histories 内の各エントリ
            if (Array.isArray(history.histories)) {
                history.histories.forEach(h => {
                    if (h.client_id) {
                        ids.add(Number(h.client_id));
                    }
                });
            }
        });
    }

    return ids;
}


/**
 * ラストタッチ更新ステップの開始
 * - メインの企業情報更新後に呼び出される想定
 * - カーソルを初期化し、トリガーを設定する
 * @param {boolean} [isDiffMode=false] - trueなら日次差分更新(updateLastTouchDailyDiffStep)を実行
 */
function startLastTouchStep(batchSize, intervalMinutes, isDiffMode) {
    const props = PropertiesService.getScriptProperties();
    const batch = Number(batchSize) || LT_BATCH_DEFAULT;
    const interval = Number(intervalMinutes) || 5;

    if (isDiffMode) {
        // 日次差分モード：新ロジック
        // 既存トリガー停止
        stopStepTrigger("updateLastTouchDailyDiffStep");
        // 単発実行に近いが、トリガーで起動（コンテキスト分離）
        ScriptApp.newTrigger("updateLastTouchDailyDiffStep")
            .timeBased()
            .everyMinutes(interval) // 念のためインターバルにするが、1回で終われば自身で消す
            .create();

        // ログ
        appendRunLog("START_LT_DIFF", "日次差分更新を開始します");
        return;
    }

    // --- 以下、従来（全件走査）モード ---
    // カーソル初期化
    props.setProperty("LAST_TOUCH_CURSOR", "0");
    props.setProperty("LAST_TOUCH_BATCH", String(batch));

    // 既存トリガー停止＆新規作成
    stopStepTrigger("updateLastTouchStep");
    ScriptApp.newTrigger("updateLastTouchStep")
        .timeBased()
        .everyMinutes(interval)
        .create();

    // 統計プロパティの初期化（完了通知時に正確な時間と件数を表示するため）
    props.setProperty("LAST_TOUCH_START_TIME", new Date().toISOString());
    props.setProperty("LAST_TOUCH_UPDATED_COUNT", "0");
}

/**
 * ラストタッチ更新のバッチ実行関数
 * - トリガーから呼び出される
 */
function updateLastTouchStep() {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) {
        Logger.log("別実行中のため中断（LastTouch）");
        return;
    }

    const startTime = new Date();
    // タイムアウト対策: 実行時間がこの時間を超えたら、そこまでの結果を保存して終了する
    const TIME_LIMIT_MS = 280 * 1000; // 4分40秒 (トリガー上限は約6分)

    const props = PropertiesService.getScriptProperties();

    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("企業情報");
        if (!sheet) throw new Error("「企業情報」シートが見つかりません");

        // ヘッダ特定
        const lastCol = sheet.getLastColumn();
        const lastRow = sheet.getLastRow();
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        const h2i = {};
        headers.forEach((h, i) => { h2i[String(h).trim()] = i + 1; });

        const colName = h2i["会社名"];
        const colNewId = h2i["新会社ID"];
        const colOldId = h2i["会社ID"]; // 旧会社ID列を追加
        let colLastTouch = h2i["ラストタッチ"];
        let colLastTouchUrl = h2i["ラストタッチURL"];

        if (!colName || !colNewId) throw new Error("必須列（会社名/新会社ID）が見つかりません");

        // ラストタッチ列がなければ追加（運用上は upsertClients で既にあるはずだが念のため）
        if (!colLastTouch) {
            // 再取得して末尾に追加
            const currentLastCol = sheet.getLastColumn();
            colLastTouch = currentLastCol + 1;
            sheet.getRange(1, colLastTouch).setValue("ラストタッチ");
        }

        // ラストタッチURL列がなければ追加
        if (!colLastTouchUrl) {
            const currentLastCol = sheet.getLastColumn();
            colLastTouchUrl = currentLastCol + 1;
            sheet.getRange(1, colLastTouchUrl).setValue("ラストタッチURL");
        }

        // カーソルとバッチサイズ
        let cursor = Number(props.getProperty("LAST_TOUCH_CURSOR")) || 0;
        const batch = Number(props.getProperty("LAST_TOUCH_BATCH")) || LT_BATCH_DEFAULT;

        // データ範囲: 行番号は 2 から始まる (cursor=0 のとき row=2)
        // cursor は「処理済みの行数（ヘッダ除く）」とみなす
        const dataStartRow = 2 + cursor;

        // 残り行数が無い場合
        if (dataStartRow > lastRow) {
            finishLastTouchStep(props);
            return;
        }

        // 今回処理する行数を計算
        const rowsToProcess = Math.min(batch, lastRow - dataStartRow + 1);
        if (rowsToProcess <= 0) {
            finishLastTouchStep(props);
            return;
        }

        // 対象範囲の読み込み（会社名・新会社ID・会社ID）
        const names = sheet.getRange(dataStartRow, colName, rowsToProcess, 1).getValues();
        const ids = sheet.getRange(dataStartRow, colNewId, rowsToProcess, 1).getValues();
        const oldIds = colOldId ? sheet.getRange(dataStartRow, colOldId, rowsToProcess, 1).getValues() : null;

        // 並列取得用のリクエスト作成
        // null/空の会社名はスキップ対象だが、行ズレしないように配列インデックスを維持する
        const tasks = [];
        const apiKey = getApiKey();
        const url = "https://hammock.hot-profile.com/rest_api/v1/daily_reports/get_entry_list";

        for (let i = 0; i < rowsToProcess; i++) {
            const cName = String(names[i][0]).trim();
            const cId = Number(ids[i][0]);
            const cOldId = oldIds ? Number(oldIds[i][0]) : 0; // 旧会社ID
            // 会社名とIDが有効な場合のみタスク化
            if (cName && cId > 0) {
                // ★API最適化: search[client_name] + order[visit_on desc] + limit 10
                const payload = {
                    api_key: apiKey,
                    search: { client_name: cName },
                    page: { display_number: 10, number: 1 },
                    order: { key: "visit_on", type: "desc" }
                };
                tasks.push({
                    index: i,
                    targetClientId: cId,
                    targetOldClientId: cOldId > 0 ? cOldId : null, // 旧会社IDを追加
                    req: {
                        url: url,
                        method: "post",
                        contentType: "application/json",
                        payload: JSON.stringify(payload),
                        muteHttpExceptions: true
                    }
                });
            }
        }

        // tasks をチャンク処理 (並列実行)
        // 結果を入れる配列（サイズは rowsToProcess）
        const updatesDate = new Array(rowsToProcess).fill(null); // 日時用
        const updatesUrl = new Array(rowsToProcess).fill(null);  // URL用
        let processedCount = 0; // 実際にAPIリクエストまで処理が進んだ件数（スキップ含む）のインデックス追跡用
        let timeOutOccurred = false;

        // tasks が空でも rowsToProcess 分は進める必要があるが、
        // tasks がある場合は tasks の最後の index まで処理したとみなす
        // ここでは単純に i ループがどこまで進んだかで processedRows を決める

        let lastProcessedTaskIndex = -1;

        for (let i = 0; i < tasks.length; i += LT_CONCURRENCY) {
            // ▼タイムアウトチェック
            const elapsed = new Date().getTime() - startTime.getTime();
            if (elapsed > TIME_LIMIT_MS) {
                Logger.log(`時間制限超過のため中断します (経過: ${elapsed}ms)`);
                timeOutOccurred = true;
                break;
            }

            const chunk = tasks.slice(i, i + LT_CONCURRENCY);
            const requests = chunk.map(t => t.req);

            try {
                const responses = UrlFetchApp.fetchAll(requests);

                responses.forEach((res, idx) => {
                    const task = chunk[idx];
                    const status = res.getResponseCode();
                    const body = res.getContentText();

                    // 処理が進んだタスクのインデックスを更新
                    lastProcessedTaskIndex = Math.max(lastProcessedTaskIndex, i + idx);

                    if (status >= 200 && status < 300) {
                        try {
                            const data = JSON.parse(body);
                            const reports = Array.isArray(data.daily_reports) ? data.daily_reports : [];
                            // 新会社ID または 旧会社ID のどちらかと一致すればマッチ
                            // client_histories も含めてチェック
                            const found = reports.find(r => {
                                const ids = getAllClientIdsFromReport(r);
                                return ids.has(task.targetClientId) ||
                                    (task.targetOldClientId && ids.has(task.targetOldClientId));
                            });

                            if (found && found.visit_on) {
                                updatesDate[task.index] = found.visit_on;
                                // IDがあればURL生成
                                if (found.id) {
                                    updatesUrl[task.index] = `https://005108.hammock.hot-profile.com/daily_reports/${found.id}`;
                                }
                            } else {
                                // 見つからない、またはID不一致の場合は更新しない（既存データを維持するため null のままにする）
                                // updates[task.index] = ""; 
                            }
                        } catch (e) {
                            Logger.log(`LastTouch parsing error: ${e}`);
                        }
                    } else {
                        Logger.log(`LastTouch HTTP error (idx=${task.index}): ${status}`);
                    }
                });

            } catch (e) {
                Logger.log(`LastTouch fetchAll error: ${e}`);
            }

            // スロットリング
            if (i + LT_CONCURRENCY < tasks.length) {
                Utilities.sleep(LT_CHUNK_SLEEP_MS);
            }
        }

        // コミットする行数を決定
        let rowsToCommit = 0;

        if (tasks.length === 0) {
            // タスクがそもそもない（全行スキップなど）場合は、全行コミットしてよい
            rowsToCommit = rowsToProcess;
        } else if (timeOutOccurred) {
            // タイムアウト時は、最後に処理したタスクの行番号までをコミット
            // tasks[lastProcessedTaskIndex] の .index が、元の配列(rowsToProcess)での位置
            if (lastProcessedTaskIndex >= 0) {
                rowsToCommit = tasks[lastProcessedTaskIndex].index + 1;
            } else {
                rowsToCommit = 0; // まだ1つも処理できなかった
            }
        } else {
            // 正常完了時は全部コミット
            rowsToCommit = rowsToProcess;
        }

        // ログ出力削除：Logger.logを削除

        if (rowsToCommit > 0) {
            // シートへ結果書き込み

            // 1. ラストタッチ日時
            const rangeOutDate = sheet.getRange(dataStartRow, colLastTouch, rowsToCommit, 1);
            const existingValsDate = rangeOutDate.getValues();
            let commitUpdatedCount = 0; // 実際に値が変わった件数

            const outValuesDate = existingValsDate.map((row, i) => {
                const newVal = updatesDate[i];
                // APIからの取得値がない場合は既存のまま
                if (newVal === null || newVal === undefined) return row;

                const oldVal = row[0];
                let oldValStr = "";
                if (oldVal instanceof Date) {
                    oldValStr = Utilities.formatDate(oldVal, "Asia/Tokyo", "yyyy-MM-dd");
                } else {
                    // 文字列の場合、区切り文字の違い（/と-）を吸収して比較
                    oldValStr = String(oldVal).trim().replace(/\//g, "-");
                }

                // 既存値とAPI値が異なる場合のみ更新＆カウント
                if (oldValStr !== newVal) {
                    commitUpdatedCount++;
                    return [newVal];
                }
                return row; // 変更なし
            });
            rangeOutDate.setValues(outValuesDate);

            // 2. ラストタッチURL
            const rangeOutUrl = sheet.getRange(dataStartRow, colLastTouchUrl, rowsToCommit, 1);
            const existingValsUrl = rangeOutUrl.getValues();
            const outValuesUrl = existingValsUrl.map((row, i) => {
                const newVal = updatesUrl[i];
                if (newVal !== null && newVal !== undefined) return [newVal];
                return row;
            });
            rangeOutUrl.setValues(outValuesUrl);

            // カーソル更新
            const nextCursor = cursor + rowsToCommit;
            props.setProperty("LAST_TOUCH_CURSOR", String(nextCursor));

            // 更新件数を集計（実際に値が変わった件数）
            const actualUpdatedCount = commitUpdatedCount;
            const prevCount = Number(props.getProperty("LAST_TOUCH_UPDATED_COUNT")) || 0;
            props.setProperty("LAST_TOUCH_UPDATED_COUNT", String(prevCount + actualUpdatedCount));

            // 取得ログへの出力
            appendRunLog("LAST_TOUCH", "", {
                cursor: nextCursor,
                total: lastRow - 1,
                progress: `${nextCursor}件処理完了 (残り: ${Math.max(0, lastRow - 1 - nextCursor)}件)、今回更新: ${actualUpdatedCount}件`
            });

            // タイムアウトした場合はここで終了（次回トリガーで続きから）
            if (timeOutOccurred) {
                return;
            }

            // 終了判定（最後まで行き着いたか）
            if (nextCursor >= lastRow - 1) {
                finishLastTouchStep(props);
            }
        } else {
            // 何も処理できなかった場合（即タイムアウト？）
            // 無限ループ防止のため、もしタイムアウトでかつ処理数0なら、
            // 強制的に少し待つか、あるいはエラーとして抜けるなどの対策が必要だが、
            // 通常4分あれば1件は処理できるはず。
            if (timeOutOccurred) {
                Logger.log("処理件数0でタイムアウトしました。");
            }
        }

    } catch (e) {
        Logger.log("LastTouch Error: " + e);
        // エラーでも停止せず、ログに残して通知（必要なら）
        safeNotify({ title: "ラストタッチ更新エラー", markdown: String(e), level: "error" });
    } finally {
        lock.releaseLock();
    }
}

/**
 * 終了処理: トリガー削除とプロパティ削除、完了通知
 */
function finishLastTouchStep(props) {
    stopStepTrigger("updateLastTouchStep");

    // 統計情報を取得
    const startTimeStr = props.getProperty("LAST_TOUCH_START_TIME");
    const startTime = startTimeStr ? new Date(startTimeStr) : new Date();
    const endTime = new Date();
    const durationSec = Math.floor((endTime - startTime) / 1000);
    const totalUpdated = Number(props.getProperty("LAST_TOUCH_UPDATED_COUNT")) || 0;

    // プロパティ削除
    props.deleteProperty("LAST_TOUCH_CURSOR");
    props.deleteProperty("LAST_TOUCH_BATCH");
    props.deleteProperty("LAST_TOUCH_START_TIME");
    props.deleteProperty("LAST_TOUCH_UPDATED_COUNT");

    // 週次（全件走査）完了時点を日次差分の起点として保存
    // これにより、翌日のDaily Diffはここからの差分のみを取得するようになる
    const nowStr = Utilities.formatDate(endTime, "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
    props.setProperty("LAST_TOUCH_DIFF_RUN", nowStr);

    // 取得ログへの出力
    appendRunLog("LAST_TOUCH_DONE", "", {
        progress: `ラストタッチ更新完了：総更新件数=${totalUpdated}`
    });

    safeNotifySummary({
        summaries: [{ target: 'ラストタッチ', added: 0, updated: totalUpdated, notFound: 0 }],
        startTime: startTime,
        endTime: endTime,
        durationSec: durationSec,
        isError: false,
        mode: "last_touch_done"
    });
}

/**
 * 日報APIから指定期間に更新されたデータを全件取得する（ページネーション対応）
 */
function fetchDailyReportsByUpdatedRange(apiKey, fromUpdatedOn, toUpdatedOn, pageSize) {
    if (!pageSize) pageSize = 100;
    const url = "https://hammock.hot-profile.com/rest_api/v1/daily_reports/get_entry_list";
    const allReports = [];
    let page = 1;

    while (true) {
        // search[from_updated_on] / search[to_updated_on] を使用
        // page[display_number] / page[number] を使用
        const payload = {
            api_key: apiKey,
            search: {
                from_datetime_updated_on: fromUpdatedOn,
                to_datetime_updated_on: toUpdatedOn
            },
            page: {
                display_number: pageSize,
                number: page
            },
            // order: { key: "updated_on", type: "asc" } // 500エラー回避のため一旦無効化
        };

        try {
            // 共通.jsのfetchWithRetryを利用
            const res = fetchWithRetry(url, {
                method: "post",
                contentType: "application/json",
                payload: JSON.stringify(payload)
            });

            const data = JSON.parse(res.getContentText());
            const reports = Array.isArray(data.daily_reports) ? data.daily_reports : [];

            allReports.push(...reports);

            // 次ページ判定
            if (reports.length < pageSize) break;
            page++;

            // 安全装置: 無限ループ防止（万が一データ数が膨大な場合）
            if (page > 30) {
                Logger.log("fetchDailyReportsByUpdatedRange: ページ上限(30)に達しました");
                break;
            }

        } catch (e) {
            Logger.log(`fetchDailyReportsByUpdatedRange error (page=${page}): ${e}`);
            throw e; // エラーは上位でハンドリング
        }
    }
    return allReports;
}

/**
 * ラストタッチ更新の日次差分処理（API差分ベース）
 * - 前回実行(LAST_TOUCH_DIFF_RUN)以降に更新された日報を取得
 * - 該当する企業のラストタッチを更新
 */
function updateLastTouchDailyDiffStep() {
    const lock = LockService.getScriptLock();
    // 日次差分は一発で終わらせる想定だが、念のためロック
    if (!lock.tryLock(30000)) {
        Logger.log("別実行中のため中断（LastTouch Daily Diff）");
        return;
    }

    const startTime = new Date();
    const props = PropertiesService.getScriptProperties();

    try {
        // 1. 期間取得
        const apiKey = getApiKey();
        if (!apiKey) throw new Error("APIキーが未設定です");

        let lastRun = props.getProperty("LAST_TOUCH_DIFF_RUN");
        // 初回などで未設定の場合は「昨日」から（あるいは適当な過去）
        if (!lastRun) {
            const d = new Date();
            d.setDate(d.getDate() - 1);
            d.setHours(0, 0, 0, 0);
            lastRun = Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");
        }

        const now = new Date();
        const nowStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");

        // 2. 差分取得 (Batch Size 200)
        const reports = fetchDailyReportsByUpdatedRange(apiKey, lastRun, nowStr, 200);

        // 更新対象がなければ終了
        if (reports.length === 0) {
            Logger.log(`ラストタッチ差分なし: ${lastRun} ~ ${nowStr}`);
            // プロパティ更新（次回はこの時点から）
            props.setProperty("LAST_TOUCH_DIFF_RUN", nowStr);

            safeNotifySummary({
                summaries: [{ target: 'ラストタッチ(差分)', added: 0, updated: 0, notFound: 0 }],
                startTime, endTime: new Date(), durationSec: 0, isError: false, mode: "last_touch_daily"
            });
            return;
        }

        // 3. シート読み込み & マップ作成 (New Company ID -> Row Index)
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("企業情報");
        if (!sheet) throw new Error("「企業情報」シートが見つかりません");

        const lastRow = sheet.getLastRow();
        if (lastRow < 2) return; // データなし

        // ヘッダ特定 & 列インデックス取得
        const lastCol = sheet.getLastColumn();
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        const h2i = {};
        headers.forEach((h, i) => { h2i[String(h).trim()] = i + 1; });

        const colName = h2i["会社名"];
        const colNewId = h2i["新会社ID"];
        const colOldId = h2i["会社ID"]; // 旧会社ID列を追加
        const colLastTouch = h2i["ラストタッチ"];
        const colLastTouchUrl = h2i["ラストタッチURL"];

        // 必須列チェック (ラストタッチURLは任意とするが、ラストタッチ列は更新先として必須)
        if (!colName || !colNewId || !colLastTouch) {
            throw new Error("必須列（会社名/新会社ID/ラストタッチ）が見つかりません。先に「週間データ取得」などを実行して列を作成してください。");
        }

        // rowMap作成 (Company ID -> [{ rowIndex, name }, ...])
        // ★修正: 同じ会社IDに複数の行が紐づく場合があるため、値を配列に変更
        // シート全データをメモリに読み込んでマップ化
        const dataValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
        const rowMap = new Map();

        for (let i = 0; i < dataValues.length; i++) {
            const row = dataValues[i];
            // 配列インデックスは colIndex - 1
            const cNewId = Number(row[colNewId - 1]);
            const cOldId = colOldId ? Number(row[colOldId - 1]) : 0; // 旧会社ID
            const cName = row[colName - 1];
            const rowInfo = {
                rowIndex: i + 2, // シート上の行番号 (ヘッダ1行 + i(=0始まり) + 1)
                name: cName
            };

            // 新会社IDでマップ登録（配列に追加）
            if (cNewId > 0) {
                if (!rowMap.has(cNewId)) {
                    rowMap.set(cNewId, []);
                }
                rowMap.get(cNewId).push(rowInfo);
            }
            // 旧会社IDでもマップ登録（新会社IDと異なる場合のみ、配列に追加）
            if (cOldId > 0 && cOldId !== cNewId) {
                if (!rowMap.has(cOldId)) {
                    rowMap.set(cOldId, []);
                }
                rowMap.get(cOldId).push(rowInfo);
            }
        }

        const updatesMap = new Map();

        reports.forEach(r => {
            // client_histories も含めて全ての関連IDを抽出
            const ids = getAllClientIdsFromReport(r);
            if (ids.size === 0) return;

            // 各IDについて、より新しい日付のレポートであれば更新
            ids.forEach(cid => {
                if (isNaN(cid) || cid <= 0) return;

                // 既存の候補より日付が新しければ更新
                if (updatesMap.has(cid)) {
                    const curr = updatesMap.get(cid);
                    if (r.visit_on > curr.visit_on) {
                        updatesMap.set(cid, r);
                    }
                } else {
                    updatesMap.set(cid, r);
                }
            });
        });

        let updateCount = 0;

        // ラストタッチ列も一括で読んでおく
        const lastTouchValues = sheet.getRange(2, colLastTouch, lastRow - 1, 1).getValues();

        const requests = [];
        const logEntries = [];

        updatesMap.forEach((report, cid) => {
            if (!rowMap.has(cid)) return; // シートにない企業は無視

            // ★修正: rowMap は配列を返すので、該当するすべての行に対してループ処理
            const entries = rowMap.get(cid);
            entries.forEach(entry => {
                const rowIndex = entry.rowIndex;
                const companyName = entry.name;
                const arrayIdx = rowIndex - 2;

                const newDate = report.visit_on;
                if (!newDate) return;

                const currentVal = lastTouchValues[arrayIdx][0];

                // 比較 (YYYY-MM-DD vs current)
                let currentStr = "";
                if (currentVal instanceof Date) {
                    currentStr = Utilities.formatDate(currentVal, "Asia/Tokyo", "yyyy-MM-dd");
                } else {
                    currentStr = String(currentVal).trim().replace(/\//g, "-");
                }

                if (currentStr !== newDate) {
                    // 更新必要
                    const rowData = { rowIndex: rowIndex, date: newDate, url: null };
                    if (report.id && colLastTouchUrl) {
                        rowData.url = `https://005108.hammock.hot-profile.com/daily_reports/${report.id}`;
                    }
                    requests.push(rowData);

                    // 詳細ログ追加
                    // logEntries形式: [日時, 操作, 会社ID, 会社名, 更新日時, 進捗状況]
                    logEntries.push([
                        nowStr,
                        "LAST_TOUCH",
                        cid,
                        companyName,
                        newDate, // 5列目: 更新日時（ラストタッチ日時）
                        `以前: ${currentStr || "(空)"} → 新: ${newDate}` // 6列目: 進捗状況
                    ]);
                }
            }); // entries.forEach
        });

        // ログシートへ追記
        appendLogEntries(logEntries);

        // 書き込み実行 (1件ずつだが、対象数が少なければOK)
        requests.forEach(req => {
            sheet.getRange(req.rowIndex, colLastTouch).setValue(req.date);
            if (req.url && colLastTouchUrl) {
                sheet.getRange(req.rowIndex, colLastTouchUrl).setValue(req.url);
            }
        });

        updateCount = requests.length;

        // 5. 状態保存 & ログ
        props.setProperty("LAST_TOUCH_DIFF_RUN", nowStr);

        appendRunLog("LAST_TOUCH_DIFF", "", {
            total: updatesMap.size,
            progress: `対象${updatesMap.size}社中、${updateCount}件更新 (from=${lastRun})`
        });

        safeNotifySummary({
            summaries: [{ target: 'ラストタッチ(差分)', added: 0, updated: updateCount, notFound: 0 }],
            startTime, endTime: new Date(), durationSec: Math.floor((new Date() - startTime) / 1000),
            isError: false, mode: "last_touch_daily"
        });

    } catch (e) {
        Logger.log(`updateLastTouchDailyDiffStep error: ${e}`);
        safeNotifySummary({
            summaries: [`エラー: ${e.message}`, `stack: ${e.stack}`],
            startTime, endTime: new Date(), durationSec: 0, isError: true, mode: "last_touch_daily"
        });
        throw e;
    } finally {
        // トリガー停止（1回で終わる想定だが、繰り返し実行防止）
        stopStepTrigger("updateLastTouchDailyDiffStep");
        lock.releaseLock();
    }
}

/**
 * 企業情報が更新されたレコードのラストタッチを即座に再取得する関数
 * - upsertClients で UPDATE された行に対して呼び出される想定
 * - 並列APIコール（UrlFetchApp.fetchAll）で効率的に取得
 *
 * @param {Array<{rowIndex: number, companyName: string, companyId: number, oldCompanyId?: number}>} targets
 *   - rowIndex: シート上の行番号（1始まり）
 *   - companyName: 会社名（API検索キー）
 *   - companyId: 新会社ID（マッチング用）
 *   - oldCompanyId: 旧会社ID（マッチング用、任意）
 */
function forceUpdateLastTouch(targets) {
    if (!Array.isArray(targets) || targets.length === 0) return;

    const apiKey = getApiKey();
    if (!apiKey) {
        Logger.log("forceUpdateLastTouch: API_KEY が未設定のためスキップ");
        return;
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("企業情報");
    if (!sheet) {
        Logger.log("forceUpdateLastTouch: 「企業情報」シートが見つかりません");
        return;
    }

    // ヘッダから列位置を取得
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const h2i = {};
    headers.forEach((h, i) => { h2i[String(h).trim()] = i + 1; });

    const colLastTouch = h2i["ラストタッチ"];
    const colLastTouchUrl = h2i["ラストタッチURL"];

    if (!colLastTouch) {
        Logger.log("forceUpdateLastTouch: ラストタッチ列が見つかりません");
        return;
    }

    const url = "https://hammock.hot-profile.com/rest_api/v1/daily_reports/get_entry_list";

    // 有効なターゲットのみフィルタ（会社名が空のものはスキップ）
    const validTargets = targets.filter(t => t.companyName && String(t.companyName).trim() && t.companyId > 0);
    if (validTargets.length === 0) return;

    // 並列APIリクエスト作成
    const tasks = validTargets.map(t => ({
        target: t,
        req: {
            url: url,
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify({
                api_key: apiKey,
                search: { client_name: String(t.companyName).trim() },
                page: { display_number: 10, number: 1 },
                order: { key: "visit_on", type: "desc" }
            }),
            muteHttpExceptions: true
        }
    }));

    // チャンク処理（並列数: LT_CONCURRENCY）
    let updatedCount = 0;

    for (let i = 0; i < tasks.length; i += LT_CONCURRENCY) {
        const chunk = tasks.slice(i, i + LT_CONCURRENCY);
        const requests = chunk.map(t => t.req);

        try {
            const responses = UrlFetchApp.fetchAll(requests);

            responses.forEach((res, idx) => {
                const task = chunk[idx];
                const status = res.getResponseCode();
                const body = res.getContentText();

                if (status >= 200 && status < 300) {
                    try {
                        const data = JSON.parse(body);
                        const reports = Array.isArray(data.daily_reports) ? data.daily_reports : [];

                        // 新会社ID または 旧会社ID のどちらかと一致すればマッチ
                        const found = reports.find(r => {
                            const ids = getAllClientIdsFromReport(r);
                            return ids.has(task.target.companyId) ||
                                (task.target.oldCompanyId && ids.has(task.target.oldCompanyId));
                        });

                        if (found && found.visit_on) {
                            // ラストタッチ日時を書き込み
                            sheet.getRange(task.target.rowIndex, colLastTouch).setValue(found.visit_on);
                            // ラストタッチURLを書き込み（列がある場合のみ）
                            if (colLastTouchUrl && found.id) {
                                sheet.getRange(task.target.rowIndex, colLastTouchUrl)
                                    .setValue(`https://005108.hammock.hot-profile.com/daily_reports/${found.id}`);
                            }
                            updatedCount++;
                        }
                        // マッチしない場合は何もしない（ラストタッチなしとしてそのまま）
                    } catch (e) {
                        Logger.log(`forceUpdateLastTouch parse error (row=${task.target.rowIndex}): ${e}`);
                    }
                } else {
                    Logger.log(`forceUpdateLastTouch HTTP error (row=${task.target.rowIndex}): ${status}`);
                }
            });
        } catch (e) {
            Logger.log(`forceUpdateLastTouch fetchAll error: ${e}`);
        }

        // スロットリング
        if (i + LT_CONCURRENCY < tasks.length) {
            Utilities.sleep(LT_CHUNK_SLEEP_MS);
        }
    }

    Logger.log(`forceUpdateLastTouch: ${validTargets.length}件中 ${updatedCount}件のラストタッチを再取得しました`);
}
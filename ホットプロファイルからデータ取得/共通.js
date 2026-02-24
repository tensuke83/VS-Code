/**
 * APIキーを取得する関数
 * - プロジェクトのスクリプトプロパティに保存しておくこと
 */
function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty("API_KEY");
}

/**
 * Teams Webhook URL を取得
 * - プロジェクトのスクリプトプロパティに保存しておくこと
 */
function getTeamsWebhookUrl() {
  return PropertiesService.getScriptProperties().getProperty("TEAMS_WEBHOOK_URL");
}

/**
 * Adaptive Card を Teams に送信
 * payload.body が指定されている場合はそれを直接 body として使用（リッチカード対応）
 * payload.title + payload.markdown の場合は従来の TextBlock 2段構成
 */
function sendTeamsAdaptiveCard(payload) {
  const url = getTeamsWebhookUrl();
  if (!url) throw new Error("TEAMS_WEBHOOK_URL が未設定です");

  // body の組み立て
  let cardBody;
  if (Array.isArray(payload.body)) {
    // リッチカード: body 配列がそのまま渡された場合
    cardBody = payload.body;
  } else {
    // 従来互換: title + markdown テキスト
    cardBody = [
      {
        type: 'TextBlock',
        text: `**${payload.title || ''}**`,
        wrap: true,
        weight: 'bolder',
        size: 'medium'
      },
      {
        type: 'TextBlock',
        text: payload.markdown || '',
        wrap: true
      }
    ];
  }

  const adaptive = {
    type: 'message',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: {
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
        type: 'AdaptiveCard',
        version: '1.4',
        msteams: { width: "Full" },
        body: cardBody
      }
    }]
  };

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify(adaptive),
    muteHttpExceptions: true
  });

  const status = res.getResponseCode();
  const resBody = res.getContentText();
  Logger.log(`[Teams] status=${status} body=${truncateString(resBody, 1500)}`);

  if (status < 200 || status >= 300) {
    throw new Error(`Teams webhook error: HTTP ${status} - ${resBody}`);
  }
}

/**
 * safeNotify: 統一フォーマットで Teams に通知
 */
function safeNotify({ title, markdown, level }) {
  try {
    sendTeamsAdaptiveCard({ title, markdown });
  } catch (e) {
    Logger.log("Teams通知失敗: " + e);
  }
}



/**
 * Teams通知用のサマリ関数（Adaptive Card リッチレイアウト版）
 * - FactSet で数値情報を整列表示
 * - Container で対象ごとにグループ化
 * - エラー時は Attention カラーで強調
 * - 時刻は JST + 24時間表記に統一
 */
function safeNotifySummary({ summaries, startTime, endTime, durationSec, isError, mode }) {
  // JST + 24時間表記
  const fmtJst24h = (d) => Utilities.formatDate(d, "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss");

  // mode 表示
  let modeLabel;
  if (mode === 'weekly') {
    modeLabel = '週次更新';
  } else if (mode === 'last_touch_done') {
    modeLabel = 'ラストタッチ更新';
  } else {
    modeLabel = '日次更新';
  }

  // タイトル
  const title = isError
    ? `エラー発生（${modeLabel}）`
    : `処理完了（${modeLabel}）`;

  // Adaptive Card body の組み立て
  const body = [];

  // --- タイトル ---
  body.push({
    type: 'TextBlock',
    text: title,
    wrap: true,
    weight: 'bolder',
    size: 'large',
    color: isError ? 'attention' : 'default'
  });

  // --- サブタイトル（モード表示） ---
  body.push({
    type: 'TextBlock',
    text: `ホットプロファイルデータ　${modeLabel}`,
    wrap: true,
    size: 'small',
    isSubtle: true,
    spacing: 'none'
  });

  // --- 区切り線 ---
  body.push({ type: 'TextBlock', text: '─────────────────────────────', spacing: 'small', isSubtle: true });

  if (isError) {
    // エラー時：エラーメッセージを赤色で表示
    (summaries || []).forEach(msg => {
      body.push({
        type: 'TextBlock',
        text: String(msg),
        wrap: true,
        color: 'attention'
      });
    });
  } else {
    // 成功時：対象別サマリ
    (summaries || []).forEach(s => {
      const totalChanged = (s.added || 0) + (s.updated || 0);

      // 対象名の見出し
      body.push({
        type: 'TextBlock',
        text: `■ ${s.target}`,
        wrap: true,
        weight: 'bolder',
        spacing: 'medium'
      });

      // 数値情報を FactSet で整列表示
      body.push({
        type: 'FactSet',
        separator: false,
        facts: [
          { title: '追加件数', value: String(s.added || 0) },
          { title: '更新件数', value: String(s.updated || 0) },
          { title: '見つからなかった件数', value: String(s.notFound || 0) },
          { title: '追加＋更新 合計', value: String(totalChanged) }
        ]
      });
    });

    // --- 区切り線 ---
    body.push({ type: 'TextBlock', text: '─────────────────────────────', spacing: 'medium', isSubtle: true });

    // --- 時刻情報を FactSet で整列表示 ---
    const minutes = Math.floor(durationSec / 60);
    const seconds = durationSec % 60;
    const durationStr = minutes > 0 ? `${minutes}分${seconds}秒` : `${seconds}秒`;

    body.push({
      type: 'FactSet',
      facts: [
        { title: '処理開始時刻', value: fmtJst24h(startTime) },
        { title: '処理終了時刻', value: fmtJst24h(endTime) },
        { title: '総所要時間', value: durationStr }
      ]
    });
  }

  // Teams通知（失敗しても処理を止めない）
  try {
    sendTeamsAdaptiveCard({ body: body });
  } catch (e) {
    Logger.log("Teams通知失敗: " + e);
  }
}

/**
 * truncateString: 長い文字列を切り詰める
 */
function truncateString(str, maxLength) {
  if (!str) return '';
  return (str.length > maxLength) ? str.substring(0, maxLength) + '...' : str;
}

/**
 * 見出し名から列番号を取得
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {string} headerName - 見出し名
 * @return {number} 列番号（1始まり）、見つからなければ -1
 */
function getColumnByHeader(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(headerName) + 1;
}

/**
 * 設定シートを参照して、各シートの指定列（見出し名）に式を再投入する
 * 設定シート構成例:
 * | シート名 | 列名   | 式テンプレート |
 * |----------|--------|----------------|
 * | 企業情報 | リンク | "https://005108.hammock.hot-profile.com/clients/"&B2 |
 * | 企業情報 | ステータス | IF(C2="","未処理","完了") |
 */
function reapplyFormulasFromSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("設定");
  if (!settingsSheet) return;

  const settings = settingsSheet.getDataRange().getValues();
  for (let i = 1; i < settings.length; i++) { // 1行目はヘッダ
    const [sheetName, headerName, template] = settings[i];
    if (!sheetName || !headerName || !template) continue;

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;

    const col = getColumnByHeader(sheet, headerName);
    if (col < 1) continue;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) continue;

    for (let r = 2; r <= lastRow; r++) {
      const formula = template.replace(/2/g, r); // 行番号を置換
      const finalFormula = formula.startsWith("=") ? formula : "=" + formula;
      sheet.getRange(r, col).setFormula(finalFormula);
    }
  }
}

/**
 * "yyyy/MM/dd HH:mm:ss" / Date / その他文字列をエポック(ms)に変換する
 * 変換できない場合は 0 を返す
 */
function toEpochMs(val) {
  if (!val) return 0;

  // すでに Date 型
  if (val instanceof Date) return val.getTime();

  const s = String(val).trim();
  if (!s) return 0;

  // HotProfileの updated_at 想定形式（例: 2025/04/04 15:38:15）
  try {
    const d = Utilities.parseDate(s, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
    return d ? d.getTime() : 0;
  } catch (e) {
    // フォールバック（形式が揺れた時）
    const d = new Date(s);
    return isNaN(d.getTime()) ? 0 : d.getTime();
  }
}

/**
 * UrlFetchApp.fetch をラップし、エラー時に自動リトライを行う
 * @param {string} url - リクエストURL
 * @param {Object} options - UrlFetchApp.fetch に渡すオプション
 * @param {Object} [retryOptions] - リトライ設定
 * @param {number} [retryOptions.maxRetries=3] - 最大リトライ回数
 * @param {number} [retryOptions.baseDelayMs=1000] - 初回リトライ待機時間（ミリ秒）
 * @param {number[]} [retryOptions.retryStatusCodes=[429, 500, 502, 503, 504]] - リトライ対象のHTTPステータスコード
 * @returns {GoogleAppsScript.URL_Fetch.HTTPResponse} - レスポンス
 * @throws {Error} - 最大リトライ回数を超えた場合
 */
function fetchWithRetry(url, options, retryOptions) {
  const maxRetries = (retryOptions && retryOptions.maxRetries) || 3;
  const baseDelayMs = (retryOptions && retryOptions.baseDelayMs) || 1000;
  const retryStatusCodes = (retryOptions && retryOptions.retryStatusCodes) || [429, 500, 502, 503, 504];

  // muteHttpExceptions を強制的に true に（ステータスコードを自前でハンドリングするため）
  const fetchOptions = Object.assign({}, options, { muteHttpExceptions: true });

  let lastError = null;
  let lastResponse = null;

  for (let attempt = 1; attempt <= maxRetries + 1; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, fetchOptions);
      const status = response.getResponseCode();

      // 成功（2xx）
      if (status >= 200 && status < 300) {
        return response;
      }

      // リトライ対象のステータスコード
      if (retryStatusCodes.indexOf(status) >= 0) {
        lastResponse = response;
        lastError = new Error(`HTTP ${status}: ${truncateString(response.getContentText(), 200)}`);

        if (attempt <= maxRetries) {
          const delayMs = baseDelayMs * Math.pow(2, attempt - 1); // 指数バックオフ
          Logger.log(`APIリトライ実行中 (${attempt}/${maxRetries}): HTTP ${status} - 待機 ${delayMs}ms`);
          Utilities.sleep(delayMs);
          continue;
        }
      }

      // リトライ対象外のエラー（4xx など）はそのまま返す
      return response;

    } catch (e) {
      // ネットワークエラー等の例外
      lastError = e;

      if (attempt <= maxRetries) {
        const delayMs = baseDelayMs * Math.pow(2, attempt - 1);
        Logger.log(`APIリトライ実行中 (${attempt}/${maxRetries}): ${e.message || e} - 待機 ${delayMs}ms`);
        Utilities.sleep(delayMs);
        continue;
      }
    }
  }

  // 全リトライ失敗
  Logger.log(`API呼び出し失敗（リトライ上限超過）: ${lastError}`);
  throw lastError || new Error("fetchWithRetry: 最大リトライ回数を超えました");
}
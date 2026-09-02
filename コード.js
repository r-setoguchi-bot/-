/**
 * Google Chatからメッセージを受け取ったときに実行される関数（30秒ルール完全回避版）
 */
function onMessage(event) {
  // デバッグ用：届いたデータの形をログに記録する
  console.log("受信イベントデータ: " + JSON.stringify(event));

  // 0. メッセージ以外と確実に判明しているイベント（Botのスペース追加・削除、ボタン操作等）のみ
  //    処理対象外としてスキップする（未知の値は誤検知を避けるため通常処理に進める）
  const KNOWN_NON_MESSAGE_EVENT_TYPES = ["ADDED_TO_SPACE", "REMOVED_FROM_SPACE", "CARD_CLICKED"];
  const eventType = event.chat?.type || event.type || "";
  if (KNOWN_NON_MESSAGE_EVENT_TYPES.indexOf(eventType) !== -1) {
    console.log("非メッセージイベントのためスキップ: " + eventType);
    return createChatResponse("");
  }

  // 1. ユーザー名を取得（アドオン形式の深い階層から確実に抽出）
  const userName = event.chat?.messagePayload?.message?.sender?.displayName ||
                   event.chat?.user?.displayName ||
                   "瀬戸口さん";

  // 2. メッセージ本文を取得（アドオン形式の深い階層から確実に抽出）
  const userInput = (event.chat?.messagePayload?.message?.text ||
                    event.message?.text ||
                    "").trim();

  // 3. 非同期プッシュ返信用にスペース名（住所）を取得
  const spaceName = event.chat?.messagePayload?.space?.name ||
                    event.chat?.messagePayload?.message?.space?.name ||
                    event.space?.name ||
                    event.message?.space?.name ||
                    "";

  if (!userInput) {
    return createChatResponse("⚠️ メッセージがうまく読み取れませんでした。");
  }

  // 4. 【30秒制限回避の核心】メッセージを一旦「処理キュー」シートに保存する
  const sheetUrl = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_URL");
  const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  const queueSheet = getRequiredSheet(spreadsheet, "処理キュー");

  const now = new Date();
  const timestamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  const queueId = generateUniqueId("Q", now);

  // 処理キューシートに「Pending（未処理）」状態でタスクを預ける
  queueSheet.appendRow([
    queueId,     // A: キューID
    timestamp,   // B: 受付日時
    userName,    // C: ユーザー名
    userInput,   // D: メッセージ
    spaceName,   // E: スペース名
    "Pending"    // F: ステータス
  ]);

  // 5. チャット画面には「1秒」で受付完了を返し、30秒タイムアウトを完全に回避する
  return createChatResponse("🤖 瀬戸口さん、ご指示を承りました。ただいま秘書が内容を確認し、手帳の整理と解析を行っております。少々お待ちくださいませ…");
}

/**
 * Google Chat（アドオン形式）で正しく表示されるよう、返信データを包む関数
 */
function createChatResponse(text) {
  return {
    "hostAppDataAction": {
      "chatDataAction": {
        "createMessageAction": {
          "message": {
            "text": text
          }
        }
      }
    }
  };
}

/**
 * 指定したシート名でシートを取得する。見つからない場合は例外を投げる
 * （シートの並び替え・削除に対する耐性のため、インデックス指定は使わない）
 */
function getRequiredSheet(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート「${sheetName}」が見つかりません。`);
  return sheet;
}

/**
 * プレフィックス＋タイムスタンプ＋乱数でID文字列を生成する（同一秒内の衝突防止）
 */
function generateUniqueId(prefix, now) {
  const timestamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyyMMddHHmmss");
  const randomSuffix = Math.floor(Math.random() * 900) + 100; // 100〜999の3桁乱数
  return prefix + timestamp + randomSuffix;
}

/**
 * 処理が「Processing」のまま一定時間放置されたキューを「Pending」に戻し、再試行対象にする
 * （GASの実行時間上限などで処理が中断され、行が取り残されるケースの救済）
 */
function recoverStuckQueueRows(queueSheet, data, now) {
  const STUCK_THRESHOLD_MS = 10 * 60 * 1000; // 10分

  for (let i = 1; i < data.length; i++) {
    if (data[i][5] !== "Processing") continue;

    const enqueuedAt = new Date(data[i][1]); // B列: 受付日時
    if (now.getTime() - enqueuedAt.getTime() > STUCK_THRESHOLD_MS) {
      const rowNum = i + 1;
      console.error(`スタックしたキューをPendingに戻して再試行します: 行${rowNum}`);
      queueSheet.getRange(rowNum, 6).setValue("Pending");
      data[i][5] = "Pending"; // 同一ループ内ですぐ処理対象にするためメモリ上のデータも更新
    }
  }
}

/**
 * ⏰ 1分ごとに定期実行され、「処理キュー」シートを巡回するバッチ関数
 */
function processQueue() {
  // 二重起動防止のロックを取得（30秒待機）
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;

  try {
    const sheetUrl = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_URL");
    const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
    const queueSheet = getRequiredSheet(spreadsheet, "処理キュー");
    const lastRow = queueSheet.getLastRow();

    if (lastRow < 2) return; // データがなければ終了

    const data = queueSheet.getDataRange().getValues();
    const now = new Date();

    // 前回実行が中断されて「Processing」のまま止まっている行を再試行対象に戻す
    recoverStuckQueueRows(queueSheet, data, now);

    // 溜まっている Pending 状態のキューを処理する
    for (let i = 1; i < data.length; i++) {
      const status = data[i][5]; // F列: ステータス

      if (status !== "Pending") continue;

      const rowNum = i + 1;
      const userName = data[i][2];  // C列: ユーザー名
      const userInput = data[i][3]; // D列: メッセージ
      const spaceName = data[i][4]; // E列: スペース名

      // ステータスを「Processing（処理中）」に更新
      queueSheet.getRange(rowNum, 6).setValue("Processing");
      SpreadsheetApp.flush();

      try {
        // AI解析メインエンジンを呼び出し、処理結果のテキストを得る
        const replyText = handleUserIntent(userInput, userName, spaceName);

        // サービスアカウントを使ってチャットスペースへ非同期プッシュ送信（返信）
        sendPushReply(spaceName, replyText);

        // ステータスを「Done（完了）」に更新
        queueSheet.getRange(rowNum, 6).setValue("Done");

      } catch (error) {
        console.error("キュー処理エラー: " + error.message);
        queueSheet.getRange(rowNum, 6).setValue("Error: " + error.message);
        try {
          sendPushReply(spaceName, "⚠️ 申し訳ありません。処理中にエラーが発生しました。少し時間をおいて再度お試しください。");
        } catch (notifyError) {
          console.error("エラー通知の送信にも失敗: " + notifyError.message);
        }
      }
      SpreadsheetApp.flush();
    }
  } finally {
    // ロックを解放
    lock.releaseLock();
  }
}

/**
 * 期限が本日・明日・または超過しているタスクについて、Google Chatへリマインドを送信する関数
 * 時間主導型トリガー（1日1回）での実行を想定
 */
function sendTaskReminders() {
  const sheetUrl = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_URL");
  const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  const taskSheet = getRequiredSheet(spreadsheet, "Task_List");
  const data = taskSheet.getDataRange().getValues();

  const now = new Date();
  const todayStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd");
  const tomorrowStr = Utilities.formatDate(new Date(now.getTime() + 24 * 60 * 60 * 1000), "Asia/Tokyo", "yyyy/MM/dd");

  // スペース（送信先）ごとにリマインド内容をまとめるためのバケツ
  const remindersBySpace = {};

  for (let i = 1; i < data.length; i++) {
    const status = data[i][5];         // F列: ステータス
    const deadlineDate = data[i][3];   // D列: 期限日付
    const lastRemindedAt = data[i][6]; // G列: 最終リマインド日時
    const spaceName = data[i][9];      // J列: スペース名

    if (status === "完了" || !deadlineDate || !spaceName) continue;

    // 同じ日に既にリマインド済みならスキップ（重複送信防止）
    const lastRemindedDateStr = lastRemindedAt
      ? Utilities.formatDate(new Date(lastRemindedAt), "Asia/Tokyo", "yyyy/MM/dd")
      : "";
    if (lastRemindedDateStr === todayStr) continue;

    let label = "";
    if (deadlineDate < todayStr) {
      label = "⚠️ 期限超過";
    } else if (deadlineDate === todayStr) {
      label = "⏰ 本日期限";
    } else if (deadlineDate === tomorrowStr) {
      label = "📅 明日期限";
    } else {
      continue; // まだ期限に余裕があるタスクは対象外
    }

    const taskContent = data[i][2];    // C列: タスク内容
    const assignedPerson = data[i][7]; // H列: 担当者

    if (!remindersBySpace[spaceName]) remindersBySpace[spaceName] = [];
    remindersBySpace[spaceName].push({
      rowNum: i + 1,
      line: `${label}：「${taskContent}」（期限 ${deadlineDate}／担当：${assignedPerson}）`
    });
  }

  for (const spaceName in remindersBySpace) {
    const items = remindersBySpace[spaceName];
    const message = "🤖 秘書より、本日のタスクリマインドです。\n\n" + items.map(item => item.line).join("\n");

    try {
      sendPushReply(spaceName, message);
      // 送信に成功したタスクのみ、最終リマインド日時を更新する
      const timestamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
      items.forEach(item => taskSheet.getRange(item.rowNum, 7).setValue(timestamp));
    } catch (e) {
      console.error("リマインド送信に失敗しました（" + spaceName + "）: " + e.message);
    }
  }
}

/**
 * sendTaskRemindersを1日1回（午前9時）自動実行するトリガーを設定する関数
 * GASエディタでこの関数を1回だけ手動実行してください
 */
function setupReminderTrigger() {
  // 重複登録を防ぐため、既存の同名トリガーを削除してから作り直す
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "sendTaskReminders") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("sendTaskReminders")
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
}

/**
 * ====================================================================
 * kintone連携：日報自動生成
 * 前日と本日のkintoneデータを比較し、差分（＝今日動きがあった内容）を検出して日報を組み立てる
 * ====================================================================
 */

// 日報の対象とするkintoneアプリ一覧。それぞれのアプリID・APIトークンはスクリプトプロパティで管理する
const KINTONE_APPS = [
  { key: "案件管理",     idProp: "KINTONE_ANKEN_APP_ID",     tokenProp: "KINTONE_ANKEN_API_TOKEN",     snapshotSheet: "kintone_案件管理_前日" },
  { key: "契約管理",     idProp: "KINTONE_KEIYAKU_APP_ID",   tokenProp: "KINTONE_KEIYAKU_API_TOKEN",   snapshotSheet: "kintone_契約管理_前日" },
  { key: "粗大ゴミ管理", idProp: "KINTONE_SODAIGOMI_APP_ID", tokenProp: "KINTONE_SODAIGOMI_API_TOKEN", snapshotSheet: "kintone_粗大ゴミ管理_前日" },
  { key: "解約リスト",   idProp: "KINTONE_KAIYAKU_APP_ID",   tokenProp: "KINTONE_KAIYAKU_API_TOKEN",   snapshotSheet: "kintone_解約リスト_前日" }
];

/**
 * kintoneアプリの全レコードをREST APIで取得する（$idベースでページング）
 */
function fetchKintoneRecords(subdomain, appId, apiToken) {
  let allRecords = [];
  let lastId = 0;
  const limit = 500;

  while (true) {
    const query = encodeURIComponent(`$id > ${lastId} order by $id asc limit ${limit}`);
    const url = `https://${subdomain}.cybozu.com/k/v1/records.json?app=${appId}&query=${query}`;
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "X-Cybozu-API-Token": apiToken },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`kintone取得失敗 (app=${appId}, HTTP ${response.getResponseCode()}): ${response.getContentText()}`);
    }

    const records = JSON.parse(response.getContentText()).records || [];
    if (records.length === 0) break;

    allRecords = allRecords.concat(records);
    lastId = Number(records[records.length - 1].$id.value);
    if (records.length < limit) break;
  }

  return allRecords;
}

/**
 * kintoneレコード1件を「フィールドコード→値（文字列化）」のマップに変換する（差分比較用）
 */
function recordToValueMap(record) {
  const valueMap = {};
  Object.keys(record).forEach(fieldCode => {
    if (fieldCode.indexOf("$") === 0) return; // $id, $revisionなどのシステムフィールドは除外
    const field = record[fieldCode];
    if (!field || field.value === undefined || field.value === null || field.value === "") return;
    valueMap[fieldCode] = typeof field.value === "object" ? JSON.stringify(field.value) : String(field.value);
  });
  return valueMap;
}

/**
 * 前日スナップショットをシートから読み込む。シートが無ければ「初回実行」としてnullを返す
 */
function loadKintoneSnapshot(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  const snapshot = {};
  for (let i = 1; i < data.length; i++) {
    const recordId = data[i][0];
    const fieldCode = data[i][1];
    const value = data[i][2];
    if (!recordId) continue;
    if (!snapshot[recordId]) snapshot[recordId] = {};
    snapshot[recordId][fieldCode] = value;
  }
  return snapshot;
}

/**
 * 今回取得したスナップショットをシートへ保存する（次回実行時の比較対象として上書きする）
 */
function saveKintoneSnapshot(spreadsheet, sheetName, recordsById) {
  const sheet = getOrCreateSheet(spreadsheet, sheetName);
  sheet.clear();

  const rows = [["recordId", "fieldCode", "value"]];
  Object.keys(recordsById).forEach(recordId => {
    const valueMap = recordsById[recordId];
    Object.keys(valueMap).forEach(fieldCode => {
      rows.push([recordId, fieldCode, valueMap[fieldCode]]);
    });
  });

  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
}

/**
 * 指定したシート名のシートを取得し、無ければ新規作成する
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  return sheet;
}

/**
 * 1アプリ分の新旧データを比較し、「新規追加」「フィールド変更」を検出してテキスト行にする
 */
function diffKintoneRecords(oldSnapshot, newRecords) {
  const lines = [];
  const newRecordsById = {};

  newRecords.forEach(record => {
    const recordId = record.$id.value;
    const valueMap = recordToValueMap(record);
    newRecordsById[recordId] = valueMap;

    const displayName = valueMap["屋号名"] || valueMap["契約店舗名称"] || valueMap["会社名"] || `レコード#${recordId}`;
    const oldValueMap = oldSnapshot[recordId];

    if (!oldValueMap) {
      lines.push(`【新規】${displayName}`);
      return;
    }

    const changedFields = [];
    Object.keys(valueMap).forEach(fieldCode => {
      const oldValue = oldValueMap[fieldCode] || "";
      const newValue = valueMap[fieldCode];
      if (oldValue !== newValue) {
        changedFields.push(`${fieldCode}: ${oldValue || "(空)"} → ${newValue}`);
      }
    });

    if (changedFields.length > 0) {
      lines.push(`【更新】${displayName}\n  ` + changedFields.join("\n  "));
    }
  });

  return { lines, newRecordsById };
}

/**
 * 処理キューシートから、直近で使われたスペース名（日報の送信先）を取得する
 */
function getDefaultSpaceName(spreadsheet) {
  const queueSheet = getRequiredSheet(spreadsheet, "処理キュー");
  const lastRow = queueSheet.getLastRow();
  if (lastRow < 2) return "";

  const data = queueSheet.getRange(2, 5, lastRow - 1, 1).getValues(); // E列: スペース名
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0]) return data[i][0];
  }
  return "";
}

/**
 * kintoneの4アプリ（案件管理・契約管理・粗大ゴミ管理・解約リスト）の前日比差分と、
 * 秘書bot自身の本日の対応履歴をまとめて日報としてChatへ送信する関数
 * 時間主導型トリガー（1日1回）での実行を想定
 */
function sendDailyReport() {
  const subdomain = PropertiesService.getScriptProperties().getProperty("KINTONE_SUBDOMAIN");
  if (!subdomain) {
    console.error("KINTONE_SUBDOMAINが未設定のため日報生成をスキップしました。");
    return;
  }

  const sheetUrl = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_URL");
  const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  const sections = [];

  KINTONE_APPS.forEach(app => {
    const appId = PropertiesService.getScriptProperties().getProperty(app.idProp);
    const apiToken = PropertiesService.getScriptProperties().getProperty(app.tokenProp);
    if (!appId || !apiToken) {
      console.log(`${app.key}: アプリID/APIトークンが未設定のためスキップします。`);
      return;
    }

    try {
      const oldSnapshot = loadKintoneSnapshot(spreadsheet, app.snapshotSheet);
      const newRecords = fetchKintoneRecords(subdomain, appId, apiToken);

      if (oldSnapshot === null) {
        // 初回実行：比較対象がまだ無いため、今回はスナップショットの作成のみ行う
        const initialSnapshot = {};
        newRecords.forEach(record => {
          initialSnapshot[record.$id.value] = recordToValueMap(record);
        });
        saveKintoneSnapshot(spreadsheet, app.snapshotSheet, initialSnapshot);
        console.log(`${app.key}: 初回スナップショットを作成しました（${newRecords.length}件）。次回から差分を検出します。`);
        return;
      }

      const { lines, newRecordsById } = diffKintoneRecords(oldSnapshot, newRecords);
      if (lines.length > 0) {
        sections.push(`■${app.key}\n` + lines.join("\n"));
      }

      saveKintoneSnapshot(spreadsheet, app.snapshotSheet, newRecordsById);
    } catch (e) {
      console.error(`${app.key}の日報生成中にエラーが発生しました: ` + e.message);
      sections.push(`■${app.key}\n⚠️ データ取得に失敗しました（${e.message}）`);
    }
  });

  // 秘書bot経由の本日の対応履歴も日報に追加する
  const logSheet = getRequiredSheet(spreadsheet, "対応履歴ログ");
  const logData = logSheet.getDataRange().getValues();
  const todayStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
  const todayLogs = [];
  for (let i = 1; i < logData.length; i++) {
    const recordedAt = String(logData[i][2]); // C列: 記録日時
    if (recordedAt.indexOf(todayStr) === 0) {
      todayLogs.push(`・${logData[i][3]}（担当: ${logData[i][4]}）`); // D列: 対応内容・経緯, E列: 担当者
    }
  }
  if (todayLogs.length > 0) {
    sections.push("■秘書bot対応履歴\n" + todayLogs.join("\n"));
  }

  if (sections.length === 0) {
    console.log("本日は日報に含める動きがありませんでした。");
    return;
  }

  const spaceName = getDefaultSpaceName(spreadsheet);
  if (!spaceName) {
    console.error("送信先スペースが特定できなかったため日報を送信できませんでした。");
    return;
  }

  const message = `🤖 本日（${todayStr}）の日報です。\n\n` + sections.join("\n\n");
  sendPushReply(spaceName, message);
}

/**
 * sendDailyReportを1日1回（午後6時）自動実行するトリガーを設定する関数
 * GASエディタでこの関数を1回だけ手動実行してください
 */
function setupDailyReportTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "sendDailyReport") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("sendDailyReport")
    .timeBased()
    .atHour(18)
    .everyDays(1)
    .create();
}

/**
 * ====================================================================
 * kintone連携：請求単価チェック（見積り添付ファイルとの突合）
 * 契約管理アプリの各レコードについて、添付された見積りファイル（PDF/Excel/画像）から
 * AIで金額を読み取り、現在の「請求単価」フィールドと一致しているかを確認する。
 * ※このチェックはkintoneへの書き込みは一切行わず、結果をChatとログシートへ出力するのみ。
 * ====================================================================
 */

// このチェックで参照するスクリプトプロパティ名（kintone管理画面でフィールドコードを確認して設定する）
const BILLING_RATE_CHECK_CONFIG = {
  appIdProp: "KINTONE_KEIYAKU_APP_ID",
  apiTokenProp: "KINTONE_KEIYAKU_API_TOKEN",
  mitsumoriFieldProp: "KINTONE_KEIYAKU_MITSUMORI_FIELD", // 見積り添付ファイルのフィールドコード
  tankaFieldProp: "KINTONE_KEIYAKU_TANKA_FIELD",         // 請求単価のフィールドコード
  resultSheetName: "請求単価チェック結果"
};

/**
 * 契約管理アプリの全レコードを対象に、見積り添付ファイルの金額と請求単価フィールドを突合する関数
 * 時間主導型トリガー、またはGASエディタからの手動実行を想定
 */
function checkBillingRates() {
  const props = PropertiesService.getScriptProperties();
  const subdomain = props.getProperty("KINTONE_SUBDOMAIN");
  const appId = props.getProperty(BILLING_RATE_CHECK_CONFIG.appIdProp);
  const apiToken = props.getProperty(BILLING_RATE_CHECK_CONFIG.apiTokenProp);
  const mitsumoriField = props.getProperty(BILLING_RATE_CHECK_CONFIG.mitsumoriFieldProp);
  const tankaField = props.getProperty(BILLING_RATE_CHECK_CONFIG.tankaFieldProp);
  const geminiApiKey = props.getProperty("GEMINI_API_KEY");

  if (!subdomain || !appId || !apiToken || !mitsumoriField || !tankaField || !geminiApiKey) {
    console.error("請求単価チェックに必要なスクリプトプロパティが不足しています（KINTONE_SUBDOMAIN / " +
      BILLING_RATE_CHECK_CONFIG.appIdProp + " / " + BILLING_RATE_CHECK_CONFIG.apiTokenProp + " / " +
      BILLING_RATE_CHECK_CONFIG.mitsumoriFieldProp + " / " + BILLING_RATE_CHECK_CONFIG.tankaFieldProp + " / GEMINI_API_KEY）。");
    return;
  }

  const records = fetchKintoneRecords(subdomain, appId, apiToken);
  const results = [];

  records.forEach(record => {
    const recordId = record.$id.value;
    const displayName = (record["屋号名"] && record["屋号名"].value) ||
                         (record["契約店舗名称"] && record["契約店舗名称"].value) ||
                         (record["会社名"] && record["会社名"].value) ||
                         `レコード#${recordId}`;

    const currentTankaRaw = record[tankaField] ? record[tankaField].value : "";
    const files = (record[mitsumoriField] && record[mitsumoriField].value) || [];

    if (files.length === 0) {
      results.push({ recordId, displayName, status: "見積り未添付", current: currentTankaRaw, extracted: "", note: "" });
      return;
    }

    try {
      const file = files[0]; // 複数添付されている場合は先頭（最新想定）のみをチェック対象にする
      const blob = fetchKintoneFile(subdomain, file.fileKey, apiToken);
      const extraction = extractEstimateAmount(blob, file.contentType, geminiApiKey);

      if (extraction.amount === null) {
        results.push({ recordId, displayName, status: "抽出失敗", current: currentTankaRaw, extracted: "", note: extraction.note || "" });
        return;
      }

      const currentTankaNum = parseAmount(currentTankaRaw);
      if (currentTankaNum === null) {
        results.push({ recordId, displayName, status: "請求単価未入力", current: currentTankaRaw, extracted: extraction.amount, note: extraction.note || "" });
      } else if (currentTankaNum === extraction.amount) {
        results.push({ recordId, displayName, status: "一致", current: currentTankaRaw, extracted: extraction.amount, note: "" });
      } else {
        results.push({ recordId, displayName, status: "不一致", current: currentTankaRaw, extracted: extraction.amount, note: extraction.note || "" });
      }
    } catch (e) {
      console.error(`請求単価チェック中にエラー（レコード#${recordId}）: ` + e.message);
      results.push({ recordId, displayName, status: "エラー", current: currentTankaRaw, extracted: "", note: e.message });
    }
  });

  saveBillingRateCheckResults(results);
  sendBillingRateCheckReport(results);
}

/**
 * kintoneの添付ファイルを1件ダウンロードする（file.jsonエンドポイント）
 */
function fetchKintoneFile(subdomain, fileKey, apiToken) {
  const url = `https://${subdomain}.cybozu.com/k/v1/file.json?fileKey=${fileKey}`;
  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { "X-Cybozu-API-Token": apiToken },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`kintoneファイル取得失敗 (HTTP ${response.getResponseCode()}): ${response.getContentText()}`);
  }
  return response.getBlob();
}

/**
 * 見積りファイル（PDF/画像/Excel）から請求すべき金額をAIで抽出する
 * PDF・画像はGeminiのマルチモーダル入力でそのまま読み取り、
 * Excel（.xlsx）はセルの文字列を抽出してテキストとしてAIに渡す
 */
function extractEstimateAmount(blob, contentType, apiKey) {
  const promptBase = "これは取引先への見積書です。この見積書に記載されている、実際に請求すべき単価・金額" +
    "（値引き後の最終的な金額。消費税込みの合計金額を優先）を1つだけ特定してください。" +
    "複数の商品/サービスが並んでいる場合は、個別の内訳ではなく合計金額（税込）を優先してください。" +
    "金額がどうしても読み取れない場合はamountをnullにしてください。" +
    "説明文などは一切含めず、次のJSON形式の文字列のみを出力してください:\n" +
    '{"amount": 数値（円、カンマなし。税込金額）またはnull, "note": "抽出時に気になった点があれば一言（無ければ空文字）"}';

  let parts;
  if (contentType === "application/pdf" || contentType.indexOf("image/") === 0) {
    parts = [
      { text: promptBase },
      { inlineData: { mimeType: contentType, data: Utilities.base64Encode(blob.getBytes()) } }
    ];
  } else if (contentType.indexOf("spreadsheetml") !== -1) {
    const sheetText = extractTextFromXlsx(blob);
    if (!sheetText) return { amount: null, note: "Excelファイルの内容を読み取れませんでした。" };
    parts = [{ text: promptBase + "\n\n【見積書の内容（セルの値を抽出したもの）】\n" + sheetText }];
  } else {
    return { amount: null, note: `未対応のファイル形式です（${contentType}）。` };
  }

  return callGeminiForAmount(parts, apiKey);
}

/**
 * Gemini APIへ画像/PDF/テキストを渡し、金額抽出結果のJSONを解析して返す
 */
function callGeminiForAmount(parts, apiKey) {
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash-lite:generateContent?key=" + apiKey;
  const payload = {
    "contents": [{ "parts": parts }],
    "generationConfig": { "responseMimeType": "application/json" }
  };
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    const text = json.candidates && json.candidates[0].content.parts[0].text;
    if (!text) return { amount: null, note: "AIから金額を読み取れませんでした。" };

    const cleanJsonStr = text.replace(/```json/g, "").replace(/```/g, "").trim();
    const parsed = JSON.parse(cleanJsonStr);
    const amount = (parsed.amount === null || parsed.amount === undefined) ? null : Number(parsed.amount);
    return { amount: (amount === null || isNaN(amount)) ? null : amount, note: parsed.note || "" };
  } catch (e) {
    console.error("見積り金額抽出のGemini呼び出しでエラー: " + e.message);
    return { amount: null, note: "AI呼び出し中にエラーが発生しました。" };
  }
}

/**
 * カンマ・円マークなどを含む文字列を数値に変換する（変換できない場合はnull）
 */
function parseAmount(raw) {
  if (raw === "" || raw === null || raw === undefined) return null;
  const cleaned = String(raw).replace(/[^\d.]/g, "");
  if (cleaned === "") return null;
  const num = Number(cleaned);
  return isNaN(num) ? null : num;
}

/**
 * xlsx（zip形式）から共有文字列・シートXMLを取り出し、セルの値をタブ区切りテキストへ変換する
 * Drive APIの追加権限を使わずに済むよう、Utilities.unzipで直接パースする
 */
function extractTextFromXlsx(blob) {
  try {
    const entries = Utilities.unzip(blob);
    let sharedStrings = [];
    const sheetTexts = [];

    entries.forEach(entry => {
      if (entry.getName() === "xl/sharedStrings.xml") {
        sharedStrings = parseSharedStringsXml(entry.getDataAsString());
      }
    });

    entries.forEach(entry => {
      if (/^xl\/worksheets\/sheet\d+\.xml$/.test(entry.getName())) {
        sheetTexts.push(parseSheetXml(entry.getDataAsString(), sharedStrings));
      }
    });

    return sheetTexts.join("\n").trim();
  } catch (e) {
    console.error("xlsx解析エラー: " + e.message);
    return "";
  }
}

function parseSharedStringsXml(xml) {
  const result = [];
  const siRegex = /<si>([\s\S]*?)<\/si>/g;
  let siMatch;
  while ((siMatch = siRegex.exec(xml)) !== null) {
    const tRegex = /<t[^>]*>([\s\S]*?)<\/t>/g;
    let text = "";
    let tMatch;
    while ((tMatch = tRegex.exec(siMatch[1])) !== null) {
      text += tMatch[1];
    }
    result.push(decodeXmlEntities(text));
  }
  return result;
}

function parseSheetXml(xml, sharedStrings) {
  const lines = [];
  const rowRegex = /<row\b[^>]*>([\s\S]*?)<\/row>/g;
  let rowMatch;
  while ((rowMatch = rowRegex.exec(xml)) !== null) {
    const cellRegex = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
    const cells = [];
    let cellMatch;
    while ((cellMatch = cellRegex.exec(rowMatch[1])) !== null) {
      const attrs = cellMatch[1];
      const cellBody = cellMatch[2];
      const typeMatch = /\st="([^"]*)"/.exec(attrs);
      const cellType = typeMatch ? typeMatch[1] : null;

      let value = "";
      if (cellType === "s") {
        const vMatch = /<v>([\s\S]*?)<\/v>/.exec(cellBody);
        if (vMatch) value = sharedStrings[Number(vMatch[1])] || "";
      } else if (cellType === "inlineStr") {
        const tMatch = /<t[^>]*>([\s\S]*?)<\/t>/.exec(cellBody);
        if (tMatch) value = decodeXmlEntities(tMatch[1]);
      } else {
        const vMatch = /<v>([\s\S]*?)<\/v>/.exec(cellBody);
        if (vMatch) value = vMatch[1];
      }
      if (value !== "") cells.push(value);
    }
    if (cells.length > 0) lines.push(cells.join("\t"));
  }
  return lines.join("\n");
}

function decodeXmlEntities(str) {
  return str
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

/**
 * チェック結果をスプレッドシートへ保存する（毎回上書き。監査・詳細確認用）
 */
function saveBillingRateCheckResults(results) {
  const sheetUrl = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_URL");
  const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  const sheet = getOrCreateSheet(spreadsheet, BILLING_RATE_CHECK_CONFIG.resultSheetName);
  sheet.clear();

  const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  const rows = [["チェック日時", "レコードID", "契約先", "ステータス", "現在の請求単価", "見積りから読み取った金額", "備考"]];
  results.forEach(r => {
    rows.push([now, r.recordId, r.displayName, r.status, r.current, r.extracted, r.note]);
  });
  sheet.getRange(1, 1, rows.length, 7).setValues(rows);
}

/**
 * チェック結果のうち「一致」以外（要確認）の項目のみをChatへ要約送信する
 */
function sendBillingRateCheckReport(results) {
  const sheetUrl = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_URL");
  const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  const spaceName = getDefaultSpaceName(spreadsheet);
  if (!spaceName) {
    console.error("送信先スペースが特定できなかったため請求単価チェック結果を送信できませんでした。");
    return;
  }

  const needsAttention = results.filter(r => r.status !== "一致");
  const okCount = results.length - needsAttention.length;
  const MAX_LINES = 20;

  let message = `🤖 請求単価チェック結果（契約管理アプリ）\n対象 ${results.length}件中、一致 ${okCount}件・要確認 ${needsAttention.length}件でした。\n\n`;

  if (needsAttention.length === 0) {
    message += "すべて一致していました。特に対応は不要です。";
  } else {
    message += needsAttention.slice(0, MAX_LINES).map(r => {
      let line = `【${r.status}】${r.displayName}（現在: ${r.current || "(空)"}`;
      if (r.extracted !== "") line += ` / 見積り: ${r.extracted}`;
      line += "）";
      if (r.note) line += `\n  備考: ${r.note}`;
      return line;
    }).join("\n");

    if (needsAttention.length > MAX_LINES) {
      message += `\n…ほか${needsAttention.length - MAX_LINES}件`;
    }
    message += `\n\n詳細は「${BILLING_RATE_CHECK_CONFIG.resultSheetName}」シートをご確認ください。`;
  }

  try {
    sendPushReply(spaceName, message);
  } catch (e) {
    console.error("請求単価チェック結果の送信に失敗しました: " + e.message);
  }
}

/**
 * checkBillingRatesを1日1回（午前8時）自動実行するトリガーを設定する関数
 * 自動チェックを開始したいタイミングでGASエディタからこの関数を1回だけ手動実行してください
 */
function setupBillingRateCheckTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "checkBillingRates") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("checkBillingRates")
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
}

/**
 * Gemini APIを呼び出してユーザーの意図（新規タスク・進捗更新・要約・会話）を判断し、適切に処理する関数
 */
function handleUserIntent(userInput, userName, spaceName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) return "⚠️ エラー: APIキーが設定されていません。";

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash-lite:generateContent?key=" + apiKey;

  // 「明日」「来週」などの言葉を理解できるよう、現在の日時を教える
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy年MM月dd日 HH:mm");

  const systemPrompt = `
あなたは瀬戸口さんの優秀な専属秘書botです。
現在の日時は【${todayStr}】です。

瀬戸口さん（${userName}）からのメッセージ:
「${userInput}」

【指示】
瀬戸口さんのメッセージの意図を以下の4つのいずれかに分類し、必ず指定されたJSONフォーマットの文字列のみを出力してください（余計な解説文などは一切含めないでください）。

1. メッセージが「新しいタスクの依頼や追加」である場合（例：「明日までに資料作成」「タスクに追加して」など）
{
  "action": "task",
  "taskContent": "タスクの具体的な内容",
  "deadlineDate": "YYYY/MM/DD (期限が不明な場合は空文字)",
  "deadlineTime": "HH:mm (時間が不明な場合は空文字)",
  "assignedPerson": "このタスクの担当者（部下の名前や瀬戸口さんの名前など。明示がなければ「瀬戸口さん」）",
  "statusSummary": "新規タスク登録",
  "replyMessage": "タスクを新規登録した旨の、秘書らしい丁寧な返答"
}

2. メッセージが「既存タスクの進捗更新・経緯の追記・完了報告」である場合（例：「資料作成の件、A社から来週まで待ってと連絡あった」「週報タスク終わったよ」「〇〇の件は今山田くんがボール持ってる」など）
{
  "action": "update_task",
  "targetTaskKeyword": "どの既存タスクに関する進捗か特定するためのキーワード（タスク内容に含まれていそうな象徴的な言葉）",
  "progressDetail": "今回新しく記録する対応内容や経緯の具体的なディテール",
  "assignedPerson": "現在のボール保持者・担当者（誰のところで止まっているか、誰が今動くべきか）",
  "statusSummary": "最新の状況を一言で表すサマリー（例：「先方連絡待ち」「完了」「山田さん対応中」など）",
  "taskStatus": "タスクの最新ステータス（「完了」または「未完了」または「先方返信待ち」など）",
  "replyMessage": "進捗履歴をしっかりとノートに記録した旨の、秘書らしい丁寧な報告返答"
}

3. メッセージが「文章の要約や下書きの作成」を求めている場合（例：「この文章を要約して」「返信メールを作って」など）
{
  "action": "summary_or_draft",
  "replyMessage": "しっかり要約・下書きした内容を、秘書らしく丁寧に見やすく整えた文章"
}

4. それ以外の「ただの会話や質問」の場合
{
  "action": "chat",
  "replyMessage": "秘書らしい丁寧で明るい返答"
}
`;

  const payload = {
    "contents": [{ "parts": [{ "text": systemPrompt }] }],
    "generationConfig": {
      "responseMimeType": "application/json"
    }
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    if (json.candidates && json.candidates[0].content.parts[0].text) {
      const aiResponse = json.candidates[0].content.parts[0].text;
      
      try {
        // AIが返したJSON文字列をプログラムで使える形（オブジェクト）に変換
        const cleanJsonStr = aiResponse.replace(/```json/g, "").replace(/```/g, "").trim();
        const resultData = JSON.parse(cleanJsonStr);

        // 新規タスク登録の処理
        if (resultData.action === "task") {
          writeTaskToSheet(resultData, userName, spaceName);
          return resultData.replyMessage;
        }
        
        // 既存タスクの進捗更新処理（新機能！）
        if (resultData.action === "update_task") {
          return updateTaskInSheets(resultData, userName);
        }
        
        // 要約結果や下書き、会話などそのままチャットに返す処理
        return resultData.replyMessage;

      } catch (parseError) {
        console.error("AI応答のJSON解析に失敗しました: " + parseError.message + " / 応答内容: " + aiResponse);
        return "⚠️ 申し訳ありません。AIからの応答をうまく読み取れませんでした。お手数ですが、もう一度お試しください。";
      }
    } else {
      console.error("Geminiの応答に想定した内容が含まれていませんでした: " + JSON.stringify(json));
      return "⚠️ 申し訳ありません。ただいまAIとの通信で問題が発生しております。少し時間をおいて再度お試しください。";
    }
  } catch (e) {
    console.error("Gemini API呼び出しでエラーが発生しました: " + e.toString());
    return "⚠️ 申し訳ありません。ただいまシステムに問題が発生しております。少し時間をおいて再度お試しください。";
  }
}

/**
 * スプレッドシートに新規タスクを書き込む関数
 */
function writeTaskToSheet(taskData, userName, spaceName) {
  const sheetUrl = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_URL");
  const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  const sheet = getRequiredSheet(spreadsheet, "Task_List");

  const now = new Date();
  const timestamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

  // タスクID（T + 年月日時間 + 乱数 で自動生成）
  const taskId = generateUniqueId("T", now);

  // スプレッドシートの各列（A〜J列）に合わせてデータを追加
  sheet.appendRow([
    taskId,                              // A: タスクID
    timestamp,                           // B: 登録日時
    taskData.taskContent,                // C: タスク内容
    taskData.deadlineDate,               // D: 期限（日付）
    taskData.deadlineTime,               // E: 期限（時間）
    "未完了",                            // F: ステータス
    "",                                  // G: 最終リマインド日時
    taskData.assignedPerson || userName, // H: 担当者（ボール保持者）
    taskData.statusSummary || "新規登録",// I: 最新の状況サマリー
    spaceName || ""                      // J: スペース名（リマインド送信先）
  ]);
}

/**
 * タスク一覧のデータからキーワードに一致する行を全て探す（配列インデックスの配列を返す）。
 * 最新のタスクを優先するため下の行から逆順に検索する。
 * onlyIncomplete=true の場合は「完了」ステータスの行を除外して検索する（誤爆防止）。
 */
function findMatchingTaskRowIndexes(taskDataRange, keyword, onlyIncomplete) {
  const matches = [];
  for (let i = taskDataRange.length - 1; i >= 1; i--) {
    const status = String(taskDataRange[i][5]); // F列: ステータス
    if (onlyIncomplete && status === "完了") continue;

    const taskContent = String(taskDataRange[i][2]).toLowerCase(); // C列: タスク内容
    if (taskContent.indexOf(keyword) !== -1) {
      matches.push(i);
    }
  }
  return matches;
}

/**
 * 既存タスクを検索し、進捗の上書きと履歴シートへの追記を行う関数
 */
function updateTaskInSheets(updateData, userName) {
  const sheetUrl = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_URL");
  const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  const taskSheet = getRequiredSheet(spreadsheet, "Task_List");
  const logSheet = getRequiredSheet(spreadsheet, "対応履歴ログ");

  const taskDataRange = taskSheet.getDataRange().getValues();
  const keyword = updateData.targetTaskKeyword.toLowerCase();

  // 誤爆防止のため、まず「未完了」タスクの中から探し、
  // 見つからなかった場合のみ完了済みタスクも含めて再検索する
  let matches = findMatchingTaskRowIndexes(taskDataRange, keyword, true);
  if (matches.length === 0) {
    matches = findMatchingTaskRowIndexes(taskDataRange, keyword, false);
  }

  if (matches.length === 0) {
    // 検索キーワードに一致するタスクがなかった場合
    return `⚠️ 秘書より報告です。「${updateData.targetTaskKeyword}」に該当するタスクが一覧から見つけられませんでした。お手数ですが、正確な名前でご指示いただくか、新規タスクとしてご登録ください。`;
  }

  if (matches.length > 1) {
    // 候補が複数ある場合は誤爆を避けるため、自動で決めつけずユーザーに確認を求める
    const candidateList = matches.slice(0, 5)
      .map(i => `・${taskDataRange[i][0]}: ${taskDataRange[i][2]}`)
      .join("\n");
    return `⚠️ 秘書より確認です。「${updateData.targetTaskKeyword}」に該当するタスクが複数見つかりました。誤って別のタスクを更新しないよう、より具体的なキーワードで改めてご指示いただけますか。\n\n候補:\n${candidateList}`;
  }

  const matchedIndex = matches[0];
  const foundRowIndex = matchedIndex + 1; // 行番号（1始まり）
  const taskId = taskDataRange[matchedIndex][0]; // A列: タスクID

  const now = new Date();
  const timestamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

  // 1. 1枚目の「タスク一覧」シートの最新情報を上書き更新
  if (updateData.taskStatus) {
    taskSheet.getRange(foundRowIndex, 6).setValue(updateData.taskStatus);     // F列: ステータス
  }
  if (updateData.assignedPerson) {
    taskSheet.getRange(foundRowIndex, 8).setValue(updateData.assignedPerson); // H列: 担当者（ボール保持者）
  }
  if (updateData.statusSummary) {
    taskSheet.getRange(foundRowIndex, 9).setValue(updateData.statusSummary);  // I列: 最新の状況サマリー
  }

  // 2. 2枚目の「対応履歴ログ」シートに歴史（タイムライン）を1行追記
  const logId = generateUniqueId("L", now);
  logSheet.appendRow([
    logId,                     // A: 履歴ID
    taskId,                    // B: タスクID（紐付け用）
    timestamp,                 // C: 記録日時
    updateData.progressDetail, // D: 対応内容・経緯
    updateData.assignedPerson  // E: 担当者（ボール保持者）
  ]);

  return updateData.replyMessage;
}

/**
 * 外部の処理からGoogle Chatへメッセージを非同期プッシュ送信する関数（過去の成功コードから移植）
 */
function sendPushReply(spaceName, text) {
  if (!spaceName) return;
  const token = getBotToken();
  const payload = {
    "text": text
  };
  const response = UrlFetchApp.fetch(`https://chat.googleapis.com/v1/${spaceName}/messages`, {
    method: "post",
    headers: { "Authorization": "Bearer " + token },
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const responseCode = response.getResponseCode();
  if (responseCode >= 300) {
    throw new Error(`Chat送信失敗 (HTTP ${responseCode}): ${response.getContentText()}`);
  }
}

/**
 * サービスアカウントのJSONから認証トークンを取得する関数（過去の成功コードから移植）
 */
function getBotToken() {
  const jsonStr = PropertiesService.getScriptProperties().getProperty("SERVICE_ACCOUNT_JSON");
  if (!jsonStr) throw new Error("スクリプトプロパティ 'SERVICE_ACCOUNT_JSON' が未設定です。");
  const key = JSON.parse(jsonStr);
  const service = OAuth2.createService('ChatBotService')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setPrivateKey(key.private_key)
    .setIssuer(key.client_email)
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setScope('https://www.googleapis.com/auth/chat.bot');
  if (!service.hasAccess()) throw new Error("Bot認証に失敗しました: " + service.getLastError());
  return service.getAccessToken();
}
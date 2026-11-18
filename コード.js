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
        // （「もう一回やって」などに対応できるよう、直前のやり取りも渡す）
        const recentConversation = getRecentConversation(data, i, spaceName, new Date());
        let replyText = handleUserIntent(userInput, userName, spaceName, recentConversation);

        // 万一返答が空でも、Chatへの送信が失敗しないよう定型文に置き換える
        if (!replyText || !String(replyText).trim()) {
          replyText = "⚠️ 申し訳ありません。ご指示の内容をうまく読み取れませんでした。お手数ですが、言い方を変えてもう一度お試しください。";
        }

        // サービスアカウントを使ってチャットスペースへ非同期プッシュ送信（返信）
        // 長い返答は、Chatの文字数上限に収まるよう複数通に分けて送る
        splitMessageForChat(String(replyText), 3800).forEach(part => sendPushReply(spaceName, part));

        // ステータスを「Done（完了）」に更新
        queueSheet.getRange(rowNum, 6).setValue("Done");
        data[i][5] = "Done"; // 同じ回の後続メッセージが「直前のやり取り」として参照できるように

      } catch (error) {
        console.error("キュー処理エラー: " + error.message);
        queueSheet.getRange(rowNum, 6).setValue("Error: " + error.message);
        data[i][5] = "Error: " + error.message;
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
 * ====================================================================
 * Googleカレンダー連携：日付まわりの共通関数
 * （スクリプトのタイムゾーンは appsscript.json で Asia/Tokyo に設定済み）
 * ====================================================================
 */

const WEEKDAY_LABELS = ["日", "月", "火", "水", "木", "金", "土"];

/**
 * 指定した日時の「その日の0時0分」を返す
 */
function startOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/**
 * 指定した日時から days 日後の日時を返す（時刻はそのまま）
 */
function addDays(date, days) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() + days, date.getHours(), date.getMinutes());
}

/**
 * 日付を「11/20(金)」の形式にする
 */
function formatDateWithWeekday(date) {
  return `${date.getMonth() + 1}/${date.getDate()}(${WEEKDAY_LABELS[date.getDay()]})`;
}

/**
 * 時刻を「14:00」の形式にする
 */
function formatTime(date) {
  return Utilities.formatDate(date, "Asia/Tokyo", "HH:mm");
}

/**
 * 「YYYY/MM/DD」と「HH:mm」の文字列から日時を組み立てる。
 * 時刻が空なら0時0分とする。日付として成立しない場合は null を返す。
 */
function parseDateTime(dateStr, timeStr) {
  const dateMatch = String(dateStr || "").trim().match(/^(\d{4})(?:\/|-)(\d{1,2})(?:\/|-)(\d{1,2})$/);
  if (!dateMatch) return null;

  const year = Number(dateMatch[1]);
  const month = Number(dateMatch[2]);
  const day = Number(dateMatch[3]);

  let hours = 0;
  let minutes = 0;
  const timeMatch = String(timeStr || "").trim().match(/^(\d{1,2}):(\d{2})$/);
  if (timeMatch) {
    hours = Number(timeMatch[1]);
    minutes = Number(timeMatch[2]);
    if (hours > 23 || minutes > 59) return null;
  }

  const result = new Date(year, month - 1, day, hours, minutes);
  // 「2/30」のような存在しない日付は自動で繰り上がってしまうため、ここで弾く
  if (result.getMonth() !== month - 1 || result.getDate() !== day) return null;
  return result;
}

/**
 * 「HH:mm」形式の時刻文字列かどうかを判定する
 */
function isTimeString(timeStr) {
  return /^\d{1,2}:\d{2}$/.test(String(timeStr || "").trim());
}

/**
 * 「明日」から「来週の日曜日」までの期間を返す（今週・来週の予定の見通し用）。
 * end は期間に含まれない終端（来週日曜の翌日0時）。
 */
function getUpcomingRange(now) {
  const today = startOfDay(now);
  const daysUntilSunday = (7 - today.getDay()) % 7; // 今日が日曜なら0
  return {
    start: addDays(today, 1),
    end: addDays(today, daysUntilSunday + 8)
  };
}

/**
 * Googleカレンダー（実行ユーザーのメインカレンダー）から、指定期間の予定を取得する。
 * 「欠席」と回答済みの予定は除外し、開始日時の早い順に並べて返す。
 */
function getCalendarEvents(rangeStart, rangeEnd) {
  const events = CalendarApp.getDefaultCalendar().getEvents(rangeStart, rangeEnd);
  return events
    .filter(event => event.getMyStatus() !== CalendarApp.GuestStatus.NO)
    .map(event => {
      const isAllDay = event.isAllDayEvent();
      return {
        title: event.getTitle() || "（件名なし）",
        start: isAllDay ? event.getAllDayStartDate() : event.getStartTime(),
        end: isAllDay ? event.getAllDayEndDate() : event.getEndTime(),
        isAllDay: isAllDay,
        location: event.getLocation() || ""
      };
    })
    .sort((a, b) => a.start.getTime() - b.start.getTime());
}

/**
 * 予定1件を1行のテキストにする。withDate=true の場合は先頭に日付を付ける
 */
function formatEventLine(calendarEvent, withDate) {
  const datePart = withDate ? formatDateWithWeekday(calendarEvent.start) + " " : "";
  const timePart = calendarEvent.isAllDay
    ? "終日"
    : `${formatTime(calendarEvent.start)}〜${formatTime(calendarEvent.end)}`;
  const locationPart = calendarEvent.location ? `（${calendarEvent.location}）` : "";
  return `・${datePart}${timePart}　${calendarEvent.title}${locationPart}`;
}

/**
 * 予定の一覧を、最大件数で区切ってテキスト行の配列にする（超過分は「ほか◯件」と表示）
 */
function formatEventLines(calendarEvents, withDate, maxLines) {
  const lines = calendarEvents.slice(0, maxLines).map(calendarEvent => formatEventLine(calendarEvent, withDate));
  if (calendarEvents.length > maxLines) {
    lines.push(`　ほか${calendarEvents.length - maxLines}件`);
  }
  return lines;
}

/**
 * 「出勤情報」として共有されている終日予定かどうかを判定する。
 * 件名に「10:30-15:00」のような時間帯が入っている終日予定（各メンバーの勤務時間）と、
 * 件名が「オフィス」の終日予定（勤務場所）を出勤情報とみなす。
 */
function isAttendanceEvent(calendarEvent) {
  if (!calendarEvent.isAllDay) return false;
  const title = String(calendarEvent.title).trim();
  if (title === "オフィス") return true;
  // 時刻の区切り（: ：）と、範囲を表すさまざまなハイフン・波線に対応する
  return /\d{1,2}[:：]\d{2}\s*[-‐‑–—―−ーｰ~〜～]\s*\d{1,2}[:：]\d{2}/.test(title);
}

/**
 * 今週・来週の予定と未完了タスクをAIに渡し、「何があるから、何の用意が必要か」の助言文を作る。
 * 予定が無い場合やAIの呼び出しに失敗した場合は空文字を返す（朝のご報告自体は止めない）。
 */
function buildPreparationAdvice(upcomingEvents, openTaskLines, now) {
  if (upcomingEvents.length === 0) return "";

  const nowStr = `${formatDateWithWeekday(now)} ${formatTime(now)}`;
  const eventText = formatEventLines(upcomingEvents, true, 40).join("\n");
  const taskText = openTaskLines.length > 0 ? openTaskLines.slice(0, 20).join("\n") : "（なし）";

  const prompt = `
あなたは瀬戸口さんの優秀な専属秘書です。現在の日時は【${nowStr}】です。
以下は瀬戸口さんの今週・来週の予定と、未完了のタスクです。
予定を見渡し、事前の用意が必要そうなものだけを選んで、「いつ何があるので、何をいつまでに用意しておくとよいか」を助言してください。

【ルール】
- 最大5件。日付が近く、重要そうなものを優先する。
- 1件につき1〜2行で、「・11/20(金) 予定名 → 用意しておくこと」の形式にする。
- 予定の件名・場所・タスクから読み取れる範囲で助言し、書かれていない事実を作らない。
- 関連しそうな未完了タスクがあれば、その予定と結び付けて触れる。
- 用意が不要そうな予定（定例の集まり、休み、移動など）は書かない。
- 該当が1件も無ければ「特に事前の用意が必要な予定は見当たりません。」とだけ書く。
- 見出し・前置き・締めの挨拶は書かず、記号の「*」や「#」も使わない。本文のみを出力する。

【今週・来週の予定】
${eventText}

【未完了のタスク】
${taskText}
`;

  try {
    const advice = callGemini(prompt, false);
    return advice ? advice.trim() : "";
  } catch (e) {
    console.error("準備の助言の生成に失敗しました: " + e.toString());
    return "";
  }
}

/**
 * ====================================================================
 * 管理部のスプレッドシート連携
 * 朝のご報告に「管理部シートの追いかけ案件」と「保証金返金の状況」を載せる
 * （どちらも読むだけで、シートには何も書き込まない）
 * ====================================================================
 */

// 【クリメン】管理部シート（スクリプトプロパティ ADMIN_SHEET_ID で差し替え可能）
const DEFAULT_ADMIN_SHEET_ID = "1PtG_fx5TpxyYh6kt_Gi6FxSnt0W3rGkzeY1tu08WTPM";
// 保証金返金対応シート（2026.9~）（スクリプトプロパティ DEPOSIT_SHEET_ID で差し替え可能）
const DEFAULT_DEPOSIT_SHEET_ID = "187NIXrBvRpLXm7a1mKCbfGFveJrZeyoh_keUSPZo77k";

// 管理部シートのうち、追いかけ表として読むタブ
const ADMIN_FOLLOWUP_TABS = ["保証金返金", "入金名義相違", "粗大ゴミ追いかけ", "管理部業務(その他)"];
// 「対応状況」「対応済み」列がこの言葉なら、終わった案件として扱う
const ADMIN_DONE_WORDS = ["対応済", "対応済み", "完了"];

/**
 * 文字列の空白・改行をまとめ、指定の長さで切り詰める（チャットで1行に収めるため）
 */
function compactText(value, maxLength) {
  const text = String(value || "").replace(/[\s　]+/g, " ").trim();
  return text.length > maxLength ? text.substring(0, maxLength) + "…" : text;
}

/**
 * 複数行のセルから、空でない最初の数行を取り出して1行にまとめる
 */
function firstLines(value, lineCount, maxLength) {
  const lines = String(value || "").split(/\r?\n/).map(line => line.trim()).filter(line => line);
  return compactText(lines.slice(0, lineCount).join(" "), maxLength);
}

/**
 * 年・月・日から日付を作る。存在しない日付（2/30など）は null を返す
 */
function makeValidDate(year, month, day) {
  const result = new Date(year, month - 1, day);
  if (result.getFullYear() !== year || result.getMonth() !== month - 1 || result.getDate() !== day) return null;
  return result;
}

/**
 * シートのセルの値（日付型、「2026/10/21」「10/21」「10月21日」などの文字列）を日付に変換する。
 * 年が書かれていない場合は今年とみなし、それが90日以上先になる場合は去年とみなす。
 * 読み取れない場合は null を返す。
 */
function parseSheetDate(value, now) {
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return isNaN(value.getTime()) ? null : startOfDay(value);
  }
  const text = String(value || "").trim();
  if (!text) return null;

  let match = text.match(/^(\d{4})\s*(?:\/|-|年)\s*(\d{1,2})\s*(?:\/|-|月)\s*(\d{1,2})/);
  if (match) return makeValidDate(Number(match[1]), Number(match[2]), Number(match[3]));

  match = text.match(/^(\d{1,2})\s*(?:\/|月)\s*(\d{1,2})/);
  if (!match) return null;

  const month = Number(match[1]);
  const day = Number(match[2]);
  let result = makeValidDate(now.getFullYear(), month, day);
  if (result && result.getTime() > addDays(startOfDay(now), 90).getTime()) {
    result = makeValidDate(now.getFullYear() - 1, month, day);
  }
  return result;
}

/**
 * 日付を「10/21」の形式にする
 */
function formatMonthDay(date) {
  return `${date.getMonth() + 1}/${date.getDate()}`;
}

/**
 * 見出し行の中から、候補の言葉で始まる列の位置を探す（見つからなければ -1）
 */
function findColumnIndex(headers, candidates) {
  const normalized = headers.map(header => String(header || "").replace(/[\s　]+/g, ""));
  for (const candidate of candidates) {
    const index = normalized.findIndex(header => header.indexOf(candidate) === 0);
    if (index !== -1) return index;
  }
  return -1;
}

/**
 * 管理部シートの追いかけ表（1タブ分）から、まだ終わっていない案件を取り出す。
 * 見出し行は「担当」という列がある行とみなし、列の位置は見出しの言葉で探す（列の並び替えに強くするため）。
 */
function readAdminFollowupTab(sheet, tabName, now) {
  const data = sheet.getDataRange().getValues();
  const headerIndex = data.slice(0, 8).findIndex(row => findColumnIndex(row, ["担当"]) !== -1);
  if (headerIndex === -1) return [];

  const headers = data[headerIndex];
  const col = {
    caseName: findColumnIndex(headers, ["案件", "店舗名"]),
    detail: findColumnIndex(headers, ["現状", "現在ステータス", "品目"]),
    owner: findColumnIndex(headers, ["担当"]),
    nextAction: findColumnIndex(headers, ["次にやること"]),
    nextCheck: findColumnIndex(headers, ["次回確認日"]),
    status: findColumnIndex(headers, ["対応状況", "対応済"]),
    note: findColumnIndex(headers, ["備考"])
  };
  const cell = (row, index) => (index === -1 ? "" : row[index]);

  const items = [];
  for (let i = headerIndex + 1; i < data.length; i++) {
    const row = data[i];
    const status = String(cell(row, col.status) || "").trim();
    if (ADMIN_DONE_WORDS.indexOf(status) !== -1) continue;
    if (String(cell(row, col.note) || "").trim().indexOf("例") === 0) continue; // 記入例の行

    // 案件欄が空、または「ー」だけの場合は、現状欄の書き出しを案件名の代わりに使う
    const caseName = firstLines(cell(row, col.caseName), 2, 40);
    const isPlaceholder = /^[ー\-－―─]*$/.test(caseName);
    const label = (isPlaceholder ? "" : caseName) || firstLines(cell(row, col.detail), 2, 40);
    const owner = compactText(cell(row, col.owner), 10);
    if (!label && !owner) continue; // 空行

    items.push({
      tabName: tabName,
      label: label || "（案件名なし）",
      owner: owner,
      nextAction: firstLines(cell(row, col.nextAction), 1, 30),
      nextCheck: parseSheetDate(cell(row, col.nextCheck), now),
      hasNextCheckColumn: col.nextCheck !== -1
    });
  }
  return items;
}

/**
 * 朝のご報告の「管理部シートの追いかけ」欄を組み立てる。
 * 次回確認日が今日以前の案件を一覧にし、確認日が未設定の案件はタブごとの件数だけ示す。
 */
function buildAdminFollowupSection(now) {
  const sheetId = PropertiesService.getScriptProperties().getProperty("ADMIN_SHEET_ID") || DEFAULT_ADMIN_SHEET_ID;
  const spreadsheet = SpreadsheetApp.openById(sheetId);
  const today = startOfDay(now);

  let openItems = [];
  ADMIN_FOLLOWUP_TABS.forEach(tabName => {
    const sheet = spreadsheet.getSheetByName(tabName);
    if (!sheet) {
      console.log(`管理部シートにタブ「${tabName}」が見つからないため、読み飛ばしました。`);
      return;
    }
    openItems = openItems.concat(readAdminFollowupTab(sheet, tabName, now));
  });

  if (openItems.length === 0) {
    return "■ 管理部シートの追いかけ\n・対応中の案件はありません";
  }

  const dueItems = openItems
    .filter(item => item.nextCheck && item.nextCheck.getTime() <= today.getTime())
    .sort((a, b) => a.nextCheck.getTime() - b.nextCheck.getTime());

  const lines = [`■ 管理部シートの追いかけ（対応中 ${openItems.length}件）`];

  if (dueItems.length > 0) {
    lines.push("確認日が今日・過ぎているもの：");
    dueItems.slice(0, 10).forEach(item => {
      const isToday = item.nextCheck.getTime() === today.getTime();
      const datePart = isToday ? "今日" : `${formatMonthDay(item.nextCheck)}予定`;
      const ownerPart = item.owner ? `担当 ${item.owner}／` : "";
      const actionPart = item.nextAction ? ` → ${item.nextAction}` : "";
      lines.push(`・［${item.tabName}］${item.label}（${ownerPart}${datePart}）${actionPart}`);
    });
    if (dueItems.length > 10) lines.push(`　ほか${dueItems.length - 10}件`);
  } else {
    lines.push("・確認日が今日・過ぎている案件はありません");
  }

  // 確認日が入っていない案件は、タブごとの件数だけ知らせる（日付を入れてもらうきっかけにする）
  const undatedCounts = {};
  openItems
    .filter(item => !item.nextCheck)
    .forEach(item => { undatedCounts[item.tabName] = (undatedCounts[item.tabName] || 0) + 1; });
  const undatedParts = Object.keys(undatedCounts).map(tabName => `${tabName} ${undatedCounts[tabName]}件`);
  if (undatedParts.length > 0) {
    lines.push(`確認日が未記入の対応中案件：${undatedParts.join("、")}`);
  }

  return lines.join("\n");
}

/**
 * 解約日から「最終支払日」（解約日の翌月末日）を求める
 */
function getFinalPaymentDate(cancelDate) {
  return new Date(cancelDate.getFullYear(), cancelDate.getMonth() + 2, 0);
}

/**
 * 朝のご報告の「保証金返金」欄を組み立てる。
 * 保証金返金対応シートのシート1から、非表示にしていない行のうち、
 * 分類が①②③⑥の契約を「最終支払日を過ぎたもの」「今月末が最終支払日のもの」に分けて数える。
 */
function buildDepositSection(now) {
  const sheetId = PropertiesService.getScriptProperties().getProperty("DEPOSIT_SHEET_ID") || DEFAULT_DEPOSIT_SHEET_ID;
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("シート1");
  if (!sheet) throw new Error("保証金返金対応シートに「シート1」が見つかりません。");

  const data = sheet.getDataRange().getValues();
  const headerIndex = data.slice(0, 10).findIndex(row => findColumnIndex(row, ["契約番号"]) !== -1);
  if (headerIndex === -1) throw new Error("保証金返金対応シートの見出し行が見つかりません。");

  const headers = data[headerIndex];
  const col = {
    contractNo: findColumnIndex(headers, ["契約番号"]),
    storeName: findColumnIndex(headers, ["契約店舗名称", "店舗"]),
    category: findColumnIndex(headers, ["分類"]),
    cancelDate: findColumnIndex(headers, ["解約日"]),
    handler: findColumnIndex(headers, ["対応者"])
  };
  if (col.contractNo === -1 || col.category === -1 || col.cancelDate === -1) {
    throw new Error("保証金返金対応シートの列（契約番号・分類・解約日）が見つかりません。");
  }

  const today = startOfDay(now);
  const endOfThisMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  const overdueByCategory = {};
  const unhandled = [];
  let dueThisMonthCount = 0;

  for (let i = headerIndex + 1; i < data.length; i++) {
    const row = data[i];
    const category = String(row[col.category] || "").trim();
    const categoryMark = category.charAt(0);
    if (["①", "②", "③", "⑥"].indexOf(categoryMark) === -1) continue;

    // 非表示にした行は「対応の目処が立った」ものとして扱い、数えない（行番号は1始まり）
    if (sheet.isRowHiddenByUser(i + 1)) continue;

    const cancelDate = parseSheetDate(row[col.cancelDate], now);
    const finalPaymentDate = cancelDate ? getFinalPaymentDate(cancelDate) : null;

    // ⑥（契約管理が未解約）は解約日がないため、常に要対応とする
    const isOverdue = categoryMark === "⑥" || (finalPaymentDate && finalPaymentDate.getTime() < today.getTime());
    if (!isOverdue) {
      if (finalPaymentDate && finalPaymentDate.getTime() <= endOfThisMonth.getTime()) dueThisMonthCount++;
      continue;
    }

    overdueByCategory[category] = (overdueByCategory[category] || 0) + 1;

    const handler = col.handler === -1 ? "" : String(row[col.handler] || "").trim();
    if (!handler) {
      const storeName = col.storeName === -1 ? "" : compactText(row[col.storeName], 20);
      const cancelPart = cancelDate ? `／解約 ${formatMonthDay(cancelDate)}` : "";
      unhandled.push(`・${row[col.contractNo]} ${storeName}（${categoryMark}${cancelPart}）`);
    }
  }

  const overdueTotal = Object.keys(overdueByCategory).reduce((sum, key) => sum + overdueByCategory[key], 0);
  const lines = ["■ 保証金返金（非表示の行を除く）"];

  if (overdueTotal === 0) {
    lines.push("・最終支払日を過ぎて対応中の契約はありません");
  } else {
    lines.push(`最終支払日を過ぎて対応中：${overdueTotal}件`);
    Object.keys(overdueByCategory).sort().forEach(category => {
      lines.push(`・${category}：${overdueByCategory[category]}件`);
    });
  }
  if (dueThisMonthCount > 0) {
    lines.push(`今月末に最終支払日を迎える契約：${dueThisMonthCount}件`);
  }
  if (unhandled.length > 0) {
    lines.push("対応者が未記入のもの：");
    unhandled.slice(0, 5).forEach(line => lines.push(line));
    if (unhandled.length > 5) lines.push(`　ほか${unhandled.length - 5}件`);
  }

  return lines.join("\n");
}

/**
 * 朝のご報告のうち、管理部のシートに関する欄をまとめて組み立てる。
 * 片方の読み取りに失敗しても、もう片方とご報告全体は送れるようにする。
 */
function buildAdminBriefingSections(now) {
  const sections = [];
  try {
    sections.push(buildAdminFollowupSection(now));
  } catch (e) {
    console.error("管理部シートの読み取りに失敗しました: " + e.toString());
    sections.push("■ 管理部シートの追いかけ\n⚠️ シートを読み取れませんでした。");
  }
  try {
    sections.push(buildDepositSection(now));
  } catch (e) {
    console.error("保証金返金対応シートの読み取りに失敗しました: " + e.toString());
    sections.push("■ 保証金返金\n⚠️ シートを読み取れませんでした。");
  }
  return sections;
}

/**
 * 朝のご報告のうち、カレンダーに関する部分（今日の予定／今週・来週の予定／準備の助言）を組み立てる
 */
function buildCalendarBriefing(now, openTaskLines) {
  const today = startOfDay(now);
  const allTodayEvents = getCalendarEvents(today, addDays(today, 1));
  // 今日の分だけは出勤情報も知らせる（予定とは別枠で表示する）
  const todayAttendance = allTodayEvents.filter(calendarEvent => isAttendanceEvent(calendarEvent));
  const todayEvents = allTodayEvents.filter(calendarEvent => !isAttendanceEvent(calendarEvent));

  // 今週・来週の見通しからは、出勤情報を除く
  const upcomingRange = getUpcomingRange(now);
  const upcomingEvents = getCalendarEvents(upcomingRange.start, upcomingRange.end)
    .filter(calendarEvent => !isAttendanceEvent(calendarEvent));
  const advice = buildPreparationAdvice(upcomingEvents, openTaskLines, now);

  const attendanceSection = todayAttendance.length > 0
    ? "■ 今日の出勤\n" + todayAttendance.slice(0, 15).map(calendarEvent => `・${calendarEvent.title}`).join("\n")
    : "";

  const todaySection = [
    attendanceSection,
    "■ 今日の予定\n" +
      (todayEvents.length > 0 ? formatEventLines(todayEvents, false, 15).join("\n") : "・予定は入っておりません")
  ].filter(section => section).join("\n\n");

  const upcomingSection = "■ 今週・来週の予定\n" +
    (upcomingEvents.length > 0 ? formatEventLines(upcomingEvents, true, 25).join("\n") : "・予定は入っておりません");

  const adviceSection = advice ? "■ 用意しておくとよいこと\n" + advice : "";

  return { todaySection, upcomingSection, adviceSection };
}

/**
 * ⏰ 朝のご報告：今日の予定＋期限が近いタスク＋今週・来週の予定と、そのための用意をまとめてChatへ送信する関数
 * 時間主導型トリガー（1日1回・朝）での実行を想定
 *
 * - カレンダーの内容は、直近でやり取りのあったスペース（瀬戸口さんとの会話）にのみ送る
 * - タスクのリマインドは、これまで通りタスクを登録したスペースごとに送る
 */
function sendMorningBriefing() {
  const sheetUrl = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_URL");
  const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  const taskSheet = getRequiredSheet(spreadsheet, "Task_List");
  const data = taskSheet.getDataRange().getValues();

  const now = new Date();
  const todayStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd");
  const tomorrowStr = Utilities.formatDate(addDays(now, 1), "Asia/Tokyo", "yyyy/MM/dd");

  // スペース（送信先）ごとにリマインド内容をまとめるためのバケツ
  const remindersBySpace = {};
  // AIに「用意しておくとよいこと」を考えてもらう材料として、未完了タスクの一覧も集める
  const openTaskLines = [];

  for (let i = 1; i < data.length; i++) {
    const status = data[i][5];         // F列: ステータス
    const taskContent = data[i][2];    // C列: タスク内容
    const assignedPerson = data[i][7]; // H列: 担当者
    const lastRemindedAt = data[i][6]; // G列: 最終リマインド日時
    const spaceName = data[i][9];      // J列: スペース名

    // D列: 期限日付（セルが日付型になっていても文字列でも比較できるよう、文字列に揃える）
    const rawDeadline = data[i][3];
    const deadlineDate = Object.prototype.toString.call(rawDeadline) === "[object Date]"
      ? Utilities.formatDate(rawDeadline, "Asia/Tokyo", "yyyy/MM/dd")
      : String(rawDeadline || "");

    if (isClosedTaskStatus(status) || !taskContent) continue;

    openTaskLines.push(`・${taskContent}（期限 ${deadlineDate || "未定"}／担当：${assignedPerson || "未定"}／状況：${status}）`);

    if (!deadlineDate || !spaceName) continue;

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

    if (!remindersBySpace[spaceName]) remindersBySpace[spaceName] = [];
    remindersBySpace[spaceName].push({
      rowNum: i + 1,
      line: `${label}：「${taskContent}」（期限 ${deadlineDate}／担当：${assignedPerson}）`
    });
  }

  // カレンダーの内容を送るスペースを決め、カレンダー部分を組み立てる
  const calendarSpaceName = getDefaultSpaceName(spreadsheet);
  let calendarBriefing = null;
  let adminSections = [];
  if (calendarSpaceName) {
    // 管理部シートの追いかけ案件と、保証金返金の状況（瀬戸口さんとの会話にのみ送る）
    adminSections = buildAdminBriefingSections(now);

    try {
      calendarBriefing = buildCalendarBriefing(now, openTaskLines);
    } catch (e) {
      console.error("カレンダーの読み取りに失敗しました: " + e.toString());
      calendarBriefing = {
        todaySection: "■ 今日の予定\n⚠️ カレンダーを読み取れませんでした。",
        upcomingSection: "",
        adviceSection: ""
      };
    }
    // タスクのリマインドが無い日でも、カレンダーの報告は送る
    if (!remindersBySpace[calendarSpaceName]) remindersBySpace[calendarSpaceName] = [];
  } else {
    console.log("送信先スペースが特定できなかったため、カレンダーの報告は省略しました。");
  }

  for (const spaceName in remindersBySpace) {
    const items = remindersBySpace[spaceName];
    const taskLines = items.map(item => item.line).join("\n");
    let message;

    if (spaceName === calendarSpaceName && calendarBriefing) {
      const sections = [
        calendarBriefing.todaySection,
        "■ 期限が近いタスク\n" + (items.length > 0 ? taskLines : "・期限が近いタスクはありません"),
        ...adminSections,
        calendarBriefing.upcomingSection,
        calendarBriefing.adviceSection
      ].filter(section => section);
      message = `🤖 瀬戸口さん、おはようございます。${formatDateWithWeekday(now)} の朝のご報告です。\n\n` + sections.join("\n\n");
    } else {
      message = "🤖 秘書より、本日のタスクリマインドです。\n\n" + taskLines;
    }

    try {
      // Chatの1メッセージあたりの文字数上限に収まるよう、長い場合は欄の区切りで複数通に分けて送る
      splitMessageForChat(message, 3800).forEach(part => sendPushReply(spaceName, part));
      // 送信に成功したタスクのみ、最終リマインド日時を更新する
      const timestamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
      items.forEach(item => taskSheet.getRange(item.rowNum, 7).setValue(timestamp));
    } catch (e) {
      console.error("朝のご報告の送信に失敗しました（" + spaceName + "）: " + e.message);
    }
  }
}

/**
 * 長いメッセージを、Chatの文字数上限に収まるよう複数に分ける。
 * なるべく空行（欄の区切り）で分け、1つの欄が長すぎる場合は行の区切り、それでも長ければ文字数で分ける。
 */
function splitMessageForChat(message, maxLength) {
  if (message.length <= maxLength) return [message];

  const parts = [];
  let current = "";
  const pushCurrent = () => {
    if (current) parts.push(current);
    current = "";
  };
  const appendPiece = (piece, separator) => {
    if (!current) {
      current = piece;
    } else if (current.length + separator.length + piece.length <= maxLength) {
      current += separator + piece;
    } else {
      pushCurrent();
      current = piece;
    }
  };

  message.split("\n\n").forEach(block => {
    if (block.length <= maxLength) {
      appendPiece(block, "\n\n");
      return;
    }
    pushCurrent(); // 長い欄は、新しい1通の先頭から始める
    block.split("\n").forEach(line => {
      for (let start = 0; start < line.length; start += maxLength) {
        appendPiece(line.substring(start, start + maxLength), "\n");
      }
    });
    pushCurrent();
  });
  pushCurrent();
  return parts;
}

/**
 * 以前の関数名（タスクのリマインドのみだった頃の名前）で登録済みのトリガーがあっても動くよう、
 * 朝のご報告へ処理を引き継ぐ
 */
function sendTaskReminders() {
  sendMorningBriefing();
}

/**
 * sendMorningBriefingを1日1回（午前9時）自動実行するトリガーを設定する関数
 * GASエディタでこの関数を1回だけ手動実行してください
 */
function setupReminderTrigger() {
  // 重複登録を防ぐため、既存の同じ役割のトリガーを削除してから作り直す
  ScriptApp.getProjectTriggers().forEach(trigger => {
    const handler = trigger.getHandlerFunction();
    if (handler === "sendMorningBriefing" || handler === "sendTaskReminders") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("sendMorningBriefing")
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .create();
}

/**
 * チャットで依頼された予定をGoogleカレンダーに登録し、登録結果の報告文を返す関数
 * - 日付や件名が読み取れない場合、過去の日付の場合は、登録せずに聞き返す
 * - 開始時刻が無い場合は終日の予定として登録する
 * - 終了時刻が無い場合は開始から1時間の予定として登録する
 * - 同じ件名・同じ開始日時の予定が既にある場合は二重登録しない
 */
function addCalendarEventFromIntent(eventData, now) {
  const title = String(eventData.eventTitle || "").trim();
  const startDay = parseDateTime(eventData.startDate, "");

  if (!title || !startDay) {
    return "⚠️ 秘書より確認です。カレンダーに登録する予定の「日付」または「件名」が読み取れませんでした。お手数ですが、「11/20の14時からA社打ち合わせをカレンダーに登録して」のようにご指示いただけますか。";
  }

  if (startDay.getTime() < startOfDay(now).getTime()) {
    return `⚠️ 秘書より確認です。「${title}」の日付が ${formatDateWithWeekday(startDay)}（${startDay.getFullYear()}年）と過去の日付になっております。誤って登録しないよう保留にしました。お手数ですが、日付を改めてご指示いただけますか。`;
  }

  const calendar = CalendarApp.getDefaultCalendar();
  const options = {};
  if (eventData.location) options.location = String(eventData.location);
  if (eventData.description) options.description = String(eventData.description);
  const locationNote = eventData.location ? `\n・場所：${eventData.location}` : "";

  // 開始時刻が無い場合：終日の予定として登録する
  if (!isTimeString(eventData.startTime)) {
    const endDay = parseDateTime(eventData.endDate, "");
    const isMultiDay = endDay && endDay.getTime() > startDay.getTime();
    const lastDay = isMultiDay ? endDay : startDay;

    const isDuplicate = calendar.getEvents(startDay, addDays(lastDay, 1)).some(existing =>
      existing.isAllDayEvent() &&
      existing.getTitle() === title &&
      existing.getAllDayStartDate().getTime() === startDay.getTime()
    );
    const whenText = isMultiDay
      ? `${formatDateWithWeekday(startDay)}〜${formatDateWithWeekday(endDay)} 終日`
      : `${formatDateWithWeekday(startDay)} 終日`;

    if (isDuplicate) {
      return `📅 秘書より報告です。${whenText}「${title}」は、すでにカレンダーに登録されておりましたので、二重登録は控えました。`;
    }

    if (isMultiDay) {
      calendar.createAllDayEvent(title, startDay, addDays(endDay, 1), options); // 終了日は「含まれない終端」で指定する
    } else {
      calendar.createAllDayEvent(title, startDay, options);
    }

    return `📅 承知いたしました。カレンダーに登録いたしました。\n・${whenText}　「${title}」${locationNote}\n\n時刻のご指定がなかったため、終日の予定として登録しております。時刻が決まりましたらお申し付けください。`;
  }

  // 開始時刻がある場合：時間帯つきの予定として登録する
  const start = parseDateTime(eventData.startDate, eventData.startTime);
  if (!start) {
    return `⚠️ 秘書より確認です。「${title}」の開始時刻がうまく読み取れませんでした。お手数ですが、「14時から」のように時刻を改めてご指示いただけますか。`;
  }

  let end = isTimeString(eventData.endTime)
    ? parseDateTime(eventData.endDate || eventData.startDate, eventData.endTime)
    : null;
  let endNote = "";
  if (!end || end.getTime() <= start.getTime()) {
    end = new Date(start.getTime() + 60 * 60 * 1000);
    endNote = "\n\n終了時刻のご指定がなかったため、1時間の予定として登録しております。";
  }

  const isSameDay = startOfDay(start).getTime() === startOfDay(end).getTime();
  const whenText = isSameDay
    ? `${formatDateWithWeekday(start)} ${formatTime(start)}〜${formatTime(end)}`
    : `${formatDateWithWeekday(start)} ${formatTime(start)}〜${formatDateWithWeekday(end)} ${formatTime(end)}`;

  const isDuplicate = calendar.getEvents(start, end).some(existing =>
    !existing.isAllDayEvent() &&
    existing.getTitle() === title &&
    existing.getStartTime().getTime() === start.getTime()
  );
  if (isDuplicate) {
    return `📅 秘書より報告です。${whenText}「${title}」は、すでにカレンダーに登録されておりましたので、二重登録は控えました。`;
  }

  calendar.createEvent(title, start, end, options);

  return `📅 承知いたしました。カレンダーに登録いたしました。\n・${whenText}　「${title}」${locationNote}${endNote}`;
}

/**
 * チャットで尋ねられた期間の予定をGoogleカレンダーから読み取り、一覧の報告文を返す関数
 */
function answerCalendarQuery(queryData, now) {
  const MAX_RANGE_DAYS = 31;

  const rangeStart = parseDateTime(queryData.rangeStartDate, "") || startOfDay(now);
  let rangeEnd = parseDateTime(queryData.rangeEndDate, "") || rangeStart;
  if (rangeEnd.getTime() < rangeStart.getTime()) rangeEnd = rangeStart;
  if (rangeEnd.getTime() > addDays(rangeStart, MAX_RANGE_DAYS - 1).getTime()) {
    rangeEnd = addDays(rangeStart, MAX_RANGE_DAYS - 1);
  }

  const isSingleDay = rangeStart.getTime() === rangeEnd.getTime();
  const rangeLabel = isSingleDay
    ? formatDateWithWeekday(rangeStart)
    : `${formatDateWithWeekday(rangeStart)}〜${formatDateWithWeekday(rangeEnd)}`;

  // 1日だけの確認では出勤情報も含め、複数日の確認では出勤情報を除いて見やすくする
  const calendarEvents = getCalendarEvents(rangeStart, addDays(rangeEnd, 1))
    .filter(calendarEvent => isSingleDay || !isAttendanceEvent(calendarEvent));
  if (calendarEvents.length === 0) {
    return `📅 ${rangeLabel} のご予定は、カレンダーには入っておりません。`;
  }

  return `📅 ${rangeLabel} のご予定は以下の${calendarEvents.length}件です。\n\n` +
    formatEventLines(calendarEvents, !isSingleDay, 30).join("\n");
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
 * Gemini APIを呼び出してユーザーの意図（新規タスク・進捗更新・要約・会話）を判断し、適切に処理する関数
 */
function handleUserIntent(userInput, userName, spaceName, recentConversation) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) return "⚠️ エラー: APIキーが設定されていません。";

  // 「明日」「来週の月曜」などの言葉を理解できるよう、現在の日時と曜日を教える
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy年MM月dd日") +
                   `（${WEEKDAY_LABELS[now.getDay()]}曜日）` +
                   Utilities.formatDate(now, "Asia/Tokyo", " HH:mm");

  // 「もう一回やって」などに対応できるよう、直前のやり取りも伝える
  const historyText = recentConversation && recentConversation.length > 0
    ? recentConversation.join("\n")
    : "（直前のやり取りはありません）";

  const systemPrompt = `
あなたは瀬戸口さんの優秀な専属秘書botです。
現在の日時は【${todayStr}】です。

【直前のやり取り（古い順。参考情報）】
${historyText}

瀬戸口さん（${userName}）からの今回のメッセージ:
「${userInput}」

【指示】
今回のメッセージに含まれる依頼を、以下の8種類の「アクション」に分けてください。
- 1通のメッセージに依頼が複数ある場合は、依頼ごとに1つずつアクションを作り、書かれている順に並べてください。
- 「もう一回やって」「さっきのお願い」のように直前のやり取りを指している場合は、直前のやり取りのうち「エラーで未完了」になっている依頼の内容から、改めてアクションを作ってください。「処理済み」の依頼は、明確に頼まれない限り繰り返さないでください。
- 何を指しているのか判断できない場合は、"chat" アクションで、何をすればよいかを丁寧に聞き返してください。
- replyMessage は必ず空でない文章にしてください。

出力は、次の形のJSONの文字列のみとしてください（余計な解説文などは一切含めないでください）。
{
  "actions": [ アクション, アクション, ... ]
}

アクションの種類:

1. 「新しいタスクの依頼や追加」（例：「明日までに資料作成」「タスクに追加して」など）
{
  "action": "task",
  "taskContent": "タスクの具体的な内容",
  "deadlineDate": "YYYY/MM/DD (期限が不明な場合は空文字)",
  "deadlineTime": "HH:mm (時間が不明な場合は空文字)",
  "assignedPerson": "このタスクの担当者（部下の名前や瀬戸口さんの名前など。明示がなければ「瀬戸口さん」）",
  "statusSummary": "新規タスク登録",
  "replyMessage": "タスクを新規登録した旨の、秘書らしい丁寧な返答"
}

2. 「既存タスクの進捗更新・経緯の追記・完了報告」（例：「資料作成の件、A社から来週まで待ってと連絡あった」「週報タスク終わったよ」「〇〇の件は今山田くんがボール持ってる」など）
{
  "action": "update_task",
  "targetTaskKeyword": "どの既存タスクに関する進捗か特定するためのキーワード（タスク内容に含まれていそうな象徴的な言葉。「」で括られた言葉があればそのまま使う）",
  "progressDetail": "今回新しく記録する対応内容や経緯の具体的なディテール",
  "assignedPerson": "現在のボール保持者・担当者（誰のところで止まっているか、誰が今動くべきか）",
  "statusSummary": "最新の状況を一言で表すサマリー（例：「先方連絡待ち」「完了」「山田さん対応中」など）",
  "taskStatus": "タスクの最新ステータス（「完了」または「未完了」または「先方返信待ち」など）",
  "replyMessage": "進捗履歴をしっかりとノートに記録した旨の、秘書らしい丁寧な報告返答"
}

3. 「既存タスクの削除・取り消し」（例：「〇〇のタスク消して」「〇〇はもうやらなくていいから外して」など。終わったという報告ではなく、一覧から外したい場合）
{
  "action": "delete_task",
  "targetTaskKeyword": "削除するタスクを特定するためのキーワード（「」で括られた言葉があればそのまま使う）",
  "replyMessage": "タスクを削除する旨の、秘書らしい丁寧な返答"
}

4. 「文章の要約や下書きの作成」（例：「この文章を要約して」「返信メールを作って」など）
{
  "action": "summary_or_draft",
  "replyMessage": "しっかり要約・下書きした内容を、秘書らしく丁寧に見やすく整えた文章"
}

5. それ以外の「ただの会話や質問」
{
  "action": "chat",
  "replyMessage": "秘書らしい丁寧で明るい返答"
}

6. 「Googleカレンダーへの新しい予定の登録」（例：「20日の14時からA社打ち合わせ、カレンダーに登録しといて」「来週月曜に歯医者の予定を入れて」など。「カレンダー」「予定を入れて」「スケジュール登録」といった言葉がある場合はタスクではなくこちら）
{
  "action": "calendar_add",
  "eventTitle": "予定の件名",
  "startDate": "YYYY/MM/DD (開始日。不明な場合は空文字)",
  "startTime": "HH:mm (開始時刻。終日の予定や時刻が不明な場合は空文字)",
  "endDate": "YYYY/MM/DD (複数日にまたがる場合のみ終了日。それ以外は空文字)",
  "endTime": "HH:mm (終了時刻。不明な場合は空文字)",
  "location": "場所（不明な場合は空文字）",
  "description": "予定に添えるメモ（なければ空文字）",
  "replyMessage": "予定を登録する旨の、秘書らしい丁寧な返答"
}

7. 「Googleカレンダーにすでにある予定の日時の変更」（例：「アンケート締切を29日に変更して」「A社打ち合わせを15時からにずらして」「明日の面談を来週月曜に移して」など）
{
  "action": "calendar_update",
  "targetEventKeyword": "変更する予定を特定するための、予定の件名に含まれていそうな言葉",
  "targetDate": "YYYY/MM/DD (変更前の予定の日付が分かる場合のみ。不明なら空文字)",
  "newStartDate": "YYYY/MM/DD (変更後の日付。日付は変えず時刻だけ変える場合は、変更前の日付)",
  "newStartTime": "HH:mm (変更後の開始時刻。時刻の指定がなければ空文字)",
  "newEndTime": "HH:mm (変更後の終了時刻。指定がなければ空文字)",
  "replyMessage": "予定を変更する旨の、秘書らしい丁寧な返答"
}

8. 「Googleカレンダーの予定の確認」（例：「今日の予定は？」「来週の予定を教えて」「明日なにがあったっけ」など）
{
  "action": "calendar_query",
  "rangeStartDate": "YYYY/MM/DD (確認したい期間の最初の日)",
  "rangeEndDate": "YYYY/MM/DD (確認したい期間の最後の日。1日だけなら最初の日と同じ)",
  "replyMessage": "予定を確認する旨の、秘書らしい丁寧な返答"
}
`;

  let aiResponse;
  try {
    aiResponse = callGemini(systemPrompt, true);
  } catch (e) {
    console.error("Gemini API呼び出しでエラーが発生しました: " + e.toString());
    return "⚠️ 申し訳ありません。ただいまシステムに問題が発生しております。少し時間をおいて再度お試しください。";
  }

  if (!aiResponse) {
    return "⚠️ 申し訳ありません。ただいまAIとの通信で問題が発生しております。少し時間をおいて再度お試しください。";
  }

  let resultData;
  try {
    // AIが返したJSON文字列をプログラムで使える形（オブジェクト）に変換
    const cleanJsonStr = aiResponse.replace(/```json/g, "").replace(/```/g, "").trim();
    resultData = JSON.parse(cleanJsonStr);
  } catch (parseError) {
    console.error("AI応答のJSON解析に失敗しました: " + parseError.message + " / 応答内容: " + aiResponse);
    return "⚠️ 申し訳ありません。AIからの応答をうまく読み取れませんでした。お手数ですが、もう一度お試しください。";
  }

  const actions = normalizeIntentActions(resultData);
  if (actions.length === 0) {
    console.error("AI応答にアクションが含まれていませんでした: " + aiResponse);
    return "⚠️ 申し訳ありません。ご指示の内容をうまく読み取れませんでした。お手数ですが、言い方を変えてもう一度お試しください。";
  }

  // 1通に依頼が複数ある場合も、1つずつ順に処理する（1つが失敗しても残りは続ける）
  const MAX_ACTIONS = 5;
  const replies = actions.slice(0, MAX_ACTIONS).map(action => {
    try {
      return executeIntentAction(action, userName, spaceName, now);
    } catch (actionError) {
      console.error(`「${action.action}」の処理でエラーが発生しました: ` + actionError.toString());
      return "⚠️ 申し訳ありません。ご指示の一部の処理に失敗しました。少し時間をおいて再度お試しください。";
    }
  }).filter(reply => reply && String(reply).trim());

  if (actions.length > MAX_ACTIONS) {
    replies.push(`⚠️ 一度に承れるご指示は${MAX_ACTIONS}件までのため、残りの${actions.length - MAX_ACTIONS}件は改めてお送りください。`);
  }

  return replies.length > 0
    ? replies.join("\n\n")
    : "⚠️ 申し訳ありません。ご指示の内容をうまく読み取れませんでした。お手数ですが、言い方を変えてもう一度お試しください。";
}

/**
 * AIの応答（{"actions":[...]} の形、アクション1つだけの形、配列の形のいずれか）を、アクションの配列にそろえる
 */
function normalizeIntentActions(resultData) {
  let actions;
  if (Array.isArray(resultData)) {
    actions = resultData;
  } else if (resultData && Array.isArray(resultData.actions)) {
    actions = resultData.actions;
  } else if (resultData && resultData.action) {
    actions = [resultData];
  } else {
    actions = [];
  }
  return actions.filter(action => action && typeof action === "object" && action.action);
}

/**
 * アクション1件を実行し、チャットに返す報告文を返す
 */
function executeIntentAction(action, userName, spaceName, now) {
  switch (action.action) {
    case "task":
      writeTaskToSheet(action, userName, spaceName);
      return action.replyMessage || `📝 承知いたしました。タスク「${action.taskContent}」を登録いたしました。`;

    case "update_task":
      return updateTaskInSheets(action, userName);

    case "delete_task":
      return deleteTaskFromIntent(action, userName);

    case "calendar_add":
      return addCalendarEventFromIntent(action, now);

    case "calendar_update":
      return updateCalendarEventFromIntent(action, now);

    case "calendar_query":
      return answerCalendarQuery(action, now);

    default:
      // 要約・下書き・会話などは、AIの返答をそのまま返す
      return action.replyMessage;
  }
}

/**
 * Gemini APIを呼び出し、応答テキストを返す共通関数。
 * useJsonMode=true の場合はJSONモードを強制する。
 * 応答に本文が含まれていない場合は null を返し、通信自体の失敗は例外を投げる。
 */
function callGemini(promptText, useJsonMode) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) throw new Error("スクリプトプロパティ 'GEMINI_API_KEY' が未設定です。");

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash-lite:generateContent?key=" + apiKey;

  const payload = {
    "contents": [{ "parts": [{ "text": promptText }] }]
  };
  if (useJsonMode) {
    payload.generationConfig = { "responseMimeType": "application/json" };
  }

  const response = UrlFetchApp.fetch(url, {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  });

  const json = JSON.parse(response.getContentText());
  const text = json.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) {
    console.error("Geminiの応答に想定した内容が含まれていませんでした: " + JSON.stringify(json));
    return null;
  }
  return text;
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

// タスクの「終わった」扱いのステータス（リマインドや検索の対象から外す）
const CLOSED_TASK_STATUSES = ["完了", "削除"];

/**
 * タスクのステータスが「完了」「削除」など、もう追いかけなくてよいものかを判定する
 */
function isClosedTaskStatus(status) {
  return CLOSED_TASK_STATUSES.indexOf(String(status || "").trim()) !== -1;
}

/**
 * タスク一覧のデータからキーワードに一致する行を全て探す（配列インデックスの配列を返す）。
 * 最新のタスクを優先するため下の行から逆順に検索する。
 * onlyIncomplete=true の場合は「完了」「削除」の行を除外して検索する（誤爆防止）。
 * 「削除」の行は、onlyIncomplete=false でも対象にしない。
 */
function findMatchingTaskRowIndexes(taskDataRange, keyword, onlyIncomplete) {
  const matches = [];
  for (let i = taskDataRange.length - 1; i >= 1; i--) {
    const status = String(taskDataRange[i][5]).trim(); // F列: ステータス
    if (status === "削除") continue;
    if (onlyIncomplete && isClosedTaskStatus(status)) continue;

    const taskContent = String(taskDataRange[i][2]).toLowerCase(); // C列: タスク内容
    if (taskContent.indexOf(keyword) !== -1) {
      matches.push(i);
    }
  }
  return matches;
}

/**
 * キーワードから対象のタスクを1件に絞り込む。
 * - まず未完了のタスクから探し、無ければ完了済みも含めて探す
 * - 候補が複数あっても、タスク内容がキーワードと完全に一致するものが1件だけなら、それを選ぶ
 * 1件に絞れた場合は { index }、絞れない場合は { message }（聞き返しの文）を返す。
 */
function findTargetTask(taskDataRange, keywordText, verb) {
  const keyword = String(keywordText || "").replace(/[「」『』]/g, "").trim().toLowerCase();
  if (!keyword) {
    return { message: `⚠️ 秘書より確認です。${verb}タスクがどれか読み取れませんでした。お手数ですが、タスク名を添えて改めてご指示いただけますか。` };
  }

  let matches = findMatchingTaskRowIndexes(taskDataRange, keyword, true);
  if (matches.length === 0) {
    matches = findMatchingTaskRowIndexes(taskDataRange, keyword, false);
  }

  if (matches.length > 1) {
    const exactMatches = matches.filter(i => String(taskDataRange[i][2]).trim().toLowerCase() === keyword);
    if (exactMatches.length === 1) matches = exactMatches;
  }

  if (matches.length === 0) {
    return { message: `⚠️ 秘書より報告です。「${keywordText}」に該当するタスクが一覧から見つけられませんでした。お手数ですが、正確な名前でご指示いただけますか。` };
  }

  if (matches.length > 1) {
    // 候補が複数ある場合は誤爆を避けるため、自動で決めつけずユーザーに確認を求める
    const candidateList = matches.slice(0, 5)
      .map(i => `・${taskDataRange[i][2]}（期限 ${formatTaskDeadline(taskDataRange[i][3]) || "未定"}）`)
      .join("\n");
    return { message: `⚠️ 秘書より確認です。「${keywordText}」に該当するタスクが複数見つかりました。誤って別のタスクを${verb.replace(/する$/, "")}しないよう、どのタスクか改めてご指示いただけますか。\n\n候補:\n${candidateList}` };
  }

  return { index: matches[0] };
}

/**
 * タスクの期限（D列）を「2026/08/15」の形の文字列にする（日付型のセルにも対応）
 */
function formatTaskDeadline(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return Utilities.formatDate(value, "Asia/Tokyo", "yyyy/MM/dd");
  }
  return String(value || "").trim();
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
  const target = findTargetTask(taskDataRange, updateData.targetTaskKeyword, "更新する");
  if (target.message) return target.message;

  const matchedIndex = target.index;
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

  return updateData.replyMessage ||
    `📝 承知いたしました。タスク「${taskDataRange[matchedIndex][2]}」の進捗を記録いたしました。`;
}

/**
 * チャットで頼まれたタスクを一覧から外す関数。
 * 行そのものは消さず、ステータスを「削除」にする（後から経緯を確認でき、元にも戻せるようにするため）。
 */
function deleteTaskFromIntent(deleteData, userName) {
  const sheetUrl = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_URL");
  const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  const taskSheet = getRequiredSheet(spreadsheet, "Task_List");
  const logSheet = getRequiredSheet(spreadsheet, "対応履歴ログ");

  const taskDataRange = taskSheet.getDataRange().getValues();
  const target = findTargetTask(taskDataRange, deleteData.targetTaskKeyword, "削除する");
  if (target.message) return target.message;

  const row = taskDataRange[target.index];
  const rowNum = target.index + 1;
  const now = new Date();
  const timestamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

  taskSheet.getRange(rowNum, 6).setValue("削除");                     // F列: ステータス
  taskSheet.getRange(rowNum, 9).setValue("チャットの指示で削除");     // I列: 最新の状況サマリー

  logSheet.appendRow([
    generateUniqueId("L", now), // A: 履歴ID
    row[0],                     // B: タスクID
    timestamp,                  // C: 記録日時
    "チャットの指示でタスクを削除しました", // D: 対応内容・経緯
    userName                    // E: 担当者
  ]);

  const deadline = formatTaskDeadline(row[3]);
  return `🗑️ 承知いたしました。タスク「${row[2]}」${deadline ? `（期限 ${deadline}）` : ""}を一覧から外しました。今後のリマインドにも出てまいりません。\n※表の行は残し、ステータスを「削除」にしております。`;
}

/**
 * チャットで頼まれた、カレンダーにすでにある予定の日時を変更する関数
 * - 件名に指定の言葉を含む予定を探し、1件に絞れた場合だけ変更する
 * - 繰り返しの予定や、他の方から招待された予定は変更せず、ご自身での変更をお願いする
 * - 時刻の指定がなければ、元の時刻・長さのまま日付だけを動かす
 */
function updateCalendarEventFromIntent(updateData, now) {
  const keyword = String(updateData.targetEventKeyword || "").replace(/[「」『』]/g, "").trim();
  if (!keyword) {
    return "⚠️ 秘書より確認です。変更する予定がどれか読み取れませんでした。お手数ですが、予定の件名を添えて改めてご指示いただけますか。";
  }

  const newDay = parseDateTime(updateData.newStartDate, "");
  if (!newDay) {
    return `⚠️ 秘書より確認です。「${keyword}」の変更後の日付が読み取れませんでした。お手数ですが、「29日に変更して」のように日付を添えてご指示いただけますか。`;
  }
  const today = startOfDay(now);
  if (newDay.getTime() < today.getTime()) {
    return `⚠️ 秘書より確認です。「${keyword}」の変更後の日付が ${formatDateWithWeekday(newDay)}（${newDay.getFullYear()}年）と過去の日付になっております。誤って変更しないよう保留にしました。日付を改めてご指示いただけますか。`;
  }

  // 探す期間：元の日付が分かればその日、分からなければ1週間前から4か月先まで
  const targetDay = parseDateTime(updateData.targetDate, "");
  const searchStart = targetDay || addDays(today, -7);
  const searchEnd = targetDay ? addDays(targetDay, 1) : addDays(today, 120);

  const calendar = CalendarApp.getDefaultCalendar();
  const lowerKeyword = keyword.toLowerCase();
  const candidates = calendar.getEvents(searchStart, searchEnd)
    .filter(event => String(event.getTitle() || "").toLowerCase().indexOf(lowerKeyword) !== -1);

  const describe = event => {
    if (event.isAllDayEvent()) return `${formatDateWithWeekday(event.getAllDayStartDate())} 終日`;
    return `${formatDateWithWeekday(event.getStartTime())} ${formatTime(event.getStartTime())}〜${formatTime(event.getEndTime())}`;
  };

  if (candidates.length === 0) {
    return `⚠️ 秘書より報告です。「${keyword}」を件名に含む予定が、カレンダーに見つかりませんでした。お手数ですが、予定の件名を確認のうえ改めてご指示いただけますか。`;
  }

  // 繰り返しの予定は、1つの言葉で全ての回が見つかるため、まとめて1件として扱う
  const isSeries = candidates.length > 1 && candidates.every(event => event.isRecurringEvent());
  if (candidates.length > 1 && !isSeries) {
    const list = candidates.slice(0, 5).map(event => `・${describe(event)}　${event.getTitle()}`).join("\n");
    return `⚠️ 秘書より確認です。「${keyword}」を件名に含む予定が複数ありました。誤って別の予定を動かさないよう、変更前の日付も添えて改めてご指示いただけますか。\n\n候補:\n${list}`;
  }

  const event = candidates[0];
  const title = event.getTitle();

  if (event.isRecurringEvent()) {
    return `⚠️ 秘書より報告です。「${title}」は繰り返しの予定のため、誤って全ての回を動かさないよう、変更は控えました。お手数ですが、Googleカレンダーで直接ご変更ください。`;
  }
  if (typeof event.isOwnedByMe === "function" && !event.isOwnedByMe()) {
    return `⚠️ 秘書より報告です。「${title}」は他の方から招待された予定のため、こちらでは変更できませんでした。主催者の方にご依頼いただくか、Googleカレンダーで直接ご確認ください。`;
  }

  const beforeText = describe(event);
  const ONE_DAY_MS = 24 * 60 * 60 * 1000;

  if (event.isAllDayEvent() && !isTimeString(updateData.newStartTime)) {
    // 終日の予定：日数（何日間か）はそのままに、開始日を動かす
    const oldStart = event.getAllDayStartDate();
    const oldEnd = event.getAllDayEndDate(); // 終了日の翌日（含まれない終端）
    const dayCount = Math.max(1, Math.round((oldEnd.getTime() - oldStart.getTime()) / ONE_DAY_MS));
    if (dayCount === 1) {
      event.setAllDayDate(newDay);
    } else {
      event.setAllDayDates(newDay, addDays(newDay, dayCount));
    }
    const afterText = dayCount === 1
      ? `${formatDateWithWeekday(newDay)} 終日`
      : `${formatDateWithWeekday(newDay)}〜${formatDateWithWeekday(addDays(newDay, dayCount - 1))} 終日`;
    return `📅 承知いたしました。カレンダーの予定を変更いたしました。\n・「${title}」\n・変更前：${beforeText}\n・変更後：${afterText}`;
  }

  // 時間帯つきの予定（または終日の予定に時刻が指定された場合）
  const oldDurationMs = event.isAllDayEvent()
    ? 60 * 60 * 1000
    : event.getEndTime().getTime() - event.getStartTime().getTime();
  const startTimeText = isTimeString(updateData.newStartTime)
    ? updateData.newStartTime
    : formatTime(event.getStartTime());
  const newStart = parseDateTime(updateData.newStartDate, startTimeText);
  if (!newStart) {
    return `⚠️ 秘書より確認です。「${title}」の変更後の時刻が読み取れませんでした。お手数ですが、「15時から」のように時刻を改めてご指示いただけますか。`;
  }
  let newEnd = isTimeString(updateData.newEndTime) ? parseDateTime(updateData.newStartDate, updateData.newEndTime) : null;
  if (!newEnd || newEnd.getTime() <= newStart.getTime()) {
    newEnd = new Date(newStart.getTime() + oldDurationMs);
  }

  event.setTime(newStart, newEnd);
  const afterText = `${formatDateWithWeekday(newStart)} ${formatTime(newStart)}〜${formatTime(newEnd)}`;
  return `📅 承知いたしました。カレンダーの予定を変更いたしました。\n・「${title}」\n・変更前：${beforeText}\n・変更後：${afterText}`;
}

/**
 * 処理キューから、同じスペースでの直前のやり取り（30分以内・最大3件）を取り出す。
 * 「もう一回やって」のような指示を、AIが理解できるようにするために使う。
 */
function getRecentConversation(queueData, currentIndex, spaceName, now) {
  const WINDOW_MS = 30 * 60 * 1000;
  const lines = [];
  for (let j = currentIndex - 1; j >= 1 && lines.length < 3; j--) {
    if (queueData[j][4] !== spaceName) continue; // E列: スペース名
    const receivedAt = new Date(queueData[j][1]); // B列: 受付日時
    if (isNaN(receivedAt.getTime()) || now.getTime() - receivedAt.getTime() > WINDOW_MS) break;

    const status = String(queueData[j][5] || ""); // F列: ステータス
    const result = status === "Done" ? "処理済み" : status.indexOf("Error") === 0 ? "エラーで未完了" : "処理中";
    lines.unshift(`・${formatTime(receivedAt)}「${compactText(queueData[j][3], 300)}」→ ${result}`);
  }
  return lines;
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
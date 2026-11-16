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
 * 朝のご報告のうち、カレンダーに関する部分（今日の予定／今週・来週の予定／準備の助言）を組み立てる
 */
function buildCalendarBriefing(now, openTaskLines) {
  const today = startOfDay(now);
  const todayEvents = getCalendarEvents(today, addDays(today, 1));

  const upcomingRange = getUpcomingRange(now);
  const upcomingEvents = getCalendarEvents(upcomingRange.start, upcomingRange.end);
  const advice = buildPreparationAdvice(upcomingEvents, openTaskLines, now);

  const todaySection = "■ 今日の予定\n" +
    (todayEvents.length > 0 ? formatEventLines(todayEvents, false, 15).join("\n") : "・予定は入っておりません");

  const upcomingSection = "■ 今週・来週の予定\n" +
    (upcomingEvents.length > 0 ? formatEventLines(upcomingEvents, true, 15).join("\n") : "・予定は入っておりません");

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

    if (status === "完了" || !taskContent) continue;

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
  if (calendarSpaceName) {
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
        calendarBriefing.upcomingSection,
        calendarBriefing.adviceSection
      ].filter(section => section);
      message = `🤖 瀬戸口さん、おはようございます。${formatDateWithWeekday(now)} の朝のご報告です。\n\n` + sections.join("\n\n");
    } else {
      message = "🤖 秘書より、本日のタスクリマインドです。\n\n" + taskLines;
    }

    // Chatの1メッセージあたりの文字数上限に収まるよう、長すぎる場合は末尾を切り詰める
    if (message.length > 3900) {
      message = message.substring(0, 3900) + "\n…（長いため以降は省略しました）";
    }

    try {
      sendPushReply(spaceName, message);
      // 送信に成功したタスクのみ、最終リマインド日時を更新する
      const timestamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
      items.forEach(item => taskSheet.getRange(item.rowNum, 7).setValue(timestamp));
    } catch (e) {
      console.error("朝のご報告の送信に失敗しました（" + spaceName + "）: " + e.message);
    }
  }
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

  const calendarEvents = getCalendarEvents(rangeStart, addDays(rangeEnd, 1));
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
function handleUserIntent(userInput, userName, spaceName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) return "⚠️ エラー: APIキーが設定されていません。";

  // 「明日」「来週の月曜」などの言葉を理解できるよう、現在の日時と曜日を教える
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy年MM月dd日") +
                   `（${WEEKDAY_LABELS[now.getDay()]}曜日）` +
                   Utilities.formatDate(now, "Asia/Tokyo", " HH:mm");

  const systemPrompt = `
あなたは瀬戸口さんの優秀な専属秘書botです。
現在の日時は【${todayStr}】です。

瀬戸口さん（${userName}）からのメッセージ:
「${userInput}」

【指示】
瀬戸口さんのメッセージの意図を以下の6つのいずれかに分類し、必ず指定されたJSONフォーマットの文字列のみを出力してください（余計な解説文などは一切含めないでください）。

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

5. メッセージが「Googleカレンダーへの予定の登録」を求めている場合（例：「20日の14時からA社打ち合わせ、カレンダーに登録しといて」「来週月曜に歯医者の予定を入れて」など。「カレンダー」「予定を入れて」「スケジュール登録」といった言葉がある場合はタスクではなくこちら）
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

6. メッセージが「Googleカレンダーの予定の確認」を求めている場合（例：「今日の予定は？」「来週の予定を教えて」「明日なにがあったっけ」など）
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

  // 新規タスク登録の処理
  if (resultData.action === "task") {
    writeTaskToSheet(resultData, userName, spaceName);
    return resultData.replyMessage;
  }

  // 既存タスクの進捗更新処理
  if (resultData.action === "update_task") {
    return updateTaskInSheets(resultData, userName);
  }

  // カレンダーへの予定登録・予定確認の処理
  if (resultData.action === "calendar_add" || resultData.action === "calendar_query") {
    try {
      return resultData.action === "calendar_add"
        ? addCalendarEventFromIntent(resultData, now)
        : answerCalendarQuery(resultData, now);
    } catch (calendarError) {
      console.error("カレンダー操作でエラーが発生しました: " + calendarError.toString());
      return "⚠️ 申し訳ありません。カレンダーの操作に失敗しました。少し時間をおいて再度お試しください。";
    }
  }

  // 要約結果や下書き、会話などそのままチャットに返す処理
  return resultData.replyMessage;
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
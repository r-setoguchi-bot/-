/**
 * Google Chatからメッセージを受け取ったときに実行される関数（30秒ルール完全回避版）
 */
function onMessage(event) {
  // デバッグ用：届いたデータの形をログに記録する
  console.log("受信イベントデータ: " + JSON.stringify(event));

  // 1. ユーザー名を取得（アドオン形式の深い階層から確実に抽出）
  const userName = event.chat?.messagePayload?.message?.sender?.displayName || 
                   event.chat?.user?.displayName || 
                   "瀬戸口さん";
  
  // 2. メッセージ本文を取得（アドオン形式の深い階層から確実に抽出）
  const userInput = event.chat?.messagePayload?.message?.text || 
                    event.message?.text || 
                    "";
                    
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
  const queueSheet = spreadsheet.getSheetByName("処理キュー") || spreadsheet.getSheets()[2]; // 3枚目のシートを取得
  
  const now = new Date();
  const timestamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  const queueId = "Q" + Utilities.formatDate(now, "Asia/Tokyo", "yyyyMMddHHmmss");
  
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
  return createChatResponse("🤖 瀬戸口部長、ご指示を承りました。ただいま秘書が内容を確認し、手帳の整理と解析を行っております。少々お待ちくださいませ…");
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
 * ⏰ 1分ごとに定期実行され、「処理キュー」シートを巡回するバッチ関数
 */
function processQueue() {
  // 二重起動防止のロックを取得（30秒待機）
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;

  try {
    const sheetUrl = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_URL");
    const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
    const queueSheet = spreadsheet.getSheetByName("処理キュー") || spreadsheet.getSheets()[2];
    const lastRow = queueSheet.getLastRow();
    
    if (lastRow < 2) return; // データがなければ終了

    const data = queueSheet.getDataRange().getValues();

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
        const replyText = handleUserIntent(userInput, userName);

        // サービスアカウントを使ってチャットスペースへ非同期プッシュ送信（返信）
        sendPushReply(spaceName, replyText);

        // ステータスを「Done（完了）」に更新
        queueSheet.getRange(rowNum, 6).setValue("Done");

      } catch (error) {
        console.error("キュー処理エラー: " + error.message);
        queueSheet.getRange(rowNum, 6).setValue("Error: " + error.message);
        sendPushReply(spaceName, "⚠️ 申し訳ありません。処理中にエラーが発生しました。\n詳細: " + error.message);
      }
      SpreadsheetApp.flush();
    }
  } finally {
    // ロックを解放
    lock.releaseLock();
  }
}

/**
 * Gemini APIを呼び出してユーザーの意図（新規タスク・進捗更新・要約・会話）を判断し、適切に処理する関数
 */
function handleUserIntent(userInput, userName) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) return "⚠️ エラー: APIキーが設定されていません。";

  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=" + apiKey;

  // 「明日」「来週」などの言葉を理解できるよう、現在の日時を教える
  const now = new Date();
  const todayStr = Utilities.formatDate(now, "Asia/Tokyo", "yyyy年MM月dd日 HH:mm");

  const systemPrompt = `
あなたは瀬戸口部長の優秀な専属秘書botです。
現在の日時は【${todayStr}】です。

瀬戸口部長（${userName}）からのメッセージ:
「${userInput}」

【指示】
瀬戸口部長のメッセージの意図を以下の4つのいずれかに分類し、必ず指定されたJSONフォーマットの文字列のみを出力してください（余計な解説文などは一切含めないでください）。

1. メッセージが「新しいタスクの依頼や追加」である場合（例：「明日までに資料作成」「タスクに追加して」など）
{
  "action": "task",
  "taskContent": "タスクの具体的な内容",
  "deadlineDate": "YYYY/MM/DD (期限が不明な場合は空文字)",
  "deadlineTime": "HH:mm (時間が不明な場合は空文字)",
  "assignedPerson": "このタスクの担当者（部下の名前や瀬戸口部長の名前など。明示がなければ「瀬戸口部長」）",
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
    "contents": [{ "parts": [{ "text": systemPrompt }] }]
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
          writeTaskToSheet(resultData, userName);
          return resultData.replyMessage;
        }
        
        // 既存タスクの進捗更新処理（新機能！）
        if (resultData.action === "update_task") {
          return updateTaskInSheets(resultData, userName);
        }
        
        // 要約結果や下書き、会話などそのままチャットに返す処理
        return resultData.replyMessage;

      } catch (parseError) {
        return "⚠️ 解析エラー: AIの応答をデータ化できませんでした。\n" + aiResponse;
      }
    } else {
      // 💡 エラーの原因をチャット画面に直接あぶり出すよう修正！
      return "⚠️ 申し訳ありません。Geminiの応答をうまく解析できませんでした。\n詳細: " + JSON.stringify(json);
    }
  } catch (e) {
    return "⚠️ 通信エラー: " + e.toString();
  }
}

/**
 * スプレッドシートに新規タスクを書き込む関数
 */
function writeTaskToSheet(taskData, userName) {
  const sheetUrl = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_URL");
  const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  const sheet = spreadsheet.getSheets()[0]; // 1つ目のシート（一番左のシート）を取得

  const now = new Date();
  const timestamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  
  // タスクID（T + 年月日時間 で自動生成）
  const taskId = "T" + Utilities.formatDate(now, "Asia/Tokyo", "yyyyMMddHHmmss");

  // スプレッドシートの各列（A〜I列）に合わせてデータを追加
  sheet.appendRow([
    taskId,                              // A: タスクID
    timestamp,                           // B: 登録日時
    taskData.taskContent,                // C: タスク内容
    taskData.deadlineDate,               // D: 期限（日付）
    taskData.deadlineTime,               // E: 期限（時間）
    "未完了",                            // F: ステータス
    "",                                  // G: 最終リマインド日時
    taskData.assignedPerson || userName, // H: 担当者（ボール保持者）
    taskData.statusSummary || "新規登録" // I: 最新の状況サマリー
  ]);
}

/**
 * 既存タスクを検索し、進捗の上書きと履歴シートへの追記を行う関数
 */
function updateTaskInSheets(updateData, userName) {
  const sheetUrl = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_URL");
  const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);
  const taskSheet = spreadsheet.getSheets()[0]; // 1枚目：タスク一覧
  const logSheet = spreadsheet.getSheets()[1];  // 2枚目：対応履歴ログ

  const taskDataRange = taskSheet.getDataRange().getValues();
  const keyword = updateData.targetTaskKeyword.toLowerCase();
  let foundRowIndex = -1;
  let taskId = "";

  // 最新のタスクから優先して見つけるため、下の行から逆順に検索
  for (let i = taskDataRange.length - 1; i >= 1; i--) {
    const taskContent = String(taskDataRange[i][2]).toLowerCase(); // C列: タスク内容
    if (taskContent.indexOf(keyword) !== -1) {
      foundRowIndex = i + 1;        // 行番号を取得（1始まり）
      taskId = taskDataRange[i][0]; // A列: タスクIDを取得
      break;
    }
  }

  const now = new Date();
  const timestamp = Utilities.formatDate(now, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

  // 対象のタスクが見つかった場合の処理
  if (foundRowIndex !== -1) {
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
    const logId = "L" + Utilities.formatDate(now, "Asia/Tokyo", "yyyyMMddHHmmss");
    logSheet.appendRow([
      logId,                     // A: 履歴ID
      taskId,                    // B: タスクID（紐付け用）
      timestamp,                 // C: 記録日時
      updateData.progressDetail, // D: 対応内容・経緯
      updateData.assignedPerson  // E: 担当者（ボール保持者）
    ]);

    return updateData.replyMessage;
  } else {
    // 検索キーワードに一致するタスクがなかった場合
    return `⚠️ 秘書より報告です。「${updateData.targetTaskKeyword}」に該当するタスクが一覧から見つけられませんでした。お手数ですが、正確な名前でご指示いただくか、新規タスクとしてご登録ください。`;
  }
}

/**
 * 外部の処理からGoogle Chatへメッセージを非同期プッシュ送信する関数（過去の成功コードから移植）
 */
function sendPushReply(spaceName, text) {
  if (!spaceName) return;
  try {
    const token = getBotToken(); 
    const payload = {
      "text": text
    };
    UrlFetchApp.fetch(`https://chat.googleapis.com/v1/${spaceName}/messages`, {
      method: "post",
      headers: { "Authorization": "Bearer " + token },
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error("プッシュ送信失敗: " + e.message);
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
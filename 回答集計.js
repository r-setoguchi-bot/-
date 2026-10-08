// ===== フォームの回答を1つのシートに集める =====
// 「リスト」シートに記録されたフォームID（M列）を元に、作成済みの全フォームの回答を読み取り、
// 店舗ごとに1行の「回答集計」シートにまとめる。フォームはそれぞれ独立しているため、定期的に各フォームを読みに行く。
// 1回の実行は RESP_TIME_LIMIT_MS で打ち切り、残りは時間主導トリガーで自動的に続きを実行する。
//
// 使い方:
//   1. setupResponseCollectionTrigger を1回だけ実行 → RESP_INTERVAL_HOURS 時間ごとに自動で集計される
//   2. 今すぐ集計したいときは startResponseCollection を実行
//   3. 自動集計を止めるときは removeResponseCollectionTrigger を実行
//
// スプレッドシートの指定は フォーム作成.js と同じ（FORM_SPREADSHEET_URL、無ければ紐づいているスプレッドシート）。
// このファイルだけで動く（質問文は フォーム作成.js と同じものを下に書いている）。

const RESP_LIST_SHEET_NAME = "リスト";
const RESP_OUTPUT_SHEET_NAME = "回答集計";
const RESP_STATUS_SENT = "送信済";

const RESP_INTERVAL_HOURS = 2;                // 自動で集計する間隔（時間）
const RESP_TIME_LIMIT_MS = 4.5 * 60 * 1000;   // 1回あたりの処理時間上限（6分制限対策）
const RESP_RETRY_DELAY_MS = 60 * 1000;        // 続きを実行するまでの待ち時間
const RESP_CURSOR_PROPERTY = "RESP_CURSOR";   // 何番目のフォームまで読んだか
const RESP_START_HANDLER = "startResponseCollection";
const RESP_CONTINUE_HANDLER = "collectFormResponses";

// フォーム作成.js の質問文・選択肢と同じ文言にしておくこと
const RESP_Q_CONTACT_NAME = "ご担当者様のお名前";
const RESP_Q_PHONE = "ご連絡可能なお電話番号";
const RESP_Q_COMMENT = "ご不明点・ご要望（任意）";
const RESP_Q_WANT = "特別回収期間の回収について";
const RESP_CHOICE_WANT = "特別回収期間の回収を希望する（有料）";
const RESP_CHOICE_NOT_WANT = "すべての日程において回収を希望しない";
const RESP_ANSWER_YES = "希望する";
const RESP_ANSWER_NO = "希望しない";
const RESP_DATE_QUESTION_PATTERN = /^【(.+)】(.+)の収集を希望しますか？$/;

// 「リスト」シートの列（0始まり）
const RESP_COL_STORE = 0;   // A列: 店舗名
const RESP_COL_EMAIL = 1;   // B列: メールアドレス
const RESP_COL_DATE_START = 3;  // D列〜K列: 希望日（「○」の日だけ設問がある）
const RESP_COL_DATE_END = 10;
const RESP_COL_FORM_ID = 12;    // M列: フォームID
const RESP_COL_STATUS = 13;     // N列: 送信状況

function respGetSpreadsheet() {
  const url = PropertiesService.getScriptProperties().getProperty("FORM_SPREADSHEET_URL");
  if (url) {
    return SpreadsheetApp.openByUrl(url);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("スクリプトプロパティ FORM_SPREADSHEET_URL を設定してください。");
  }
  return ss;
}

// 自動集計のタイマーを登録する（1回だけ実行する）
function setupResponseCollectionTrigger() {
  deleteTriggersByHandler(RESP_START_HANDLER);
  ScriptApp.newTrigger(RESP_START_HANDLER).timeBased().everyHours(RESP_INTERVAL_HOURS).create();
  console.log(`${RESP_INTERVAL_HOURS}時間ごとの自動集計を登録しました。`);
  startResponseCollection();
}

// 自動集計を止める（集計済みの「回答集計」シートはそのまま残る）
function removeResponseCollectionTrigger() {
  deleteTriggersByHandler(RESP_START_HANDLER);
  deleteTriggersByHandler(RESP_CONTINUE_HANDLER);
  PropertiesService.getScriptProperties().deleteProperty(RESP_CURSOR_PROPERTY);
  console.log("自動集計を停止しました。");
}

// 最初のフォームから集計を始める（手動実行・定期実行はこちら）
function startResponseCollection() {
  PropertiesService.getScriptProperties().deleteProperty(RESP_CURSOR_PROPERTY);
  collectFormResponses();
}

// 前回の続きから集計する（時間切れで残ったときの自動続行はこちら）
function collectFormResponses() {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) {
    console.log("別の集計が処理中のためスキップしました。");
    return;
  }

  const startTime = Date.now();
  let processedCount = 0;
  let errorCount = 0;
  let nextIndex = 0;
  let totalForms = 0;

  try {
    deleteTriggersByHandler(RESP_CONTINUE_HANDLER);

    const ss = respGetSpreadsheet();
    const listSheet = ss.getSheetByName(RESP_LIST_SHEET_NAME);
    if (!listSheet) {
      throw new Error(`シート「${RESP_LIST_SHEET_NAME}」が見つかりません。`);
    }
    const data = listSheet.getDataRange().getDisplayValues();
    const headers = data[0].slice(RESP_COL_DATE_START, RESP_COL_DATE_END + 1);
    const forms = respGroupFormsFromList(data);
    totalForms = forms.length;

    const props = PropertiesService.getScriptProperties();
    const savedCursor = Number(props.getProperty(RESP_CURSOR_PROPERTY) || 0);
    nextIndex = savedCursor < forms.length ? savedCursor : 0;

    const rowsByStore = respLoadExistingRows(ss);

    while (nextIndex < forms.length) {
      if (Date.now() - startTime > RESP_TIME_LIMIT_MS) break;

      const form = forms[nextIndex];
      try {
        const rows = respReadFormResponses(form, headers);
        rows.forEach(row => { rowsByStore[row[0]] = row; });
        processedCount++;
      } catch (e) {
        console.error(`フォームの回答を読み取れませんでした: ${form.formId} / ${e}`);
        errorCount++;
      }
      nextIndex++;
    }

    respWriteOutputSheet(ss, headers, rowsByStore);

    if (nextIndex < forms.length) {
      props.setProperty(RESP_CURSOR_PROPERTY, String(nextIndex));
      ScriptApp.newTrigger(RESP_CONTINUE_HANDLER).timeBased().after(RESP_RETRY_DELAY_MS).create();
      console.log(`${nextIndex} / ${totalForms} 件のフォームを確認しました（エラー ${errorCount} 件）。残りは自動で続行します。`);
    } else {
      props.deleteProperty(RESP_CURSOR_PROPERTY);
      console.log(`全 ${totalForms} 件のフォームを確認しました（今回 ${processedCount} 件、エラー ${errorCount} 件）。回答のある店舗: ${Object.keys(rowsByStore).length} 件`);
      respRefreshContractStatusList();
    }
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      console.warn(`ロック解除に失敗しました（処理には影響しません）: ${e}`);
    }
  }
}

// 「リスト」から、フォームごとの店舗（と、設問のある日）を取り出す
function respGroupFormsFromList(data) {
  const formsById = {};
  const forms = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const formId = row[RESP_COL_FORM_ID];
    if (!formId || row[RESP_COL_STATUS] !== RESP_STATUS_SENT) continue;

    if (!formsById[formId]) {
      formsById[formId] = { formId: formId, stores: [] };
      forms.push(formsById[formId]);
    }
    formsById[formId].stores.push({
      name: row[RESP_COL_STORE],
      email: row[RESP_COL_EMAIL],
      offered: row.slice(RESP_COL_DATE_START, RESP_COL_DATE_END + 1).map(mark => mark === "○")
    });
  }
  return forms;
}

// フォームの最新の回答を読み取り、店舗ごとの行にする。回答が無ければ空の配列を返す
function respReadFormResponses(form, headers) {
  const responses = FormApp.openById(form.formId).getResponses();
  if (responses.length === 0) return [];

  const latest = responses.reduce((a, b) => (a.getTimestamp() > b.getTimestamp() ? a : b));
  const parsed = respParseResponse(latest);
  const answeredAt = Utilities.formatDate(latest.getTimestamp(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

  const wantLabel = parsed.want === RESP_CHOICE_WANT ? RESP_ANSWER_YES
    : parsed.want === RESP_CHOICE_NOT_WANT ? RESP_ANSWER_NO
    : parsed.want;

  return form.stores.map(store => {
    const dateCells = [];
    headers.forEach((header, i) => {
      if (!header) return;                                    // 見出しの無い列（期間が短いときの余り）は使わない
      if (!store.offered[i]) { dateCells.push(""); return; }  // 設問が無かった日
      if (parsed.want === RESP_CHOICE_NOT_WANT) { dateCells.push("×"); return; }  // 「希望しない」で送信した場合は、全日程が「×」
      const answer = (parsed.dateAnswers[store.name] || {})[header];
      dateCells.push(answer === RESP_ANSWER_YES ? "○" : answer === RESP_ANSWER_NO ? "×" : "");
    });
    return [store.name, answeredAt, wantLabel].concat(dateCells,
      [parsed.contactName, parsed.phone, parsed.comment, store.email, form.formId]);
  });
}

// 1件の回答を、質問の文言をもとに分類する
function respParseResponse(response) {
  const parsed = { want: "", contactName: "", phone: "", comment: "", dateAnswers: {} };

  response.getItemResponses().forEach(itemResponse => {
    const title = itemResponse.getItem().getTitle();
    const answer = String(itemResponse.getResponse());

    if (title === RESP_Q_WANT) parsed.want = answer;
    else if (title === RESP_Q_CONTACT_NAME) parsed.contactName = answer;
    else if (title === RESP_Q_PHONE) parsed.phone = answer;
    else if (title === RESP_Q_COMMENT) parsed.comment = answer;
    else {
      const match = RESP_DATE_QUESTION_PATTERN.exec(title);
      if (match) {
        const storeName = match[1];
        const dateHeader = match[2];
        if (!parsed.dateAnswers[storeName]) parsed.dateAnswers[storeName] = {};
        parsed.dateAnswers[storeName][dateHeader] = answer;
      }
    }
  });
  return parsed;
}

// 前回までに集計した行を、店舗名をキーにして読み込む
function respLoadExistingRows(ss) {
  const rowsByStore = {};
  const sheet = ss.getSheetByName(RESP_OUTPUT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return rowsByStore;

  const values = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < values.length; i++) {
    rowsByStore[values[i][0]] = values[i];
  }
  return rowsByStore;
}

function respWriteOutputSheet(ss, headers, rowsByStore) {
  let sheet = ss.getSheetByName(RESP_OUTPUT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RESP_OUTPUT_SHEET_NAME);
  }
  const header = ["店舗名", "回答日時", "特別回収"].concat(headers.filter(h => h),
    ["ご担当者名", "電話番号", "コメント", "メールアドレス", "フォームID"]);
  const rows = Object.keys(rowsByStore).map(store => rowsByStore[store])
    .sort((a, b) => (a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0)); // 回答が古い順

  sheet.clear();
  sheet.getRange(1, 1, rows.length + 1, header.length).setValues([header].concat(rows));
}

// kintoneリスト作成.js があれば、集計が一巡したときに「契約状況一覧」の回答状況を最新にする
function respRefreshContractStatusList() {
  try {
    if (typeof updateContractStatusList === "function") {
      updateContractStatusList();
    }
  } catch (e) {
    console.warn(`契約状況一覧の更新に失敗しました: ${e}`);
  }
}

function deleteTriggersByHandler(handlerName) {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

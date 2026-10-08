// ===== フォームの回答を1つのシートに集める =====
// 「リスト」シートに記録されたフォームID（M列）を元に、作成済みの全フォームの回答を読み取り、
// 店舗ごとに1行の「回答集計」シートにまとめる。フォームはそれぞれ独立しているため、定期的に各フォームを読みに行く。
// 1回の実行は RESP_TIME_LIMIT_MS で打ち切り、残りは時間主導トリガーで自動的に続きを実行する。
//
// 回答の変更に気づく・止める仕組み:
//   ・再回答       … 同じフォームに2回以上回答があると、「再回答」に「あり（内容は同じ／内容が変わった）」と出る（黄色の行）
//   ・業者依頼済   … 業者に依頼した店舗に「○」を入れると、そのフォームの受付を自動で止める
//                    （同じフォームの全店舗に「○」が付いたときだけ止める。「○」を消すと受付を再開する）
//   ・依頼後の回答 … 「○」を確認した後に回答が来ていたら「あり」と出る（赤色の行）
//   手入力するのは「業者依頼済」の列だけ。ほかの列は集計のたびに書き換わる。
//
// 使い方:
//   1. setupResponseCollectionTrigger を1回だけ実行 → RESP_INTERVAL_HOURS 時間ごとに自動で集計され、
//      1時間ごとに「業者依頼済」の印を見て受付を止める／再開する
//   2. 今すぐ集計したいときは startResponseCollection、今すぐ受付を止めたいときは closeRequestedForms を実行
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
const RESP_CLOSE_HANDLER = "closeRequestedForms";

// 受付を止めたフォームを開いた人に表示するメッセージ
const RESP_CLOSED_MESSAGE = "このアンケートの回答受付は終了しました。内容の変更をご希望の場合は、お手数ですが株式会社クリメンまでご連絡ください。";
const RESP_CLOSED_CONTACT = "";               // 連絡先（電話番号など）。空のままなら表示しない

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

// 「回答集計」シートの、手入力・自動記録の列（集計のたびに書き換えず、前の値を引き継ぐ）
const RESP_HEADER_REQUESTED = "業者依頼済";
const RESP_HEADER_REQUESTED_AT = "依頼確認日時";
const RESP_HEADER_ACCEPTING = "フォーム受付";
const RESP_KEEP_HEADERS = [RESP_HEADER_REQUESTED, RESP_HEADER_REQUESTED_AT, RESP_HEADER_ACCEPTING];
const RESP_STATE_CLOSED = "停止済";
const RESP_STATE_WAITING = "他の店舗が未依頼のため受付中";
const RESP_STATE_ERROR = "停止できませんでした";

const RESP_RERESPONSE_SAME = "あり（内容は同じ）";
const RESP_RERESPONSE_CHANGED = "あり（内容が変わった）";
const RESP_COLOR_CHANGED = "#fff2cc";   // 再回答で内容が変わった行（黄色）
const RESP_COLOR_AFTER_REQUEST = "#f4cccc"; // 依頼後に回答があった行（赤色）
const RESP_COLOR_NONE = "#ffffff";

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

function respFormatNow() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
}

// 自動集計と、受付の自動停止のタイマーを登録する（1回だけ実行する）
function setupResponseCollectionTrigger() {
  deleteTriggersByHandler(RESP_START_HANDLER);
  deleteTriggersByHandler(RESP_CLOSE_HANDLER);
  ScriptApp.newTrigger(RESP_START_HANDLER).timeBased().everyHours(RESP_INTERVAL_HOURS).create();
  ScriptApp.newTrigger(RESP_CLOSE_HANDLER).timeBased().everyHours(1).create();
  console.log(`${RESP_INTERVAL_HOURS}時間ごとの自動集計と、1時間ごとの受付停止の確認を登録しました。`);
  startResponseCollection();
}

// 自動集計と受付の自動停止を止める（集計済みの「回答集計」シートはそのまま残る）
function removeResponseCollectionTrigger() {
  deleteTriggersByHandler(RESP_START_HANDLER);
  deleteTriggersByHandler(RESP_CONTINUE_HANDLER);
  deleteTriggersByHandler(RESP_CLOSE_HANDLER);
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
    const outputHeader = respBuildOutputHeader(headers);
    const forms = respGroupFormsFromList(data);
    totalForms = forms.length;

    const props = PropertiesService.getScriptProperties();
    const savedCursor = Number(props.getProperty(RESP_CURSOR_PROPERTY) || 0);
    nextIndex = savedCursor < forms.length ? savedCursor : 0;

    const rowsByStore = respLoadExistingRows(ss, outputHeader);

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

    respWriteOutputSheet(ss, outputHeader, rowsByStore);

    if (nextIndex < forms.length) {
      props.setProperty(RESP_CURSOR_PROPERTY, String(nextIndex));
      ScriptApp.newTrigger(RESP_CONTINUE_HANDLER).timeBased().after(RESP_RETRY_DELAY_MS).create();
      console.log(`${nextIndex} / ${totalForms} 件のフォームを確認しました（エラー ${errorCount} 件）。残りは自動で続行します。`);
    } else {
      props.deleteProperty(RESP_CURSOR_PROPERTY);
      console.log(`全 ${totalForms} 件のフォームを確認しました（今回 ${processedCount} 件、エラー ${errorCount} 件）。回答のある店舗: ${Object.keys(rowsByStore).length} 件`);
      respApplyRequestedStatus(ss, forms);
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

// 「業者依頼済」の印を見て、フォームの受付を止める／再開する（手動実行・1時間ごとの定期実行はこちら）
function closeRequestedForms() {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(30 * 1000)) {
    console.log("集計が処理中のためスキップしました。");
    return;
  }
  try {
    const ss = respGetSpreadsheet();
    const listSheet = ss.getSheetByName(RESP_LIST_SHEET_NAME);
    if (!listSheet) {
      throw new Error(`シート「${RESP_LIST_SHEET_NAME}」が見つかりません。`);
    }
    respApplyRequestedStatus(ss, respGroupFormsFromList(listSheet.getDataRange().getDisplayValues()));
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      console.warn(`ロック解除に失敗しました（処理には影響しません）: ${e}`);
    }
  }
}

// 「回答集計」シートの見出し（日付の列は、「リスト」の見出しのうち空でないもの）
function respBuildOutputHeader(headers) {
  return ["店舗名", "回答日時", "特別回収"].concat(headers.filter(h => h),
    ["ご担当者名", "電話番号", "コメント", "メールアドレス", "フォームID", "回答回数", "再回答", "依頼後の回答"],
    RESP_KEEP_HEADERS);
}

// 見出しから、店舗ごとの行のうち「フォームの回答から作る部分」の列数を求める（店舗名〜再回答）
function respBaseLength(outputHeader) {
  return outputHeader.length - 1 - RESP_KEEP_HEADERS.length;
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

// フォームの回答を読み取り、店舗ごとの行（店舗名〜再回答）にする。回答が無ければ空の配列を返す
function respReadFormResponses(form, headers) {
  const responses = FormApp.openById(form.formId).getResponses();
  if (responses.length === 0) return [];

  const sorted = responses.slice().sort((a, b) => a.getTimestamp() - b.getTimestamp());
  const latest = sorted[sorted.length - 1];
  const parsed = respParseResponse(latest);
  const previous = sorted.length > 1 ? respParseResponse(sorted[sorted.length - 2]) : null;
  const answeredAt = Utilities.formatDate(latest.getTimestamp(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

  return form.stores.map(store => {
    const current = respBuildStoreAnswer(parsed, store, headers);

    let reresponse = "";
    if (previous) {
      const before = respBuildStoreAnswer(previous, store, headers);
      const same = before.wantLabel === current.wantLabel && before.dateCells.join("|") === current.dateCells.join("|");
      reresponse = same ? RESP_RERESPONSE_SAME : RESP_RERESPONSE_CHANGED;
    }

    return [store.name, answeredAt, current.wantLabel].concat(current.dateCells,
      [parsed.contactName, parsed.phone, parsed.comment, store.email, form.formId, responses.length, reresponse]);
  });
}

// 1件の回答から、ある店舗の「特別回収の希望」と、日付ごとの「○／×」を作る
function respBuildStoreAnswer(parsed, store, headers) {
  const wantLabel = parsed.want === RESP_CHOICE_WANT ? RESP_ANSWER_YES
    : parsed.want === RESP_CHOICE_NOT_WANT ? RESP_ANSWER_NO
    : parsed.want;

  const dateCells = [];
  headers.forEach((header, i) => {
    if (!header) return;                                    // 見出しの無い列（期間が短いときの余り）は使わない
    if (!store.offered[i]) { dateCells.push(""); return; }  // 設問が無かった日
    if (parsed.want === RESP_CHOICE_NOT_WANT) { dateCells.push("×"); return; }  // 「希望しない」で送信した場合は、全日程が「×」
    const answer = (parsed.dateAnswers[store.name] || {})[header];
    dateCells.push(answer === RESP_ANSWER_YES ? "○" : answer === RESP_ANSWER_NO ? "×" : "");
  });
  return { wantLabel: wantLabel, dateCells: dateCells };
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

// 前回までに集計した行を、店舗名をキーにして読み込む。見出しが今と違うシートは、列がずれるので読み込まない
function respLoadExistingRows(ss, outputHeader, quiet) {
  const rowsByStore = {};
  const sheet = ss.getSheetByName(RESP_OUTPUT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return rowsByStore;

  const values = sheet.getDataRange().getDisplayValues();
  if (values[0].join("\t") !== outputHeader.join("\t")) {
    if (!quiet) console.warn("「回答集計」シートの見出しが今のものと違うため、前回までの集計は引き継がずに作り直します。");
    return rowsByStore;
  }
  for (let i = 1; i < values.length; i++) {
    rowsByStore[values[i][0]] = values[i];
  }
  return rowsByStore;
}

function respWriteOutputSheet(ss, outputHeader, rowsByStore) {
  let sheet = ss.getSheetByName(RESP_OUTPUT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(RESP_OUTPUT_SHEET_NAME);
  }
  const baseLength = respBaseLength(outputHeader);

  // 書き込む直前のシートから、手入力・自動記録の列（業者依頼済など）を取り込む
  const current = respLoadExistingRows(ss, outputHeader, true);

  const rows = Object.keys(rowsByStore).map(store => {
    const base = rowsByStore[store].slice(0, baseLength);
    const kept = current[store] ? current[store].slice(baseLength + 1) : [];
    const keep = RESP_KEEP_HEADERS.map((_, i) => kept[i] || "");
    const requestedAt = keep[1];
    const answeredAt = base[1];
    const afterRequest = requestedAt && answeredAt > requestedAt ? "あり" : "";
    return base.concat([afterRequest], keep);
  }).sort((a, b) => (a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0)); // 回答が古い順

  sheet.clear();
  sheet.getRange(1, 1, rows.length + 1, outputHeader.length).setValues([outputHeader].concat(rows));

  if (rows.length > 0) {
    const backgrounds = rows.map(row => {
      const color = row[baseLength] === "あり" ? RESP_COLOR_AFTER_REQUEST
        : row[baseLength - 1] === RESP_RERESPONSE_CHANGED ? RESP_COLOR_CHANGED
        : RESP_COLOR_NONE;
      return outputHeader.map(() => color);
    });
    sheet.getRange(2, 1, rows.length, outputHeader.length).setBackgrounds(backgrounds);
  }
}

// 「業者依頼済」の印を見て、フォームの受付を止める／再開する。
// 同じフォームの全店舗に印があるときだけ止め、印が外れたら再開する。受付の状態は「フォーム受付」の列に書く。
function respApplyRequestedStatus(ss, forms) {
  const sheet = ss.getSheetByName(RESP_OUTPUT_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return;

  const values = sheet.getDataRange().getDisplayValues();
  const header = values[0];
  const colRequested = header.indexOf(RESP_HEADER_REQUESTED);
  const colRequestedAt = header.indexOf(RESP_HEADER_REQUESTED_AT);
  const colAccepting = header.indexOf(RESP_HEADER_ACCEPTING);
  if (colRequested < 0 || colRequestedAt !== colRequested + 1 || colAccepting !== colRequested + 2) {
    throw new Error("「回答集計」シートの見出しが想定と違います。startResponseCollection を実行して作り直してください。");
  }

  const rowByStore = {};
  for (let i = 1; i < values.length; i++) rowByStore[values[i][0]] = i;
  const isRequested = store => String(values[rowByStore[store.name]][colRequested]).trim() !== "";

  const now = respFormatNow();
  let closedCount = 0;
  let reopenedCount = 0;
  let errorCount = 0;

  forms.forEach(form => {
    const stores = form.stores.filter(store => rowByStore[store.name] !== undefined); // 回答のある店舗
    if (stores.length === 0) return;

    const shouldClose = stores.length === form.stores.length && stores.every(isRequested);
    const wasClosed = stores.some(store => values[rowByStore[store.name]][colAccepting] === RESP_STATE_CLOSED);

    let failed = false;
    if (shouldClose !== wasClosed) {
      try {
        const target = FormApp.openById(form.formId);
        if (shouldClose) {
          target.setAcceptingResponses(false);
          target.setCustomClosedFormMessage(RESP_CLOSED_MESSAGE + (RESP_CLOSED_CONTACT ? "\n" + RESP_CLOSED_CONTACT : ""));
          closedCount++;
        } else {
          target.setAcceptingResponses(true);
          reopenedCount++;
        }
      } catch (e) {
        console.error(`フォームの受付を切り替えられませんでした: ${form.formId} / ${e}`);
        failed = true;
        errorCount++;
      }
    }

    stores.forEach(store => {
      const row = values[rowByStore[store.name]];
      if (!isRequested(store)) {
        row[colRequestedAt] = "";
        row[colAccepting] = failed ? RESP_STATE_ERROR : "";
        return;
      }
      row[colRequestedAt] = row[colRequestedAt] || now;   // 最初に印を確認した日時（依頼後の回答かどうかの基準）
      row[colAccepting] = failed ? RESP_STATE_ERROR : (shouldClose ? RESP_STATE_CLOSED : RESP_STATE_WAITING);
    });
  });

  // 自動記録の2列（依頼確認日時・フォーム受付）だけを書き戻す。手入力の「業者依頼済」の列には触れない
  const autoColumns = values.slice(1).map(row => [row[colRequestedAt], row[colAccepting]]);
  sheet.getRange(2, colRequestedAt + 1, autoColumns.length, 2).setValues(autoColumns);

  console.log(`受付を止めたフォーム: ${closedCount} 件、再開したフォーム: ${reopenedCount} 件、失敗: ${errorCount} 件`);
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

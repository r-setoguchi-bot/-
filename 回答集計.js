// ===== フォームの回答を1つのシートに集める =====
// 「リスト」シートに記録されたフォームID（J列）を元に、作成済みの全フォームの回答を読み取り、
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
// 回答の控えメール（エビデンス用）:
//   回答を読み取ったときに、回答内容をまとめたメールをお客様（フォームのURLを送ったアドレス）へ送る。
//   スクリプトプロパティ RESP_MAIL_MODE で切り替える。未設定なら送らない。
//     off  … 送らない（初期値）
//     test … お客様には送らず、RESP_MAIL_TEST_TO のアドレスだけに送る（動作確認用）
//     on   … お客様へ送る。RESP_MAIL_BCC があれば、そのアドレスにも同じメールを送る（社内の控え）
//   同じ回答には1回だけ送る。再回答があれば、その回答の分をもう一度送る。送信の結果は「控えメール」の列に残る。
//   メールには、変更するためのフォームのURLと、お客様に案内する変更の期限（RESP_CUSTOMER_DEADLINE_TEXT）を書く。
//   社内の控え（RESP_MAIL_BCC）は、必要なときだけ設定する。回答はフォームの中に日時つきで全件残る。
//   メールはスクリプトを実行しているアカウントから送られ、送信できる数には1日の上限がある。
//
// 使い方:
//   1. setupResponseCollectionTrigger を1回だけ実行 → RESP_INTERVAL_HOURS 時間ごとに自動で集計され、
//      1時間ごとに「業者依頼済」の印を見て受付を止める／再開する
//   2. 今すぐ集計したいときは startResponseCollection、今すぐ受付を止めたいときは closeRequestedForms を実行
//      控えメールの文面を確かめたいときは sendSampleReceiptMail を実行（RESP_MAIL_TEST_TO 宛てにサンプルを送る）
//   3. 自動集計を止めるときは removeResponseCollectionTrigger を実行
//
// 変更の最終日時:
//   RESP_CUTOFF_AT に日時を入れると、その日時を過ぎたあと、全フォームの受付を自動で止める（1時間ごとに確認）。
//   お客様は、それまでは同じフォームから何度でも回答し直せる（最新の回答が有効）。空のままなら自動では止めない。
//   お客様に案内する期限（RESP_CUSTOMER_DEADLINE_TEXT / FORM_CHANGE_DEADLINE_TEXT）より後ろの日時にしておくと、少し余裕が持てる。
//
// スプレッドシートの指定は フォーム作成.js と同じ（FORM_SPREADSHEET_URL、無ければ紐づいているスプレッドシート）。
// このファイルだけで動く（質問文は フォーム作成.js と同じものを下に書いている）。

const RESP_LIST_SHEET_NAME = "リスト";
const RESP_OUTPUT_SHEET_NAME = "回答集計";
const RESP_STATUS_SENT = "送信済";

const RESP_INTERVAL_HOURS = 2;                // 自動で集計する間隔（時間）
const RESP_TIME_LIMIT_MS = 4.5 * 60 * 1000;   // 1回あたりの処理時間上限（6分制限対策）
const RESP_RETRY_DELAY_MS = 60 * 1000;        // 続きを実行するまでの待ち時間
const RESP_SAFETY_DELAY_MS = 7 * 60 * 1000;   // 実行が強制終了した場合に再開するまでの待ち時間（6分制限より長く）
const RESP_CURSOR_PROPERTY = "RESP_CURSOR";   // 何番目のフォームまで読んだか
const RESP_START_HANDLER = "startResponseCollection";
const RESP_CONTINUE_HANDLER = "collectFormResponses";
const RESP_CLOSE_HANDLER = "closeRequestedForms";

// 実際にフォームの受付を止める日時（例: "2026/11/30 23:59"）。過ぎたら全フォームの受付を止める。空のままなら止めない。
const RESP_CUTOFF_AT = "2026/11/30 23:59";
// お客様に案内する「変更の期限」の表記（控えメールに書く）。受付を止める日時とは別。フォーム作成.js の FORM_CHANGE_DEADLINE_TEXT と同じにしておくこと。
// 空のままなら、メールに期限の文を入れない
const RESP_CUSTOMER_DEADLINE_TEXT = "2026年11月15日(日)";
const RESP_CUTOFF_HANDLER = "closeAllFormsAtCutoff";
const RESP_CUTOFF_CONTINUE_HANDLER = "continueCloseAllFormsAtCutoff";
const RESP_CUTOFF_CURSOR_PROPERTY = "RESP_CUTOFF_CURSOR";
const RESP_CUTOFF_DONE_PROPERTY = "RESP_CUTOFF_DONE";

// 受付を止めたフォームを開いた人に表示するメッセージ
const RESP_CLOSED_MESSAGE = "このアンケートの回答受付は終了しました。内容の変更をご希望の場合は、お手数ですが株式会社クリメンまでご連絡ください。";
const RESP_CLOSED_CONTACT = "";               // 連絡先（電話番号など）。空のままなら表示しない

// 控えメール
const RESP_MAIL_MODE_PROPERTY = "RESP_MAIL_MODE";           // off / test / on
const RESP_MAIL_TEST_TO_PROPERTY = "RESP_MAIL_TEST_TO";     // test のときの宛先
const RESP_MAIL_FROM_PROPERTY = "MAIL_FROM";                // 送信元アドレス（会社のアドレス。案内メールと共通）。未設定なら実行アカウントのアドレス
const RESP_MAIL_BCC_PROPERTY = "RESP_MAIL_BCC";             // on のときに、社内の控えとして同じメールを送るアドレス（任意）
const RESP_MAIL_SENDER_NAME = "株式会社クリメン";
const RESP_MAIL_SUBJECT = "【株式会社クリメン】年末年始廃棄物回収アンケート ご回答内容の控え";
const RESP_MAIL_PERIOD_TEXT = "2026年12月30日(水)から2027年1月3日(日)";
const RESP_MAIL_CONTACT = "";                               // 内容の変更の連絡先（電話番号など）。空のままなら案内の文だけ入れる

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
const RESP_COL_DATE_START = 3;  // D列〜H列: 希望日（「○」の日だけ設問がある）
const RESP_COL_DATE_END = 7;
const RESP_COL_FORM_URL = 8;    // I列: フォームURL
const RESP_COL_FORM_ID = 9;     // J列: フォームID
const RESP_COL_STATUS = 10;     // K列: フォーム作成状況
const RESP_COL_ADDRESS = 17;    // R列: 店舗の住所（控えメールに載せる）

// 「回答集計」シートの、手入力・自動記録の列（集計のたびに書き換えず、前の値を引き継ぐ）
const RESP_HEADER_REQUESTED = "業者依頼済";
const RESP_HEADER_REQUESTED_AT = "依頼確認日時";
const RESP_HEADER_ACCEPTING = "フォーム受付";
const RESP_HEADER_MAILED = "控えメール";   // 控えメールを送った回答の日時（テスト送信は「テスト:」、失敗は「失敗:」が付く）
const RESP_KEEP_HEADERS = [RESP_HEADER_REQUESTED, RESP_HEADER_REQUESTED_AT, RESP_HEADER_ACCEPTING, RESP_HEADER_MAILED];
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
  deleteTriggersByHandler(RESP_CUTOFF_HANDLER);
  ScriptApp.newTrigger(RESP_CUTOFF_HANDLER).timeBased().everyHours(1).create();
  console.log(`${RESP_INTERVAL_HOURS}時間ごとの自動集計と、1時間ごとの受付停止の確認を登録しました。`);
  startResponseCollection();
}

// 自動集計と受付の自動停止を止める（集計済みの「回答集計」シートはそのまま残る）
function removeResponseCollectionTrigger() {
  deleteTriggersByHandler(RESP_START_HANDLER);
  deleteTriggersByHandler(RESP_CONTINUE_HANDLER);
  deleteTriggersByHandler(RESP_CLOSE_HANDLER);
  deleteTriggersByHandler(RESP_CUTOFF_HANDLER);
  deleteTriggersByHandler(RESP_CUTOFF_CONTINUE_HANDLER);
  PropertiesService.getScriptProperties().deleteProperty(RESP_CURSOR_PROPERTY);
  PropertiesService.getScriptProperties().deleteProperty(RESP_CUTOFF_CURSOR_PROPERTY);
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
    // 強制終了や、続きの予約の失敗に備えた、保険の再開トリガー
    const safetyTrigger = respScheduleTrigger(RESP_CONTINUE_HANDLER, RESP_SAFETY_DELAY_MS);

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
    const mailedByStore = {};                       // 今回、控えメールを送った（または送ろうとした）店舗
    const mailSettings = respGetMailSettings();
    const mailedColumn = outputHeader.indexOf(RESP_HEADER_MAILED);
    const dateHeaders = headers.filter(h => h);

    while (nextIndex < forms.length) {
      if (Date.now() - startTime > RESP_TIME_LIMIT_MS) break;

      const form = forms[nextIndex];
      try {
        const rows = respReadFormResponses(form, headers);
        if (rows.length > 0) {
          const before = rowsByStore[rows[0][0]];
          const previousMailed = before && before[mailedColumn] ? before[mailedColumn] : "";
          const mark = respSendReceiptIfNeeded(form, rows, dateHeaders, previousMailed, mailSettings);
          rows.forEach(row => {
            if (mark) mailedByStore[row[0]] = mark;
          });
        }
        rows.forEach(row => { rowsByStore[row[0]] = row; });
        processedCount++;
      } catch (e) {
        console.error(`フォームの回答を読み取れませんでした: ${form.formId} / ${e}`);
        errorCount++;
      }
      nextIndex++;
    }

    respWriteOutputSheet(ss, outputHeader, rowsByStore, mailedByStore);

    if (nextIndex < forms.length) {
      props.setProperty(RESP_CURSOR_PROPERTY, String(nextIndex));
      // 先に続きの予約を作る（失敗しても、保険の予約が残る）。そのあとで、保険の予約を消す
      respScheduleTrigger(RESP_CONTINUE_HANDLER, RESP_RETRY_DELAY_MS);
      try {
        ScriptApp.deleteTrigger(safetyTrigger);
      } catch (e) {
        console.warn(`保険の予約を消せませんでした（処理には影響しません）: ${e}`);
      }
      console.log(`${nextIndex} / ${totalForms} 件のフォームを確認しました（エラー ${errorCount} 件）。残りは自動で続行します。`);
    } else {
      deleteTriggersByHandler(RESP_CONTINUE_HANDLER);
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
      formsById[formId] = { formId: formId, formUrl: row[RESP_COL_FORM_URL], stores: [] };
      forms.push(formsById[formId]);
    }
    formsById[formId].stores.push({
      name: row[RESP_COL_STORE],
      email: row[RESP_COL_EMAIL],
      address: String(row[RESP_COL_ADDRESS] || "").trim(),
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
  const dateCells = [];
  headers.forEach((header, i) => {
    if (!header) return;                                    // 見出しの無い列（期間が短いときの余り）は使わない
    if (!store.offered[i]) { dateCells.push(""); return; }  // 設問が無かった日
    if (parsed.want === RESP_CHOICE_NOT_WANT) { dateCells.push("×"); return; }  // 旧形式のフォーム：「希望しない」で送信した場合は、全日程が「×」
    const answer = (parsed.dateAnswers[store.name] || {})[header];
    dateCells.push(answer === RESP_ANSWER_YES ? "○" : answer === RESP_ANSWER_NO ? "×" : "");
  });

  // 「特別回収」の列: 旧形式のフォームは最初の選択肢から、今の形式は日ごとの回答から決める（1日でも「希望する」があれば「希望する」）
  let wantLabel = parsed.want === RESP_CHOICE_WANT ? RESP_ANSWER_YES
    : parsed.want === RESP_CHOICE_NOT_WANT ? RESP_ANSWER_NO
    : parsed.want;
  if (!parsed.want) {
    wantLabel = dateCells.indexOf("○") !== -1 ? RESP_ANSWER_YES : dateCells.indexOf("×") !== -1 ? RESP_ANSWER_NO : "";
  }
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

function respWriteOutputSheet(ss, outputHeader, rowsByStore, mailedByStore) {
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
    if (mailedByStore && mailedByStore[store]) keep[RESP_KEEP_HEADERS.indexOf(RESP_HEADER_MAILED)] = mailedByStore[store];
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
  if (respIsAfterCutoff()) {
    console.log("変更の最終日時を過ぎているため、「業者依頼済」による受付の切り替えは行いません。");
    return;
  }

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
          target.setCustomClosedFormMessage(respClosedMessage());
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

// 控えメールの設定を読む。modeが on / test 以外のときは送らない
function respGetMailSettings() {
  const props = PropertiesService.getScriptProperties();
  const rawMode = String(props.getProperty(RESP_MAIL_MODE_PROPERTY) || "off").trim().toLowerCase();
  const testTo = String(props.getProperty(RESP_MAIL_TEST_TO_PROPERTY) || "").trim();
  const bcc = String(props.getProperty(RESP_MAIL_BCC_PROPERTY) || "").trim();

  const from = String(props.getProperty(RESP_MAIL_FROM_PROPERTY) || "").trim();

  let mode = rawMode === "on" || rawMode === "test" ? rawMode : "off";
  if (mode !== "off" && from) {
    const own = String(Session.getEffectiveUser().getEmail() || "").toLowerCase();
    const aliases = GmailApp.getAliases().map(alias => String(alias).toLowerCase());
    if (from.toLowerCase() !== own && aliases.indexOf(from.toLowerCase()) === -1) {
      console.warn(`${RESP_MAIL_FROM_PROPERTY}（${from}）はこのアカウントから送信できるアドレスではないため、控えメールは送りません。`);
      mode = "off";
    }
  }
  if (mode === "test" && !testTo) {
    console.warn(`${RESP_MAIL_MODE_PROPERTY} が test ですが、${RESP_MAIL_TEST_TO_PROPERTY} が未設定のため、控えメールは送りません。`);
    mode = "off";
  }
  return { mode: mode, testTo: testTo, bcc: bcc, from: from };
}

// 回答1件につき、控えメールを1回だけ送る。送った（または送れなかった）印の文字列を返す。何もしなかったときは空文字
//   印: 回答日時のみ=お客様に送信済 / 「テスト:」+回答日時=テスト送信済 / 「失敗:」+回答日時=送信を試みて失敗（自動では再送しない）
function respSendReceiptIfNeeded(form, rows, dateHeaders, previousMailed, mail) {
  if (mail.mode === "off") return "";

  const answeredAt = rows[0][1];
  const sentMark = answeredAt;
  const testMark = "テスト:" + answeredAt;
  const failMark = "失敗:" + answeredAt;

  if (previousMailed === sentMark || previousMailed === failMark) return "";
  if (mail.mode === "test" && previousMailed === testMark) return "";

  const customerEmail = String((form.stores[0] || {}).email || "").trim();
  const to = mail.mode === "test" ? mail.testTo : customerEmail;
  if (!/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(to)) {
    console.error(`控えメールの宛先が正しくありません: ${form.formId} / ${to}`);
    return mail.mode === "test" ? "" : failMark;
  }

  const recipientCount = 1 + (mail.mode === "on" && mail.bcc ? 1 : 0);
  if (MailApp.getRemainingDailyQuota() < recipientCount) {
    console.warn("メールの1日の送信上限に達したため、控えメールの送信を見送りました。次の集計で送ります。");
    return "";
  }

  const addresses = {};
  form.stores.forEach(store => { addresses[store.name] = store.address; });
  const message = respBuildReceiptMail(rows, dateHeaders, form.formUrl, addresses);
  const subjectPrefix = mail.mode === "test" ? "【テスト】" : "";
  const bodyPrefix = mail.mode === "test" ? `（テスト送信です。本来の宛先: ${customerEmail}）\n\n` : "";
  const options = { to: to, subject: subjectPrefix + message.subject, body: bodyPrefix + message.body, name: RESP_MAIL_SENDER_NAME };
  if (message.htmlBody) {
    options.htmlBody = (mail.mode === "test" ? `<p>（テスト送信です。本来の宛先: ${customerEmail.replace(/&/g, "&amp;").replace(/</g, "&lt;")}）</p>` : "") + message.htmlBody;
  }
  if (mail.mode === "on" && mail.bcc) options.bcc = mail.bcc;

  try {
    respSendMail(options, mail.from);
    return mail.mode === "test" ? testMark : sentMark;
  } catch (e) {
    console.error(`控えメールを送れませんでした: ${form.formId} / ${e}`);
    return mail.mode === "test" ? "" : failMark;
  }
}

// メールを送る。送信元アドレスがあれば Gmail から（別のアドレスとして登録済みのもの）、無ければ実行アカウントから送る
function respSendMail(options, from) {
  if (!from) {
    MailApp.sendEmail(options);
    return;
  }
  const gmailOptions = { name: options.name, from: from };
  if (options.bcc) gmailOptions.bcc = options.bcc;
  if (options.htmlBody) gmailOptions.htmlBody = options.htmlBody;
  GmailApp.sendEmail(options.to, options.subject, options.body, gmailOptions);
}

// 1つのフォームの回答（店舗ごとの行）から、控えメールの件名と本文を作る
function respBuildReceiptMail(rows, dateHeaders, formUrl, addresses) {
  const D = dateHeaders.length;
  const first = rows[0];
  const answeredAt = first[1];
  const contactName = first[3 + D];
  const phone = first[4 + D];
  const comment = first[5 + D];
  const responseCount = Number(first[8 + D]) || 1;

  const notice = "※収集業者の都合などにより、回収日や料金を変更せざるを得ない場合は、改めてご連絡いたします。";
  const esc = text => String(text).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const NO_COLLECTION = "回収がありません";
  const mark = cell => (cell === "○" || cell === "×" ? cell : NO_COLLECTION);

  // 店舗ごとに、特別回収期間のすべての日を載せる（回収日でない日は「回収がありません」）
  const stores = rows.map(row => ({
    name: row[0],
    address: (addresses || {})[row[0]] || "",
    want: row[2],
    days: dateHeaders.map((header, i) => ({ header: header, cell: row[3 + i] }))
  }));

  const lines = [];
  lines.push(`${contactName || "ご担当者"} 様`, "");
  lines.push("このたびは「年末年始廃棄物回収に関するアンケート」にご回答いただき、ありがとうございました。");
  lines.push("下記の内容で承りました。ご確認ください。", "");
  lines.push(`■回答日時：${answeredAt}`);
  lines.push(`■特別回収期間：${RESP_MAIL_PERIOD_TEXT}`);
  lines.push(`■ご担当者様：${contactName}`);
  lines.push(`■お電話番号：${phone}`, "");
  lines.push("【店舗ごとの回収のご希望】");
  stores.forEach(store => {
    lines.push("", `▼${store.name}`);
    if (store.address) lines.push(`  ご住所：${store.address}`);
    store.days.forEach(d => lines.push(`  ${d.header}　${mark(d.cell)}`));
  });
  if (comment) lines.push("", `■ご不明点・ご要望：${comment}`);
  lines.push("");
  lines.push(notice);
  lines.push("※本メールは、ご回答内容の控えとしてお送りしています。");
  lines.push("※回答内容を変更する場合は、下記のフォームから再度ご回答ください。最新のご回答を有効とします。");
  if (RESP_CUSTOMER_DEADLINE_TEXT) lines.push(`  変更のご希望は${RESP_CUSTOMER_DEADLINE_TEXT}までにお願いいたします。`);
  if (formUrl) lines.push(`  ${formUrl}`);
  if (RESP_MAIL_CONTACT) lines.push(`※ご不明な点は、${RESP_MAIL_CONTACT}までご連絡ください。`);
  lines.push("", RESP_MAIL_SENDER_NAME);

  // HTML版: 店舗ごとに、日付を横に並べた表にする（○＝希望する、×＝希望しない）
  const cellStyle = "border:1px solid #999;padding:6px 14px;text-align:center;";
  const h = [];
  h.push('<div style="font-family:sans-serif;font-size:14px;line-height:1.7;color:#222;">');
  h.push(`<p>${esc(contactName || "ご担当者")} 様</p>`);
  h.push("<p>このたびは「年末年始廃棄物回収に関するアンケート」にご回答いただき、ありがとうございました。<br>下記の内容で承りました。ご確認ください。</p>");
  h.push(`<p>■回答日時：${esc(answeredAt)}<br>■特別回収期間：${esc(RESP_MAIL_PERIOD_TEXT)}<br>■ご担当者様：${esc(contactName)}<br>■お電話番号：${esc(phone)}</p>`);
  h.push("<p><b>【店舗ごとの回収のご希望】</b><br>○＝回収を希望した　／　×＝回収を希望しない　／　回収がありません＝回収日ではない日</p>");
  stores.forEach(store => {
    h.push(`<p style="margin:16px 0 4px;"><b>▼${esc(store.name)}</b>${store.address ? `<br><span style="font-size:13px;">ご住所：${esc(store.address)}</span>` : ""}</p>`);
    h.push('<table style="border-collapse:collapse;">');
    h.push("<tr>" + store.days.map(d => `<th style="${cellStyle}background:#eee;">${esc(d.header)}</th>`).join("") + "</tr>");
    h.push("<tr>" + store.days.map(d => `<td style="${cellStyle}${d.cell === "○" ? "font-size:18px;font-weight:bold;color:#1a7f37;" : d.cell === "×" ? "font-size:18px;color:#666;" : "font-size:12px;color:#999;"}">${esc(mark(d.cell))}</td>`).join("") + "</tr>");
    h.push("</table>");
  });
  if (comment) h.push(`<p>■ご不明点・ご要望：${esc(comment)}</p>`);
  h.push(`<p style="margin-top:20px;">${esc(notice)}<br>※本メールは、ご回答内容の控えとしてお送りしています。<br>※回答内容を変更する場合は、下記のフォームから再度ご回答ください。最新のご回答を有効とします。`);
  if (RESP_CUSTOMER_DEADLINE_TEXT) h.push(`<br>　変更のご希望は${esc(RESP_CUSTOMER_DEADLINE_TEXT)}までにお願いいたします。`);
  if (formUrl) h.push(`<br>　<a href="${esc(formUrl)}">${esc(formUrl)}</a>`);
  if (RESP_MAIL_CONTACT) h.push(`<br>※ご不明な点は、${esc(RESP_MAIL_CONTACT)}までご連絡ください。`);
  h.push(`</p><p>${esc(RESP_MAIL_SENDER_NAME)}</p></div>`);

  const subject = responseCount > 1 ? RESP_MAIL_SUBJECT.replace("の控え", "の控え（再回答）") : RESP_MAIL_SUBJECT;
  return { subject: subject, body: lines.join("\n"), htmlBody: h.join("") };
}

// 控えメールの文面を確かめるため、サンプルのメールを RESP_MAIL_TEST_TO に送る（お客様には送らない）
function sendSampleReceiptMail() {
  const testTo = String(PropertiesService.getScriptProperties().getProperty(RESP_MAIL_TEST_TO_PROPERTY) || "").trim();
  if (!testTo) {
    throw new Error(`スクリプトプロパティ ${RESP_MAIL_TEST_TO_PROPERTY} にテスト用のメールアドレスを設定してください。`);
  }
  const dateHeaders = ["12月30日(水)", "12月31日(木)", "1月1日(金)", "1月2日(土)", "1月3日(日)"];
  const rows = [
    ["サンプル店A", "2026/11/01 10:00:00", RESP_ANSWER_YES, "○", "×", "", "○", "", "山田 太郎", "03-0000-0000", "朝に来ますか？", "sample@example.com", "FORMID", 1, ""],
    ["サンプル店B", "2026/11/01 10:00:00", RESP_ANSWER_NO, "×", "×", "", "", "", "山田 太郎", "03-0000-0000", "朝に来ますか？", "sample@example.com", "FORMID", 1, ""]
  ];
  const message = respBuildReceiptMail(rows, dateHeaders, "https://docs.google.com/forms/d/e/（サンプル）/viewform",
    { "サンプル店A": "〒130-0013 東京都墨田区錦糸1-2-3 サンプルビル1F", "サンプル店B": "〒130-0022 東京都墨田区江東橋4-5-6" });
  const from = String(PropertiesService.getScriptProperties().getProperty(RESP_MAIL_FROM_PROPERTY) || "").trim();
  respSendMail({ to: testTo, subject: "【テスト】" + message.subject, body: "（サンプルです。実際の回答ではありません）\n\n" + message.body,
    htmlBody: "<p>（サンプルです。実際の回答ではありません）</p>" + message.htmlBody, name: RESP_MAIL_SENDER_NAME }, from);
  console.log(`サンプルのメールを ${testTo} に送りました。`);
}

// 変更の最終日時を Date にする。未設定・形式が違うときは null
function respParseCutoff() {
  const text = String(RESP_CUTOFF_AT).trim();
  if (!text) return null;
  const m = /^(\d{4})\/(\d{1,2})\/(\d{1,2})(?:\s+(\d{1,2}):(\d{2}))?$/.exec(text);
  if (!m) {
    console.warn(`RESP_CUTOFF_AT の形式が違います（例: 2026/11/30 23:59）: ${text}`);
    return null;
  }
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4] || 23), Number(m[5] || 59));
}

function respIsAfterCutoff() {
  const cutoff = respParseCutoff();
  return !!cutoff && new Date() >= cutoff;
}

function respClosedMessage() {
  return RESP_CLOSED_MESSAGE + (RESP_CLOSED_CONTACT ? "\n" + RESP_CLOSED_CONTACT : "");
}

// 変更の最終日時を過ぎたら、作成済みの全フォームの受付を止める（1時間ごとの定期実行。時間切れは自動で続行する）
function closeAllFormsAtCutoff() {
  const cutoff = respParseCutoff();
  if (!cutoff || new Date() < cutoff) return;

  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(RESP_CUTOFF_DONE_PROPERTY) === RESP_CUTOFF_AT) return; // この日時の分は完了済み

  const lock = LockService.getUserLock();
  if (!lock.tryLock(30 * 1000)) {
    console.log("別の処理が動いているためスキップしました。");
    return;
  }

  const startTime = Date.now();
  let closedCount = 0;
  const failedIds = [];
  let nextIndex = 0;
  let totalForms = 0;

  try {
    deleteTriggersByHandler(RESP_CUTOFF_CONTINUE_HANDLER);
    const cutoffSafetyTrigger = respScheduleTrigger(RESP_CUTOFF_CONTINUE_HANDLER, RESP_SAFETY_DELAY_MS);

    const ss = respGetSpreadsheet();
    const listSheet = ss.getSheetByName(RESP_LIST_SHEET_NAME);
    if (!listSheet) {
      throw new Error(`シート「${RESP_LIST_SHEET_NAME}」が見つかりません。`);
    }
    const forms = respGroupFormsFromList(listSheet.getDataRange().getDisplayValues());
    totalForms = forms.length;

    nextIndex = Number(props.getProperty(RESP_CUTOFF_CURSOR_PROPERTY) || 0);
    if (nextIndex >= totalForms) nextIndex = 0;

    while (nextIndex < totalForms) {
      if (Date.now() - startTime > RESP_TIME_LIMIT_MS) break;

      try {
        const target = FormApp.openById(forms[nextIndex].formId);
        if (target.isAcceptingResponses()) {
          target.setAcceptingResponses(false);
          target.setCustomClosedFormMessage(respClosedMessage());
          closedCount++;
        }
      } catch (e) {
        console.error(`フォームの受付を止められませんでした: ${forms[nextIndex].formId} / ${e}`);
        failedIds.push(forms[nextIndex].formId);
      }
      nextIndex++;
    }

    if (nextIndex < totalForms) {
      props.setProperty(RESP_CUTOFF_CURSOR_PROPERTY, String(nextIndex));
      respScheduleTrigger(RESP_CUTOFF_CONTINUE_HANDLER, RESP_RETRY_DELAY_MS);
      try {
        ScriptApp.deleteTrigger(cutoffSafetyTrigger);
      } catch (e) {
        console.warn(`保険の予約を消せませんでした（処理には影響しません）: ${e}`);
      }
      console.log(`変更の最終日時を過ぎたため受付を止めています: ${nextIndex} / ${totalForms} 件（今回 ${closedCount} 件を停止）。残りは自動で続行します。`);
    } else {
      deleteTriggersByHandler(RESP_CUTOFF_CONTINUE_HANDLER);
      props.deleteProperty(RESP_CUTOFF_CURSOR_PROPERTY);
      props.setProperty(RESP_CUTOFF_DONE_PROPERTY, RESP_CUTOFF_AT);
      console.log(`変更の最終日時を過ぎたため、全 ${totalForms} 件のフォームの受付を止めました（今回 ${closedCount} 件、失敗 ${failedIds.length} 件）。`);
      if (failedIds.length > 0) console.warn(`止められなかったフォーム: ${failedIds.join(", ")}`);
    }
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      console.warn(`ロック解除に失敗しました（処理には影響しません）: ${e}`);
    }
  }
}

// 時間切れで残ったときの続き（自動で呼ばれる）
function continueCloseAllFormsAtCutoff() {
  closeAllFormsAtCutoff();
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

// 続きの実行を予約する。Googleの一時的なサーバーエラーで失敗することがあるので、少し待って、数回やり直す
function respScheduleTrigger(handlerName, delayMs) {
  let lastError;
  for (let attempt = 0; attempt < 5; attempt++) {
    try {
      return ScriptApp.newTrigger(handlerName).timeBased().after(delayMs).create();
    } catch (e) {
      lastError = e;
      console.warn(`続きの予約に失敗しました（${attempt + 1}回目）: ${e}`);
      Utilities.sleep(2000 * Math.pow(2, attempt));   // 2, 4, 8, 16, 32秒
    }
  }
  throw lastError;
}

function deleteTriggersByHandler(handlerName) {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

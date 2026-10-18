// ===== 案内メールを送ったお客様に、フォームの不備のお詫びと、修正のお知らせを送る =====
// 「案内メール送信記録」で「送信済」になっているフォームの宛先にだけ、1通ずつ送る（まだ案内を送っていない店舗には送らない）。
// 送信元は、案内メールと同じ（Microsoft Graph、km-customer@kurimen.co.jp）。案内メール送信.js の関数を使う。
// 1通ごとに「お詫びメール送信記録」シートに記録して、同じフォームには二重に送らない。
//
// 使い方:
//   1. sendTestApology を実行 → 先頭の2件分を、INVITE_TEST_TO のアドレスだけに送る（文面を確認する）
//   2. スクリプトプロパティ APOLOGY_MODE に on を入れて、startApology を実行（終わるまで自動で続く）
//   3. 途中で止めるときは stopApology
//   4. 送り終わったら、APOLOGY_MODE を削除する

const APO_LOG_SHEET_NAME = "お詫びメール送信記録";
const APO_MODE_PROPERTY = "APOLOGY_MODE";
const APO_HANDLER = "continueApology";
const APO_TIME_LIMIT_MS = 4.5 * 60 * 1000;
const APO_RETRY_DELAY_MS = 60 * 1000;
const APO_SAFETY_DELAY_MS = 7 * 60 * 1000;
const APO_TEST_COUNT = 2;
const APO_LOG_HEADER = ["フォームID", "送信先", "対象店舗", "結果", "日時"];

// 件名は、案内メールの件名の先頭に「【お詫び】」を付ける（別のメールとして、目に留まりやすくする）
function apoSubject_() {
  return "【お詫び】" + INV_SUBJECT;
}

// 案内メールを送ったフォームの一覧（案内メール送信記録の「送信済」）
function apoTargets_() {
  const ss = invGetSpreadsheet();
  const listSheet = ss.getSheetByName(INV_LIST_SHEET_NAME);
  if (!listSheet) throw new Error(`シート「${INV_LIST_SHEET_NAME}」が見つかりません。`);
  const forms = invGroupFormsFromList(listSheet.getDataRange().getDisplayValues());

  const sent = new Set();
  const logSheet = ss.getSheetByName(INV_LOG_SHEET_NAME);
  if (logSheet && logSheet.getLastRow() > 1) {
    logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 4).getDisplayValues().forEach(row => {
      if (row[3] === INV_RESULT_SENT) sent.add(row[0]);
    });
  }
  return forms.filter(form => sent.has(form.formId));
}

function apoBuildMail_(form) {
  const lines = [];
  lines.push("【お詫びとお知らせ】", "");
  lines.push("ご担当者様", "");
  lines.push("平素より大変お世話になっております。株式会社クリメンです。");
  lines.push("先ほどお送りしました、年末年始の特別回収に関するアンケートのメールについて、お詫びとお知らせです。", "");
  lines.push("アンケートフォームの設定に不備があり、一部のお客様で、フォームを開けない状態となっておりました。");
  lines.push("現在は修正が完了しており、先ほどのメールのリンク、または下記のリンクから、ご回答いただけます。");
  lines.push("ご迷惑をおかけし、誠に申し訳ございません。", "");
  if (form.customer) lines.push("【ご契約者】", form.customer, "");
  lines.push("【対象店舗】");
  form.stores.forEach(store => lines.push(`・${store}`));
  lines.push("");
  lines.push("■アンケートフォーム");
  lines.push(form.formUrl, "");
  lines.push(`■ご回答期限：${INV_DEADLINE_TEXT}まで`, "");
  lines.push("すでにご回答いただいている場合は、改めてのご回答は不要です。");
  lines.push("フォームが開けない場合は、このメールにご返信ください。", "");
  lines.push("お手数をおかけしますが、よろしくお願いいたします。", "", INV_SENDER_NAME);
  return { subject: apoSubject_(), body: lines.join("\n") };
}

// テスト送信: 先頭の数件分を、INVITE_TEST_TO のアドレスだけに送る（お客様には送らない。記録にも残さない）
function sendTestApology() {
  const testTo = String(PropertiesService.getScriptProperties().getProperty(INV_TEST_TO_PROPERTY) || "").trim();
  if (!testTo) throw new Error(`スクリプトプロパティ ${INV_TEST_TO_PROPERTY} にテスト用のメールアドレスを設定してください。`);
  const from = invGetFrom();
  const targets = apoTargets_();
  if (targets.length === 0) throw new Error("案内メールを送ったフォームがありません。");
  targets.slice(0, APO_TEST_COUNT).forEach(form => {
    const message = apoBuildMail_(form);
    invSend(testTo, "【テスト】" + message.subject, `（テスト送信です。本来の宛先: ${form.email}）\n\n` + message.body, from);
  });
  console.log(`${Math.min(APO_TEST_COUNT, targets.length)} 通のテストメールを ${testTo} に送りました（送信元: ${from}）。対象のフォーム: 全 ${targets.length} 件。`);
}

function startApology() {
  apoDeleteTriggers_();
  runApologyBatch_();
}

function continueApology() {
  runApologyBatch_();
}

function stopApology() {
  apoDeleteTriggers_();
  console.log("お詫びメールの自動送信を停止しました。");
}

function runApologyBatch_() {
  const props = PropertiesService.getScriptProperties();
  if (String(props.getProperty(APO_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${APO_MODE_PROPERTY} が on ではないため、送信しません。`);
    return;
  }
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) { console.log("別の送信が処理中のためスキップしました。"); return; }

  const startTime = Date.now();
  let sentCount = 0;
  let failedCount = 0;
  let remaining = 0;
  try {
    apoDeleteTriggers_();
    const safety = ScriptApp.newTrigger(APO_HANDLER).timeBased().after(APO_SAFETY_DELAY_MS).create();

    const from = invGetFrom();
    const ss = invGetSpreadsheet();
    let logSheet = ss.getSheetByName(APO_LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = ss.insertSheet(APO_LOG_SHEET_NAME);
      logSheet.appendRow(APO_LOG_HEADER);
    }
    const logged = new Set();
    if (logSheet.getLastRow() > 1) {
      logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getDisplayValues().forEach(row => logged.add(row[0]));
    }
    const targets = apoTargets_().filter(form => !logged.has(form.formId));

    for (let n = 0; n < targets.length; n++) {
      if (Date.now() - startTime > APO_TIME_LIMIT_MS) { remaining = targets.length - n; break; }
      const form = targets[n];
      let result = INV_RESULT_SENT;
      if (!/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(form.email)) {
        console.error(`宛先が正しくありません: ${form.formId}`);
        result = INV_RESULT_FAILED;
      } else {
        try {
          const message = apoBuildMail_(form);
          invSend(form.email, message.subject, message.body, from);
        } catch (e) {
          console.error(`お詫びメールを送れませんでした: ${form.formId} / ${e}`);
          result = INV_RESULT_FAILED;
        }
      }
      logSheet.appendRow([form.formId, form.email, form.stores.join("・"), result, invFormatNow()]);
      SpreadsheetApp.flush();
      if (result === INV_RESULT_SENT) { sentCount++; Utilities.sleep(INV_GRAPH_INTERVAL_MS); } else { failedCount++; }
    }

    if (remaining > 0) {
      ScriptApp.newTrigger(APO_HANDLER).timeBased().after(APO_RETRY_DELAY_MS).create();
      try { ScriptApp.deleteTrigger(safety); } catch (e) { /* 影響なし */ }
      console.log(`今回 ${sentCount} 通送信（失敗 ${failedCount} 件）。残り ${remaining} 件は自動で続行します。`);
    } else {
      apoDeleteTriggers_();
      console.log(`今回 ${sentCount} 通送信（失敗 ${failedCount} 件）。すべて完了しました。記録は「${APO_LOG_SHEET_NAME}」シートにあります。`);
    }
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

function apoDeleteTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === APO_HANDLER) ScriptApp.deleteTrigger(t); });
}

// ===== 控えメールが重複して届いたお客様に、お詫びを送る =====
// 回答の日時が午前0〜9時台のお客様は、控えメールが、2時間ごとに、繰り返し届いてしまった（時が1桁で表示される不具合）。
// 「回答集計」の「控えメール」の列が、時が1桁の日時（例: 9:41:03）の行を、対象にする。同じフォームは、1通にまとめる。
// 1通ごとに「重複お詫びメール送信記録」シートに記録して、同じフォームには二重に送らない。
//
// 使い方:
//   1. sendTestDuplicateApology を実行 → 先頭の2件分を、INVITE_TEST_TO のアドレスだけに送る（文面と、対象の件数を確認する）
//   2. スクリプトプロパティ DUP_APOLOGY_MODE に on を入れて、startDuplicateApology を実行
//   3. 送り終わったら、DUP_APOLOGY_MODE を削除する

const DUP_LOG_SHEET_NAME = "重複お詫びメール送信記録";
const DUP_MODE_PROPERTY = "DUP_APOLOGY_MODE";
const DUP_SUBJECT = "【お詫び】年末年始廃棄物回収アンケート ご回答内容の控えメールが重複して届いたことについて";

// 控えメールが重複して届いた可能性のあるフォーム（控えメールの列が、時が1桁の日時のもの）
function dupTargets_() {
  const ss = invGetSpreadsheet();
  const sheet = ss.getSheetByName("回答集計");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getDisplayValues();
  const header = values[0];
  const colMail = header.indexOf("控えメール");
  const colEmail = header.indexOf("メールアドレス");
  const colForm = header.indexOf("フォームID");
  if (colMail < 0 || colEmail < 0 || colForm < 0) throw new Error("「回答集計」シートの見出しが、想定と違います。");

  const byForm = {};
  const targets = [];
  for (let i = 1; i < values.length; i++) {
    const mark = String(values[i][colMail] || "").trim();
    if (!/^\d{4}\/\d{1,2}\/\d{1,2}\s+\d:\d{2}:\d{2}$/.test(mark)) continue;   // 時が1桁の、お客様に送信済みの印
    const formId = values[i][colForm];
    if (!formId) continue;
    if (!byForm[formId]) {
      byForm[formId] = { formId: formId, email: String(values[i][colEmail] || "").trim(), stores: [] };
      targets.push(byForm[formId]);
    }
    byForm[formId].stores.push(values[i][0]);
  }
  return targets;
}

function dupBuildMail_(target) {
  const lines = [];
  lines.push("ご担当者様", "");
  lines.push("平素より大変お世話になっております。株式会社クリメンです。");
  lines.push("このたびは、年末年始の廃棄物回収に関するアンケートにご回答いただき、ありがとうございました。", "");
  lines.push("システムの不具合により、ご回答内容の控えメールが、同じ内容で、繰り返しお手元に届いておりました。");
  lines.push("ご迷惑をおかけし、誠に申し訳ございません。", "");
  lines.push("不具合は修正済みで、今後、同じ控えメールが繰り返し届くことはございません。");
  lines.push("ご回答の内容に、変更はございません。繰り返し届いたメールは、お手数ですが、破棄していただけますと幸いです。", "");
  lines.push("【対象店舗】");
  target.stores.forEach(store => lines.push(`・${store}`));
  lines.push("");
  lines.push(`ご回答内容を変更される場合は、先にお送りしたメールのリンクから、${INV_DEADLINE_TEXT}までに、もう一度ご回答ください。`, "");
  lines.push("今後とも、よろしくお願いいたします。", "", INV_SENDER_NAME);
  return { subject: DUP_SUBJECT, body: lines.join("\n") };
}

function sendTestDuplicateApology() {
  const testTo = String(PropertiesService.getScriptProperties().getProperty(INV_TEST_TO_PROPERTY) || "").trim();
  if (!testTo) throw new Error(`スクリプトプロパティ ${INV_TEST_TO_PROPERTY} にテスト用のメールアドレスを設定してください。`);
  const from = invGetFrom();
  const targets = dupTargets_();
  console.log(`対象のフォーム: 全 ${targets.length} 件（店舗は、合計 ${targets.reduce((n, t) => n + t.stores.length, 0)} 件）`);
  if (targets.length === 0) throw new Error("対象のフォームがありません。");
  targets.slice(0, 2).forEach(target => {
    const message = dupBuildMail_(target);
    invSend(testTo, "【テスト】" + message.subject, `（テスト送信です。本来の宛先: ${target.email}）\n\n` + message.body, from);
  });
  console.log(`${Math.min(2, targets.length)} 通のテストメールを ${testTo} に送りました（送信元: ${from}）。`);
}

function startDuplicateApology() {
  const props = PropertiesService.getScriptProperties();
  if (String(props.getProperty(DUP_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${DUP_MODE_PROPERTY} が on ではないため、送信しません。`);
    return;
  }
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) { console.log("別の処理が動いているためスキップしました。少し待って、もう一度実行してください。"); return; }
  const startTime = Date.now();
  let sentCount = 0;
  let failedCount = 0;
  let remaining = 0;
  try {
    const from = invGetFrom();
    const ss = invGetSpreadsheet();
    let logSheet = ss.getSheetByName(DUP_LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = ss.insertSheet(DUP_LOG_SHEET_NAME);
      logSheet.appendRow(APO_LOG_HEADER);
    }
    const logged = new Set();
    if (logSheet.getLastRow() > 1) {
      logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getDisplayValues().forEach(row => logged.add(row[0]));
    }
    const targets = dupTargets_().filter(target => !logged.has(target.formId));
    for (let n = 0; n < targets.length; n++) {
      if (Date.now() - startTime > APO_TIME_LIMIT_MS) { remaining = targets.length - n; break; }
      const target = targets[n];
      let result = INV_RESULT_SENT;
      if (!/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(target.email)) {
        console.error(`宛先が正しくありません: ${target.formId}`);
        result = INV_RESULT_FAILED;
      } else {
        try {
          const message = dupBuildMail_(target);
          invSend(target.email, message.subject, message.body, from);
        } catch (e) {
          console.error(`お詫びメールを送れませんでした: ${target.formId} / ${e}`);
          result = INV_RESULT_FAILED;
        }
      }
      logSheet.appendRow([target.formId, target.email, target.stores.join("・"), result, invFormatNow()]);
      SpreadsheetApp.flush();
      if (result === INV_RESULT_SENT) { sentCount++; Utilities.sleep(INV_GRAPH_INTERVAL_MS); } else { failedCount++; }
    }
    console.log(`今回 ${sentCount} 通送信（失敗 ${failedCount} 件）。` + (remaining > 0 ? `残り ${remaining} 件は、もう一度 startDuplicateApology を実行してください。` : "すべて完了しました。") + `記録は「${DUP_LOG_SHEET_NAME}」シートにあります。`);
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

// ===== 控えメールが重複して届いたお客様の一覧を、シートに書き出す（お客様には、何も送らない） =====
// 「回答集計」の「控えメール」の列が、時が1桁の日時（例: 9:41:03）の行を、「重複対象一覧」シートに書き出す。
// 「概算の重複回数」は、回答の日時から、控えメールを止めた日時（DUP_STOPPED_AT）までの間に、2時間ごとの集計が、何回動いたかの目安。
const DUP_LIST_SHEET_NAME = "重複対象一覧";
const DUP_STOPPED_AT = "2026/10/18 10:30";   // RESP_MAIL_MODE を off にした、おおよその日時

function listDuplicateReceiptTargets() {
  const ss = invGetSpreadsheet();
  const sheet = ss.getSheetByName("回答集計");
  if (!sheet || sheet.getLastRow() < 2) throw new Error("「回答集計」シートに、データがありません。");
  const values = sheet.getDataRange().getDisplayValues();
  const header = values[0];
  const col = name => header.indexOf(name);
  const colMail = col("控えメール"), colAt = col("回答日時"), colName = col("ご担当者名"), colEmail = col("メールアドレス"), colForm = col("フォームID");
  if ([colMail, colAt, colName, colEmail, colForm].some(c => c < 0)) throw new Error("「回答集計」シートの見出しが、想定と違います。");

  const stoppedAt = new Date(DUP_STOPPED_AT.replace(/-/g, "/")).getTime();
  const parse = text => { const m = /^(\d{4})\/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2}):(\d{2})$/.exec(String(text).trim()); return m ? new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]).getTime() : null; };

  const byForm = {};
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const mark = String(values[i][colMail] || "").trim();
    if (!/^\d{4}\/\d{1,2}\/\d{1,2}\s+\d:\d{2}:\d{2}$/.test(mark)) continue;
    const formId = values[i][colForm];
    if (byForm[formId]) { byForm[formId][0] += "・" + values[i][0]; continue; }
    const t = parse(values[i][colAt]);
    const estimate = t ? Math.max(0, Math.floor((stoppedAt - t) / (2 * 60 * 60 * 1000))) : "";
    byForm[formId] = [values[i][0], values[i][colAt], values[i][colName], values[i][colEmail], estimate, formId];
    rows.push(byForm[formId]);
  }
  let out = ss.getSheetByName(DUP_LIST_SHEET_NAME);
  if (!out) out = ss.insertSheet(DUP_LIST_SHEET_NAME);
  out.clear();
  const head = ["対象店舗", "回答日時", "ご担当者名", "メールアドレス", "概算の重複回数", "フォームID"];
  out.getRange(1, 1, 1, head.length).setValues([head]);
  if (rows.length > 0) out.getRange(2, 1, rows.length, head.length).setValues(rows);
  console.log(`重複して届いた可能性のあるお客様: ${rows.length} 件。シート「${DUP_LIST_SHEET_NAME}」に書き出しました（お客様には、何も送っていません）。`);
}

// ===== 回答のないお客様への、催促メール =====
// 対象: 10/15・10/16 に案内メールを送った（「案内メール送信記録」）フォームのうち、いま回答がないもの。
// 除外: 最近（REM_RECENT_FROM 以降）、別のメールを送ったフォーム／受付が終了したフォーム／テスト店舗／「除外」の店舗
// 手順:
//   1. previewReminders   … 対象を数え、シート「催促メール対象一覧」に書き出す（メールは送らない）。シートがすでにあるときは、作り直さない
//   2. sendTestReminders  … メールを、INVITE_TEST_TO のアドレスだけに送る（お客様には送らない）
//   3. startReminders     … お客様に、メールを送る（スクリプトプロパティ REMIND_MAIL_MODE が on のときだけ動く）。
//                            時間がかかるので、数分おきに、自動で続きを送る。止めるときは stopReminders
//   4. 終わったら、REMIND_MAIL_MODE を削除する

const REM_MODE_PROPERTY = "REMIND_MAIL_MODE";
const REM_TARGET_SHEET = "催促メール対象一覧";
const REM_LOG_SHEET = "催促メール送信記録";
const REM_INVITE_DATES = ["2026/10/15", "2026/10/16"];   // 案内メールを送った日
const REM_RECENT_FROM = "2026/10/29";                    // この日以降に、別のメールを送ったフォームは、対象から外す
const REM_OTHER_LOG_SHEETS = ["新規契約メール送信記録", "追加日程メール送信記録", "保留3社メール送信記録", "ルアナ渋谷メール送信記録", "個別対応2グループ メール送信記録", "期間外お知らせ送信記録", "お詫びメール送信記録", "重複お詫びメール送信記録"];
const REM_HANDLER = "continueReminders";
const REM_SUBJECT = "【株式会社クリメン】年末年始の特別回収アンケートのご回答のお願い";
const REM_TIME_LIMIT_MS = 4.5 * 60 * 1000;

// 「リスト」から、フォームごとの送信先と店舗を取り出す（いまの内容）
function remForms_() {
  const data = invGetSpreadsheet().getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  const forms = {};
  const order = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[INV_COL_FORM_ID] || r[INV_COL_STATUS] !== INV_STATUS_SENT) continue;
    if (String(r[INV_COL_STORE]).indexOf("テスト店舗") === 0) continue;
    const id = r[INV_COL_FORM_ID];
    if (!forms[id]) { forms[id] = { formId: id, formUrl: r[INV_COL_FORM_URL], email: String(r[INV_COL_EMAIL]).trim(), customer: String(r[INV_COL_CUSTOMER] || "").trim(), stores: [] }; order.push(id); }
    forms[id].stores.push(r[INV_COL_STORE]);
  }
  return order.map(id => forms[id]);
}

// 最近、別のメールを送ったフォームのID
function remRecentlyMailed_() {
  const ss = invGetSpreadsheet();
  const ids = new Set();
  REM_OTHER_LOG_SHEETS.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    sheet.getDataRange().getDisplayValues().slice(1).forEach(row => {
      const date = row.find(c => /^\d{4}\/\d{1,2}\/\d{1,2}/.test(String(c)));
      if (!date) return;
      const d = String(date).split(" ")[0].split("/").map((x, k) => k === 0 ? x : ("0" + x).slice(-2)).join("/");
      if (d < REM_RECENT_FROM) return;
      row.forEach(c => { String(c).split("|").forEach(token => { if (token.length > 20) ids.add(token); }); });
    });
  });
  return ids;
}

function remInviteMap_() {
  const sheet = invGetSpreadsheet().getSheetByName(INV_LOG_SHEET_NAME);
  const map = {};
  sheet.getDataRange().getDisplayValues().slice(1).forEach(r => { if (r[3] === INV_RESULT_SENT) map[r[0]] = r[4]; });
  return map;
}

function remTargets_(liveCheck) {
  const ss = invGetSpreadsheet();
  const answered = new Set();
  const out = ss.getSheetByName(RESP_OUTPUT_SHEET_NAME).getDataRange().getDisplayValues();
  const idCol = out[0].indexOf("フォームID");
  out.slice(1).forEach(r => { if (r[idCol]) answered.add(r[idCol]); });
  const invite = remInviteMap_();
  const recent = remRecentlyMailed_();
  const counts = { 全フォーム: 0, 回答あり: 0, 案内が10月15_16日ではない: 0, 最近別のメール: 0, 対象: 0 };
  const targets = [];
  remForms_().forEach(f => {
    counts.全フォーム++;
    if (answered.has(f.formId)) { counts.回答あり++; return; }
    const sentAt = invite[f.formId] || "";
    if (!REM_INVITE_DATES.some(d => sentAt.indexOf(d) === 0)) { counts.案内が10月15_16日ではない++; return; }
    if (recent.has(f.formId)) { counts.最近別のメール++; return; }
    f.invitedAt = sentAt;
    targets.push(f);
  });
  counts.対象 = targets.length;
  return { targets: targets, counts: counts };
}

// 1. 対象の確認（メールは送らない）
function previewReminders() {
  const ss = invGetSpreadsheet();
  const r = remTargets_();
  console.log(`対象の内訳: ${JSON.stringify(r.counts)}`);
  const bad = r.targets.filter(f => !/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(f.email));
  console.log(`宛先の形が正しくないもの: ${bad.length} 件`);
  if (ss.getSheetByName(REM_TARGET_SHEET)) { console.log(`シート「${REM_TARGET_SHEET}」は、すでにあるため、作り直しません（作り直すときは、シートを削除してください）。`); return; }
  const sheet = ss.insertSheet(REM_TARGET_SHEET);
  const head = ["フォームID", "送信先", "契約者", "対象店舗", "案内を送った日時"];
  sheet.getRange(1, 1, 1, head.length).setValues([head]);
  if (r.targets.length) sheet.getRange(2, 1, r.targets.length, head.length).setValues(r.targets.map(f => [f.formId, f.email, f.customer, f.stores.join("・"), f.invitedAt]));
  console.log(`シート「${REM_TARGET_SHEET}」に ${r.targets.length} 件を書き出しました。`);
}

// 催促メールの件名と本文
function remBuildMail_(f) {
  const lines = [];
  lines.push("ご担当者様", "");
  lines.push("平素より大変お世話になっております。株式会社クリメンです。");
  lines.push("先日お送りした、年末年始の特別回収のアンケートについて、まだご回答をいただいておりません。", "");
  if (f.customer) lines.push("【ご契約者】", f.customer, "");
  lines.push("【対象店舗】");
  f.stores.forEach(s => lines.push(`・${s}`));
  lines.push("");
  lines.push(`■ご回答期限：${INV_DEADLINE_TEXT}まで`);
  lines.push("　※ご回答がない場合、年末年始の特別回収を行えない可能性があります。", "");
  lines.push("お手数ですが、下記のアンケートフォームから、ご回答をお願いいたします。");
  lines.push(f.formUrl, "");
  lines.push("【ご回答について】");
  lines.push("・特別回収を希望されない場合も、すべての日について「希望しない」を選択してご回答ください。");
  lines.push("・回答内容を変更する場合は、同じフォームから再度ご回答ください（最新のご回答を有効とします）。", "");
  lines.push("※すでにご回答いただいている場合は、本メールと行き違いです。何卒ご容赦ください。", "");
  if (INV_CONTACT) lines.push(`ご不明な点は、${INV_CONTACT}までご連絡ください。`, "");
  if (INV_NOREPLY_TEXT) lines.push(INV_NOREPLY_TEXT, "");
  lines.push("お忙しいところ恐れ入りますが、よろしくお願いいたします。", "", INV_SENDER_NAME);
  return { subject: REM_SUBJECT, body: lines.join("\n") };
}

// 2. テスト送信（お客様には送らない）
function sendTestReminders() {
  const testTo = String(PropertiesService.getScriptProperties().getProperty(INV_TEST_TO_PROPERTY) || "").trim();
  if (!testTo) throw new Error(`スクリプトプロパティ ${INV_TEST_TO_PROPERTY} にテスト用のメールアドレスを設定してください。`);
  const from = invGetFrom();
  const targets = remTargets_().targets.slice(0, INV_TEST_COUNT);
  targets.forEach(f => {
    const m = remBuildMail_(f);
    invSend(testTo, "【テスト】" + m.subject, `（テスト送信です。本来の宛先: ${f.email}）\n\n` + m.body, from);
    Utilities.sleep(INV_GRAPH_INTERVAL_MS);
  });
  console.log(`テストメール ${targets.length} 通を ${testTo} に送りました。`);
}

function remDeleteTriggers_() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === REM_HANDLER).forEach(t => ScriptApp.deleteTrigger(t));
}

function stopReminders() {
  remDeleteTriggers_();
  console.log("催促メールの自動送信を停止しました。");
}

function continueReminders() { startReminders(); }

// 3. お客様に、メールを送る
function startReminders() {
  if (String(PropertiesService.getScriptProperties().getProperty(REM_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${REM_MODE_PROPERTY} が on ではないため、送信しません。`);
    remDeleteTriggers_();
    return;
  }
  const lock = LockService.getUserLock();
  if (!lock.tryLock(30 * 1000)) {
    remDeleteTriggers_();
    ScriptApp.newTrigger(REM_HANDLER).timeBased().after(2 * 60 * 1000).create();
    console.log("別の処理が動いているため、2分後に、もう一度試します。");
    return;
  }
  const startTime = Date.now();
  let sent = 0, failed = 0, skipped = 0, remaining = 0;
  try {
    remDeleteTriggers_();
    const ss = invGetSpreadsheet();
    let log = ss.getSheetByName(REM_LOG_SHEET);
    if (!log) { log = ss.insertSheet(REM_LOG_SHEET); log.appendRow(["フォームID", "送信先", "対象店舗", "結果", "日時"]); }
    const logged = new Set();
    if (log.getLastRow() > 1) log.getRange(2, 1, log.getLastRow() - 1, 1).getDisplayValues().forEach(r => logged.add(r[0]));
    const from = invGetFrom();
    const todo = remTargets_().targets.filter(f => !logged.has(f.formId));
    for (let n = 0; n < todo.length; n++) {
      if (Date.now() - startTime > REM_TIME_LIMIT_MS) { remaining = todo.length - n; break; }
      const f = todo[n];
      let result = INV_RESULT_SENT;
      try {
        const form = FormApp.openById(f.formId);
        if (form.getResponses().length > 0) { result = "回答あり（送らない）"; }
        else if (!form.isAcceptingResponses()) { result = "受付終了（送らない）"; }
        else if (!/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(f.email)) { result = INV_RESULT_FAILED; console.error(`宛先が正しくありません: 末尾 ${f.formId.slice(-8)}`); }
        else { const m = remBuildMail_(f); invSend(f.email, m.subject, m.body, from); }
      } catch (e) {
        console.error(`メールを送れませんでした: 末尾 ${f.formId.slice(-8)} / ${e}`);
        result = INV_RESULT_FAILED;
      }
      log.appendRow([f.formId, f.email, f.stores.join("・"), result, invFormatNow()]);
      SpreadsheetApp.flush();
      if (result === INV_RESULT_SENT) { sent++; Utilities.sleep(INV_GRAPH_INTERVAL_MS); }
      else if (result === INV_RESULT_FAILED) failed++;
      else skipped++;
    }
    if (remaining > 0) {
      ScriptApp.newTrigger(REM_HANDLER).timeBased().after(60 * 1000).create();
      console.log(`今回 ${sent} 通送信（失敗 ${failed} 件、送らなかった ${skipped} 件）。残り ${remaining} 件は、1分後に自動で続けます。`);
    } else {
      console.log(`今回 ${sent} 通送信（失敗 ${failed} 件、送らなかった ${skipped} 件）。すべて完了しました。記録は「${REM_LOG_SHEET}」シートにあります。`);
    }
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

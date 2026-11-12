// ===== 回答のないお客様への、2回目・3回目の催促メール（11/12「あと3日」・11/15「本日が期限」）=====
// 対象: そのときに、回答のない全フォーム（宛先が正しくないものは、自動で飛ばす）。送る直前に、回答と受付状況を、もう一度確認する。
// 11/12用と11/15用は、関数・記録シート・安全装置（スクリプトプロパティ）が別。11/15に、11/12の記録に引っかからずに、もう一度送れる。
//
// 【11/12用】 previewFinalR12 → sendTestFinalR12 → startFinalR12（REMIND12_MAIL_MODE が on のとき）／ 止めるとき stopFinalR12
// 【11/15用 9:00】 previewFinalR15 → sendTestFinalR15 → startFinalR15（REMIND15_MAIL_MODE が on のとき）／ 止めるとき stopFinalR15
// 【11/15用 17:00】 previewFinalR15b → sendTestFinalR15b → startFinalR15b（REMIND15_MAIL_MODE が on のとき）／ 止めるとき stopFinalR15b
// 【11/15の自動】 setupAutoFinalR15 で、9:00 と 17:00 の予約を作る（cancelAutoFinalR15 で取り消し）
// 終わったら、それぞれの MODE を削除する

const REMF_ROUNDS = {
  r12: {
    key: "r12", mode: "REMIND12_MAIL_MODE", handler: "continueFinalR12",
    log: "催促メール1112 送信記録", target: "催促メール1112 対象一覧",
    subject: "【株式会社クリメン】【回答期限まであと3日】年末年始の特別回収アンケートのご回答のお願い",
    headline: "【回答期限まで、あと3日です】",
    intro: "年末年始の特別回収のアンケートについて、回答期限まで、あと3日となりました。\n現時点で、ご回答をいただけておりません。",
    deadlineLine: () => `★ご回答期限：${INV_DEADLINE_TEXT}まで（あと3日です）`,
    pushLine: "お早めのご回答をお願いいたします。"
  },
  r15: {
    key: "r15", mode: "REMIND15_MAIL_MODE", handler: "continueFinalR15",
    log: "催促メール1115 送信記録", target: "催促メール1115 対象一覧",
    subject: "【株式会社クリメン】【本日が回答期限です】年末年始の特別回収アンケートのご回答のお願い",
    headline: "【本日が、回答期限です】",
    intro: "年末年始の特別回収のアンケートについて、本日が回答期限です。\n現時点で、ご回答をいただけておりません。",
    deadlineLine: () => `★ご回答期限：本日 ${INV_DEADLINE_TEXT}まで`,
    pushLine: "本日中のご回答を、お願いいたします。"
  },
  r15b: {
    key: "r15b", mode: "REMIND15_MAIL_MODE", handler: "continueFinalR15b",
    log: "催促メール1115b 送信記録", target: "催促メール1115b 対象一覧",
    subject: "【株式会社クリメン】【本日が回答期限です・本日中にご回答をお願いします】年末年始の特別回収アンケート",
    headline: "【本日が、回答期限です。本日中にご回答をお願いします】",
    intro: "年末年始の特別回収のアンケートについて、本日が回答期限です。\n朝にもご案内しましたが、現時点で、ご回答をいただけておりません。",
    deadlineLine: () => `★ご回答期限：本日 ${INV_DEADLINE_TEXT}まで`,
    pushLine: "本日中のご回答を、お願いいたします。"
  }
};
const REMF_TIME_LIMIT_MS = 4.5 * 60 * 1000;

function remfBody_(round, f) {
  const lines = [];
  lines.push("ご担当者様", "");
  lines.push("平素より大変お世話になっております。株式会社クリメンです。", "");
  lines.push(round.headline, "");
  lines.push(round.intro, "");
  if (f.customer) lines.push("【ご契約者】", f.customer, "");
  lines.push("【対象店舗】");
  f.stores.forEach(s => lines.push(`・${s}`));
  lines.push("");
  lines.push(round.deadlineLine());
  lines.push("★ご回答がない場合、年末年始の特別回収を行えない可能性があります。", "");
  lines.push(round.pushLine);
  lines.push("下記のアンケートフォームから、ご回答をお願いいたします。");
  lines.push(f.formUrl, "");
  lines.push("【ご回答について】");
  lines.push("・特別回収を希望されない場合も、すべての日について「希望しない」を選択してご回答ください。");
  lines.push("・回答内容を変更する場合は、同じフォームから再度ご回答ください（最新のご回答を有効とします）。", "");
  lines.push("※すでにご回答いただいている場合は、本メールと行き違いです。何卒ご容赦ください。", "");
  if (INV_CONTACT) lines.push(`ご不明な点は、${INV_CONTACT}までご連絡ください。`, "");
  if (INV_NOREPLY_TEXT) lines.push(INV_NOREPLY_TEXT, "");
  lines.push("お忙しいところ恐れ入りますが、よろしくお願いいたします。", "", INV_SENDER_NAME);
  return { subject: round.subject, body: lines.join("\n") };
}

// いま回答のないフォーム（「リスト」のいまの内容と「回答集計」から）
function remfTargets_() {
  const ss = invGetSpreadsheet();
  const out = ss.getSheetByName(RESP_OUTPUT_SHEET_NAME).getDataRange().getDisplayValues();
  const idCol = out[0].indexOf("フォームID");
  const answered = new Set();
  out.slice(1).forEach(r => { if (r[idCol]) answered.add(r[idCol]); });
  const forms = remForms_();
  const targets = forms.filter(f => !answered.has(f.formId));
  return { all: forms.length, answered: forms.length - targets.length, targets: targets };
}

function remfPreview_(round) {
  const ss = invGetSpreadsheet();
  const r = remfTargets_();
  const bad = r.targets.filter(f => !/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(f.email));
  console.log(`全フォーム ${r.all} 件、回答あり ${r.answered} 件、回答なし（送る対象）${r.targets.length} 件。宛先の形が正しくないもの: ${bad.length} 件。`);
  if (ss.getSheetByName(round.target)) { console.log(`シート「${round.target}」は、すでにあるため、作り直しません。`); return; }
  const sheet = ss.insertSheet(round.target);
  const head = ["フォームID", "送信先", "契約者", "対象店舗"];
  sheet.getRange(1, 1, 1, head.length).setValues([head]);
  if (r.targets.length) sheet.getRange(2, 1, r.targets.length, head.length).setValues(r.targets.map(f => [f.formId, f.email, f.customer, f.stores.join("・")]));
  console.log(`シート「${round.target}」に ${r.targets.length} 件を書き出しました。`);
}

function remfTest_(round) {
  const testTo = String(PropertiesService.getScriptProperties().getProperty(INV_TEST_TO_PROPERTY) || "").trim();
  if (!testTo) throw new Error(`スクリプトプロパティ ${INV_TEST_TO_PROPERTY} にテスト用のメールアドレスを設定してください。`);
  const from = invGetFrom();
  const targets = remfTargets_().targets.slice(0, INV_TEST_COUNT);
  targets.forEach(f => {
    const m = remfBody_(round, f);
    invSend(testTo, "【テスト】" + m.subject, `（テスト送信です。本来の宛先: ${f.email}）\n\n` + m.body, from);
    Utilities.sleep(INV_GRAPH_INTERVAL_MS);
  });
  console.log(`テストメール ${targets.length} 通を ${testTo} に送りました。`);
}

function remfDeleteTriggers_(round) {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === round.handler).forEach(t => ScriptApp.deleteTrigger(t));
}

function remfStart_(round) {
  if (String(PropertiesService.getScriptProperties().getProperty(round.mode) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${round.mode} が on ではないため、送信しません。`);
    remfDeleteTriggers_(round);
    return;
  }
  const lock = LockService.getUserLock();
  if (!lock.tryLock(30 * 1000)) {
    remfDeleteTriggers_(round);
    ScriptApp.newTrigger(round.handler).timeBased().after(2 * 60 * 1000).create();
    console.log("別の処理が動いているため、2分後に、もう一度試します。");
    return;
  }
  const startTime = Date.now();
  let sent = 0, failed = 0, skipped = 0, remaining = 0;
  try {
    remfDeleteTriggers_(round);
    const ss = invGetSpreadsheet();
    let log = ss.getSheetByName(round.log);
    if (!log) { log = ss.insertSheet(round.log); log.appendRow(["フォームID", "送信先", "対象店舗", "結果", "日時"]); }
    const logged = new Set();
    if (log.getLastRow() > 1) log.getRange(2, 1, log.getLastRow() - 1, 1).getDisplayValues().forEach(r => logged.add(r[0]));
    const from = invGetFrom();
    const todo = remfTargets_().targets.filter(f => !logged.has(f.formId));
    for (let n = 0; n < todo.length; n++) {
      if (Date.now() - startTime > REMF_TIME_LIMIT_MS) { remaining = todo.length - n; break; }
      const f = todo[n];
      let result = INV_RESULT_SENT;
      try {
        const form = FormApp.openById(f.formId);
        if (form.getResponses().length > 0) { result = "回答あり（送らない）"; }
        else if (!form.isAcceptingResponses()) { result = "受付終了（送らない）"; }
        else if (!/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(f.email)) { result = INV_RESULT_FAILED; console.error(`宛先が正しくありません: 末尾 ${f.formId.slice(-8)}`); }
        else { const m = remfBody_(round, f); invSend(f.email, m.subject, m.body, from); }
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
      ScriptApp.newTrigger(round.handler).timeBased().after(60 * 1000).create();
      console.log(`今回 ${sent} 通送信（失敗 ${failed} 件、送らなかった ${skipped} 件）。残り ${remaining} 件は、1分後に自動で続けます。`);
    } else {
      console.log(`今回 ${sent} 通送信（失敗 ${failed} 件、送らなかった ${skipped} 件）。すべて完了しました。記録は「${round.log}」シートにあります。`);
    }
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

// ---- 11/12用 ----
function previewFinalR12() { remfPreview_(REMF_ROUNDS.r12); }
function sendTestFinalR12() { remfTest_(REMF_ROUNDS.r12); }
function startFinalR12() { remfStart_(REMF_ROUNDS.r12); }
function continueFinalR12() { remfStart_(REMF_ROUNDS.r12); }
function stopFinalR12() { remfDeleteTriggers_(REMF_ROUNDS.r12); console.log("11/12用の催促メールの自動送信を停止しました。"); }

// ---- 11/15用 ----
function previewFinalR15() { remfPreview_(REMF_ROUNDS.r15); }
function sendTestFinalR15() { remfTest_(REMF_ROUNDS.r15); }
function startFinalR15() { remfStart_(REMF_ROUNDS.r15); }
function continueFinalR15() { remfStart_(REMF_ROUNDS.r15); }
function stopFinalR15() { remfDeleteTriggers_(REMF_ROUNDS.r15); console.log("11/15用の催促メールの自動送信を停止しました。"); }

// ---- 11/15用（17:00の2回目） ----
function previewFinalR15b() { remfPreview_(REMF_ROUNDS.r15b); }
function sendTestFinalR15b() { remfTest_(REMF_ROUNDS.r15b); }
function startFinalR15b() { remfStart_(REMF_ROUNDS.r15b); }
function continueFinalR15b() { remfStart_(REMF_ROUNDS.r15b); }
function stopFinalR15b() { remfDeleteTriggers_(REMF_ROUNDS.r15b); console.log("11/15 17:00用の催促メールの自動送信を停止しました。"); }

// ---- 11/15用: 自動で送り始める予約（9:00 と 17:00）----
// setupAutoFinalR15 を実行すると、9:00 と 17:00 に、それぞれ自動で動く予約ができる。
// 動く関数は、今日が 11/15 でなければ、何もしない。REMIND15_MAIL_MODE が on でなければ、送らない。
// 取り消すときは cancelAutoFinalR15
const REMF_AUTO_R15_DATE = "2026/11/15";
const REMF_AUTO_R15_SLOTS = [
  { at: "2026/11/15 09:00", handler: "autoStartFinalR15" },
  { at: "2026/11/15 17:00", handler: "autoStartFinalR15b" }
];

function setupAutoFinalR15() {
  cancelAutoFinalR15();
  REMF_AUTO_R15_SLOTS.forEach(slot => {
    const when = Utilities.parseDate(slot.at, "Asia/Tokyo", "yyyy/MM/dd HH:mm");
    if (when.getTime() <= Date.now()) throw new Error(`${slot.at} は、すでに過ぎています。`);
    ScriptApp.newTrigger(slot.handler).timeBased().at(when).create();
    console.log(`${slot.at}（日本時間）に、自動で送り始める予約を作りました。`);
  });
  console.log("送るには、REMIND15_MAIL_MODE が on である必要があります。");
}

function cancelAutoFinalR15() {
  const handlers = REMF_AUTO_R15_SLOTS.map(s => s.handler);
  ScriptApp.getProjectTriggers().filter(t => handlers.indexOf(t.getHandlerFunction()) !== -1).forEach(t => ScriptApp.deleteTrigger(t));
  console.log("11/15用の自動送信の予約を、取り消しました（なければ、何もしていません）。");
}

function remfAutoGuard_() {
  const today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
  if (today !== REMF_AUTO_R15_DATE) { console.log(`今日は ${today} で、${REMF_AUTO_R15_DATE} ではないため、送りません。`); return false; }
  return true;
}
function autoStartFinalR15() { if (remfAutoGuard_()) remfStart_(REMF_ROUNDS.r15); }
function autoStartFinalR15b() { if (remfAutoGuard_()) remfStart_(REMF_ROUNDS.r15b); }

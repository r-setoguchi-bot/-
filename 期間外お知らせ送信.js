// ===== 特別回収の期間中に回収日がない店舗に、お知らせメールを送る =====
// 「契約状況一覧」で、理由が「期間中に回収できる日がありません」の店舗に、アドレスと契約者の組ごとに、1通ずつ送る。
// 送信元は、案内メールと同じ（Microsoft Graph、km-customer@kurimen.co.jp）。案内メール送信.js の関数を使う。
// 1通ごとに「期間外お知らせ送信記録」シートに記録して、同じ組には二重に送らない。
//
// 使い方:
//   1. listNoticeTargets を実行 → 対象を「期間外お知らせ対象一覧」シートに書き出して、固定する（お客様には送らない）。内容を目で確認する
//   2. sendTestNotice を実行 → 先頭の2件分を、INVITE_TEST_TO のアドレスだけに送る（文面を確認する）
//   3. スクリプトプロパティ NOTICE_MODE に on を入れて、startNotice を実行（時間切れで残ったら、もう一度実行）
//   4. 送り終わったら、NOTICE_MODE を削除する

const NOTICE_LIST_SHEET_NAME = "期間外お知らせ対象一覧";
const NOTICE_LOG_SHEET_NAME = "期間外お知らせ送信記録";
const NOTICE_MODE_PROPERTY = "NOTICE_MODE";
const NOTICE_REASON_PREFIX = "期間中に回収できる日がありません";
const NOTICE_TIME_LIMIT_MS = 4.5 * 60 * 1000;
const NOTICE_LOG_HEADER = ["キー", "送信先", "対象店舗", "結果", "日時"];
const NOTICE_SUBJECT = "【株式会社クリメン】年末年始（12/30〜1/3）の廃棄物回収に関するお知らせ【期間中は回収がありません】";
const NOTICE_PERIOD_TEXT = "2026年12月30日(水)から2027年1月3日(日)";

// 「契約状況一覧」から、対象を集めて、「期間外お知らせ対象一覧」に書き出す（一度作ったら、上書きしない）
function listNoticeTargets() {
  const ss = invGetSpreadsheet();
  let out = ss.getSheetByName(NOTICE_LIST_SHEET_NAME);
  if (out && out.getLastRow() > 1) throw new Error(`「${NOTICE_LIST_SHEET_NAME}」シートは、すでに作ってあります。作り直す場合は、シートを削除してから、もう一度実行してください。`);

  const src = ss.getSheetByName("契約状況一覧");
  if (!src) throw new Error("「契約状況一覧」シートが見つかりません。");
  const values = src.getDataRange().getDisplayValues();
  const head = values[0];
  const idx = name => head.indexOf(name);
  const cols = ["店舗名", "契約者", "送信先アドレス", "理由"];
  if (cols.some(c => idx(c) < 0)) throw new Error("「契約状況一覧」の見出しが、想定と違います。");

  const byKey = {};
  const rows = [];
  let noAddress = 0;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idx("理由")]).indexOf(NOTICE_REASON_PREFIX) !== 0) continue;
    const email = String(values[i][idx("送信先アドレス")] || "").trim();
    const customer = String(values[i][idx("契約者")] || "").trim();
    if (!email) noAddress++;
    const key = email + "|" + customer;
    if (!byKey[key]) {
      byKey[key] = [[], customer, email, 0, key];
      rows.push(byKey[key]);
    }
    byKey[key][0].push(values[i][idx("店舗名")]);
    byKey[key][3]++;
  }
  rows.forEach(row => { row[0] = row[0].join("\n"); });

  if (!out) out = ss.insertSheet(NOTICE_LIST_SHEET_NAME);
  out.clear();
  const outHead = ["対象店舗", "契約者", "送信先アドレス", "店舗数", "キー"];
  out.getRange(1, 1, 1, outHead.length).setValues([outHead]);
  if (rows.length > 0) out.getRange(2, 1, rows.length, outHead.length).setValues(rows);
  const storeCount = rows.reduce((n, r) => n + r[3], 0);
  console.log(`お知らせの対象: ${rows.length} 通（店舗は、合計 ${storeCount} 件）。アドレスが空の店舗: ${noAddress} 件。シート「${NOTICE_LIST_SHEET_NAME}」に書き出しました（お客様には、何も送っていません）。`);
}

function noticeTargets_() {
  const sheet = invGetSpreadsheet().getSheetByName(NOTICE_LIST_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) throw new Error(`先に listNoticeTargets を実行して、「${NOTICE_LIST_SHEET_NAME}」を作ってください。`);
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getDisplayValues()
    .filter(row => row[4])
    .map(row => ({ stores: String(row[0]).split("\n"), customer: row[1], email: String(row[2] || "").trim(), key: row[4] }));
}

function noticeBuildMail_(target) {
  const lines = [];
  lines.push("ご担当者様", "");
  lines.push("平素より大変お世話になっております。株式会社クリメンです。");
  lines.push("年末年始の回収について、お知らせいたします。", "");
  if (target.customer) lines.push("【ご契約者】", target.customer, "");
  lines.push("【対象店舗】");
  target.stores.forEach(store => lines.push(`・${store}`));
  lines.push("");
  lines.push(`■回収のない期間：${NOTICE_PERIOD_TEXT}`, "");
  lines.push("貴店は、上記の期間中、可燃ごみ・不燃ごみともに、回収はございません。");
  lines.push("※ 不燃ごみも、期間中は、回収いたしません。");
  lines.push("※ 期間中の回収のご依頼は、承ることができません。あらかじめ、ご了承ください。", "");
  lines.push("■お願い");
  lines.push("期間中のごみは、期間前までに、排出を済ませていただきますよう、お願いいたします。");
  lines.push("期間中は、回収に伺いませんので、ごみを出さないよう、お願いいたします。", "");
  lines.push("■回収の再開");
  lines.push("通常どおりの回収は、年明けの回収曜日から、再開いたします。", "");
  lines.push("何卒よろしくお願いいたします。", "", INV_SENDER_NAME);
  return { subject: NOTICE_SUBJECT, body: lines.join("\n") };
}

function sendTestNotice() {
  const testTo = String(PropertiesService.getScriptProperties().getProperty(INV_TEST_TO_PROPERTY) || "").trim();
  if (!testTo) throw new Error(`スクリプトプロパティ ${INV_TEST_TO_PROPERTY} にテスト用のメールアドレスを設定してください。`);
  const from = invGetFrom();
  const targets = noticeTargets_();
  const invalid = targets.filter(t => !/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(t.email)).length;
  console.log(`対象: 全 ${targets.length} 通（店舗は、合計 ${targets.reduce((n, t) => n + t.stores.length, 0)} 件）。宛先が正しくないもの: ${invalid} 件。`);
  if (targets.length === 0) throw new Error("対象がありません。");
  targets.slice(0, 2).forEach(target => {
    const message = noticeBuildMail_(target);
    invSend(testTo, "【テスト】" + message.subject, `（テスト送信です。本来の宛先: ${target.email}）\n\n` + message.body, from);
  });
  console.log(`${Math.min(2, targets.length)} 通のテストメールを ${testTo} に送りました（送信元: ${from}）。`);
}

function startNotice() {
  const props = PropertiesService.getScriptProperties();
  if (String(props.getProperty(NOTICE_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${NOTICE_MODE_PROPERTY} が on ではないため、送信しません。`);
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
    let logSheet = ss.getSheetByName(NOTICE_LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = ss.insertSheet(NOTICE_LOG_SHEET_NAME);
      logSheet.appendRow(NOTICE_LOG_HEADER);
    }
    const logged = new Set();
    if (logSheet.getLastRow() > 1) {
      logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getDisplayValues().forEach(row => logged.add(row[0]));
    }
    const targets = noticeTargets_().filter(target => !logged.has(target.key));
    for (let n = 0; n < targets.length; n++) {
      if (Date.now() - startTime > NOTICE_TIME_LIMIT_MS) { remaining = targets.length - n; break; }
      const target = targets[n];
      let result = INV_RESULT_SENT;
      if (!/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(target.email)) {
        console.error(`宛先が正しくありません: ${target.stores[0]}`);
        result = INV_RESULT_FAILED;
      } else {
        try {
          const message = noticeBuildMail_(target);
          invSend(target.email, message.subject, message.body, from);
        } catch (e) {
          console.error(`お知らせメールを送れませんでした: ${target.stores[0]} / ${e}`);
          result = INV_RESULT_FAILED;
        }
      }
      logSheet.appendRow([target.key, target.email, target.stores.join("・"), result, invFormatNow()]);
      SpreadsheetApp.flush();
      if (result === INV_RESULT_SENT) { sentCount++; Utilities.sleep(INV_GRAPH_INTERVAL_MS); } else { failedCount++; }
    }
    console.log(`今回 ${sentCount} 通送信（失敗 ${failedCount} 件）。` + (remaining > 0 ? `残り ${remaining} 件は、もう一度 startNotice を実行してください。` : "すべて完了しました。") + `記録は「${NOTICE_LOG_SHEET_NAME}」シートにあります。`);
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

// ===== 新しく契約したお客様への、アンケートのご案内メール =====
// 「リスト」の NEWC_FIRST_ROW 行目以降（新規契約の店舗で、今回フォームを作った行）を、フォームごとに、1通ずつ送る。
// 送信元は、案内メールと同じ（Microsoft Graph、km-customer@kurimen.co.jp）。案内メール送信.js の関数を使う。
//   ・保留のお客様（EXMAIL_HOLD_CUSTOMERS）は、日程変更の店舗と、まとめて送るため、ここでは送らない
//   ・同じ会社・同じアドレスで、すでに別のフォームを送っているお客様には、「別の店舗のアンケートです」の1行を入れる
// 1通ごとに「新規契約メール送信記録」シートに記録して、同じフォームには、二重に送らない。
//
// 使い方:
//   1. listNewContractMailTargets を実行 → 対象を「新規契約メール対象一覧」シートに書き出して、固定する（お客様には送らない）。内容を目で確認する
//   2. sendTestNewContractMails を実行 → 「別フォームの注記あり」と「なし」の、それぞれ先頭1件を、INVITE_TEST_TO のアドレスだけに送る
//   3. スクリプトプロパティ NEWC_MAIL_MODE に on を入れて、startNewContractMails を実行（時間切れで残ったら、もう一度実行）
//   4. 送り終わったら、NEWC_MAIL_MODE を削除する

const NEWCMAIL_LIST_SHEET_NAME = "新規契約メール対象一覧";
const NEWCMAIL_LOG_SHEET_NAME = "新規契約メール送信記録";
const NEWCMAIL_MODE_PROPERTY = "NEWC_MAIL_MODE";
const NEWCMAIL_TIME_LIMIT_MS = 4.5 * 60 * 1000;
const NEWCMAIL_LOG_HEADER = ["キー", "送信先", "対象店舗", "注記", "結果", "日時"];
const NEWCMAIL_SUBJECT = "【株式会社クリメン】年末年始（12/30〜1/3）の特別回収に関するアンケートのお願い";
const NEWCMAIL_NOTE = "※以前、同じ会社の別の店舗について、アンケートをお送りしております。下記の店舗は、以前のフォームではなく、こちらのフォームでご回答ください。";

// 1. 対象を集めて、「新規契約メール対象一覧」に書き出す（一度作ったら、上書きしない）
function listNewContractMailTargets() {
  const ss = invGetSpreadsheet();
  let out = ss.getSheetByName(NEWCMAIL_LIST_SHEET_NAME);
  if (out && out.getLastRow() > 1) throw new Error(`「${NEWCMAIL_LIST_SHEET_NAME}」シートは、すでに作ってあります。作り直す場合は、シートを削除してから、もう一度実行してください。`);

  const list = ss.getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  const hold = EXMAIL_HOLD_CUSTOMERS.map(extraNorm_);
  // 新しい行より前に、同じ会社（契約者）で、フォームがある店舗
  const hadForm = new Set();
  for (let i = 1; i < NEWC_FIRST_ROW - 1 && i < list.length; i++) {
    if (list[i][9] && list[i][10] === "送信済" && list[i][16]) hadForm.add(extraNorm_(list[i][16]));   // 同じ契約者（会社）。アドレスが違っていても、同じ会社なら「あり」
  }
  const byForm = {};
  const rows = [];
  for (let i = NEWC_FIRST_ROW - 1; i < list.length; i++) {
    const row = list[i];
    const formId = row[9];
    if (!formId || row[10] !== "送信済") continue;
    if (!byForm[formId]) {
      byForm[formId] = { stores: [], customer: row[16], email: String(row[1] || "").trim(), formId: formId, url: row[8] };
      rows.push(byForm[formId]);
    }
    byForm[formId].stores.push(row[0]);
  }
  const outRows = rows.map(t => [
    t.stores.join("\n"), t.customer, t.email,
    hadForm.has(extraNorm_(t.customer)) ? "あり" : "",
    hold.some(h => extraNorm_(t.customer).indexOf(h) !== -1) ? "保留" : "",
    t.formId, t.url
  ]);

  if (!out) out = ss.insertSheet(NEWCMAIL_LIST_SHEET_NAME);
  out.clear();
  const head = ["対象店舗", "契約者", "送信先アドレス", "別フォームの注記", "保留", "フォームID", "フォームURL"];
  out.getRange(1, 1, 1, head.length).setValues([head]);
  if (outRows.length > 0) out.getRange(2, 1, outRows.length, head.length).setValues(outRows);
  console.log(`新規契約のご案内: ${outRows.length} 通（店舗は、合計 ${rows.reduce((n, t) => n + t.stores.length, 0)} 件）。うち、別フォームの注記あり ${outRows.filter(r => r[3]).length} 通、保留 ${outRows.filter(r => r[4]).length} 通（保留は送りません）。シート「${NEWCMAIL_LIST_SHEET_NAME}」に書き出しました（お客様には、何も送っていません）。`);
}

function newcMailTargets_() {
  const sheet = invGetSpreadsheet().getSheetByName(NEWCMAIL_LIST_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) throw new Error(`先に listNewContractMailTargets を実行して、「${NEWCMAIL_LIST_SHEET_NAME}」を作ってください。`);
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getDisplayValues()
    .filter(row => row[5] && row[4] !== "保留")
    .map(row => ({ stores: String(row[0]).split("\n"), customer: row[1], email: String(row[2] || "").trim(), note: row[3] === "あり", formId: row[5], url: row[6], key: row[5] }));
}

function newcMailBuild_(target) {
  const lines = [];
  lines.push("ご担当者様", "");
  lines.push("平素より大変お世話になっております。株式会社クリメンです。");
  lines.push("この度は、新たにご契約をいただき、誠にありがとうございます。");
  lines.push("年末年始の特別回収につきまして、ご希望をお伺いするアンケートをお送りします。", "");
  if (target.note) lines.push(NEWCMAIL_NOTE, "");
  if (target.customer) lines.push("【ご契約者】", target.customer, "");
  lines.push("【対象店舗】");
  target.stores.forEach(store => lines.push(`・${store}`));
  lines.push("");
  lines.push(`■特別回収期間：${INV_PERIOD_TEXT}`);
  lines.push("　※1月1日は回収できません。");
  lines.push("■回収対象：可燃ごみのみ");
  lines.push("■回収時間：各回収日の当日の早朝から午前中");
  lines.push("■特別回収料金：店舗ごとに異なります。アンケート内の各店舗の設問でご確認ください。", "");
  lines.push(`■ご回答期限：${INV_DEADLINE_TEXT}まで`, "");
  lines.push("下記のアンケートフォームから、ご回答をお願いいたします。");
  lines.push(target.url, "");
  lines.push("【ご回答について】");
  lines.push("・フォームの次の画面で、店舗ごとに、各日の回収を「希望する」「希望しない」からお選びください。");
  lines.push("・特別回収を希望されない場合も、すべての日について「希望しない」を選択してご回答ください。");
  lines.push("・通常の可燃回収曜日以外は回収できません。希望日は、回収に伺う日付でお選びください。");
  lines.push("・ご回答のあと、回答内容の控えをこのメールアドレスにお送りします。");
  lines.push("・回答内容を変更する場合は、同じフォームから再度ご回答ください（最新のご回答を有効とします）。");
  lines.push(`　変更のご希望は${INV_DEADLINE_TEXT}までにお願いいたします。`, "");
  if (INV_CONTACT) lines.push(`ご不明な点は、${INV_CONTACT}までご連絡ください。`, "");
  if (INV_NOREPLY_TEXT) lines.push(INV_NOREPLY_TEXT, "");
  lines.push("お忙しいところ恐れ入りますが、よろしくお願いいたします。", "", INV_SENDER_NAME);
  return { subject: NEWCMAIL_SUBJECT, body: lines.join("\n") };
}

// 2. テスト送信: 注記あり・なしの、先頭1件ずつを、INVITE_TEST_TO のアドレスだけに送る（お客様には送らない。記録にも残さない）
function sendTestNewContractMails() {
  const testTo = String(PropertiesService.getScriptProperties().getProperty(INV_TEST_TO_PROPERTY) || "").trim();
  if (!testTo) throw new Error(`スクリプトプロパティ ${INV_TEST_TO_PROPERTY} にテスト用のメールアドレスを設定してください。`);
  const from = invGetFrom();
  const targets = newcMailTargets_();
  const invalid = targets.filter(t => !/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(t.email)).length;
  console.log(`対象: 全 ${targets.length} 通（保留を除く）。宛先が正しくないもの: ${invalid} 件。`);
  if (targets.length === 0) throw new Error("対象がありません。");
  const withNote = targets.filter(t => t.note)[0];
  const withoutNote = targets.filter(t => !t.note)[0];
  [withNote, withoutNote].filter(Boolean).forEach(target => {
    const message = newcMailBuild_(target);
    invSend(testTo, "【テスト】" + message.subject, `（テスト送信です。本来の宛先: ${target.email}）\n\n` + message.body, from);
    Utilities.sleep(INV_GRAPH_INTERVAL_MS);
  });
  console.log(`テストメールを ${testTo} に送りました（注記あり: ${withNote ? "送った" : "対象なし"}、注記なし: ${withoutNote ? "送った" : "対象なし"}）。送信元: ${from}`);
}

// 3. 本番の送信
function startNewContractMails() {
  const props = PropertiesService.getScriptProperties();
  if (String(props.getProperty(NEWCMAIL_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${NEWCMAIL_MODE_PROPERTY} が on ではないため、送信しません。`);
    return;
  }
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) { console.log("別の処理が動いているためスキップしました。少し待って、もう一度実行してください。"); return; }
  const startTime = Date.now();
  let sentCount = 0, failedCount = 0, remaining = 0;
  try {
    const from = invGetFrom();
    const ss = invGetSpreadsheet();
    let logSheet = ss.getSheetByName(NEWCMAIL_LOG_SHEET_NAME);
    if (!logSheet) { logSheet = ss.insertSheet(NEWCMAIL_LOG_SHEET_NAME); logSheet.appendRow(NEWCMAIL_LOG_HEADER); }
    const logged = new Set();
    if (logSheet.getLastRow() > 1) logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getDisplayValues().forEach(row => logged.add(row[0]));
    const targets = newcMailTargets_().filter(t => !logged.has(t.key));
    for (let n = 0; n < targets.length; n++) {
      if (Date.now() - startTime > NEWCMAIL_TIME_LIMIT_MS) { remaining = targets.length - n; break; }
      const target = targets[n];
      let result = INV_RESULT_SENT;
      if (!/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(target.email)) {
        console.error(`宛先が正しくありません: ${target.stores[0]}`);
        result = INV_RESULT_FAILED;
      } else {
        try {
          const message = newcMailBuild_(target);
          invSend(target.email, message.subject, message.body, from);
        } catch (e) {
          console.error(`メールを送れませんでした: ${target.stores[0]} / ${e}`);
          result = INV_RESULT_FAILED;
        }
      }
      logSheet.appendRow([target.key, target.email, target.stores.join("・"), target.note ? "別フォームの注記あり" : "", result, invFormatNow()]);
      SpreadsheetApp.flush();
      if (result === INV_RESULT_SENT) { sentCount++; Utilities.sleep(INV_GRAPH_INTERVAL_MS); } else { failedCount++; }
    }
    console.log(`今回 ${sentCount} 通送信（失敗 ${failedCount} 件）。` + (remaining > 0 ? `残り ${remaining} 件は、もう一度 startNewContractMails を実行してください。` : "すべて完了しました。") + `記録は「${NEWCMAIL_LOG_SHEET_NAME}」シートにあります。`);
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

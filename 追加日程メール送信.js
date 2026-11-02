// ===== 日程の追加（12/29・1/4）に伴う、お客様へのメール =====
// 「リスト」の S列・T列に「○」が入っている店舗（追加日程フォーム.js の applyExtraDates で、設問を足した店舗）の契約者に、
// フォームごとに、1通ずつ送る。送信元は、案内メールと同じ（Microsoft Graph、km-customer@kurimen.co.jp）。
//   ・すでに回答があるフォーム … 前回の回答を入力済みにしたリンクを付けて、お詫びと、追加のご回答のお願いを送る
//   ・まだ回答がないフォーム … 状況の説明と、いつものフォームのリンクを送る
//   ・保留のお客様（EXMAIL_HOLD_CUSTOMERS）… 新規契約の店舗と、まとめて送るため、今は送らない
// 1通ごとに「追加日程メール送信記録」シートに記録して、同じフォームには、二重に送らない。
//
// 使い方:
//   1. listExtraDateMailTargets を実行 → 対象を「追加日程メール対象一覧」シートに書き出して、固定する（お客様には送らない）。内容を目で確認する
//   2. sendTestExtraDateMails を実行 → 回答済み・未回答の、それぞれ先頭1件分を、INVITE_TEST_TO のアドレスだけに送る（文面を確認する）
//   3. スクリプトプロパティ EXTRA_MAIL_MODE に on を入れて、startExtraDateMails を実行（時間切れで残ったら、もう一度実行）
//   4. 送り終わったら、EXTRA_MAIL_MODE を削除する

const EXMAIL_LIST_SHEET_NAME = "追加日程メール対象一覧";
const EXMAIL_LOG_SHEET_NAME = "追加日程メール送信記録";
const EXMAIL_MODE_PROPERTY = "EXTRA_MAIL_MODE";
const EXMAIL_TIME_LIMIT_MS = 4.5 * 60 * 1000;
const EXMAIL_LOG_HEADER = ["キー", "送信先", "対象店舗", "種類", "結果", "日時"];
// 新規契約の店舗と、まとめて送るお客様（今は送らない）。全角半角・空白の違いは、無視する
const EXMAIL_HOLD_CUSTOMERS = ["集Ｅｎｔｅｒｐｒｉｓｅ", "エイブルジャパン", "ＩＦＲＥＡ ＤＩＮＩＮＧ"];

function exmailFee_(header, vendor) {
  const fees = EXTRA_DATE_VENDOR_FEES[header] || {};
  const key = Object.keys(fees).filter(k => extraNorm_(vendor).indexOf(k) !== -1)[0];
  return key ? fees[key] : (FORM_VENDOR_FEES[vendor] || FORM_DEFAULT_FEE);
}

// 1. 対象を集めて、「追加日程メール対象一覧」に書き出す（一度作ったら、上書きしない）
function listExtraDateMailTargets() {
  const ss = invGetSpreadsheet();
  let out = ss.getSheetByName(EXMAIL_LIST_SHEET_NAME);
  if (out && out.getLastRow() > 1) throw new Error(`「${EXMAIL_LIST_SHEET_NAME}」シートは、すでに作ってあります。作り直す場合は、シートを削除してから、もう一度実行してください。`);

  const values = ss.getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  const hold = EXMAIL_HOLD_CUSTOMERS.map(extraNorm_);
  const byForm = {};
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const formId = row[9];
    if (!formId || row[10] !== "送信済") continue;
    EXTRA_DATE_COLS.forEach((col, k) => {
      if (row[col] !== "○") return;
      if (!byForm[formId]) {
        byForm[formId] = { stores: [], dates: [], fees: [], customer: row[16], email: String(row[1] || "").trim(), formId: formId };
        rows.push(byForm[formId]);
      }
      byForm[formId].stores.push(row[0]);
      byForm[formId].dates.push(EXTRA_DATE_HEADERS[k]);
      byForm[formId].fees.push(exmailFee_(EXTRA_DATE_HEADERS[k], row[2]));
    });
  }

  const outRows = rows.map(t => {
    const held = hold.some(h => extraNorm_(t.customer).indexOf(h) !== -1);
    let answered = "";
    try { answered = FormApp.openById(t.formId).getResponses().length > 0 ? "回答あり" : "回答なし"; } catch (e) { answered = "確認できず"; }
    return [t.stores.join("\n"), t.dates.join("\n"), t.fees.join("\n"), t.customer, t.email, answered, held ? "保留" : "", t.formId];
  });

  if (!out) out = ss.insertSheet(EXMAIL_LIST_SHEET_NAME);
  out.clear();
  const head = ["対象店舗", "追加の日", "料金", "契約者", "送信先アドレス", "回答状況（作成時）", "保留", "フォームID"];
  out.getRange(1, 1, 1, head.length).setValues([head]);
  if (outRows.length > 0) out.getRange(2, 1, outRows.length, head.length).setValues(outRows);
  const count = f => outRows.filter(f).length;
  console.log(`追加日程メールの対象: ${outRows.length} 通（店舗は、合計 ${rows.reduce((n, t) => n + t.stores.length, 0)} 件）。うち、回答あり ${count(r => r[5] === "回答あり")} 通、回答なし ${count(r => r[5] === "回答なし")} 通、確認できず ${count(r => r[5] === "確認できず")} 通、保留 ${count(r => r[6] === "保留")} 通（保留は送りません）。シート「${EXMAIL_LIST_SHEET_NAME}」に書き出しました（お客様には、何も送っていません）。`);
}

function exmailTargets_(includeHold) {
  const sheet = invGetSpreadsheet().getSheetByName(EXMAIL_LIST_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) throw new Error(`先に listExtraDateMailTargets を実行して、「${EXMAIL_LIST_SHEET_NAME}」を作ってください。`);
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getDisplayValues()
    .filter(row => row[7] && (includeHold || row[6] !== "保留"))
    .map(row => ({
      stores: String(row[0]).split("\n"), dates: String(row[1]).split("\n"), fees: String(row[2]).split("\n"),
      customer: row[3], email: String(row[4] || "").trim(), formId: row[7], key: row[7]
    }));
}

// 送る時点の、回答の有無とリンクを調べる。回答があれば、最新の回答を入力済みにしたリンクにする
function exmailResolve_(target) {
  const form = FormApp.openById(target.formId);
  const responses = form.getResponses();
  if (responses.length > 0) return { answered: true, url: responses[responses.length - 1].toPrefilledUrl() };
  return { answered: false, url: form.getPublishedUrl() };
}

function exmailBuildMail_(target, resolved) {
  const uniqueDates = target.dates.filter((d, i) => target.dates.indexOf(d) === i);
  const lines = [];
  lines.push("ご担当者様", "");
  lines.push("平素より大変お世話になっております。株式会社クリメンです。");
  if (resolved.answered) lines.push("先日は、年末年始の特別回収のアンケートにご回答いただき、ありがとうございました。", "");
  else lines.push("年末年始の特別回収につきまして、ご希望をお伺いするアンケートをお送りします。", "");
  lines.push("大変申し訳ございません。収集業者の年末年始の営業の都合により、下記の店舗では、通常の収集ができなくなった日がございます。");
  lines.push("この日に収集をご希望の場合は、有料の特別回収となります。");
  if (resolved.answered) lines.push("すでにご回答いただいたにもかかわらず、追加でのご回答をお願いすることになり、重ねてお詫び申し上げます。");
  lines.push("");
  if (target.customer) lines.push("【ご契約者】", target.customer, "");
  lines.push("【対象店舗と、対象の日】");
  target.stores.forEach((store, i) => lines.push(`・${store}：${target.dates[i]}（特別回収料金：${target.fees[i]}/1日）`));
  lines.push("");
  lines.push("■回収対象：可燃ごみのみ");
  lines.push("　※通常の可燃回収曜日以外は回収できません。");
  lines.push(`■ご回答期限：${INV_DEADLINE_TEXT}まで`);
  lines.push("※ご回答がない場合、上記の日の収集は行いません。", "");
  if (resolved.answered) {
    lines.push("下のリンクには、前回のご回答が入力済みです。上記の日の設問だけ、「希望する」「希望しない」をお選びいただき、送信してください。");
  } else {
    lines.push("下記のアンケートフォームから、ご回答をお願いいたします。フォームには、上記の日の設問も入っています。");
  }
  lines.push(resolved.url, "");
  lines.push("【ご回答について】");
  if (!resolved.answered) {
    lines.push("・フォームの次の画面で、店舗ごとに、各日の回収を「希望する」「希望しない」からお選びください。");
    lines.push("・特別回収を希望されない場合も、すべての日について「希望しない」を選択してご回答ください。");
    lines.push("・通常の可燃回収曜日以外は回収できません。希望日は、回収に伺う日付でお選びください。");
  } else {
    lines.push("・前回のご回答の内容は、そのままで構いません。変更したい場合も、同じリンクで修正できます。");
  }
  lines.push("・ご回答のあと、回答内容の控えをこのメールアドレスにお送りします。");
  lines.push("・回答内容を変更する場合は、同じフォームから再度ご回答ください（最新のご回答を有効とします）。");
  lines.push(`　変更のご希望は${INV_DEADLINE_TEXT}までにお願いいたします。`, "");
  if (INV_CONTACT) lines.push(`ご不明な点は、${INV_CONTACT}までご連絡ください。`, "");
  if (INV_NOREPLY_TEXT) lines.push(INV_NOREPLY_TEXT, "");
  lines.push("ご迷惑をおかけし、誠に申し訳ございません。何卒よろしくお願いいたします。", "", INV_SENDER_NAME);
  const subject = resolved.answered
    ? `【株式会社クリメン】【重要】年末年始の特別回収：${uniqueDates.join("・")}の収集について（追加のご回答のお願い）`
    : `【株式会社クリメン】【重要】年末年始の特別回収のご案内（${uniqueDates.join("・")}の収集について）`;
  return { subject: subject, body: lines.join("\n") };
}

// 2. テスト送信: 回答あり・回答なしの、先頭1件ずつを、INVITE_TEST_TO のアドレスだけに送る（お客様には送らない。記録にも残さない）
function sendTestExtraDateMails() {
  const testTo = String(PropertiesService.getScriptProperties().getProperty(INV_TEST_TO_PROPERTY) || "").trim();
  if (!testTo) throw new Error(`スクリプトプロパティ ${INV_TEST_TO_PROPERTY} にテスト用のメールアドレスを設定してください。`);
  const from = invGetFrom();
  const targets = exmailTargets_(false);
  const invalid = targets.filter(t => !/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(t.email)).length;
  console.log(`対象: 全 ${targets.length} 通（保留を除く）。宛先が正しくないもの: ${invalid} 件。`);
  if (targets.length === 0) throw new Error("対象がありません。");
  const sent = { answered: false, unanswered: false };
  for (let i = 0; i < targets.length && !(sent.answered && sent.unanswered); i++) {
    const resolved = exmailResolve_(targets[i]);
    const kind = resolved.answered ? "answered" : "unanswered";
    if (sent[kind]) continue;
    const message = exmailBuildMail_(targets[i], resolved);
    invSend(testTo, "【テスト】" + message.subject, `（テスト送信です。本来の宛先: ${targets[i].email}）\n\n` + message.body, from);
    sent[kind] = true;
    Utilities.sleep(INV_GRAPH_INTERVAL_MS);
  }
  console.log(`テストメールを ${testTo} に送りました（回答あり用: ${sent.answered ? "送った" : "対象なし"}、回答なし用: ${sent.unanswered ? "送った" : "対象なし"}）。送信元: ${from}`);
}

// 3. 本番の送信
function startExtraDateMails() {
  const props = PropertiesService.getScriptProperties();
  if (String(props.getProperty(EXMAIL_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${EXMAIL_MODE_PROPERTY} が on ではないため、送信しません。`);
    return;
  }
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) { console.log("別の処理が動いているためスキップしました。少し待って、もう一度実行してください。"); return; }
  const startTime = Date.now();
  let sentCount = 0, failedCount = 0, remaining = 0;
  try {
    const from = invGetFrom();
    const ss = invGetSpreadsheet();
    let logSheet = ss.getSheetByName(EXMAIL_LOG_SHEET_NAME);
    if (!logSheet) { logSheet = ss.insertSheet(EXMAIL_LOG_SHEET_NAME); logSheet.appendRow(EXMAIL_LOG_HEADER); }
    const logged = new Set();
    if (logSheet.getLastRow() > 1) logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getDisplayValues().forEach(row => logged.add(row[0]));
    const targets = exmailTargets_(false).filter(t => !logged.has(t.key));
    for (let n = 0; n < targets.length; n++) {
      if (Date.now() - startTime > EXMAIL_TIME_LIMIT_MS) { remaining = targets.length - n; break; }
      const target = targets[n];
      let result = INV_RESULT_SENT;
      let kind = "";
      if (!/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(target.email)) {
        console.error(`宛先が正しくありません: ${target.stores[0]}`);
        result = INV_RESULT_FAILED;
      } else {
        try {
          const resolved = exmailResolve_(target);
          kind = resolved.answered ? "回答あり" : "回答なし";
          const message = exmailBuildMail_(target, resolved);
          invSend(target.email, message.subject, message.body, from);
        } catch (e) {
          console.error(`メールを送れませんでした: ${target.stores[0]} / ${e}`);
          result = INV_RESULT_FAILED;
        }
      }
      logSheet.appendRow([target.key, target.email, target.stores.join("・"), kind, result, invFormatNow()]);
      SpreadsheetApp.flush();
      if (result === INV_RESULT_SENT) { sentCount++; Utilities.sleep(INV_GRAPH_INTERVAL_MS); } else { failedCount++; }
    }
    console.log(`今回 ${sentCount} 通送信（失敗 ${failedCount} 件）。` + (remaining > 0 ? `残り ${remaining} 件は、もう一度 startExtraDateMails を実行してください。` : "すべて完了しました。") + `記録は「${EXMAIL_LOG_SHEET_NAME}」シートにあります。`);
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

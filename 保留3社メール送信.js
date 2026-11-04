// ===== 保留の3社（日程変更の店舗と、新店舗の、両方があるお客様）への、まとめのメール =====
// 対象: 「追加日程メール対象一覧」で「保留」の行（集Ｅｎｔｅｒｐｒｉｓｅ・ＩＦＲＥＡ ＤＩＮＩＮＧ・エイブルジャパン）。
// 新店舗は、先に applyAddStores で、今あるフォームに足しておく（「リスト」の新店舗の行が、今あるフォームに付け替わる）。
// 送り先は、今あるフォームの、元のアドレス。前回の回答が入力済みのリンクを付ける。
//   1. previewHoldMails        … 対象（日程変更の店舗と、新店舗）を、シート「保留3社メール対象」に書き出す（お客様には送らない）
//   2. sendTestHoldMails       … 3通を、INVITE_TEST_TO のアドレスだけに送る
//   3. スクリプトプロパティ HOLD_MAIL_MODE に on を入れて、startHoldMails を実行
//   4. 送り終わったら、HOLD_MAIL_MODE を削除する

const HOLDMAIL_SHEET_NAME = "保留3社メール対象";
const HOLDMAIL_LOG_SHEET_NAME = "保留3社メール送信記録";
const HOLDMAIL_MODE_PROPERTY = "HOLD_MAIL_MODE";
const HOLDMAIL_LOG_HEADER = ["キー", "送信先", "対象店舗", "結果", "日時"];

// 対象を集める: 「追加日程メール対象一覧」の保留の行 + 「リスト」で、同じフォームに付け替わった新店舗
function holdMailTargets_() {
  const ss = invGetSpreadsheet();
  const sheet = ss.getSheetByName(EXMAIL_LIST_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) throw new Error(`「${EXMAIL_LIST_SHEET_NAME}」シートがありません。`);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getDisplayValues().filter(r => r[7] && r[6] === "保留");
  const list = ss.getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  return rows.map(r => {
    const formId = r[7];
    const newStores = [];
    for (let i = NEWC_FIRST_ROW - 1; i < list.length; i++) if (list[i][9] === formId && list[i][10] === "送信済") newStores.push(list[i][0]);
    return { stores: String(r[0]).split("\n"), dates: String(r[1]).split("\n"), fees: String(r[2]).split("\n"), customer: r[3], email: String(r[4] || "").trim(), formId: formId, newStores: newStores, key: formId };
  });
}

function previewHoldMails() {
  const ss = invGetSpreadsheet();
  const targets = holdMailTargets_();
  const rows = targets.map(t => {
    let answered = "確認できず";
    try { answered = FormApp.openById(t.formId).getResponses().length > 0 ? "回答あり" : "回答なし"; } catch (e) { /* 確認できず */ }
    return [t.customer, t.email, t.stores.map((s, i) => `${s}：${t.dates[i]}`).join("\n"), t.newStores.join("\n"), answered, t.formId];
  });
  let out = ss.getSheetByName(HOLDMAIL_SHEET_NAME);
  if (!out) out = ss.insertSheet(HOLDMAIL_SHEET_NAME);
  out.clear();
  const head = ["契約者", "送信先アドレス", "日程が変わった店舗：日", "新店舗（今あるフォームに追加済み）", "回答状況", "フォームID"];
  out.getRange(1, 1, 1, head.length).setValues([head]);
  if (rows.length > 0) out.getRange(2, 1, rows.length, head.length).setValues(rows);
  console.log(`保留の対象: ${rows.length} 通。新店舗が、今あるフォームに入っているもの: ${targets.filter(t => t.newStores.length > 0).length} 通。シート「${HOLDMAIL_SHEET_NAME}」に書き出しました（お客様には、何も送っていません）。`);
}

function holdMailBuild_(target, resolved) {
  const lines = [];
  lines.push("ご担当者様", "");
  lines.push("平素より大変お世話になっております。株式会社クリメンです。");
  if (resolved.answered) lines.push("先日は、年末年始の特別回収のアンケートにご回答いただき、ありがとうございました。");
  lines.push("今回、次の2点について、ご回答をお願いしたく、ご連絡いたしました。", "");
  if (target.customer) lines.push("【ご契約者】", target.customer, "");
  lines.push("■1　日程が変わった店舗");
  lines.push("大変申し訳ございません。収集業者の年末年始の営業の都合により、下記の店舗では、通常の収集ができなくなった日がございます。");
  lines.push("この日に収集をご希望の場合は、有料の特別回収となります。");
  if (resolved.answered) lines.push("すでにご回答いただいたにもかかわらず、追加でのご回答をお願いすることになり、重ねてお詫び申し上げます。");
  target.stores.forEach((store, i) => lines.push(`・${store}：${target.dates[i]}（特別回収料金：${target.fees[i]}/1日）`));
  lines.push("※ご回答がない場合、上記の日の収集は行いません。", "");
  if (target.newStores.length > 0) {
    lines.push("■2　新しい店舗");
    lines.push("この度は、新たな店舗のご契約をいただき、誠にありがとうございます。");
    lines.push("新しい店舗は、以前お送りしたアンケートと同じフォームに、追加しました。フォームの最後の、新しい店舗のページに、ご回答をお願いいたします。");
    target.newStores.forEach(store => lines.push(`・${store}`));
    lines.push("");
  }
  lines.push(`■特別回収期間：${INV_PERIOD_TEXT}`);
  lines.push("　※1月1日は回収できません。");
  lines.push("■回収対象：可燃ごみのみ");
  lines.push("　※通常の可燃回収曜日以外は回収できません。");
  lines.push("■回収時間：各回収日の当日の早朝から午前中");
  lines.push(`■ご回答期限：${INV_DEADLINE_TEXT}まで`, "");
  if (resolved.answered) {
    lines.push("下のリンクには、前回のご回答が入力済みです。上記の、対象の日の設問と、新しい店舗のページだけ、「希望する」「希望しない」をお選びいただき、送信してください。");
  } else {
    lines.push("下記のアンケートフォームから、ご回答をお願いいたします。");
  }
  lines.push(resolved.url, "");
  lines.push("【ご回答について】");
  lines.push("・前回のご回答の内容は、そのままで構いません。変更したい場合も、同じリンクで修正できます。");
  lines.push("・ご回答のあと、回答内容の控えをこのメールアドレスにお送りします。");
  lines.push("・回答内容を変更する場合は、同じフォームから再度ご回答ください（最新のご回答を有効とします）。");
  lines.push(`　変更のご希望は${INV_DEADLINE_TEXT}までにお願いいたします。`, "");
  if (INV_CONTACT) lines.push(`ご不明な点は、${INV_CONTACT}までご連絡ください。`, "");
  if (INV_NOREPLY_TEXT) lines.push(INV_NOREPLY_TEXT, "");
  lines.push("ご迷惑をおかけし、誠に申し訳ございません。何卒よろしくお願いいたします。", "", INV_SENDER_NAME);
  const uniqueDates = target.dates.filter((d, i) => target.dates.indexOf(d) === i);
  return { subject: `【株式会社クリメン】【重要】年末年始の特別回収：${uniqueDates.join("・")}の収集について${target.newStores.length > 0 ? "と、新しい店舗のご案内" : ""}（ご回答のお願い）`, body: lines.join("\n") };
}

function sendTestHoldMails() {
  const testTo = String(PropertiesService.getScriptProperties().getProperty(INV_TEST_TO_PROPERTY) || "").trim();
  if (!testTo) throw new Error(`スクリプトプロパティ ${INV_TEST_TO_PROPERTY} にテスト用のメールアドレスを設定してください。`);
  const from = invGetFrom();
  const targets = holdMailTargets_();
  if (targets.length === 0) throw new Error("対象がありません。");
  targets.forEach(target => {
    const message = holdMailBuild_(target, exmailResolve_(target));
    invSend(testTo, "【テスト】" + message.subject, `（テスト送信です。本来の宛先: ${target.email}）\n\n` + message.body, from);
    Utilities.sleep(INV_GRAPH_INTERVAL_MS);
  });
  console.log(`テストメール ${targets.length} 通を ${testTo} に送りました（送信元: ${from}）。`);
}

function startHoldMails() {
  if (String(PropertiesService.getScriptProperties().getProperty(HOLDMAIL_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${HOLDMAIL_MODE_PROPERTY} が on ではないため、送信しません。`);
    return;
  }
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) { console.log("別の処理が動いているためスキップしました。少し待って、もう一度実行してください。"); return; }
  let sentCount = 0, failedCount = 0;
  try {
    const from = invGetFrom();
    const ss = invGetSpreadsheet();
    let logSheet = ss.getSheetByName(HOLDMAIL_LOG_SHEET_NAME);
    if (!logSheet) { logSheet = ss.insertSheet(HOLDMAIL_LOG_SHEET_NAME); logSheet.appendRow(HOLDMAIL_LOG_HEADER); }
    const logged = new Set();
    if (logSheet.getLastRow() > 1) logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getDisplayValues().forEach(row => logged.add(row[0]));
    holdMailTargets_().filter(t => !logged.has(t.key)).forEach(target => {
      let result = INV_RESULT_SENT;
      if (!/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(target.email)) {
        console.error(`宛先が正しくありません: ${target.stores[0]}`);
        result = INV_RESULT_FAILED;
      } else {
        try {
          const message = holdMailBuild_(target, exmailResolve_(target));
          invSend(target.email, message.subject, message.body, from);
        } catch (e) {
          console.error(`メールを送れませんでした: ${target.stores[0]} / ${e}`);
          result = INV_RESULT_FAILED;
        }
      }
      logSheet.appendRow([target.key, target.email, target.stores.concat(target.newStores).join("・"), result, invFormatNow()]);
      SpreadsheetApp.flush();
      if (result === INV_RESULT_SENT) { sentCount++; Utilities.sleep(INV_GRAPH_INTERVAL_MS); } else { failedCount++; }
    });
    console.log(`今回 ${sentCount} 通送信（失敗 ${failedCount} 件）。記録は「${HOLDMAIL_LOG_SHEET_NAME}」シートにあります。`);
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

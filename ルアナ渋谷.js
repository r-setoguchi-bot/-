// ===== ルアナ渋谷（野口興業。期間中は回収がなく、1/4(月)だけ特別回収が必要な店舗）を、今あるフォームに足して、お客様にお知らせする =====
// 同じ契約者（ホームディフェンステック株式会社）・同じアドレスの店舗が、すでに1つのフォームにまとまっているので、そのフォームに、ルアナ渋谷のページ（1/4の設問だけ）を足す。
// 「リスト」にも、ルアナ渋谷の行を足す（状況は送信済、1/4の「○」あり、フォームは今あるフォーム）。
//   1. previewLuana        … 計画を、ログに出す（何も変更しない）
//   2. applyLuana          … 「リスト」に行を足し、今あるフォームに、ページと1/4の設問を足す（スクリプトプロパティ LUANA_MODE が on のときだけ動く）
//   3. sendTestLuanaMail   … メールを、INVITE_TEST_TO のアドレスだけに送る
//   4. startLuanaMail      … お客様に、メールを1通送る（スクリプトプロパティ LUANA_MAIL_MODE が on のときだけ動く）
//   5. 終わったら、LUANA_MODE と LUANA_MAIL_MODE を削除する

const LUANA_STORE = "ルアナ渋谷";
const LUANA_MODE_PROPERTY = "LUANA_MODE";
const LUANA_MAIL_MODE_PROPERTY = "LUANA_MAIL_MODE";
const LUANA_LOG_SHEET_NAME = "ルアナ渋谷メール送信記録";

function luanaInfo_() {
  const ss = invGetSpreadsheet();
  const iv = ss.getSheetByName("業者影響確認").getDataRange().getDisplayValues();
  const h = iv[0];
  const c = n => h.indexOf(n);
  const found = iv.slice(1).find(r => extraNorm_(r[c("店舗名")]) === extraNorm_(LUANA_STORE));
  if (!found) throw new Error(`「業者影響確認」に「${LUANA_STORE}」がありません。`);
  const store = found[c("店舗名")], customer = found[c("契約者")], email = String(found[c("送信先アドレス")] || "").trim(), address = found[c("住所")];
  const list = ss.getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  let existing = null, vendor = "";
  for (let i = 1; i < list.length; i++) {
    const r = list[i];
    if (!existing && r[9] && r[10] === "送信済" && String(r[1]).trim().toLowerCase() === email.toLowerCase() && extraNorm_(r[16]) === extraNorm_(customer)) existing = { formId: r[9], url: r[8] };
    if (!vendor && String(r[2]).indexOf("野口興業") !== -1) vendor = r[2];
  }
  const rowAlready = list.some(r => extraNorm_(r[0]) === extraNorm_(store) && extraNorm_(r[16]) === extraNorm_(customer));
  return { store: store, customer: customer, email: email, address: address, existing: existing, vendor: vendor, rowAlready: rowAlready, list: list };
}

// 1. 計画の確認（何も変更しない）
function previewLuana() {
  const info = luanaInfo_();
  let responses = "確認できず";
  try { if (info.existing) responses = FormApp.openById(info.existing.formId).getResponses().length; } catch (e) { /* 確認できず */ }
  console.log(`【確認のみ・何も変更していません】店舗: ${info.store} / 契約者: ${info.customer} / 住所あり: ${info.address ? "はい" : "いいえ"} / 今あるフォーム: ${info.existing ? "あり（末尾 " + info.existing.formId.slice(-8) + "、回答 " + responses + " 件）" : "なし"} / 業者名: ${info.vendor || "（見つからない）"} / 「リスト」にすでに行がある: ${info.rowAlready ? "はい" : "いいえ"}`);
}

// 2. 「リスト」に行を足し、今あるフォームに、ページと1/4の設問を足す
function applyLuana() {
  if (String(PropertiesService.getScriptProperties().getProperty(LUANA_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${LUANA_MODE_PROPERTY} が on ではないため、変更しません。`);
    return;
  }
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) { console.log("別の処理が動いているためスキップしました。少し待って、もう一度実行してください。"); return; }
  try {
    const info = luanaInfo_();
    if (!info.existing) throw new Error("今あるフォームが見つかりません。");
    if (!info.vendor) throw new Error("業者名が見つかりません。");
    const form = FormApp.openById(info.existing.formId);
    const sectionTitle = `【${info.store}】の収集希望について`;
    const hasSection = form.getItems().some(item => item.getType() === FormApp.ItemType.SECTION_HEADER && item.getTitle() === sectionTitle);
    const fee = FORM_VENDOR_FEES[info.vendor] || FORM_DEFAULT_FEE;
    if (!hasSection) {
      form.addPageBreakItem().setTitle("希望日の選択（新しい店舗）");
      form.addSectionHeaderItem().setTitle(sectionTitle).setHelpText(`ご住所：${String(info.address || "").trim() || "（登録がありません）"}\n特別回収料金：${fee}/1日`);
      const desc = form.getDescription();
      const m = desc.match(/【対象店舗】\n([^\n]*)/);
      if (m && m[1].indexOf(info.store) === -1) form.setDescription(desc.replace(m[0], () => m[0] + "・" + info.store));
      const o = addExtraDateQuestion_(info.existing.formId, info.store, EXTRA_DATE_HEADERS[1], info.vendor, false, new Set());
      if (o.indexOf("追加した") !== 0) throw new Error(`1/4の設問を足せませんでした: ${o}`);
    }
    if (!info.rowAlready) {
      const row = new Array(20).fill("");
      row[0] = info.store; row[1] = info.email; row[2] = info.vendor;
      row[8] = info.existing.url; row[9] = info.existing.formId; row[10] = "送信済";
      row[16] = info.customer; row[17] = info.address;
      row[EXTRA_DATE_COLS[1]] = "○";
      const sheet = invGetSpreadsheet().getSheetByName(INV_LIST_SHEET_NAME);
      sheet.getRange(sheet.getLastRow() + 1, 1, 1, 20).setValues([row]);
    }
    console.log(`${info.store}: ${hasSection ? "フォームには、すでにページがありました" : "フォームに、ページと1/4の設問を足しました"}。「リスト」の行: ${info.rowAlready ? "すでにありました" : "足しました"}。`);
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

function luanaMailBuild_(info, resolved) {
  const fee = FORM_VENDOR_FEES[info.vendor] || FORM_DEFAULT_FEE;
  const lines = [];
  lines.push("ご担当者様", "");
  lines.push("平素より大変お世話になっております。株式会社クリメンです。");
  lines.push(`先日、「${info.store}」について、年末年始（${INV_PERIOD_TEXT}）の期間中は回収がない旨のお知らせをお送りしました。`);
  lines.push("その後、収集業者の営業日程を確認し、1月4日(月)の収集について、あらためてご案内いたします。", "");
  if (info.customer) lines.push("【ご契約者】", info.customer, "");
  lines.push("【対象店舗】", `・${info.store}`, "");
  lines.push("■1月4日(月)の収集について");
  lines.push("通常の収集日である1月4日(月)は、収集業者の年末年始の営業の都合により、通常の収集ができません。");
  lines.push(`この日に収集をご希望の場合は、有料の特別回収（特別回収料金：${fee}/1日）となります。`);
  lines.push("※ご回答がない場合、1月4日(月)の収集は行いません。");
  lines.push(`※${INV_PERIOD_TEXT}の期間中は、先日お知らせしたとおり、回収がございません（不燃ごみも回収いたしません）。`, "");
  lines.push("■回収対象：可燃ごみのみ");
  lines.push("■回収時間：当日の早朝から午前中");
  lines.push(`■ご回答期限：${INV_DEADLINE_TEXT}まで`, "");
  lines.push(`「${info.store}」のページを、以前お送りしたアンケートと同じフォームに、追加しました。`);
  if (resolved.answered) {
    lines.push("下のリンクには、他の店舗の、前回のご回答が入力済みです。フォームの最後にある「" + info.store + "」のページだけ、1月4日(月)の収集を「希望する」「希望しない」からお選びいただき、送信してください。");
  } else {
    lines.push("下記のアンケートフォームの、最後にある「" + info.store + "」のページに、ご回答をお願いいたします。以前のアンケートに、まだご回答いただいていない場合は、すべての店舗のページに、ご回答ください。");
  }
  lines.push(resolved.url, "");
  lines.push("【ご回答について】");
  lines.push("・ご回答のあと、回答内容の控えをこのメールアドレスにお送りします。");
  lines.push("・回答内容を変更する場合は、同じフォームから再度ご回答ください（最新のご回答を有効とします）。");
  lines.push(`　変更のご希望は${INV_DEADLINE_TEXT}までにお願いいたします。`, "");
  if (INV_CONTACT) lines.push(`ご不明な点は、${INV_CONTACT}までご連絡ください。`, "");
  if (INV_NOREPLY_TEXT) lines.push(INV_NOREPLY_TEXT, "");
  lines.push("先日のお知らせと内容が変わり、ご迷惑をおかけし、申し訳ございません。何卒よろしくお願いいたします。", "", INV_SENDER_NAME);
  return { subject: `【株式会社クリメン】【重要】年末年始の特別回収：${info.store}の1月4日(月)の収集について（先日のお知らせの補足とご回答のお願い）`, body: lines.join("\n") };
}

// 3. テスト送信（お客様には送らない）
function sendTestLuanaMail() {
  const testTo = String(PropertiesService.getScriptProperties().getProperty(INV_TEST_TO_PROPERTY) || "").trim();
  if (!testTo) throw new Error(`スクリプトプロパティ ${INV_TEST_TO_PROPERTY} にテスト用のメールアドレスを設定してください。`);
  const info = luanaInfo_();
  if (!info.existing) throw new Error("今あるフォームが見つかりません。");
  const message = luanaMailBuild_(info, exmailResolve_({ formId: info.existing.formId }));
  invSend(testTo, "【テスト】" + message.subject, `（テスト送信です。本来の宛先: ${info.email}）\n\n` + message.body, invGetFrom());
  console.log(`テストメールを ${testTo} に送りました。`);
}

// 4. お客様に、メールを1通送る
function startLuanaMail() {
  if (String(PropertiesService.getScriptProperties().getProperty(LUANA_MAIL_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${LUANA_MAIL_MODE_PROPERTY} が on ではないため、送信しません。`);
    return;
  }
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) { console.log("別の処理が動いているためスキップしました。少し待って、もう一度実行してください。"); return; }
  try {
    const ss = invGetSpreadsheet();
    let log = ss.getSheetByName(LUANA_LOG_SHEET_NAME);
    if (!log) { log = ss.insertSheet(LUANA_LOG_SHEET_NAME); log.appendRow(["キー", "送信先", "対象店舗", "結果", "日時"]); }
    const info = luanaInfo_();
    if (!info.existing) throw new Error("今あるフォームが見つかりません。");
    if (!info.rowAlready) throw new Error("先に applyLuana を実行してください。");
    const key = info.existing.formId + "|" + info.store;
    if (log.getLastRow() > 1 && log.getRange(2, 1, log.getLastRow() - 1, 1).getDisplayValues().some(r => r[0] === key)) { console.log("すでに送信済みです（記録あり）。"); return; }
    if (!/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(info.email)) throw new Error("宛先が正しくありません。");
    let result = INV_RESULT_SENT;
    try {
      const message = luanaMailBuild_(info, exmailResolve_({ formId: info.existing.formId }));
      invSend(info.email, message.subject, message.body, invGetFrom());
    } catch (e) {
      console.error(`メールを送れませんでした: ${e}`);
      result = INV_RESULT_FAILED;
    }
    log.appendRow([key, info.email, info.store, result, invFormatNow()]);
    console.log(`ルアナ渋谷のメール: ${result}。記録は「${LUANA_LOG_SHEET_NAME}」シートにあります。`);
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

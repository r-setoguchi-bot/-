// ===== あなたが個別に対応していた2グループの、日程が変わった店舗（東小金井店・筑前屋 王子店）を、フォームに足して、お知らせする =====
//   ・それゆけ！鶏ヤロー！ 東小金井店 … 12/29(火)が、通常の収集ができず、特別回収（光栄和・8,000円/1日）
//   ・筑前屋 王子店 … 1/4(月)が、通常の収集ができず、特別回収（野口興業・7,000円/1日）
// 手順:
//   1. prepareIndiv           … 「リスト」の、この2店舗の行に、S列・T列の「○」を入れる
//   2. previewExtraDates      … （追加日程フォーム.js）フォームに足す内容を、確認する（フォームは変更しない）
//   3. applyExtraDates        … （追加日程フォーム.js。スクリプトプロパティ EXTRA_DATE_MODE が on のとき）今あるフォームに、設問を足す
//   4. previewIndivMails      … メールの対象を、ログに出す（送らない）
//   5. sendTestIndivMails     … メールを、INVITE_TEST_TO のアドレスだけに送る
//   6. startIndivMails        … お客様に、メールを送る（スクリプトプロパティ INDIV_MAIL_MODE が on のときだけ動く）
//   7. 終わったら、EXTRA_DATE_MODE と INDIV_MAIL_MODE を削除する

const INDIV_STORES = [
  { store: "それゆけ！鶏ヤロー！　東小金井店", customer: "株式会社鶏ヤロー", dateIndex: 0 },
  { store: "筑前屋　王子店", customer: "ＤｉｎｅＶｉｔａ　Ｇｒｏｕｐ株式会社", dateIndex: 1 }
];
const INDIV_MAIL_MODE_PROPERTY = "INDIV_MAIL_MODE";
const INDIV_LOG_SHEET_NAME = "個別対応2グループ メール送信記録";

function indivRows_() {
  const list = invGetSpreadsheet().getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  return INDIV_STORES.map(s => {
    let found = null;
    for (let i = 1; i < list.length; i++) {
      if (extraNorm_(list[i][0]) === extraNorm_(s.store) && extraNorm_(list[i][16]) === extraNorm_(s.customer)) { found = { rowNumber: i + 1, row: list[i] }; break; }
    }
    if (!found) throw new Error(`「リスト」に見つかりません: ${s.store}`);
    return { def: s, rowNumber: found.rowNumber, row: found.row };
  });
}

// 1. 「リスト」の、この2店舗の行に、「○」を入れる
function prepareIndiv() {
  const sheet = invGetSpreadsheet().getSheetByName(INV_LIST_SHEET_NAME);
  const lines = [];
  indivRows_().forEach(x => {
    if (!x.row[9] || x.row[10] !== "送信済") throw new Error(`フォームがない、または状況が送信済ではありません: ${x.def.store}`);
    const col = EXTRA_DATE_COLS[x.def.dateIndex];
    const cell = sheet.getRange(x.rowNumber, col + 1);
    if (cell.getDisplayValue() !== "○") cell.setValue("○");
    let responses = "確認できず";
    try { responses = FormApp.openById(x.row[9]).getResponses().length; } catch (e) { /* 確認できず */ }
    lines.push(`${x.def.store}（${x.rowNumber}行目）: ${EXTRA_DATE_HEADERS[x.def.dateIndex]} の「○」を入れました。フォームの回答: ${responses} 件。`);
  });
  lines.forEach(l => console.log(l));
}

function indivTargets_() {
  return indivRows_().map(x => {
    const header = EXTRA_DATE_HEADERS[x.def.dateIndex];
    return {
      stores: [x.row[0]], dates: [header], fees: [exmailFee_(header, x.row[2])],
      customer: x.row[16], email: String(x.row[1] || "").trim(), formId: x.row[9], key: x.row[9] + "|" + x.row[0]
    };
  });
}

// 4. メールの対象を、ログに出す（送らない）
function previewIndivMails() {
  indivTargets_().forEach(t => {
    let responses = "確認できず";
    try { responses = FormApp.openById(t.formId).getResponses().length; } catch (e) { /* 確認できず */ }
    console.log(`${t.stores[0]} / ${t.dates[0]} / ${t.fees[0]} / 契約者: ${t.customer} / 宛先の形: ${/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(t.email) ? "正しい" : "正しくない"} / フォームの回答: ${responses} 件`);
  });
}

// 5. テスト送信（お客様には送らない）
function sendTestIndivMails() {
  const testTo = String(PropertiesService.getScriptProperties().getProperty(INV_TEST_TO_PROPERTY) || "").trim();
  if (!testTo) throw new Error(`スクリプトプロパティ ${INV_TEST_TO_PROPERTY} にテスト用のメールアドレスを設定してください。`);
  const from = invGetFrom();
  indivTargets_().forEach(t => {
    const message = exmailBuildMail_(t, exmailResolve_(t));
    invSend(testTo, "【テスト】" + message.subject, `（テスト送信です。本来の宛先: ${t.email}）\n\n` + message.body, from);
    Utilities.sleep(INV_GRAPH_INTERVAL_MS);
  });
  console.log(`テストメール 2 通を ${testTo} に送りました。`);
}

// 6. お客様に、メールを送る
function startIndivMails() {
  if (String(PropertiesService.getScriptProperties().getProperty(INDIV_MAIL_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${INDIV_MAIL_MODE_PROPERTY} が on ではないため、送信しません。`);
    return;
  }
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) { console.log("別の処理が動いているためスキップしました。少し待って、もう一度実行してください。"); return; }
  let sent = 0, failed = 0;
  try {
    const ss = invGetSpreadsheet();
    let log = ss.getSheetByName(INDIV_LOG_SHEET_NAME);
    if (!log) { log = ss.insertSheet(INDIV_LOG_SHEET_NAME); log.appendRow(["キー", "送信先", "対象店舗", "結果", "日時"]); }
    const logged = new Set();
    if (log.getLastRow() > 1) log.getRange(2, 1, log.getLastRow() - 1, 1).getDisplayValues().forEach(r => logged.add(r[0]));
    const from = invGetFrom();
    indivTargets_().filter(t => !logged.has(t.key)).forEach(t => {
      let result = INV_RESULT_SENT;
      if (!/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(t.email)) {
        console.error(`宛先が正しくありません: ${t.stores[0]}`);
        result = INV_RESULT_FAILED;
      } else {
        try {
          const message = exmailBuildMail_(t, exmailResolve_(t));
          invSend(t.email, message.subject, message.body, from);
        } catch (e) {
          console.error(`メールを送れませんでした: ${t.stores[0]} / ${e}`);
          result = INV_RESULT_FAILED;
        }
      }
      log.appendRow([t.key, t.email, t.stores[0], result, invFormatNow()]);
      SpreadsheetApp.flush();
      if (result === INV_RESULT_SENT) { sent++; Utilities.sleep(INV_GRAPH_INTERVAL_MS); } else { failed++; }
    });
    console.log(`今回 ${sent} 通送信（失敗 ${failed} 件）。記録は「${INDIV_LOG_SHEET_NAME}」シートにあります。`);
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

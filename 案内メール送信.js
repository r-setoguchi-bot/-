// ===== フォームの案内メールを、顧客ごとに1通ずつ送る =====
// 「リスト」シートで作成済み（フォーム作成状況が「送信済」）のフォームごとに、案内メールを1通送る。
// 宛先は「リスト」のB列のアドレス。同じフォームの店舗は、1通にまとめて書く。
// 1回の実行は INV_TIME_LIMIT_MS で打ち切り、残りがあれば時間主導トリガーで自動的に続きを実行する。
// 送った結果は「案内メール送信記録」シートに残り、同じフォームには二重に送らない（「リスト」を作り直しても記録は残る）。
//
// 送信元:
//   スクリプトプロパティ MAIL_FROM に、会社のアドレスを入れる。
//   そのアドレスが、実行するアカウントのGmailの「別のアドレスから送信」に登録されていること（回答の控えメールも同じアドレスから送る）。
//   未設定なら、スクリプトを実行しているアカウントのアドレスから送る。
//
// 安全のしくみ（お客様に誤って送らないため）:
//   ・スクリプトプロパティ INVITE_MODE が on のときだけ本番送信する（未設定は送らない）。
//   ・INV_NOT_BEFORE より前は、本番送信を行わない。
//   ・テスト送信（sendTestInvitations）は、INVITE_TEST_TO のアドレスだけに送り、記録には残さない。
//
// 使い方:
//   1. スクリプトプロパティ MAIL_FROM（会社のアドレス）、INVITE_TEST_TO（テスト用の自分のアドレス）を設定
//   2. sendTestInvitations を実行 → 先頭の数件分の案内を、テスト用のアドレスに送る（文面とURLを確認する）
//   3. 当日、スクリプトプロパティ INVITE_MODE を on にして startInvitations を実行（終わるまで自動で続く）
//   4. 途中で止めるときは stopInvitations
//   5. 送れなかった行（「失敗」）は、記録シートのその行を消してから、もう一度 startInvitations を実行すると送り直す
//
// スプレッドシートの指定は フォーム作成.js と同じ（FORM_SPREADSHEET_URL、無ければ紐づいているスプレッドシート）。
// このファイルだけで動く。

const INV_LIST_SHEET_NAME = "リスト";
const INV_LOG_SHEET_NAME = "案内メール送信記録";
const INV_STATUS_SENT = "送信済";           // 「リスト」のフォーム作成状況
const INV_MODE_PROPERTY = "INVITE_MODE";     // on のときだけ本番送信
const INV_TEST_TO_PROPERTY = "INVITE_TEST_TO";
const INV_FROM_PROPERTY = "MAIL_FROM";

const INV_BATCH_HANDLER = "continueInvitations";
const INV_TIME_LIMIT_MS = 4.5 * 60 * 1000;
const INV_RETRY_DELAY_MS = 60 * 1000;          // 続きを実行するまでの待ち時間
const INV_QUOTA_WAIT_MS = 3 * 60 * 60 * 1000;  // 1日の送信上限に達したときに、再開するまでの待ち時間
const INV_TEST_COUNT = 3;                      // テスト送信で送る件数

// この日時より前は、本番送信しない（誤って早く送らないための安全装置）
const INV_NOT_BEFORE = "2026/10/15 09:00";

// 「リスト」シートの列（0始まり）
const INV_COL_STORE = 0;     // A列: 店舗名
const INV_COL_EMAIL = 1;     // B列: メールアドレス
const INV_COL_FORM_URL = 8;  // I列: フォームURL
const INV_COL_FORM_ID = 9;   // J列: フォームID
const INV_COL_STATUS = 10;   // K列: フォーム作成状況

// 案内メールの文面
const INV_SENDER_NAME = "株式会社クリメン";
const INV_SUBJECT = "【株式会社クリメン】年末年始（12/30〜1/3）廃棄物回収のご希望に関するアンケートのお願い【11/15までにご回答ください】";
const INV_PERIOD_TEXT = "2026年12月30日(水)から2027年1月3日(日)";
const INV_DEADLINE_TEXT = "2026年11月15日(日)";   // フォーム作成.js の FORM_DEADLINE_TEXT と同じにしておく
const INV_CONTACT = "";                           // 問い合わせ先（電話番号・メールなど）。空のままなら案内の行を入れない
const INV_REPLY_TO = "";                          // 返信先のアドレス。空のままなら、送信元に返信される

const INV_LOG_HEADER = ["フォームID", "送信先", "対象店舗", "結果", "日時"];
const INV_RESULT_SENT = "送信済";
const INV_RESULT_FAILED = "失敗";

function invGetSpreadsheet() {
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

function invFormatNow() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
}

// 送信元アドレスを確認する。未設定なら空文字（実行アカウントのアドレスから送る）。使えないアドレスなら、送る前にエラーで止める
function invGetFrom() {
  const from = String(PropertiesService.getScriptProperties().getProperty(INV_FROM_PROPERTY) || "").trim();
  if (!from) return "";
  const own = String(Session.getEffectiveUser().getEmail() || "").toLowerCase();
  const aliases = GmailApp.getAliases().map(alias => String(alias).toLowerCase());
  if (from.toLowerCase() !== own && aliases.indexOf(from.toLowerCase()) === -1) {
    throw new Error(`${INV_FROM_PROPERTY}（${from}）は、このアカウントから送信できるアドレスではありません。` +
      "Gmailの設定「アカウント」→「他のメールアドレスを追加」で登録してください。");
  }
  return from;
}

function invSend(to, subject, body, from) {
  const options = { name: INV_SENDER_NAME };
  if (from) options.from = from;
  if (INV_REPLY_TO) options.replyTo = INV_REPLY_TO;
  GmailApp.sendEmail(to, subject, body, options);
}

// 「リスト」から、作成済みフォームごとの宛先と店舗を取り出す
function invGroupFormsFromList(data) {
  const formsById = {};
  const forms = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const formId = row[INV_COL_FORM_ID];
    if (!formId || row[INV_COL_STATUS] !== INV_STATUS_SENT) continue;

    if (!formsById[formId]) {
      formsById[formId] = { formId: formId, formUrl: row[INV_COL_FORM_URL], email: String(row[INV_COL_EMAIL]).trim(), stores: [] };
      forms.push(formsById[formId]);
    }
    formsById[formId].stores.push(row[INV_COL_STORE]);
  }
  return forms;
}

// 案内メールの件名と本文を作る
function invBuildMail(form) {
  const lines = [];
  lines.push("ご担当者様", "");
  lines.push("平素より大変お世話になっております。株式会社クリメンです。");
  lines.push("年末年始の特別回収につきまして、ご希望をお伺いするアンケートをお送りします。", "");
  lines.push("【対象店舗】");
  form.stores.forEach(store => lines.push(`・${store}`));
  lines.push("");
  lines.push(`■特別回収期間：${INV_PERIOD_TEXT}`);
  lines.push("　※1月1日は回収できません。");
  lines.push("■回収対象：可燃ごみのみ");
  lines.push("■特別回収料金：店舗ごとに異なります。アンケート内の各店舗の設問でご確認ください。", "");
  lines.push(`■ご回答期限：${INV_DEADLINE_TEXT}まで`, "");
  lines.push("下記のアンケートフォームから、ご回答をお願いいたします。");
  lines.push(form.formUrl, "");
  lines.push("【ご回答について】");
  lines.push("・特別回収を希望される場合は、「特別回収期間の回収を希望する（有料）」を選択し、次の画面で希望日をお選びください。");
  lines.push("・希望されない場合は、「すべての日程において回収を希望しない」を選択してください。");
  lines.push("・通常の可燃回収曜日以外は回収できません。希望日は、回収に伺う日付でお選びください。");
  lines.push("・ご回答のあと、回答内容の控えをこのメールアドレスにお送りします。");
  lines.push("・回答内容を変更する場合は、同じフォームから再度ご回答ください（最新のご回答を有効とします）。");
  lines.push(`　変更のご希望は${INV_DEADLINE_TEXT}までにお願いいたします。`, "");
  if (INV_CONTACT) lines.push(`ご不明な点は、${INV_CONTACT}までご連絡ください。`, "");
  lines.push("お忙しいところ恐れ入りますが、よろしくお願いいたします。", "", INV_SENDER_NAME);
  return { subject: INV_SUBJECT, body: lines.join("\n") };
}

// テスト送信: 先頭の数件分の案内を、INVITE_TEST_TO のアドレスだけに送る（お客様には送らない。記録にも残さない）
function sendTestInvitations() {
  const props = PropertiesService.getScriptProperties();
  const testTo = String(props.getProperty(INV_TEST_TO_PROPERTY) || "").trim();
  if (!testTo) {
    throw new Error(`スクリプトプロパティ ${INV_TEST_TO_PROPERTY} にテスト用のメールアドレスを設定してください。`);
  }
  const from = invGetFrom();
  const ss = invGetSpreadsheet();
  const listSheet = ss.getSheetByName(INV_LIST_SHEET_NAME);
  if (!listSheet) throw new Error(`シート「${INV_LIST_SHEET_NAME}」が見つかりません。`);

  const forms = invGroupFormsFromList(listSheet.getDataRange().getDisplayValues());
  if (forms.length === 0) throw new Error("フォーム作成済みの店舗がありません。");

  forms.slice(0, INV_TEST_COUNT).forEach(form => {
    const message = invBuildMail(form);
    invSend(testTo, "【テスト】" + message.subject,
      `（テスト送信です。本来の宛先: ${form.email}）\n\n` + message.body, from);
  });
  console.log(`${Math.min(INV_TEST_COUNT, forms.length)} 通のテストメールを ${testTo} に送りました（送信元: ${from || "実行アカウント"}）。`);
}

function startInvitations() {
  deleteInvitationTriggers();
  runInvitationBatch();
}

function continueInvitations() {
  runInvitationBatch();
}

function stopInvitations() {
  deleteInvitationTriggers();
  console.log("案内メールの自動送信を停止しました。");
}

function runInvitationBatch() {
  const props = PropertiesService.getScriptProperties();
  if (String(props.getProperty(INV_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${INV_MODE_PROPERTY} が on ではないため、送信しません。`);
    return;
  }
  const notBefore = new Date(INV_NOT_BEFORE.replace(/-/g, "/"));
  if (new Date() < notBefore) {
    console.log(`${INV_NOT_BEFORE} より前なので、送信しません。`);
    return;
  }

  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) {
    console.log("別の送信が処理中のためスキップしました。");
    return;
  }

  const startTime = Date.now();
  let sentCount = 0;
  let failedCount = 0;
  let stoppedByQuota = false;
  let remaining = 0;

  try {
    deleteInvitationTriggers();

    const from = invGetFrom();
    const ss = invGetSpreadsheet();
    const listSheet = ss.getSheetByName(INV_LIST_SHEET_NAME);
    if (!listSheet) throw new Error(`シート「${INV_LIST_SHEET_NAME}」が見つかりません。`);
    const forms = invGroupFormsFromList(listSheet.getDataRange().getDisplayValues());

    let logSheet = ss.getSheetByName(INV_LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = ss.insertSheet(INV_LOG_SHEET_NAME);
      logSheet.appendRow(INV_LOG_HEADER);
    }
    const logged = new Set();
    if (logSheet.getLastRow() > 1) {
      logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 1).getDisplayValues().forEach(row => logged.add(row[0]));
    }

    const targets = forms.filter(form => !logged.has(form.formId));

    for (let n = 0; n < targets.length; n++) {
      if (Date.now() - startTime > INV_TIME_LIMIT_MS) {
        remaining = targets.length - n;
        break;
      }
      if (MailApp.getRemainingDailyQuota() < 1) {
        stoppedByQuota = true;
        remaining = targets.length - n;
        break;
      }

      const form = targets[n];
      const storeNames = form.stores.join("・");
      let result = INV_RESULT_SENT;
      if (!/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(form.email)) {
        console.error(`宛先が正しくありません: ${form.formId} / ${form.email}`);
        result = INV_RESULT_FAILED;
      } else {
        try {
          const message = invBuildMail(form);
          invSend(form.email, message.subject, message.body, from);
        } catch (e) {
          console.error(`案内メールを送れませんでした: ${form.formId} / ${form.email} / ${e}`);
          result = INV_RESULT_FAILED;
        }
      }
      // 1通ごとに記録して、途中で止まっても二重に送らないようにする
      logSheet.appendRow([form.formId, form.email, storeNames, result, invFormatNow()]);
      SpreadsheetApp.flush();
      if (result === INV_RESULT_SENT) sentCount++;
      else failedCount++;
    }

    if (remaining > 0) {
      ScriptApp.newTrigger(INV_BATCH_HANDLER).timeBased()
        .after(stoppedByQuota ? INV_QUOTA_WAIT_MS : INV_RETRY_DELAY_MS).create();
      console.log(`今回 ${sentCount} 通送信（失敗 ${failedCount} 件）。残り ${remaining} 件は` +
        (stoppedByQuota ? "1日の送信上限に達したため、しばらく待ってから" : "") + "自動で続行します。");
    } else {
      console.log(`今回 ${sentCount} 通送信（失敗 ${failedCount} 件）。すべて完了しました。送信記録は「${INV_LOG_SHEET_NAME}」シートにあります。`);
    }
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      console.warn(`ロック解除に失敗しました（処理には影響しません）: ${e}`);
    }
  }
}

function deleteInvitationTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === INV_BATCH_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

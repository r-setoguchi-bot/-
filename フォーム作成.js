// ===== 年末年始回収アンケート フォーム一括作成 =====
// 「リスト」シートの未送信行をメールアドレスごとにまとめてフォームを作成する。
// 1回の実行は FORM_TIME_LIMIT_MS で打ち切り、残りがあれば時間主導トリガーで自動的に続きを実行する。
//
// 使い方:
//   1. startFormCreation を1回だけ手動実行（以降は全件完了まで自動で続行）
//   2. 途中で止めたい場合は stopFormCreation を実行
//
// スプレッドシートの指定:
//   スクリプトプロパティ FORM_SPREADSHEET_URL があればそのスプレッドシートを使う。
//   無ければスクリプトが紐づいているスプレッドシート（コンテナバインド時）を使う。

const FORM_LIST_SHEET_NAME = "リスト";
const FORM_SEND_SHEET_NAME = "メール送信用";
const FORM_BATCH_HANDLER = "createFormsInBatches";
const FORM_TIME_LIMIT_MS = 4.5 * 60 * 1000; // 1回あたりの処理時間上限（6分制限対策）
const FORM_RETRY_DELAY_MS = 60 * 1000;      // 続きを実行するまでの待ち時間
const FORM_SAFETY_DELAY_MS = 7 * 60 * 1000; // 実行が強制終了した場合に再開するまでの待ち時間（6分制限より長く）

const FORM_STATUS_SENT = "送信済";
const FORM_STATUS_ERROR = "エラー";

// 列番号（1始まり）
const FORM_COL_URL = 12;    // L列: フォームURL
const FORM_COL_STATUS = 14; // N列: 送信状況

// 希望日の列（0始まり、D〜K列）
const FORM_DATE_COL_START = 3;
const FORM_DATE_COL_END = 10;

const FORM_TITLE = '【株式会社クリメン】2026年～2027年 年末年始廃棄物回収に関するアンケート';
const FORM_PERIOD_TEXT = '2026年12月28日(月)から2027年1月4日(月)';

const FORM_CHOICE_WANT = '特別回収期間の回収を希望する（有料）';
const FORM_CHOICE_NOT_WANT = 'すべての日程において回収を希望しない';

const FORM_DEFAULT_FEE = "7,000円";
const FORM_VENDOR_FEES = {
  "有限会社長澤商事": "10,000円",
  "株式会社光栄和": "20,000円",
  "鍵本産業株式会社": "10,000円"
};

function startFormCreation() {
  deleteFormBatchTriggers();
  createFormsInBatches();
}

function stopFormCreation() {
  deleteFormBatchTriggers();
  console.log("フォーム自動作成を停止しました。");
}

function createFormsInBatches() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10 * 1000)) {
    console.log("別の実行が処理中のためスキップしました。");
    return;
  }

  const startTime = Date.now();
  let formCount = 0;
  let errorCount = 0;
  let remaining = 0;

  try {
    // 先に保険の再開トリガーを予約しておく。6分制限などで強制終了しても、これで処理が再開する
    deleteFormBatchTriggers();
    scheduleFormBatch(FORM_SAFETY_DELAY_MS);

    const ss = getFormSpreadsheet();
    const sheet = ss.getSheetByName(FORM_LIST_SHEET_NAME);
    if (!sheet) {
      throw new Error(`シート「${FORM_LIST_SHEET_NAME}」が見つかりません。`);
    }
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data[0];
    const sendSheet = getFormSendSheet(ss);

    const groupedData = groupUnsentRowsByEmail(data);
    const emails = Object.keys(groupedData);

    for (let n = 0; n < emails.length; n++) {
      if (Date.now() - startTime > FORM_TIME_LIMIT_MS) {
        remaining = emails.length - n;
        break;
      }

      const email = emails[n];
      const stores = groupedData[email];

      try {
        const result = createFormForStores(stores, headers);

        // 1件ごとに書き込み、途中で止まっても同じフォームを二重に作らないようにする
        stores.forEach(store => {
          sheet.getRange(store.rowNumber, FORM_COL_URL, 1, 3)
               .setValues([[result.formUrl, result.formId, FORM_STATUS_SENT]]);
        });
        sendSheet.appendRow([email, result.storeNamesText, result.formUrl]);
        SpreadsheetApp.flush();
        formCount++;
      } catch (e) {
        // エラー行は対象外にして、同じ行で失敗し続けて止まらなくなるのを防ぐ
        console.error(`フォーム作成に失敗しました: ${email} / ${e}`);
        stores.forEach(store => {
          sheet.getRange(store.rowNumber, FORM_COL_STATUS).setValue(FORM_STATUS_ERROR);
        });
        SpreadsheetApp.flush();
        errorCount++;
      }
    }

    deleteFormBatchTriggers();
    if (remaining > 0) {
      scheduleFormBatch(FORM_RETRY_DELAY_MS);
      console.log(`今回 ${formCount} 件作成（エラー ${errorCount} 件）。残り ${remaining} 件は自動で続行します。`);
    } else {
      console.log(`今回 ${formCount} 件作成（エラー ${errorCount} 件）。すべて完了しました。`);
      refreshContractStatusList();
    }
  } finally {
    // ロック解除時に一時的なサーバーエラーが出ることがある。実行終了時に自動で解除されるので失敗扱いにしない
    try {
      lock.releaseLock();
    } catch (e) {
      console.warn(`ロック解除に失敗しました（処理には影響しません）: ${e}`);
    }
  }
}

// kintoneリスト作成.js があれば、全件完了時に「契約状況一覧」を最新にする（失敗してもフォーム作成には影響しない）
function refreshContractStatusList() {
  try {
    if (typeof updateContractStatusList === "function") {
      updateContractStatusList();
    }
  } catch (e) {
    console.warn(`契約状況一覧の更新に失敗しました: ${e}`);
  }
}

function getFormSpreadsheet() {
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

function getFormSendSheet(ss) {
  let sendSheet = ss.getSheetByName(FORM_SEND_SHEET_NAME);
  if (!sendSheet) {
    sendSheet = ss.insertSheet(FORM_SEND_SHEET_NAME);
    sendSheet.appendRow(["メールアドレス", "対象店舗", "フォームURL"]);
  }
  return sendSheet;
}

function groupUnsentRowsByEmail(data) {
  const groupedData = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const email = row[1].trim();               // B列
    const status = row[FORM_COL_STATUS - 1];   // N列
    if (status === FORM_STATUS_SENT || status === FORM_STATUS_ERROR || !email) continue;

    if (!groupedData[email]) {
      groupedData[email] = [];
    }
    groupedData[email].push({
      storeName: row[0], // A列
      vendor: row[2],    // C列
      rowNumber: i + 1,
      rowData: row
    });
  }
  return groupedData;
}

function createFormForStores(stores, headers) {
  const storeNamesText = stores.map(store => store.storeName).join('・');

  const form = FormApp.create(FORM_TITLE);

  try {
    form.setDescription(`【対象店舗】
${storeNamesText}

特別回収期間：${FORM_PERIOD_TEXT}
※1月1日は回収不可

★特別回収料金：店舗ごとに異なります。各店舗の設問をご確認ください。
★回収対象：可燃ごみのみ

◆ 特別回収期間の回収を希望 される場合（※）は
「${FORM_CHOICE_WANT}」を選択し
次の画面にて希望日にチェックをしてください 。
 ※通常、伺っております可燃回収曜日に準じます。予めご了承ください。
 ※通常回収日以外にチェックをされましても回収は出来かねますので、予めご留意ください。

◆希望日とは回収に伺う日付です。
 （例：12月29日営業分のゴミは、12月30日が回収日となります。）

◆特別回収期間の回収を希望されない場合「${FORM_CHOICE_NOT_WANT}」を選択してください。`);

    // 1ページ目：希望の有無。「希望しない」はそのまま送信、「希望する」は次ページの希望日選択へ進む
    const wantItem = form.addMultipleChoiceItem()
        .setTitle('特別回収期間の回収について')
        .setRequired(true);

    const datePage = form.addPageBreakItem().setTitle('希望日の選択');

    wantItem.setChoices([
      wantItem.createChoice(FORM_CHOICE_WANT, datePage),
      wantItem.createChoice(FORM_CHOICE_NOT_WANT, FormApp.PageNavigationType.SUBMIT)
    ]);

    // 2ページ目：店舗ごとの希望日
    stores.forEach(store => {
      const fee = FORM_VENDOR_FEES[store.vendor] || FORM_DEFAULT_FEE;

      form.addSectionHeaderItem()
          .setTitle(`【${store.storeName}】の収集希望について`)
          .setHelpText(`特別回収料金：${fee}/1日`);

      for (let col = FORM_DATE_COL_START; col <= FORM_DATE_COL_END; col++) {
        if (store.rowData[col] === "○") {
          form.addMultipleChoiceItem()
              .setTitle(`【${store.storeName}】${headers[col]}の収集を希望しますか？`)
              .setChoiceValues(['希望する', '希望しない'])
              .setRequired(true);
        }
      }
    });

    return {
      formUrl: form.getPublishedUrl(),
      formId: form.getId(),
      storeNamesText: storeNamesText
    };
  } catch (e) {
    // 作りかけのフォームがドライブに残らないようゴミ箱へ移す
    try {
      DriveApp.getFileById(form.getId()).setTrashed(true);
    } catch (trashError) {
      console.error(`作りかけのフォームを削除できませんでした: ${form.getId()} / ${trashError}`);
    }
    throw e;
  }
}

function scheduleFormBatch(delayMs) {
  ScriptApp.newTrigger(FORM_BATCH_HANDLER).timeBased().after(delayMs).create();
}

// このフォーム作成処理用のトリガーだけを削除する（リマインド・日報などのトリガーは残す）
function deleteFormBatchTriggers() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === FORM_BATCH_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

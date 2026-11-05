// ===== 「一石三鳥　韓邸」を、単独のフォームにして、10/15の回答を引き継ぐ（お客様には、何も送らない）=====
//   1. splitKantei        … 単独のフォームを作り、今の回答を、そのまま新しいフォームに入れ、「リスト」と「回答集計」を付け替える
//                           （スクリプトプロパティ KANTEI_MODE が on のときだけ動く）
//   2. trimOldForKantei   … 引き継ぎを確認したあと、古いフォーム（Bistro…と共有）から、韓邸のページを消す
//   3. 終わったら、KANTEI_MODE を削除する

const KANTEI_MODE_PROPERTY = "KANTEI_MODE";
const KANTEI_STORE = "一石三鳥　韓邸";

function kanteiFind_() {
  const data = invGetSpreadsheet().getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    if (extraNorm_(data[i][0]) === extraNorm_(KANTEI_STORE) && extraNorm_(data[i][16]) === extraNorm_(HQ_CUSTOMER) && data[i][10] === "送信済") {
      return { rowNumber: i + 1, row: data[i], header: data[0] };
    }
  }
  throw new Error("「リスト」に、韓邸の行が見つかりません。");
}

function splitKantei() {
  if (String(PropertiesService.getScriptProperties().getProperty(KANTEI_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${KANTEI_MODE_PROPERTY} が on ではないため、変更しません。`);
    return;
  }
  const lock = LockService.getUserLock();   // 集計（控えメールを送る処理）と同じロック。持っている間は、集計が動かない
  if (!lock.tryLock(30 * 1000)) { console.log("別の処理が動いているためスキップしました。少し待って、もう一度実行してください。"); return; }
  try {
    const ss = invGetSpreadsheet();
    const listSheet = ss.getSheetByName(INV_LIST_SHEET_NAME);
    const t = kanteiFind_();
    const oldId = t.row[9];
    const oldForm = FormApp.openById(oldId);
    const oldResponses = oldForm.getResponses();
    if (oldResponses.length === 0) throw new Error("今のフォームに、回答がありません。変更しません。");
    const sharedStores = oldForm.getItems().filter(i => i.getType() === FormApp.ItemType.SECTION_HEADER && /^【.+】の収集希望について$/.test(i.getTitle())).length;
    if (sharedStores < 2) { console.log("すでに単独のフォームです。変更しません。"); return; }

    // 最新の回答を取る
    const latest = oldResponses.slice().sort((a, b) => a.getTimestamp() - b.getTimestamp()).pop();
    const prefix = `【${KANTEI_STORE}】`;
    const answers = latest.getItemResponses().filter(r => {
      const title = r.getItem().getTitle();
      return title === FORM_Q_CONTACT_NAME || title === FORM_Q_PHONE || title === FORM_Q_COMMENT || title.indexOf(prefix) === 0;
    });

    // 単独のフォームを作る
    const store = { storeName: t.row[0], vendor: t.row[2], rowNumber: t.rowNumber, rowData: t.row };
    const result = createFormWithRetry([store], t.header, getFormFolder(), String(t.row[16]).trim());
    const newId = result.formId;
    try { const f = FormApp.openById(newId); if (f.requiresLogin()) f.setRequireLogin(false); } catch (e) { /* 確認できないことがある */ }
    openAccessRetry_(() => openAccessOne_(newId));
    const newForm = FormApp.openById(newId);

    // 回答を、新しいフォームに入れる（質問の題名が同じものに入れる）
    const newItems = {};
    newForm.getItems().forEach(i => { newItems[i.getTitle()] = i; });
    const newResponse = newForm.createResponse();
    let put = 0;
    answers.forEach(r => {
      const title = r.getItem().getTitle();
      const value = r.getResponse();
      if (value === "" || value === null || (Array.isArray(value) && value.length === 0)) return;
      const item = newItems[title];
      if (!item) throw new Error(`新しいフォームに、同じ題名の質問がありません: ${title.slice(0, 20)}`);
      let itemResponse;
      switch (item.getType()) {
        case FormApp.ItemType.TEXT: itemResponse = item.asTextItem().createResponse(String(value)); break;
        case FormApp.ItemType.PARAGRAPH_TEXT: itemResponse = item.asParagraphTextItem().createResponse(String(value)); break;
        case FormApp.ItemType.MULTIPLE_CHOICE: itemResponse = item.asMultipleChoiceItem().createResponse(String(value)); break;
        default: throw new Error(`扱えない質問の種類です: ${item.getType()}`);
      }
      newResponse.withItemResponse(itemResponse);
      put++;
    });
    // 必須の質問に、答えが入っているかを確認してから、送る
    const needed = newForm.getItems().filter(i => {
      const type = i.getType();
      if (type === FormApp.ItemType.TEXT) return i.asTextItem().isRequired();
      if (type === FormApp.ItemType.PARAGRAPH_TEXT) return i.asParagraphTextItem().isRequired();
      if (type === FormApp.ItemType.MULTIPLE_CHOICE) return i.asMultipleChoiceItem().isRequired();
      return false;
    }).length;
    const answeredTitles = new Set(answers.filter(r => r.getResponse() !== "").map(r => r.getItem().getTitle()));
    const missing = newForm.getItems().filter(i => {
      const type = i.getType();
      const req = (type === FormApp.ItemType.TEXT && i.asTextItem().isRequired()) || (type === FormApp.ItemType.PARAGRAPH_TEXT && i.asParagraphTextItem().isRequired()) || (type === FormApp.ItemType.MULTIPLE_CHOICE && i.asMultipleChoiceItem().isRequired());
      return req && !answeredTitles.has(i.getTitle());
    });
    if (missing.length > 0) throw new Error(`必須の質問に、引き継ぐ答えがありません（${missing.length} 件）。新しいフォーム（末尾 ${newId.slice(-8)}）は作りましたが、回答は入れていません。`);
    const submitted = newResponse.submit();
    const answeredAt = Utilities.formatDate(submitted.getTimestamp(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

    // 「リスト」を、新しいフォームに付け替える
    listSheet.getRange(t.rowNumber, 9, 1, 2).setValues([[result.formUrl, newId]]);

    // 「回答集計」の、韓邸の行を更新する。控えメールの印を、新しい回答日時にして、控えメールが届かないようにする
    const out = ss.getSheetByName(RESP_OUTPUT_SHEET_NAME);
    const vals = out.getDataRange().getDisplayValues();
    const head = vals[0];
    const col = name => { const k = head.indexOf(name); if (k < 0) throw new Error(`回答集計に、見出し「${name}」がありません。`); return k + 1; };
    let rowN = 0;
    for (let i = 1; i < vals.length; i++) if (extraNorm_(vals[i][0]) === extraNorm_(KANTEI_STORE)) { rowN = i + 1; break; }
    if (rowN === 0) throw new Error("回答集計に、韓邸の行がありません。新しいフォームは作りましたが、控えメールの印は書けていません。");
    out.getRange(rowN, col("回答日時")).setValue(answeredAt);
    out.getRange(rowN, col("フォームID")).setValue(newId);
    out.getRange(rowN, col("回答回数")).setValue(1);
    out.getRange(rowN, col("再回答")).setValue("");
    out.getRange(rowN, col("控えメール")).setValue(answeredAt);
    SpreadsheetApp.flush();
    console.log(`韓邸の単独フォームを作り、回答を ${put} 項目、引き継ぎました（必須の質問 ${needed} 件）。「リスト」と「回答集計」を付け替えました。控えメールは送りません。`);
    console.log("URL: " + result.formUrl);
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

// 2. 古いフォームから、韓邸のページを消す（引き継ぎを確認してから、実行する）
function trimOldForKantei() {
  if (String(PropertiesService.getScriptProperties().getProperty(KANTEI_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${KANTEI_MODE_PROPERTY} が on ではないため、変更しません。`);
    return;
  }
  const bistro = invGetSpreadsheet().getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues()
    .find(r => extraNorm_(r[0]) === extraNorm_(HQ_KEEP_STORES[0]) && r[10] === "送信済");
  if (!bistro) throw new Error("「リスト」に、Bistro… の行が見つかりません。");
  const kantei = kanteiFind_();
  if (kantei.row[9] === bistro[9]) throw new Error("韓邸が、まだ古いフォームを指しています。先に splitKantei を実行してください。");
  const left = hqTrimOldForm_(bistro[9], [KANTEI_STORE]);
  console.log(`古いフォームから、韓邸のページを消しました（残った店舗: ${left} 件）。`);
}

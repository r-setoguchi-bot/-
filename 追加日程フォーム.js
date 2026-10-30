// ===== 業者の日程の追加（光栄和 12/29、野口興業 1/4）に伴う、フォームへの設問の追加 =====
// 回答済みのフォームも、URLを変えずに、設問を足す（今までの回答は、そのまま残る）。
// 「リスト」シートの S列（12/29）・T列（1/4）に「○」を入れた店舗が、対象。「○」は、prepareExtraDates が入れる（手で入れてもよい）。
//
// 手順（上から順に。◎は、お客様に影響する操作）:
//   1. setupExtraDateHeaders            … 「リスト」の S1・T1 に、見出しを入れる
//   2. migrateResponseSheetForExtraDates … 「回答集計」の U1・V1 に、見出しを入れる（回答集計.js の修正を入れてから）
//   3. startResponseCollection を実行して、「回答集計」の行が減らないことを確認する
//   4. テスト店舗（フォームは自分だけのもの）の T列に「○」を入れて、previewExtraDates → applyExtraDates → フォームを開いて確認
//   5. テスト店舗の「○」を消して、prepareExtraDates … 対象の店舗に「○」を入れる
//   6. previewExtraDates … 追加する内容を確認する（フォームは変更しない）
//   7. ◎ applyExtraDates … 既存のフォームに設問を足す（スクリプトプロパティ EXTRA_DATE_MODE が on のとき、だけ動く）
//   8. listExtraDatePending … 前回回答したのに、追加の日程が未回答の店舗を一覧にする

const EXTRA_DATE_HEADERS = ["12月29日(火)", "1月4日(月)"];
const EXTRA_DATE_COLS = [18, 19];                 // 「リスト」の S列・T列（0始まり）
const EXTRA_DATE_VENDOR_FEES = { "12月29日(火)": { "光栄和": "8,000円" } };   // 日付ごとの、業者別の料金（なければ、店舗の通常の特別回収料金）
const EXTRA_DATE_NOTICE_MARK = "【日程追加のお知らせ】";
const EXTRA_DATE_FORM_NOTICE = "【日程追加のお知らせ】一部の店舗で、収集業者の都合により、特別回収の日が追加されました。該当する店舗のページに「【日程追加のお知らせ】」と書いてあります。すでにご回答いただいた方も、追加された日について、お手数ですが、改めてご回答ください（前回と同じ内容を選んだうえで、追加の日もお答えください。最新のご回答を有効とします）。";
const EXTRA_DATE_MODE_PROPERTY = "EXTRA_DATE_MODE";
const EXTRA_DATE_RECORD_SHEET = "追加日程 設問追加記録";
const EXTRA_DATE_TARGET_SHEET = "追加日程の対象";
const EXTRA_DATE_PENDING_SHEET = "追加日程 回答状況";
// あなたが個別に対応するお客様（自動の対象から外す）。全角半角・空白の違いは、無視する
const EXTRA_DATE_EXCLUDE_CUSTOMERS = ["株式会社鶏ヤロー", "ＤｉｎｅＶｉｔａ　Ｇｒｏｕｐ株式会社"];
// 「業者影響確認」と「リスト」で、店名が違う店舗（業者影響確認の「店名|契約者」→ リストの店名）
const EXTRA_DATE_STORE_ALIASES = { "炭火と酒　くろや　田町店|株式会社集Ｅｎｔｅｒｐｒｉｓｅ": "くろや" };

function extraNorm_(text) { return String(text).normalize("NFKC").replace(/\s/g, ""); }

// 1. 「リスト」の S1・T1 に、見出しを入れる
function setupExtraDateHeaders() {
  const sheet = invGetSpreadsheet().getSheetByName(INV_LIST_SHEET_NAME);
  const range = sheet.getRange(1, EXTRA_DATE_COLS[0] + 1, 1, EXTRA_DATE_HEADERS.length);
  const now = range.getDisplayValues()[0];
  if (now.every((v, i) => v === EXTRA_DATE_HEADERS[i])) { console.log("「リスト」の S1・T1 は、すでに見出しが入っています。"); return; }
  if (now.some(v => v !== "")) throw new Error("「リスト」の S1・T1 に、別の内容が入っています。確認してください。");
  range.setValues([EXTRA_DATE_HEADERS]);
  console.log(`「リスト」の S1・T1 に、見出し（${EXTRA_DATE_HEADERS.join("、")}）を入れました。`);
}

// 2. 「回答集計」の U1・V1 に、見出しを入れる。今までの A〜T列の見出しが、想定どおりのときだけ
function migrateResponseSheetForExtraDates() {
  const ss = invGetSpreadsheet();
  const list = ss.getSheetByName(INV_LIST_SHEET_NAME);
  const listHead = list.getRange(1, 1, 1, EXTRA_DATE_COLS[1] + 1).getDisplayValues()[0];
  const extraHeaders = EXTRA_DATE_COLS.map(col => listHead[col]);
  if (extraHeaders.some(h => !h)) throw new Error("先に setupExtraDateHeaders を実行してください。");
  const normalHeaders = listHead.slice(RESP_COL_DATE_START, RESP_COL_DATE_END + 1);
  const oldHeader = respBuildOutputHeader(normalHeaders.concat(["", ""]));
  const sheet = ss.getSheetByName(RESP_OUTPUT_SHEET_NAME);
  const nowHeader = sheet.getRange(1, 1, 1, oldHeader.length + extraHeaders.length).getDisplayValues()[0];
  const expectedNew = oldHeader.concat(extraHeaders);
  if (nowHeader.join("\t") === expectedNew.join("\t")) { console.log("「回答集計」は、すでに追加の見出しが入っています。"); return; }
  if (nowHeader.slice(0, oldHeader.length).join("\t") !== oldHeader.join("\t")) throw new Error("「回答集計」の見出しが、想定と違います。中身を確認してください。");
  if (nowHeader.slice(oldHeader.length).some(v => v !== "")) throw new Error(`「回答集計」の ${oldHeader.length + 1} 列目以降に、別の内容が入っています。確認してください。`);
  sheet.getRange(1, oldHeader.length + 1, 1, extraHeaders.length).setValues([extraHeaders]);
  console.log(`「回答集計」の ${oldHeader.length + 1} 列目から、見出し（${extraHeaders.join("、")}）を入れました。`);
}

// 5. 対象の店舗に「○」を入れる（「業者影響確認」シートと、「リスト」から）
function prepareExtraDates() {
  const ss = invGetSpreadsheet();
  const impactSheet = ss.getSheetByName("業者影響確認");
  if (!impactSheet) throw new Error("先に listVendorImpact を実行してください。");
  const list = ss.getSheetByName(INV_LIST_SHEET_NAME);
  const listValues = list.getDataRange().getDisplayValues();
  if (listValues[0][EXTRA_DATE_COLS[0]] !== EXTRA_DATE_HEADERS[0] || listValues[0][EXTRA_DATE_COLS[1]] !== EXTRA_DATE_HEADERS[1]) throw new Error("先に setupExtraDateHeaders を実行してください。");

  const rowByKey = {};
  for (let i = 1; i < listValues.length; i++) rowByKey[extraNorm_(listValues[i][0]) + "|" + extraNorm_(listValues[i][16])] = i;

  const iv = impactSheet.getDataRange().getDisplayValues();
  const ih = iv[0];
  const col = name => ih.indexOf(name);
  const excluded = EXTRA_DATE_EXCLUDE_CUSTOMERS.map(extraNorm_);
  const report = [];
  let setCount = 0;
  iv.slice(1).forEach(row => {
    const impact = row[col("影響")];
    if (impact === "なし") return;
    const store = row[col("店舗名")], customer = row[col("契約者")];
    const dateIndex = impact.indexOf("12/29") === 0 ? 0 : 1;
    const header = EXTRA_DATE_HEADERS[dateIndex];
    if (excluded.indexOf(extraNorm_(customer)) !== -1) { report.push([store, customer, header, "", "個別対応のため、対象外"]); return; }
    const listName = EXTRA_DATE_STORE_ALIASES[store + "|" + customer] || store;
    const rowIndex = rowByKey[extraNorm_(listName) + "|" + extraNorm_(customer)];
    if (rowIndex === undefined) { report.push([store, customer, header, "", "「リスト」に行がない（フォームを新しく作る必要あり）"]); return; }
    const listRow = listValues[rowIndex];
    if (!listRow[9] || listRow[10] !== "送信済") { report.push([store, customer, header, rowIndex + 1, "フォームがない、または状況が送信済ではない"]); return; }
    const cell = list.getRange(rowIndex + 1, EXTRA_DATE_COLS[dateIndex] + 1);
    if (cell.getDisplayValue() === "○") { report.push([store, customer, header, rowIndex + 1, "すでに「○」が入っている"]); return; }
    cell.setValue("○");
    setCount++;
    report.push([store, customer, header, rowIndex + 1, "「○」を入れた"]);
  });

  let out = ss.getSheetByName(EXTRA_DATE_TARGET_SHEET);
  if (!out) out = ss.insertSheet(EXTRA_DATE_TARGET_SHEET);
  out.clear();
  out.getRange(1, 1, 1, 5).setValues([["店舗名", "契約者", "追加する日", "リストの行", "結果"]]);
  if (report.length > 0) out.getRange(2, 1, report.length, 5).setValues(report);
  console.log(`「○」を入れた店舗: ${setCount} 件。確認が必要なもの: ${report.filter(r => r[4].indexOf("入れた") === -1 && r[4].indexOf("すでに") === -1).length} 件。結果は、シート「${EXTRA_DATE_TARGET_SHEET}」にあります。`);
}

function previewExtraDates() { runExtraDateQuestions_(true); }

// 7. 既存のフォームに、設問を足す
function applyExtraDates() {
  if (String(PropertiesService.getScriptProperties().getProperty(EXTRA_DATE_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${EXTRA_DATE_MODE_PROPERTY} が on ではないため、フォームを変更しません。`);
    return;
  }
  runExtraDateQuestions_(false);
}

function runExtraDateQuestions_(dryRun) {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) { console.log("別の処理が動いているためスキップしました。少し待って、もう一度実行してください。"); return; }
  const startTime = Date.now();
  const results = [];
  try {
    const ss = invGetSpreadsheet();
    const list = ss.getSheetByName(INV_LIST_SHEET_NAME);
    const values = list.getDataRange().getDisplayValues();

    let record = ss.getSheetByName(EXTRA_DATE_RECORD_SHEET);
    if (!record && !dryRun) { record = ss.insertSheet(EXTRA_DATE_RECORD_SHEET); record.appendRow(["フォームID", "店舗名", "追加した日", "結果", "日時"]); }
    const done = new Set();
    if (record && record.getLastRow() > 1) record.getRange(2, 1, record.getLastRow() - 1, 3).getDisplayValues().forEach(r => done.add(r.join("|")));

    const noticedForms = new Set();
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      EXTRA_DATE_COLS.forEach((col, k) => {
        if (row[col] !== "○") return;
        const formId = row[9], store = row[0], header = EXTRA_DATE_HEADERS[k];
        if (!formId) { results.push([store, header, "フォームがない"]); return; }
        if (done.has([formId, store, header].join("|"))) { results.push([store, header, "追加済み（記録あり）"]); return; }
        if (Date.now() - startTime > 4.5 * 60 * 1000) { results.push([store, header, "時間切れ（もう一度実行してください）"]); return; }
        try {
          const outcome = addExtraDateQuestion_(formId, store, header, row[2], dryRun, noticedForms);
          results.push([store, header, outcome]);
          if (!dryRun && outcome.indexOf("追加した") === 0) record.appendRow([formId, store, header, outcome, invFormatNow()]);
        } catch (e) {
          results.push([store, header, `失敗: ${e}`]);
        }
      });
    }
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
  const added = results.filter(r => r[2].indexOf("追加した") === 0 || r[2].indexOf("追加できる") === 0).length;
  const failed = results.filter(r => r[2].indexOf("失敗") === 0 || r[2].indexOf("見つからない") === 0).length;
  const skipped = results.length - added - failed;
  console.log(`${dryRun ? "【確認のみ・フォームは変更していません】" : ""}対象 ${results.length} 件: ${dryRun ? "追加できる" : "追加した"} ${added} 件、すでに追加済みなど ${skipped} 件、失敗・見つからない ${failed} 件。`);
  results.filter(r => r[2].indexOf("失敗") === 0 || r[2].indexOf("見つからない") === 0 || r[2].indexOf("時間切れ") === 0).forEach(r => console.error(`${r[0]} / ${r[1]} / ${r[2]}`));
}

// 1つの店舗・1つの日の設問を、フォームの、その店舗のページの最後に足す。戻り値は、結果の文
function addExtraDateQuestion_(formId, store, header, vendor, dryRun, noticedForms) {
  const form = FormApp.openById(formId);
  const items = form.getItems();
  const sectionTitle = `【${store}】の収集希望について`;
  const questionTitle = `【${store}】${header}の収集を希望しますか？`;

  if (items.some(item => item.getTitle() === questionTitle)) return "すでに設問がある";
  const headerIndex = items.findIndex(item => item.getType() === FormApp.ItemType.SECTION_HEADER && item.getTitle() === sectionTitle);
  if (headerIndex < 0) return "見つからない: 店舗のページ";
  let last = headerIndex;
  for (let j = headerIndex + 1; j < items.length && items[j].getType() !== FormApp.ItemType.PAGE_BREAK; j++) last = j;

  const vendorKey = Object.keys(EXTRA_DATE_VENDOR_FEES[header] || {}).filter(key => extraNorm_(vendor).indexOf(key) !== -1)[0];
  const fee = vendorKey ? EXTRA_DATE_VENDOR_FEES[header][vendorKey] : (FORM_VENDOR_FEES[vendor] || FORM_DEFAULT_FEE);
  if (dryRun) return `追加できる（${last + 2}番目の位置・${fee}/1日）`;

  const question = form.addMultipleChoiceItem()
    .setTitle(questionTitle)
    .setChoiceValues(["希望する", "希望しない"])
    .setRequired(true)
    .setHelpText(`${EXTRA_DATE_NOTICE_MARK}${header}の特別回収が追加されました。特別回収料金：${fee}/1日`);
  form.moveItem(question.getIndex(), last + 1);

  const sectionHeader = items[headerIndex].asSectionHeaderItem();
  const note = `${EXTRA_DATE_NOTICE_MARK}${header}の特別回収が追加されました（${fee}/1日）。`;
  if (sectionHeader.getHelpText().indexOf(note) === -1) sectionHeader.setHelpText(`${sectionHeader.getHelpText()}\n${note}`);

  if (!noticedForms.has(formId)) {
    noticedForms.add(formId);
    if (form.getDescription().indexOf(EXTRA_DATE_NOTICE_MARK) === -1) form.setDescription(`${EXTRA_DATE_FORM_NOTICE}\n\n${form.getDescription()}`);
  }
  return `追加した（${fee}/1日）`;
}

// 8. 追加の日程の回答状況を、一覧にする（読み取りだけ）
function listExtraDatePending() {
  const ss = invGetSpreadsheet();
  const list = ss.getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  const resp = ss.getSheetByName(RESP_OUTPUT_SHEET_NAME);
  const respValues = resp.getDataRange().getDisplayValues();
  const rh = respValues[0];
  const extraCols = EXTRA_DATE_HEADERS.map(h => rh.indexOf(h));
  const answered = {};
  for (let i = 1; i < respValues.length; i++) answered[respValues[i][0]] = respValues[i];

  const rows = [];
  for (let i = 1; i < list.length; i++) {
    EXTRA_DATE_COLS.forEach((col, k) => {
      if (list[i][col] !== "○") return;
      const store = list[i][0];
      const resp = answered[store];
      let status, previousAt = "", cell = "";
      if (!resp) status = "未回答（フォーム全体）";
      else {
        previousAt = resp[1];
        cell = extraCols[k] >= 0 ? resp[extraCols[k]] : "";
        status = cell === "○" || cell === "×" ? "追加の日も回答済み" : "回答済み・追加の日が未回答";
      }
      rows.push([status, store, list[i][16], list[i][1], EXTRA_DATE_HEADERS[k], previousAt, cell]);
    });
  }
  const order = { "回答済み・追加の日が未回答": 0, "未回答（フォーム全体）": 1, "追加の日も回答済み": 2 };
  rows.sort((a, b) => order[a[0]] - order[b[0]] || a[3].localeCompare(b[3]));
  let out = ss.getSheetByName(EXTRA_DATE_PENDING_SHEET);
  if (!out) out = ss.insertSheet(EXTRA_DATE_PENDING_SHEET);
  out.clear();
  const head = ["状況", "店舗名", "契約者", "送信先アドレス", "追加の日", "前回の回答日時", "追加の日の回答"];
  out.getRange(1, 1, 1, head.length).setValues([head]);
  if (rows.length > 0) out.getRange(2, 1, rows.length, head.length).setValues(rows);
  const count = label => rows.filter(r => r[0] === label).length;
  console.log(`追加の日程の対象 ${rows.length} 件: 回答済み・追加の日が未回答 ${count("回答済み・追加の日が未回答")} 件、未回答（フォーム全体）${count("未回答（フォーム全体）")} 件、追加の日も回答済み ${count("追加の日も回答済み")} 件。シート「${EXTRA_DATE_PENDING_SHEET}」に書き出しました。`);
}

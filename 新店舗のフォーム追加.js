// ===== 新店舗を、今あるフォームに追加する =====
// 同じ契約者・同じアドレスで、すでにフォームがあるお客様の新店舗は、新しいフォームではなく、今あるフォームに、その店舗のページを足す。
// 今回作った新店舗のフォーム（「リスト」の NEWC_FIRST_ROW 行目以降）が対象。
//   1. previewAddStores … 対象を、シート「新店舗のフォーム追加計画」に書き出す（フォームは変更しない）。内容を目で確認する
//   2. applyAddStores   … 今あるフォームに、店舗のページを足す。「リスト」のその行の、フォームURL・フォームIDを、今あるフォームに付け替える。
//                         今回作った新店舗のフォームは、回答が0件なら、ゴミ箱に移す（スクリプトプロパティ ADDSTORE_MODE が on のときだけ動く）
//   3. 送り終わったら、ADDSTORE_MODE を削除する
// お客様には、何も送らない。

const ADDSTORE_PLAN_SHEET = "新店舗のフォーム追加計画";
const ADDSTORE_RECORD_SHEET = "新店舗 フォーム追加記録";
const ADDSTORE_MODE_PROPERTY = "ADDSTORE_MODE";

// アドレスが違っても、同じ契約者なら、今あるフォームに足す（お客様には、元のアドレスに送る）契約者。全角半角・空白の違いは、無視する
const ADDSTORE_MATCH_BY_CUSTOMER = ["ＩＦＲＥＡ　ＤＩＮＩＮＧ", "エイブルジャパン"];

function addstorePlan_() {
  const list = invGetSpreadsheet().getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  const existing = {};
  const existingByCustomer = {};   // 同じ契約者の、今あるフォーム（日程変更の「○」がある行を、優先する）
  for (let i = 1; i < NEWC_FIRST_ROW - 1 && i < list.length; i++) {
    const r = list[i];
    if (!r[9] || r[10] !== "送信済" || !r[16]) continue;
    const key = String(r[1]).trim().toLowerCase() + "|" + extraNorm_(r[16]);
    if (!existing[key]) existing[key] = { formId: r[9], url: r[8] };
    const ck = extraNorm_(r[16]);
    const hasExtra = EXTRA_DATE_COLS.some(c => r[c] === "○");
    if (!existingByCustomer[ck] || (hasExtra && !existingByCustomer[ck].hasExtra)) existingByCustomer[ck] = { formId: r[9], url: r[8], hasExtra: hasExtra };
  }
  const byCustomerList = ADDSTORE_MATCH_BY_CUSTOMER.map(extraNorm_);
  const plan = [];
  for (let i = NEWC_FIRST_ROW - 1; i < list.length; i++) {
    const r = list[i];
    if (!r[9] || r[10] !== "送信済") continue;
    const key = String(r[1]).trim().toLowerCase() + "|" + extraNorm_(r[16]);
    let ex = existing[key];
    if (!ex) {
      const ck = extraNorm_(r[16]);
      if (byCustomerList.some(h => ck.indexOf(h) !== -1)) ex = existingByCustomer[ck];
    }
    if (!ex || ex.formId === r[9]) continue;
    plan.push({ rowNumber: i + 1, store: r[0], customer: r[16], email: r[1], vendor: r[2], address: r[17], newFormId: r[9], exFormId: ex.formId, exUrl: ex.url, row: r });
  }
  return { plan: plan, header: list[0] };
}

// 1. 計画の確認（フォームは変更しない）
function previewAddStores() {
  const ss = invGetSpreadsheet();
  const result = addstorePlan_();
  const rows = result.plan.map(p => {
    let responses = "確認できず";
    try { responses = FormApp.openById(p.exFormId).getResponses().length; } catch (e) { /* 確認できず */ }
    const extra = EXTRA_DATE_COLS.filter(c => p.row[c] === "○").map(c => result.header[c]).join("、");
    return [p.store, p.customer, p.email, responses, extra, p.rowNumber, p.exFormId];
  });
  let out = ss.getSheetByName(ADDSTORE_PLAN_SHEET);
  if (!out) out = ss.insertSheet(ADDSTORE_PLAN_SHEET);
  out.clear();
  const head = ["新店舗", "契約者", "送信先アドレス", "今あるフォームの回答数", "追加の日（12/29・1/4）", "リストの行", "今あるフォームID"];
  out.getRange(1, 1, 1, head.length).setValues([head]);
  if (rows.length > 0) out.getRange(2, 1, rows.length, head.length).setValues(rows);
  console.log(`【確認のみ・フォームは変更していません】今あるフォームに足す新店舗: ${rows.length} 件（今あるフォームに回答あり: ${rows.filter(r => r[3] > 0).length} 件）。シート「${ADDSTORE_PLAN_SHEET}」に書き出しました。`);
}

// 2. 今あるフォームに、新店舗のページを足す
function applyAddStores() {
  if (String(PropertiesService.getScriptProperties().getProperty(ADDSTORE_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${ADDSTORE_MODE_PROPERTY} が on ではないため、フォームを変更しません。`);
    return;
  }
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) { console.log("別の処理が動いているためスキップしました。少し待って、もう一度実行してください。"); return; }
  const startTime = Date.now();
  const results = [];
  try {
    const ss = invGetSpreadsheet();
    const listSheet = ss.getSheetByName(INV_LIST_SHEET_NAME);
    let record = ss.getSheetByName(ADDSTORE_RECORD_SHEET);
    if (!record) { record = ss.insertSheet(ADDSTORE_RECORD_SHEET); record.appendRow(["今あるフォームID", "店舗名", "結果", "日時"]); }
    const result = addstorePlan_();
    const noticedForms = new Set();
    result.plan.forEach(p => {
      if (Date.now() - startTime > 4.5 * 60 * 1000) { results.push([p.store, "時間切れ（もう一度実行してください）"]); return; }
      try {
        const form = FormApp.openById(p.exFormId);
        const sectionTitle = `【${p.store}】の収集希望について`;
        const already = form.getItems().some(item => item.getType() === FormApp.ItemType.SECTION_HEADER && item.getTitle() === sectionTitle);
        let outcome = "すでに追加済み";
        if (!already) {
          form.addPageBreakItem().setTitle("希望日の選択（新しい店舗）");
          const fee = FORM_VENDOR_FEES[p.vendor] || FORM_DEFAULT_FEE;
          form.addSectionHeaderItem().setTitle(sectionTitle).setHelpText(`ご住所：${String(p.address || "").trim() || "（登録がありません）"}\n特別回収料金：${fee}/1日`);
          for (let col = FORM_DATE_COL_START; col <= FORM_DATE_COL_END; col++) {
            if (p.row[col] === "○") {
              form.addMultipleChoiceItem().setTitle(`【${p.store}】${result.header[col]}の収集を希望しますか？`).setChoiceValues(["希望する", "希望しない"]).setRequired(true);
            }
          }
          const desc = form.getDescription();
          const m = desc.match(/【対象店舗】\n([^\n]*)/);
          if (m && m[1].indexOf(p.store) === -1) form.setDescription(desc.replace(m[0], () => m[0] + "・" + p.store));
          EXTRA_DATE_COLS.forEach((col, k) => {
            if (p.row[col] !== "○") return;
            const o = addExtraDateQuestion_(p.exFormId, p.store, EXTRA_DATE_HEADERS[k], p.vendor, false, noticedForms);
            if (o.indexOf("追加した") !== 0) throw new Error(`追加の日の設問を足せませんでした: ${o}`);
          });
          outcome = "追加した";
        }
        // 「リスト」のこの行を、今あるフォームに付け替える（状況は、送信済のまま）
        listSheet.getRange(p.rowNumber, 9, 1, 2).setValues([[p.exUrl, p.exFormId]]);
        // 今回作った新店舗のフォームは、回答が0件なら、ゴミ箱へ
        try {
          if (FormApp.openById(p.newFormId).getResponses().length === 0) DriveApp.getFileById(p.newFormId).setTrashed(true);
        } catch (e) { console.warn(`今回作ったフォームを、ゴミ箱に移せませんでした: ${p.store} / ${e}`); }
        record.appendRow([p.exFormId, p.store, outcome, invFormatNow()]);
        SpreadsheetApp.flush();
        results.push([p.store, outcome]);
      } catch (e) {
        results.push([p.store, `失敗: ${e}`]);
      }
    });
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
  const failed = results.filter(r => r[1].indexOf("失敗") === 0 || r[1].indexOf("時間切れ") === 0);
  console.log(`新店舗 ${results.length} 件: 追加した ${results.filter(r => r[1] === "追加した").length} 件、すでに追加済み ${results.filter(r => r[1] === "すでに追加済み").length} 件、失敗・時間切れ ${failed.length} 件。`);
  failed.forEach(r => console.error(`${r[0]} / ${r[1]}`));
}

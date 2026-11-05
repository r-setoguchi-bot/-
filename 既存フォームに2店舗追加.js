// ===== ＤＡＭＤＥＥ 柏店・呼炉凪来 蒲田店を、今あるフォームに足す（お客様には、何も送らない）=====
// 前提: kintoneに住所・アドレスを入れて、buildListFromKintone を実行済み（「リスト_kintone」に2店舗の行がある）
//   1. previewAddTwo … 足す内容を、ログに出す（フォーム・リストは変更しない）
//   2. applyAddTwo   … 今あるフォームに店舗のページを足し、「リスト」に行を足す
//                       （スクリプトプロパティ ADDTWO_MODE が on のときだけ動く）
//   3. 終わったら、ADDTWO_MODE を削除する

const ADDTWO_MODE_PROPERTY = "ADDTWO_MODE";
const ADDTWO_STORES = [
  { store: "ＤＡＭＤＥＥ　ＴＯＫＹＯ　柏店", customer: "株式会社ＤＡＭＤＥＥＴＯＫＹＯ" },
  { store: "呼炉凪来　蒲田店／もつ焼きディープ　蒲田店", customer: "株式会社呼炉凪来" }
];

function addtwoPlan_() {
  const ss = invGetSpreadsheet();
  const kvals = ss.getSheetByName(KLIST_OUTPUT_SHEET_NAME).getDataRange().getDisplayValues();
  const list = ss.getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  return ADDTWO_STORES.map(def => {
    let k = null;
    for (let i = 1; i < kvals.length; i++) {
      if (extraNorm_(kvals[i][0]) === extraNorm_(def.store) && extraNorm_(kvals[i][16]) === extraNorm_(def.customer)) { k = kvals[i]; break; }
    }
    if (!k) throw new Error(`「リスト_kintone」に見つかりません: ${def.store}`);
    const email = String(k[1]).trim().toLowerCase();
    let sibling = null, already = false;
    for (let i = 1; i < list.length; i++) {
      const r = list[i];
      if (extraNorm_(r[0]) === extraNorm_(def.store) && extraNorm_(r[16]) === extraNorm_(def.customer)) already = true;
      if (!sibling && r[9] && r[10] === "送信済" && String(r[1]).trim().toLowerCase() === email && extraNorm_(r[16]) === extraNorm_(def.customer)) sibling = r;
    }
    if (!sibling) throw new Error(`同じアドレス・契約者の、送信済みのフォームが「リスト」にありません: ${def.store}`);
    return { def: def, k: k, sibling: sibling, already: already, header: list[0] };
  });
}

// 1. 確認のみ
function previewAddTwo() {
  addtwoPlan_().forEach(p => {
    let responses = "確認できず";
    try { responses = FormApp.openById(p.sibling[9]).getResponses().length; } catch (e) { /* 確認できず */ }
    const dates = [];
    for (let c = FORM_DATE_COL_START; c <= FORM_DATE_COL_END; c++) if (p.k[c] === "○") dates.push(p.header[c]);
    console.log(`${p.def.store} / 業者: ${p.k[2]} / 日: ${dates.join("、") || "なし"} / 住所: ${p.k[17] ? "あり" : "空"} / 足し先フォームの回答: ${responses} 件 / リストに行: ${p.already ? "すでにある" : "なし"}`);
  });
  console.log("【確認のみ・何も変更していません】");
}

// 2. 実行
function applyAddTwo() {
  if (String(PropertiesService.getScriptProperties().getProperty(ADDTWO_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${ADDTWO_MODE_PROPERTY} が on ではないため、変更しません。`);
    return;
  }
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) { console.log("別の処理が動いているためスキップしました。少し待って、もう一度実行してください。"); return; }
  try {
    const listSheet = invGetSpreadsheet().getSheetByName(INV_LIST_SHEET_NAME);
    addtwoPlan_().forEach(p => {
      const store = p.def.store;
      const formId = p.sibling[9];
      const form = FormApp.openById(formId);
      const sectionTitle = `【${store}】の収集希望について`;
      let outcome = "ページはすでにあった";
      if (!form.getItems().some(i => i.getType() === FormApp.ItemType.SECTION_HEADER && i.getTitle() === sectionTitle)) {
        form.addPageBreakItem().setTitle("希望日の選択（新しい店舗）");
        const fee = FORM_VENDOR_FEES[p.k[2]] || FORM_DEFAULT_FEE;
        form.addSectionHeaderItem().setTitle(sectionTitle).setHelpText(`ご住所：${String(p.k[17] || "").trim() || "（登録がありません）"}\n特別回収料金：${fee}/1日`);
        let n = 0;
        for (let c = FORM_DATE_COL_START; c <= FORM_DATE_COL_END; c++) {
          if (p.k[c] === "○") {
            form.addMultipleChoiceItem().setTitle(`【${store}】${p.header[c]}の収集を希望しますか？`).setChoiceValues(["希望する", "希望しない"]).setRequired(true);
            n++;
          }
        }
        if (n === 0) throw new Error(`収集できる日がありません: ${store}`);
        const desc = form.getDescription();
        const m = desc.match(/【対象店舗】\n([^\n]*)/);
        if (m && m[1].indexOf(store) === -1) form.setDescription(desc.replace(m[0], () => m[0] + "・" + store));
        outcome = "ページを足した";
      }
      if (!p.already) {
        const row = p.k.slice(0, 18);
        while (row.length < 20) row.push("");
        row[8] = p.sibling[8]; row[9] = formId; row[10] = "送信済";
        listSheet.appendRow(row);
        outcome += "・リストに行を足した";
      }
      SpreadsheetApp.flush();
      console.log(`${store}: ${outcome}`);
    });
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

// ===== 宝華 飛田給店 のフォームを作る（お客様には、何も送らない）=====
//   createHoukaForm … 「リスト_kintone」の宝華の行を「リスト」に足し、単独のフォームを作り、「リスト」に結びつける。
//                     ほかに「フォームを作る待ちの行」があるときは、何もしないで止まる（ほかの店舗のフォームを巻き込まないため）
//   終わったら、お客様へのメールは、別に送る。

const HOUKA_STORE = "宝華　飛田給店";
const HOUKA_CUSTOMER = "株式会社宝華";

function createHoukaForm() {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(30 * 1000)) { console.log("別の処理が動いているためスキップしました。少し待って、もう一度実行してください。"); return; }
  try {
    const ss = invGetSpreadsheet();
    const listSheet = ss.getSheetByName(INV_LIST_SHEET_NAME);
    let list = listSheet.getDataRange().getDisplayValues();
    const key = (s, c) => extraNorm_(s) + "|" + extraNorm_(c);
    const target = key(HOUKA_STORE, HOUKA_CUSTOMER);

    // 1. 「リスト」に、行がなければ足す
    let rowNumber = 0;
    for (let i = 1; i < list.length; i++) if (key(list[i][0], list[i][16]) === target) { rowNumber = i + 1; break; }
    if (rowNumber === 0) {
      const kint = ss.getSheetByName("リスト_kintone").getDataRange().getDisplayValues();
      const kr = kint.find((r, i) => i > 0 && key(r[0], r[16]) === target);
      if (!kr) throw new Error("「リスト_kintone」に、宝華の行がありません。");
      const email = String(kr[1] || "").trim();
      if (!/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(email)) throw new Error("宝華の送信先アドレスの形が正しくありません。");
      if (![3, 4, 5, 6, 7].some(c => kr[c] === "○")) throw new Error("宝華に、収集できる日がありません。");
      const row = kr.slice(0, 18);
      while (row.length < 18) row.push("");
      row[8] = ""; row[9] = ""; row[10] = "";
      rowNumber = listSheet.getLastRow() + 1;
      listSheet.getRange(rowNumber, 1, 1, 20).setValues([row.concat(["", ""])]);
      SpreadsheetApp.flush();
      list = listSheet.getDataRange().getDisplayValues();
      console.log(`「リスト」の ${rowNumber} 行目に、宝華の行を足しました。`);
    }
    const r = list[rowNumber - 1];
    if (r[9] && r[10] === "送信済") { console.log("宝華は、すでにフォームがあります。何もしません。"); return; }

    // 2. ほかに、フォームを作る待ちの行がないことを確認する
    const others = [];
    for (let i = 1; i < list.length; i++) {
      if (i + 1 === rowNumber) continue;
      const s = list[i][10];
      if (s === "送信済" || s === "エラー" || s === "除外" || !String(list[i][1]).trim()) continue;
      others.push(i + 1);
    }
    if (others.length > 0) throw new Error(`宝華のほかに、フォームを作る待ちの行が ${others.length} 件あります（行: ${others.join(", ")}）。巻き込まないよう、止めました。`);

    // 3. 単独のフォームを作る
    const store = { storeName: r[0], vendor: r[2], rowNumber: rowNumber, rowData: r };
    const result = createFormWithRetry([store], list[0], getFormFolder(), String(r[16]).trim());
    try { const f = FormApp.openById(result.formId); if (f.requiresLogin()) f.setRequireLogin(false); } catch (e) { /* 確認できないことがある */ }
    openAccessRetry_(() => openAccessOne_(result.formId));
    listSheet.getRange(rowNumber, 9, 1, 3).setValues([[result.formUrl, result.formId, "送信済"]]);
    SpreadsheetApp.flush();
    console.log(`宝華の単独フォームを作り、「リスト」に結びつけました（${rowNumber}行目）。`);
    console.log("URL: " + result.formUrl);
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

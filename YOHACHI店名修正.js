// ===== YOHACHI の店名を、フォームの店名に合わせる（お客様には何も送らない）=====
//   fixYohachiName … 「リスト」A28 と「回答集計」の行の店名を「ジンギスカン羊はち　三軒茶屋店」にする（控えメールの印も引き継がれる）
function fixYohachiName() {
  const OLD = "ＹＯＨＡＣＨＩ", NEW = "ジンギスカン羊はち　三軒茶屋店";
  const FORM_ID = "1v-_1Gg9D2sdnuwucGnM2mRpgob1g0e4txbe59nQrTWw";
  const lock = LockService.getUserLock();
  if (!lock.tryLock(30 * 1000)) { console.log("別の処理が動いています。少し待って、もう一度実行してください。"); return; }
  try {
    const ss = invGetSpreadsheet();
    const ls = ss.getSheetByName(INV_LIST_SHEET_NAME);
    const ld = ls.getDataRange().getDisplayValues();
    const rows = []; ld.forEach((r, i) => { if (i > 0 && r[0] === OLD && r[9] === FORM_ID) rows.push(i + 1); });
    if (rows.length !== 1) throw new Error("「リスト」に、対象の行が1行ではありません: " + rows.join(","));
    ls.getRange(rows[0], 1).setValue(NEW);
    const rs = ss.getSheetByName("回答集計");
    const rd = rs.getDataRange().getDisplayValues();
    const rr = []; rd.forEach((r, i) => { if (i > 0 && r[0] === OLD && r[12] === FORM_ID) rr.push(i + 1); });
    rr.forEach(n => rs.getRange(n, 1).setValue(NEW));
    SpreadsheetApp.flush();
    console.log("「リスト」" + rows[0] + "行目、「回答集計」" + (rr.join(",") || "（なし）") + "行目の店名を変更しました。");
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

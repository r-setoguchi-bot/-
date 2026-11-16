// ===== YOHACHI の回答が集計に載らない原因の確認（読むだけ。何も変更しない）=====
//   diagYohachi … フォームの設問名と、最新の回答、「リスト」の店名を、ログに出す
function diagYohachi() {
  const formId = "1v-_1Gg9D2sdnuwucGnM2mRpgob1g0e4txbe59nQrTWw";
  const form = FormApp.openById(formId);
  console.log("== 設問 ==");
  form.getItems().forEach((it, i) => console.log(i + ": [" + it.getType() + "] " + it.getTitle()));
  console.log("== 回答（新しい順）==");
  form.getResponses().slice().reverse().forEach((r, n) => {
    console.log("-- " + r.getTimestamp() + " (" + (n + 1) + "件目)");
    r.getItemResponses().forEach(ir => console.log("   " + ir.getItem().getTitle() + " => " + ir.getResponse()));
  });
  console.log("== 「リスト」でこのフォームの店 ==");
  const data = invGetSpreadsheet().getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  data.forEach((r, i) => { if (r[9] === formId) console.log((i + 1) + "行目: " + r[0] + " / " + r[2] + " / D-H=" + r.slice(3, 8).join(",")); });
}

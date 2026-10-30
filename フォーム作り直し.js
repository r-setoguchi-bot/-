// ===== 業者の日程（12/29・1/4）の追加に伴う、フォームの作り直し =====
// 手順（読み取りだけの確認 → 準備 → フォーム作成 → 旧フォームの停止 → メール、の順に進める）。
//   1. 「業者影響確認」シートを作る（業者影響確認.js の listVendorImpact）
//   2. previewFormRebuild を実行 → 「フォーム作り直し計画」シートに、作り直すフォームの一覧を書き出す（読み取りだけ。何も変更しない）
//   3. 計画を目で確認する（混在・除外・回答済みの扱い）

const REBUILD_PLAN_SHEET_NAME = "フォーム作り直し計画";

// 「業者影響確認」で、影響のある店舗を含むフォームを、「リスト」から探して、一覧にする。何も変更しない。
function previewFormRebuild() {
  const ss = invGetSpreadsheet();
  const impactSheet = ss.getSheetByName("業者影響確認");
  if (!impactSheet) throw new Error("先に listVendorImpact を実行して、「業者影響確認」シートを作ってください。");
  const impactValues = impactSheet.getDataRange().getDisplayValues();
  const ih = impactValues[0];
  const ii = name => ih.indexOf(name);
  const affected = {};   // 店舗名|契約者 → 影響の内容
  impactValues.slice(1).forEach(row => {
    if (row[ii("影響")] === "なし") return;
    affected[row[ii("店舗名")] + "|" + row[ii("契約者")]] = { impact: row[ii("影響")], vendor: row[ii("業者")], state: row[ii("案内の状況")] };
  });

  const listSheet = ss.getSheetByName(INV_LIST_SHEET_NAME);
  const listValues = listSheet.getDataRange().getDisplayValues();
  const lh = listValues[0];
  const colOf = name => lh.indexOf(name);
  const COL_STORE = 0, COL_EMAIL = 1, COL_FORM = 9, COL_STATUS = 10, COL_CUSTOMER = 16;

  // 回答のある店舗
  const answered = {};
  const respSheet = ss.getSheetByName("回答集計");
  if (respSheet && respSheet.getLastRow() > 1) {
    respSheet.getRange(2, 1, respSheet.getLastRow() - 1, 2).getDisplayValues().forEach(row => { if (row[0]) answered[row[0]] = row[1]; });
  }

  const byForm = {};
  const order = [];
  const foundAffected = new Set();
  for (let i = 1; i < listValues.length; i++) {
    const row = listValues[i];
    const formId = row[COL_FORM];
    if (!formId) continue;
    if (!byForm[formId]) { byForm[formId] = { formId: formId, email: row[COL_EMAIL], customer: row[COL_CUSTOMER], stores: [] }; order.push(formId); }
    const key = row[COL_STORE] + "|" + row[COL_CUSTOMER];
    const hit = affected[key];
    if (hit) foundAffected.add(key);
    byForm[formId].stores.push({ name: row[COL_STORE], status: row[COL_STATUS], impact: hit ? hit.impact : "", answeredAt: answered[row[COL_STORE]] || "" });
  }

  const rows = [];
  order.forEach(formId => {
    const f = byForm[formId];
    if (!f.stores.some(s => s.impact)) return;
    const hit = f.stores.filter(s => s.impact);
    const rest = f.stores.filter(s => !s.impact);
    rows.push([
      f.formId, f.stores.map(s => s.status).filter((v, i, a) => a.indexOf(v) === i).join("・"), f.email, f.customer,
      f.stores.length, hit.map(s => `${s.name}（${s.impact.slice(0, 5)}）`).join("\n"), rest.map(s => s.name).join("\n"),
      f.stores.some(s => s.answeredAt) ? "あり" : "なし", rest.length > 0 ? "あり" : "なし",
      f.stores.some(s => s.status === "除外") ? "あり" : "なし"
    ]);
  });
  // フォームのない（期間外の）影響店舗
  Object.keys(affected).filter(key => !foundAffected.has(key)).forEach(key => {
    const [name, customer] = key.split("|");
    rows.push(["（フォームなし）", affected[key].state, "", customer, 1, `${name}（${affected[key].impact.slice(0, 5)}）`, "", "なし", "なし", "なし"]);
  });

  const head = ["旧フォームID", "リストの状況", "送信先アドレス", "契約者", "フォームの店舗数", "影響のある店舗", "影響のない店舗", "回答済み", "混在", "除外"];
  let out = ss.getSheetByName(REBUILD_PLAN_SHEET_NAME);
  if (!out) out = ss.insertSheet(REBUILD_PLAN_SHEET_NAME);
  out.clear();
  out.getRange(1, 1, 1, head.length).setValues([head]);
  if (rows.length > 0) out.getRange(2, 1, rows.length, head.length).setValues(rows);
  const mixed = rows.filter(r => r[8] === "あり").length;
  const excluded = rows.filter(r => r[9] === "あり").length;
  const answeredCount = rows.filter(r => r[7] === "あり").length;
  console.log(`作り直す候補のフォーム: ${rows.length} 件（回答済み ${answeredCount} 件、影響のない店舗が混在 ${mixed} 件、除外を含む ${excluded} 件）。シート「${REBUILD_PLAN_SHEET_NAME}」に書き出しました（何も変更していません）。`);
}

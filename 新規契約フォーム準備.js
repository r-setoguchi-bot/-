// ===== 新しく契約した店舗を、「リスト」に足して、フォームを作る準備 =====
// 「リスト_kintone」にあって、「リスト」にない店舗（契約状況一覧で「未作成」の店舗）が、候補。
//   1. listNewContractCandidates   … 候補を、シート「新規契約の候補」に書き出す（フォームもメールも、何もしない）。内容を目で見て、足さない行は「追加する」の○を消す
//   2. addNewContractRowsToList    … 「追加する」が○の行を、「リスト」の末尾に足す（状況は空。野口興業・光栄和の店舗には、S列・T列に「○」を入れる）
//   3. startFormCreation           … 状況が空の行のフォームを作る（フォーム作成.js）
// お客様には、何も送らない。メールは、別に送る。

const NEWC_CANDIDATE_SHEET = "新規契約の候補";
const NEWC_HEAD = ["追加する", "店舗名", "契約者", "送信先アドレス", "業者名", "回収日（12/30〜1/3）", "追加の日（12/29・1/4）", "保留（まとめて送る）", "メモ", "リスト_kintoneの行"];

function newcKey_(store, customer) { return extraNorm_(store) + "|" + extraNorm_(customer); }

// 1. 候補の一覧を作る
function listNewContractCandidates() {
  const ss = invGetSpreadsheet();
  const kint = ss.getSheetByName("リスト_kintone").getDataRange().getDisplayValues();
  const list = ss.getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  const inList = new Set();
  for (let i = 1; i < list.length; i++) inList.add(newcKey_(list[i][0], list[i][16]));

  // 業者の影響（12/29・1/4）
  const impact = {};
  const impactSheet = ss.getSheetByName("業者影響確認");
  if (impactSheet) {
    const iv = impactSheet.getDataRange().getDisplayValues();
    const ih = iv[0];
    iv.slice(1).forEach(r => {
      const text = r[ih.indexOf("影響")];
      if (text === "なし") return;
      impact[newcKey_(r[ih.indexOf("店舗名")], r[ih.indexOf("契約者")])] = text.indexOf("12/29") === 0 ? 0 : 1;
    });
  }
  const excluded = EXTRA_DATE_EXCLUDE_CUSTOMERS.map(extraNorm_);
  const hold = EXMAIL_HOLD_CUSTOMERS.map(extraNorm_);

  const rows = [];
  for (let i = 1; i < kint.length; i++) {
    const r = kint[i];
    if (!r[0] || inList.has(newcKey_(r[0], r[16]))) continue;
    const email = String(r[1] || "").trim();
    const dates = [3, 4, 5, 6, 7].filter(c => r[c] === "○").map(c => kint[0][c]);
    const extraIdx = impact[newcKey_(r[0], r[16])];
    const isExcluded = excluded.some(e => extraNorm_(r[16]).indexOf(e) !== -1);
    const isHold = hold.some(h => extraNorm_(r[16]).indexOf(h) !== -1);
    const memo = [];
    let add = "○";
    if (!email) { memo.push("アドレスなし"); add = ""; }
    if (dates.length === 0 && extraIdx === undefined) { memo.push("回収日なし"); add = ""; }
    if (isExcluded) { memo.push("個別対応のお客様"); add = ""; }
    if (r[9] || r[10]) { memo.push("すでにフォームの記録がある"); add = ""; }
    rows.push([add, r[0], r[16], email, r[2], dates.join("、"), extraIdx === undefined ? "" : EXTRA_DATE_HEADERS[extraIdx], isHold ? "保留" : "", memo.join("、"), i + 1]);
  }

  let out = ss.getSheetByName(NEWC_CANDIDATE_SHEET);
  if (out && out.getLastRow() > 1) throw new Error(`「${NEWC_CANDIDATE_SHEET}」シートは、すでに作ってあります。作り直す場合は、シートを削除してから、もう一度実行してください。`);
  if (!out) out = ss.insertSheet(NEWC_CANDIDATE_SHEET);
  out.clear();
  out.getRange(1, 1, 1, NEWC_HEAD.length).setValues([NEWC_HEAD]);
  if (rows.length > 0) out.getRange(2, 1, rows.length, NEWC_HEAD.length).setValues(rows);
  console.log(`新規契約の候補: ${rows.length} 件。うち「追加する」が○: ${rows.filter(r => r[0] === "○").length} 件、野口興業・光栄和の日程あり: ${rows.filter(r => r[6]).length} 件、保留（まとめて送る）: ${rows.filter(r => r[7]).length} 件。シート「${NEWC_CANDIDATE_SHEET}」に書き出しました（フォームもメールも、何もしていません）。`);
}

// 2. 「追加する」が○の行を、「リスト」の末尾に足す
function addNewContractRowsToList() {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) { console.log("別の処理が動いているためスキップしました。少し待って、もう一度実行してください。"); return; }
  try {
    const ss = invGetSpreadsheet();
    const cand = ss.getSheetByName(NEWC_CANDIDATE_SHEET);
    if (!cand || cand.getLastRow() < 2) throw new Error(`先に listNewContractCandidates を実行してください。`);
    const kintSheet = ss.getSheetByName("リスト_kintone");
    const kint = kintSheet.getDataRange().getDisplayValues();
    const listSheet = ss.getSheetByName(INV_LIST_SHEET_NAME);
    const list = listSheet.getDataRange().getDisplayValues();
    if (list[0][EXTRA_DATE_COLS[0]] !== EXTRA_DATE_HEADERS[0] || list[0][EXTRA_DATE_COLS[1]] !== EXTRA_DATE_HEADERS[1]) throw new Error("「リスト」の S1・T1 に、追加の日の見出しがありません。");
    const inList = new Set();
    for (let i = 1; i < list.length; i++) inList.add(newcKey_(list[i][0], list[i][16]));

    const cv = cand.getRange(2, 1, cand.getLastRow() - 1, NEWC_HEAD.length).getDisplayValues();
    const newRows = [];
    let skipped = 0;
    cv.forEach(c => {
      if (c[0] !== "○") return;
      const kr = kint[Number(c[9]) - 1];
      if (!kr || newcKey_(kr[0], kr[16]) !== newcKey_(c[1], c[2])) { console.error(`行が一致しません: ${c[1]}`); skipped++; return; }
      if (inList.has(newcKey_(kr[0], kr[16]))) { skipped++; return; }
      if (kr[9] || kr[10]) { skipped++; return; }
      const row = kr.slice(0, 18);
      while (row.length < 18) row.push("");
      row[8] = ""; row[9] = ""; row[10] = "";   // フォームURL・フォームID・状況は、空にする
      const extra = ["", ""];
      const k = EXTRA_DATE_HEADERS.indexOf(c[6]);
      if (k >= 0) extra[k] = "○";
      newRows.push(row.concat(extra));
      inList.add(newcKey_(kr[0], kr[16]));
    });
    if (newRows.length === 0) { console.log("足す行がありません。"); return; }
    const start = listSheet.getLastRow() + 1;
    listSheet.getRange(start, 1, newRows.length, 20).setValues(newRows);
    console.log(`「リスト」の ${start} 行目から、${newRows.length} 行を足しました（追加の日の「○」あり: ${newRows.filter(r => r[18] === "○" || r[19] === "○").length} 件、足さなかった行: ${skipped} 件）。状況は空です。フォームは、まだ作っていません。`);
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

// フォーム作成の前の確認: 「リスト」で、状況が空で、アドレスのある行の数と、その行の店舗名
function countRowsWaitingForForm() {
  const list = invGetSpreadsheet().getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  const waiting = [];
  for (let i = 1; i < list.length; i++) {
    const status = list[i][10];
    if (status === "送信済" || status === "エラー" || status === "除外" || !String(list[i][1]).trim()) continue;
    waiting.push(`${i + 1}行目: ${list[i][0]}`);
  }
  console.log(`フォームを作る対象（状況が空でアドレスのある行）: ${waiting.length} 行。`);
  waiting.forEach(w => console.log(w));
}

// 「契約状況一覧」で「未作成」の店舗が、「リスト_kintone」「リスト」のどこにあるかを調べて、シート「未作成の内訳」に書き出す（読み取りと、このシートへの書き出しだけ）
function diagnoseUnformedStores() {
  const ss = invGetSpreadsheet();
  const status = ss.getSheetByName("契約状況一覧").getDataRange().getDisplayValues();
  const sh = status[0];
  const iStore = sh.indexOf("店舗名"), iCust = sh.indexOf("契約者"), iState = sh.indexOf("状況"), iReason = sh.indexOf("理由");
  const kint = ss.getSheetByName("リスト_kintone").getDataRange().getDisplayValues();
  const list = ss.getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  const kintMap = {}, listMap = {};
  for (let i = 1; i < kint.length; i++) kintMap[newcKey_(kint[i][0], kint[i][16])] = i + 1;
  for (let i = 1; i < list.length; i++) listMap[newcKey_(list[i][0], list[i][16])] = i + 1;
  const rows = [];
  for (let i = 1; i < status.length; i++) {
    if (status[i][iState] !== "未作成") continue;
    const key = newcKey_(status[i][iStore], status[i][iCust]);
    const lr = listMap[key];
    rows.push([status[i][iStore], status[i][iCust], kintMap[key] ? "ある（" + kintMap[key] + "行）" : "ない", lr ? "ある（" + lr + "行）" : "ない",
      lr ? (list[lr - 1][10] || "（空）") : "", lr ? (list[lr - 1][9] ? "あり" : "なし") : "", status[i][iReason] || ""]);
  }
  let out = ss.getSheetByName("未作成の内訳");
  if (!out) out = ss.insertSheet("未作成の内訳");
  out.clear();
  const head = ["店舗名", "契約者", "リスト_kintone", "リスト", "リストの状況", "リストにフォームID", "契約状況一覧の理由"];
  out.getRange(1, 1, 1, head.length).setValues([head]);
  if (rows.length > 0) out.getRange(2, 1, rows.length, head.length).setValues(rows);
  console.log(`「未作成」の店舗: ${rows.length} 件。リスト_kintoneにない: ${rows.filter(r => r[2] === "ない").length} 件、リストにある: ${rows.filter(r => r[3] !== "ない").length} 件。シート「未作成の内訳」に書き出しました。`);
}

// ===== 新しく作ったフォームだけ、お客様が回答できる設定にする =====
// 「リスト」の NEWC_FIRST_ROW 行目以降のフォームだけ、①ログイン必須の解除、②リンクを知っている全員が回答できる設定、にする。
// 全フォームを見る startFixFormLogin / startOpenFormAccess（設定確認.js）より、はるかに少ない回数で済む。何度実行しても、同じ結果になる
const NEWC_FIRST_ROW = 1140;

function fixAccessForNewForms() {
  const startTime = Date.now();
  const list = invGetSpreadsheet().getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  const ids = [];
  for (let i = NEWC_FIRST_ROW - 1; i < list.length; i++) {
    const id = list[i][9];
    if (id && list[i][10] === "送信済" && ids.indexOf(id) === -1) ids.push(id);
  }
  let login = 0, loginAlready = 0, loginSkipped = 0, access = 0, accessAlready = 0, errors = 0, remaining = 0;
  for (let n = 0; n < ids.length; n++) {
    if (Date.now() - startTime > 4.5 * 60 * 1000) { remaining = ids.length - n; break; }
    const id = ids[n];
    // ログイン必須の確認は、いま「フォームを編集できませんでした」と出て読めないことがある。失敗しても先に進む
    try {
      const form = FormApp.openById(id);
      if (form.requiresLogin()) {
        form.setRequireLogin(false);
        login++;
      } else {
        loginAlready++;
      }
    } catch (e) {
      loginSkipped++;
    }
    // お客様の回答に必要な設定: 回答者ビューを、リンクを知っている全員に共有する
    try {
      const result = openAccessRetry_(() => openAccessOne_(id));
      if (result === "changed") access++; else accessAlready++;
    } catch (e) {
      errors++;
      console.error(`全員が回答できる設定にできませんでした: 末尾=${String(id).slice(-6)} / ${e}`);
    }
    Utilities.sleep(500);
  }
  console.log(`新しいフォーム ${ids.length} 件: 全員が回答できる設定に変更 ${access} 件（もともと全員可 ${accessAlready} 件）、エラー ${errors} 件。ログイン必須: 解除 ${login} 件、もともと不要 ${loginAlready} 件、確認できず ${loginSkipped} 件。` + (remaining > 0 ? `時間切れで ${remaining} 件が残りました。もう一度実行してください。` : ""));
}

// ===== フォームの編集ができない原因を調べる（読み取りと、テスト店舗のフォームへの「同じ内容の書き直し」だけ。お客様のフォームは、変更しない）=====
function diagnoseFormsAccess() {
  const list = invGetSpreadsheet().getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  const pick = {};
  pick["新しいフォーム（1141行目）"] = list[NEWC_FIRST_ROW][9];
  pick["古いフォーム（2行目）"] = list[1][9];
  for (let i = 1; i < list.length; i++) if (list[i][0] === "テスト店舗（削除予定）") pick["テスト店舗のフォーム"] = list[i][9];
  const me = Session.getEffectiveUser().getEmail();
  console.log(`実行しているアカウントが、ドライブの所有者と同じか、を調べます。`);
  Object.keys(pick).forEach(label => {
    const id = pick[label];
    const step = (name, fn) => {
      try { const v = fn(); console.log(`${label} / ${name}: 成功${v === undefined ? "" : "（" + v + "）"}`); return true; }
      catch (e) { console.error(`${label} / ${name}: 失敗 ${e}`); return false; }
    };
    if (!id) { console.error(`${label}: フォームIDが空です`); return; }
    let form = null;
    step("DriveApp でファイルを開く", () => DriveApp.getFileById(id).getName().slice(0, 8) + "…");
    step("所有者が実行アカウントと同じか", () => DriveApp.getFileById(id).getOwner().getEmail() === me);
    step("FormApp.openById", () => { form = FormApp.openById(id); });
    if (form) {
      step("題名を読む", () => form.getTitle().slice(0, 8) + "…");
      step("回答受付かを読む", () => form.isAcceptingResponses());
      step("ログイン必須かを読む", () => form.requiresLogin());
      if (label === "テスト店舗のフォーム") step("説明を同じ内容で書き直す（編集のテスト）", () => { form.setDescription(form.getDescription()); });
    }
  });
}

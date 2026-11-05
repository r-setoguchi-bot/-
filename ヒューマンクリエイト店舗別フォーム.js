// ===== 株式会社ＨｕｍａｎＱｒｅａｔｅ（ヒューマンクリエイト）の、1つのフォームを共有している店舗を、店舗ごとのフォームに分ける =====
// お客様には、メールを送らない（案内は、あなたが送る）。送信先アドレスの紐づけ（「リスト」のB列）は、変えない。
//   1. previewHqSplit … 分ける対象を、ログに出す（変更しない）
//   2. applyHqSplit   … 店舗ごとに新しいフォームを作り、「リスト」のI列・J列を付け替える。
//                        回答済みの2店舗は、今までの共有フォームのまま残し、他の店舗のページだけを共有フォームから消す
//                        （スクリプトプロパティ HQSPLIT_MODE が on のときだけ動く。時間切れのときは、もう一度実行する）
//   3. listHqForms    … 12店舗の、店舗名・送信先・フォームURLを、シート「ヒューマンクリエイト フォーム一覧」に書き出す
//   4. 終わったら、HQSPLIT_MODE を削除する

const HQ_CUSTOMER = "株式会社ＨｕｍａｎＱｒｅａｔｅ";
const HQ_OLD_FORM_SUFFIX = "3ocNGCQU";            // 今まで、10店舗で共有していたフォームのID（末尾）
const HQSPLIT_MODE_PROPERTY = "HQSPLIT_MODE";
const HQ_LIST_SHEET = "ヒューマンクリエイト フォーム一覧";
// 回答済みのため、今までの共有フォームのまま残す店舗
const HQ_KEEP_STORES = ["Ｂｉｓｔｒｏ　Ｕｎ　Ｔｒｏｉｓ／鮨匠　一石三鳥／鳥影　一石三鳥", "一石三鳥　韓邸"];

function hqIsTarget_(x) {
  return x.row[10] === "送信済" && String(x.row[9]).slice(-HQ_OLD_FORM_SUFFIX.length) === HQ_OLD_FORM_SUFFIX
    && !HQ_KEEP_STORES.some(k => extraNorm_(k) === extraNorm_(x.row[0]));
}

// 共有フォームから、分けた店舗のページ・設問を消し、【対象店舗】とページの題名を整える（回答のある店舗は、そのまま）
function hqTrimOldForm_(oldId, movedNames) {
  const form = FormApp.openById(oldId);
  const items = form.getItems();
  movedNames.forEach(name => {
    const prefix = `【${name}】`;
    for (let i = items.length - 1; i >= 0; i--) {
      const it = items[i];
      const title = it.getTitle();
      if (it.getType() === FormApp.ItemType.SECTION_HEADER && title === `【${name}】の収集希望について`) {
        // 直前のページ区切りも消す
        if (i > 0 && items[i - 1].getType() === FormApp.ItemType.PAGE_BREAK) { form.deleteItem(items[i - 1]); items.splice(i - 1, 1); i--; }
        form.deleteItem(it); items.splice(i, 1);
      } else if (title.indexOf(prefix) === 0) {
        form.deleteItem(it); items.splice(i, 1);
      }
    }
  });
  const keepNames = [];
  form.getItems().forEach(it => {
    const m = it.getType() === FormApp.ItemType.SECTION_HEADER && it.getTitle().match(/^【(.+)】の収集希望について$/);
    if (m) keepNames.push(m[1]);
  });
  const desc = form.getDescription();
  const dm = desc.match(/【対象店舗】\n([^\n]*)/);
  if (dm) form.setDescription(desc.replace(dm[0], () => "【対象店舗】\n" + keepNames.join("・")));
  let k = 0;
  form.getItems().forEach(it => {
    if (it.getType() === FormApp.ItemType.PAGE_BREAK && /^希望日の選択/.test(it.getTitle())) {
      k++;
      it.setTitle(keepNames.length > 1 ? `希望日の選択（${k}/${keepNames.length}店舗目）` : "希望日の選択");
    }
  });
  return keepNames.length;
}

function hqRows_() {
  const data = invGetSpreadsheet().getSheetByName(INV_LIST_SHEET_NAME).getDataRange().getDisplayValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (extraNorm_(data[i][16]) === extraNorm_(HQ_CUSTOMER)) rows.push({ rowNumber: i + 1, row: data[i] });
  }
  return { header: data[0], rows: rows };
}

function previewHqSplit() {
  const r = hqRows_();
  const targets = r.rows.filter(hqIsTarget_);
  let responses = "確認できず";
  const oldId = targets.length ? targets[0].row[9] : null;
  if (oldId) { try { responses = FormApp.openById(oldId).getResponses().length; } catch (e) { /* 確認できず */ } }
  console.log(`ヒューマンクリエイトの行: ${r.rows.length} 件。共有フォームを使っている店舗: ${targets.length} 件（分けて、新しいフォームを作る数）。共有フォームの回答: ${responses} 件（そのまま残します）。`);
  targets.forEach(t => console.log(`  ${t.rowNumber}行目 / 業者: ${t.row[2]} / 日: ${[3, 4, 5, 6, 7].filter(c => t.row[c] === "○").map(c => r.header[c]).join("、")} / 宛先の形: ${/^[^@\s,]+@[^@\s,]+\.[^@\s,]+$/.test(String(t.row[1]).trim()) ? "正しい" : "正しくない"}`));
  console.log("【確認のみ・何も変更していません】");
}

function applyHqSplit() {
  if (String(PropertiesService.getScriptProperties().getProperty(HQSPLIT_MODE_PROPERTY) || "").trim().toLowerCase() !== "on") {
    console.log(`スクリプトプロパティ ${HQSPLIT_MODE_PROPERTY} が on ではないため、変更しません。`);
    return;
  }
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) { console.log("別の処理が動いているためスキップしました。少し待って、もう一度実行してください。"); return; }
  const startTime = Date.now();
  let done = 0, failed = 0, remaining = 0, oldId = null;
  try {
    const listSheet = invGetSpreadsheet().getSheetByName(INV_LIST_SHEET_NAME);
    const folder = getFormFolder();
    const r = hqRows_();
    const targets = r.rows.filter(hqIsTarget_);
    if (targets.length) oldId = targets[0].row[9];
    const moved = [];
    for (let n = 0; n < targets.length; n++) {
      if (Date.now() - startTime > 4 * 60 * 1000) { remaining = targets.length - n; break; }
      const t = targets[n];
      try {
        const store = { storeName: t.row[0], vendor: t.row[2], rowNumber: t.rowNumber, rowData: t.row };
        const result = createFormWithRetry([store], r.header, folder, String(t.row[16]).trim());
        try {
          const form = FormApp.openById(result.formId);
          if (form.requiresLogin()) form.setRequireLogin(false);
        } catch (e) { /* 確認できないことがある。先に進む */ }
        openAccessRetry_(() => openAccessOne_(result.formId));   // リンクを知っている全員が、回答できるようにする
        listSheet.getRange(t.rowNumber, 9, 1, 2).setValues([[result.formUrl, result.formId]]);   // I列・J列だけ付け替える（B列の宛先は、そのまま）
        SpreadsheetApp.flush();
        done++;
        moved.push(t.row[0]);
        console.log(`${t.rowNumber}行目: 店舗別のフォームを作って、付け替えました。`);
      } catch (e) {
        failed++;
        console.error(`${t.rowNumber}行目: 失敗 ${e}`);
      }
    }
    // 分けた店舗のページを、共有フォームから消す（回答済みの店舗だけが残る）
    if (moved.length > 0 && oldId) {
      const left = hqTrimOldForm_(oldId, moved);
      console.log(`今までの共有フォームから、分けた ${moved.length} 店舗のページを消しました（残った店舗: ${left} 件、回答は残っています）。`);
    }
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
  console.log(`今回 ${done} 件を店舗別にしました（失敗 ${failed} 件）。` + (remaining > 0 ? `時間切れで ${remaining} 件が残りました。もう一度実行してください。` : "") + (failed > 0 ? "失敗があったので、もう一度実行してください（付け替え済みの店舗は、飛ばされます）。" : ""));
}

// 3. 12店舗の一覧をシートに書き出す（あなたが、お客様にメールするときの、元の表）
function listHqForms() {
  const ss = invGetSpreadsheet();
  const r = hqRows_();
  const rows = r.rows.filter(x => x.row[10] === "送信済").map(x => {
    let responses = "確認できず";
    try { responses = FormApp.openById(x.row[9]).getResponses().length; } catch (e) { /* 確認できず */ }
    return [x.row[0], x.row[1], x.row[8], x.row[2], responses, x.rowNumber];
  });
  let out = ss.getSheetByName(HQ_LIST_SHEET);
  if (!out) out = ss.insertSheet(HQ_LIST_SHEET);
  out.clear();
  const head = ["店舗名", "送信先アドレス（リストのB列）", "フォームURL", "業者", "フォームの回答数", "リストの行"];
  out.getRange(1, 1, 1, head.length).setValues([head]);
  if (rows.length) out.getRange(2, 1, rows.length, head.length).setValues(rows);
  console.log(`シート「${HQ_LIST_SHEET}」に ${rows.length} 店舗を書き出しました。`);
}

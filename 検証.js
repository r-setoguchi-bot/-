// ===== 二重チェック：kintoneのデータと作成したリスト・フォームを、別の処理で照合する =====
// リスト作成.js とは別の処理で、kintoneの契約管理をもう一度読み直し、アドレス・回収曜日・店舗名などを独立に判定して、
// 「リスト」（または「リスト_kintone」）と食い違いがないかを調べる。何も書き換えず、結果は「検証結果」シートに出す。
//
// 使い方:
//   1. verifyKintoneListData … 「リスト_kintone」を調べる（リストを差し替える前に実行）
//   2. verifyListData        … 「リスト」を調べる（差し替えたあと、フォームを作る前に実行）
//   3. verifyCreatedForms    … 作成したフォームを抜き取りで調べ、フォルダの中身とも照合する（フォームを作ったあと）
//   結果は「検証結果」シートの「重大度」で見る。エラー = 直してから進める、警告 = 内容を見て判断、情報 = 参考
//
// スクリプトプロパティ（リスト作成.js と同じ）: KINTONE_SUBDOMAIN / KINTONE_KEIYAKU_APP_ID / KINTONE_KEIYAKU_API_TOKEN
// 項目名の対応は、KLIST_FIELD_CODES（あれば）を優先して使う。
// スプレッドシートの指定は フォーム作成.js と同じ。このファイルだけで動く。

const VER_RESULT_SHEET_NAME = "検証結果";
const VER_REVIEW_SHEET_NAME = "要確認";
const VER_LIST_SHEET_NAME = "リスト";
const VER_KINTONE_LIST_SHEET_NAME = "リスト_kintone";
const VER_TARGET_CONTRACT_TYPE = "契約中";

// 照合の条件（リスト作成.js の設定と同じ内容を、ここにも書いて独立に判定する）
const VER_DATES = [[2026, 12, 30], [2026, 12, 31], [2027, 1, 1], [2027, 1, 2], [2027, 1, 3]];
const VER_NO_COLLECTION = ["2027-1-1"];
const VER_WEEKDAYS = "日月火水木金土";
const VER_L_CONTRACT_TYPE = "契約種別";
const VER_L_STORE = "契約店舗名称";
const VER_L_CUSTOMER = "顧客名称";
const VER_L_VENDOR = "収集業者：名称";
const VER_L_BURNABLE = "可燃収集日";
const VER_L_EMAILS = [
  "年末年始　回答アドレス①",
  "年末年始　回答アドレス②",
  "案内等送付先メールアドレス:(to)",
  "送付先別：e-mail",
  "契約時の顧客担当：メールアドレス"
];

// 名前では見つけにくい項目は、フィールドコードで指定する（項目名 → フィールドコード）
const VER_EMAIL_CODES = {
  "送付先別：e-mail": "送付先別mail1"
};

// 「リスト」の列（0始まり）
const VER_COL_STORE = 0;
const VER_COL_EMAIL = 1;
const VER_COL_VENDOR = 2;
const VER_COL_MARK_START = 3;   // D〜H列
const VER_COL_MARK_END = 7;
const VER_COL_URL = 8;
const VER_COL_ID = 9;
const VER_COL_STATUS = 10;

const VER_MAX_STORES_PER_FORM = 30;     // 1フォームの店舗数がこれを超えたら警告
const VER_FORM_SAMPLE_COUNT = 20;       // フォームの抜き取り件数
const VER_TYPO_DOMAINS = ["gmial.com", "gmai.com", "gmail.co", "gmail.con", "gmaill.com", "gamil.com", "yahoo.co.jo",
  "yahoo.cojp", "yahho.co.jp", "yahoo.ne.jp.", "hotmial.com", "hotmail.co", "outlok.com", "icloud.co", "docomo.ne.j", "ocn.ne.jo"];

function verifyKintoneListData() {
  verifyListSheet_(VER_KINTONE_LIST_SHEET_NAME);
}

function verifyListData() {
  verifyListSheet_(VER_LIST_SHEET_NAME);
}

function verGetSpreadsheet_() {
  const url = PropertiesService.getScriptProperties().getProperty("FORM_SPREADSHEET_URL");
  if (url) return SpreadsheetApp.openByUrl(url);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("スクリプトプロパティ FORM_SPREADSHEET_URL を設定してください。");
  return ss;
}

function verNormalize_(s) {
  return String(s).normalize("NFKC").replace(/\s/g, "");
}

function verFetchJson_(url, token) {
  const res = UrlFetchApp.fetch(url, { method: "get", headers: { "X-Cybozu-API-Token": token }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error(`kintone取得失敗 (HTTP ${res.getResponseCode()})`);
  return JSON.parse(res.getContentText());
}

// kintoneから、契約中の契約を読み直す（リスト作成.js とは別の読み方）
function verLoadKintone_() {
  const props = PropertiesService.getScriptProperties();
  const sub = props.getProperty("KINTONE_SUBDOMAIN");
  const app = props.getProperty("KINTONE_KEIYAKU_APP_ID");
  const token = props.getProperty("KINTONE_KEIYAKU_API_TOKEN");
  if (!sub || !app || !token) throw new Error("kintone用のスクリプトプロパティが未設定です。");

  const labelToCode = {};
  const fields = verFetchJson_(`https://${sub}.cybozu.com/k/v1/app/form/fields.json?app=${app}`, token).properties || {};
  Object.keys(fields).forEach(code => {
    const key = verNormalize_(fields[code].label);
    if (!(key in labelToCode)) labelToCode[key] = code;
  });
  const override = props.getProperty("KLIST_FIELD_CODES");
  if (override) {
    const parsed = JSON.parse(override);
    Object.keys(parsed).forEach(label => { labelToCode[verNormalize_(label)] = parsed[label]; });
  }
  const code = label => VER_EMAIL_CODES[label] || labelToCode[verNormalize_(label)];

  const required = [VER_L_CONTRACT_TYPE, VER_L_STORE, VER_L_CUSTOMER, VER_L_VENDOR, VER_L_BURNABLE];
  const missing = required.filter(label => !code(label));
  if (missing.length > 0) throw new Error(`項目が見つかりません: ${missing.join(" / ")}`);
  const emailLabels = VER_L_EMAILS.filter(label => code(label));
  const notFound = VER_L_EMAILS.filter(label => !code(label));

  const wanted = ["$id", code(VER_L_CONTRACT_TYPE), code(VER_L_STORE), code(VER_L_CUSTOMER), code(VER_L_VENDOR), code(VER_L_BURNABLE)]
    .concat(emailLabels.map(code));
  const fieldsParam = wanted.map((c, i) => `fields%5B${i}%5D=${encodeURIComponent(c)}`).join("&");

  const records = [];
  let lastId = 0;
  while (true) {
    const query = encodeURIComponent(`$id > ${lastId} order by $id asc limit 500`);
    const page = verFetchJson_(`https://${sub}.cybozu.com/k/v1/records.json?app=${app}&query=${query}&${fieldsParam}`, token).records || [];
    if (page.length === 0) break;
    page.forEach(r => records.push(r));
    lastId = Number(page[page.length - 1].$id.value);
    if (page.length < 500) break;
  }

  const text = (record, c) => {
    const f = record[c];
    if (!f || f.value === null || f.value === undefined) return "";
    return Array.isArray(f.value) ? f.value.map(v => (typeof v === "object" ? (v.name || v.code || "") : String(v))).join(",") : String(f.value);
  };

  const contracts = [];
  records.forEach(record => {
    if (text(record, code(VER_L_CONTRACT_TYPE)) !== VER_TARGET_CONTRACT_TYPE) return;
    const burnable = record[code(VER_L_BURNABLE)];
    const options = burnable && Array.isArray(burnable.value) ? burnable.value.map(String) : [];
    contracts.push({
      id: record.$id.value,
      store: text(record, code(VER_L_STORE)),
      customer: text(record, code(VER_L_CUSTOMER)),
      vendor: text(record, code(VER_L_VENDOR)),
      weekdays: options.map(o => verNormalize_(o).charAt(0)),
      emailFields: emailLabels.map(label => ({ label: label, raw: text(record, code(label)) }))
    });
  });
  return { contracts: contracts, notFoundEmailLabels: notFound };
}

// 優先順位の順にアドレスを集める。独立した判定（リスト作成.js とは別の書き方）
function verCollectEmails_(emailFields) {
  const valid = [];
  const bad = [];
  emailFields.forEach(f => {
    String(f.raw).normalize("NFKC").toLowerCase().split(/[\s,;、，；]+/).filter(Boolean).forEach(a => {
      if (/^[a-z0-9._%+\-]+@[a-z0-9\-]+(\.[a-z0-9\-]+)*\.[a-z]{2,}$/.test(a)) {
        if (valid.indexOf(a) === -1) valid.push(a);
      } else {
        bad.push(`${f.label}: ${a}`);
      }
    });
  });
  return { valid: valid, bad: bad };
}

function verExpectedMarks_(weekdays) {
  return VER_DATES.map(d => {
    const key = `${d[0]}-${d[1]}-${d[2]}`;
    if (VER_NO_COLLECTION.indexOf(key) !== -1) return "";
    const w = VER_WEEKDAYS.charAt(new Date(d[0], d[1] - 1, d[2]).getDay());
    return weekdays.indexOf(w) !== -1 ? "○" : "";
  });
}

function verifyListSheet_(sheetName) {
  const ss = verGetSpreadsheet_();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`シート「${sheetName}」が見つかりません。`);
  const data = sheet.getDataRange().getDisplayValues();
  const results = [];
  const add = (level, kind, store, message, ref) => results.push([level, kind, store || "", message, ref || ""]);

  const kintone = verLoadKintone_();
  kintone.notFoundEmailLabels.forEach(label => add("警告", "項目", "", `アドレスの項目「${label}」がkintoneに見つかりません（この項目は照合していません）`));

  // 1. kintone側で、店舗名の重複・記号の問題を調べる
  const byStore = {};
  kintone.contracts.forEach(c => {
    (byStore[c.store] = byStore[c.store] || []).push(c);
    if (!c.store) add("エラー", "店舗名", "", `店舗名が空の契約があります（レコード番号 ${c.id}）`);
    if (/[【】]/.test(c.store)) add("エラー", "店舗名", c.store, "店舗名に【 または 】が含まれています。回答の読み取りができなくなります", `レコード番号 ${c.id}`);
  });
  Object.keys(byStore).forEach(store => {
    if (store && byStore[store].length > 1) {
      add("エラー", "店舗名の重複", store, `同じ店舗名の契約が ${byStore[store].length} 件あります。回答集計が混ざります`,
        byStore[store].map(c => "レコード番号 " + c.id).join(" / "));
    }
  });

  // 2. 各契約の「あるべき姿」を独立に判定する
  const expectedList = {};   // 店舗名 → { email, marks }
  const expectedReview = {}; // 店舗名 → 理由
  kintone.contracts.forEach(c => {
    const emails = verCollectEmails_(c.emailFields);
    if (emails.valid.length === 0) {
      expectedReview[c.store] = emails.bad.length > 0 ? "アドレスの形式が正しくない" : "アドレスなし";
      return;
    }
    if (c.weekdays.length === 0 || !VER_WEEKDAYS.split("").some(w => c.weekdays.indexOf(w) !== -1)) {
      expectedReview[c.store] = "可燃の回収曜日なし";
      return;
    }
    const marks = verExpectedMarks_(c.weekdays);
    if (marks.indexOf("○") === -1) {
      expectedReview[c.store] = "期間中に回収できる日なし";
      return;
    }
    expectedList[c.store] = { email: emails.valid[0], marks: marks, contract: c, emails: emails };
  });

  // 3. リストの各行を照合する
  const listed = {};
  const emailGroups = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const store = row[VER_COL_STORE];
    if (!store) continue;
    const rowLabel = `${i + 1}行目`;
    if (listed[store]) add("エラー", "リスト", store, "リストに同じ店舗が2行あります", rowLabel);
    listed[store] = true;

    const email = String(row[VER_COL_EMAIL]).trim();
    (emailGroups[email] = emailGroups[email] || []).push(store);
    const expected = expectedList[store];

    if (!expected) {
      add("エラー", "リスト", store, expectedReview[store]
        ? `リストにありますが、kintoneの内容では対象外のはずです（${expectedReview[store]}）`
        : "リストにありますが、kintoneの契約中に見つかりません", rowLabel);
      continue;
    }
    if (email !== expected.email) {
      add("エラー", "アドレス", store, `送信先が、優先順位で選ぶ結果と違います。リスト: ${email} / 期待: ${expected.email}`, rowLabel);
    }
    const actualMarks = row.slice(VER_COL_MARK_START, VER_COL_MARK_END + 1).map(m => (m === "○" ? "○" : m === "" ? "" : "?"));
    const expectedMarks = expected.marks.concat(Array(Math.max(0, actualMarks.length - expected.marks.length)).fill(""));
    if (actualMarks.join("|") !== expectedMarks.join("|")) {
      add("エラー", "回収曜日", store, `希望日の「○」が、回収曜日から求めた結果と違います。リスト: ${actualMarks.join("")} / 期待: ${expectedMarks.join("")}`, rowLabel);
    }
    if (expected.marks.length >= 3 && row[VER_COL_MARK_START + 2] === "○") {
      add("エラー", "回収曜日", store, "回収できない日（1/1）に「○」が付いています", rowLabel);
    }
    if (String(row[VER_COL_VENDOR]).trim() !== String(expected.contract.vendor).trim()) {
      add("警告", "業者", store, `業者名がkintoneと違います。リスト: ${row[VER_COL_VENDOR]} / kintone: ${expected.contract.vendor}`, rowLabel);
    }
    if (!expected.contract.vendor) add("情報", "業者", store, "業者名が空です（料金は標準額になります）", rowLabel);
    if (expected.emails.bad.length > 0) {
      add("情報", "アドレス", store, `形式が正しくない入力があります（使用していません）: ${expected.emails.bad.join(" / ")}`, rowLabel);
    }
  }
  Object.keys(expectedList).forEach(store => {
    if (!listed[store]) add("エラー", "リスト", store, "kintoneの内容では対象のはずですが、リストにありません");
  });

  // 4. 要確認シートとの突き合わせ
  const review = ss.getSheetByName(VER_REVIEW_SHEET_NAME);
  const reviewed = {};
  if (review && review.getLastRow() > 1) {
    review.getDataRange().getDisplayValues().slice(1).forEach(row => { reviewed[row[1]] = row[2]; });
  }
  const reviewStores = Object.keys(expectedReview);
  reviewStores.forEach(store => {
    if (!(store in reviewed)) add("警告", "要確認", store, `対象外（${expectedReview[store]}）ですが、「要確認」シートにありません`);
  });
  Object.keys(reviewed).forEach(store => {
    if (!(store in expectedReview) && !listed[store]) add("警告", "要確認", store, "「要確認」にありますが、今の照合では対象外の理由が見つかりません（kintoneが更新された可能性）");
  });

  // 5. アドレス単位のチェック
  Object.keys(emailGroups).forEach(email => {
    const stores = emailGroups[email];
    if (stores.length > VER_MAX_STORES_PER_FORM) {
      add("警告", "フォーム", "", `同じアドレス ${email} に ${stores.length} 店舗が入っています。1つのフォームに入る質問が多くなります`);
    }
    const customers = {};
    stores.forEach(store => {
      const exp = expectedList[store];
      if (exp) customers[exp.contract.customer] = true;
    });
    const customerNames = Object.keys(customers);
    if (customerNames.length > 1) {
      add("警告", "別の契約者", "", `同じアドレス ${email} に、別の契約者（${customerNames.join(" / ")}）の店舗が入っています。ほかの契約者の店舗が見えます`, stores.join(" / "));
    }
    const domain = email.split("@")[1] || "";
    if (VER_TYPO_DOMAINS.indexOf(domain) !== -1 || /\.(con|cojp|c0m|comm)$/.test(domain)) {
      add("警告", "アドレス", stores.join(" / "), `アドレスの入力ミスの疑い: ${email}`);
    }
    if (!/^[\x21-\x7e]+$/.test(email)) {
      add("エラー", "アドレス", stores.join(" / "), `アドレスに全角や使えない文字が含まれています: ${email}`);
    }
  });

  verWriteResult_(ss, sheetName, `kintoneの契約中: ${kintone.contracts.length} 件 / リストの店舗: ${Object.keys(listed).length} 件`, results);
}

function verWriteResult_(ss, sheetName, countLine, results) {
  const order = { "エラー": 0, "警告": 1, "情報": 2 };
  results.sort((a, b) => order[a[0]] - order[b[0]]);
  const counts = { "エラー": 0, "警告": 0, "情報": 0 };
  results.forEach(r => { counts[r[0]]++; });

  let out = ss.getSheetByName(VER_RESULT_SHEET_NAME);
  if (!out) out = ss.insertSheet(VER_RESULT_SHEET_NAME);
  out.clear();
  const summary = [`検証した対象: 「${sheetName}」`, `日時: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss")}`,
    countLine, `エラー ${counts["エラー"]} 件 / 警告 ${counts["警告"]} 件 / 情報 ${counts["情報"]} 件`];
  const header = ["重大度", "区分", "店舗名", "内容", "参考"];
  const rows = summary.map(s => [s, "", "", "", ""]).concat([header], results.length > 0 ? results : [["問題なし", "", "", "検出された食い違いはありません", ""]]);
  out.getRange(1, 1, rows.length, 5).setValues(rows);
  out.getRange(summary.length + 1, 1, 1, 5).setFontWeight("bold");
  console.log(`検証結果（${sheetName}）: エラー ${counts["エラー"]} 件 / 警告 ${counts["警告"]} 件 / 情報 ${counts["情報"]} 件。シート「${VER_RESULT_SHEET_NAME}」を確認してください。`);
}

// 作成したフォームの抜き取り検査と、フォルダとの照合。フォームを作ったあとに実行する
function verifyCreatedForms() {
  const ss = verGetSpreadsheet_();
  const sheet = ss.getSheetByName(VER_LIST_SHEET_NAME);
  if (!sheet) throw new Error(`シート「${VER_LIST_SHEET_NAME}」が見つかりません。`);
  const data = sheet.getDataRange().getDisplayValues();
  const headers = data[0];
  const results = [];
  const add = (level, kind, store, message, ref) => results.push([level, kind, store || "", message, ref || ""]);

  const formsById = {};
  const forms = [];
  let sentRows = 0;
  let pendingRows = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[VER_COL_STATUS];
    if (status === "送信済") {
      sentRows++;
      if (!row[VER_COL_ID] || !row[VER_COL_URL]) {
        add("エラー", "フォーム", row[VER_COL_STORE], "フォーム作成状況は「送信済」ですが、URLまたはIDが空です", `${i + 1}行目`);
        continue;
      }
      if (!formsById[row[VER_COL_ID]]) {
        formsById[row[VER_COL_ID]] = { id: row[VER_COL_ID], email: row[VER_COL_EMAIL], stores: [] };
        forms.push(formsById[row[VER_COL_ID]]);
      }
      formsById[row[VER_COL_ID]].stores.push({
        name: row[VER_COL_STORE],
        dates: row.slice(VER_COL_MARK_START, VER_COL_MARK_END + 1).map((m, k) => (m === "○" ? headers[VER_COL_MARK_START + k] : "")).filter(Boolean)
      });
    } else if (status === "エラー") {
      add("エラー", "フォーム", row[VER_COL_STORE], "フォーム作成がエラーになっています", `${i + 1}行目`);
    } else if (row[VER_COL_STORE]) {
      pendingRows++;
    }
  }
  if (pendingRows > 0) add("情報", "進捗", "", `フォームがまだ作られていない店舗が ${pendingRows} 件あります`);

  // 抜き取り検査（先頭・末尾と、ランダムに選んだもの）
  const picked = [];
  if (forms.length > 0) {
    picked.push(forms[0], forms[forms.length - 1]);
    const shuffled = forms.slice(1, -1).sort(() => Math.random() - 0.5);
    shuffled.slice(0, Math.max(0, VER_FORM_SAMPLE_COUNT - picked.length)).forEach(f => picked.push(f));
  }
  const startTime = Date.now();
  let checked = 0;
  picked.forEach(f => {
    if (Date.now() - startTime > 4 * 60 * 1000) return;
    try {
      const form = FormApp.openById(f.id);
      checked++;
      const label = f.stores.map(s => s.name).join("・");
      if (!form.isAcceptingResponses()) add("警告", "フォーム", label, "フォームの受付が止まっています", f.id);
      const description = form.getDescription();
      f.stores.forEach(s => {
        if (description.indexOf(s.name) === -1) add("エラー", "フォーム", s.name, "フォームの説明文に店舗名がありません", f.id);
      });
      const titles = form.getItems().map(item => item.getTitle());
      f.stores.forEach(s => {
        if (!titles.some(t => t === `【${s.name}】の収集希望について`)) add("エラー", "フォーム", s.name, "店舗の見出しがフォームにありません", f.id);
        s.dates.forEach(d => {
          if (titles.indexOf(`【${s.name}】${d}の収集を希望しますか？`) === -1) add("エラー", "フォーム", s.name, `「${d}」の質問がフォームにありません`, f.id);
        });
      });
      const expectedQuestions = f.stores.reduce((sum, s) => sum + s.dates.length, 0);
      const actualQuestions = titles.filter(t => /^【.+】.+の収集を希望しますか？$/.test(t)).length;
      if (expectedQuestions !== actualQuestions) {
        add("エラー", "フォーム", label, `日付の質問の数が違います。期待 ${expectedQuestions} / 実際 ${actualQuestions}`, f.id);
      }
    } catch (e) {
      add("エラー", "フォーム", f.stores.map(s => s.name).join("・"), `フォームを開けません: ${e}`, f.id);
    }
  });

  // フォルダとの照合
  const folderId = PropertiesService.getScriptProperties().getProperty("FORM_FOLDER_ID");
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId.trim());
      const inFolder = {};
      const files = folder.getFilesByType(MimeType.GOOGLE_FORMS);
      while (files.hasNext()) inFolder[files.next().getId()] = true;
      forms.forEach(f => {
        if (!inFolder[f.id]) add("エラー", "フォルダ", f.stores.map(s => s.name).join("・"), "フォームが指定のフォルダにありません", f.id);
      });
      Object.keys(inFolder).forEach(id => {
        if (!formsById[id]) add("警告", "フォルダ", "", "リストにないフォームがフォルダにあります（テストで作ったものなど）。不要なら削除してください", id);
      });
    } catch (e) {
      add("警告", "フォルダ", "", `フォルダを確認できませんでした: ${e}`);
    }
  } else {
    add("情報", "フォルダ", "", "FORM_FOLDER_ID が未設定のため、フォルダとの照合はしていません");
  }

  verWriteResult_(ss, `作成済みフォーム（抜き取り ${checked} 件）`, `フォーム作成済みの店舗: ${sentRows} 件 / フォーム数: ${forms.length} 件`, results);
}

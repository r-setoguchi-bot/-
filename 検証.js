// ===== 二重チェック：kintoneのデータと作成したリスト・フォームを、別の処理で照合する =====
// リスト作成.js とは別の処理で、kintoneの契約管理をもう一度読み直し、アドレス・回収曜日・店舗名などを独立に判定して、
// 「リスト」（または「リスト_kintone」）と食い違いがないかを調べる。何も書き換えず、結果は「検証結果」シートに出す。
//
// 使い方:
//   1. verifyKintoneListData … 「リスト_kintone」を調べる（リストを差し替える前に実行）
//   2. verifyListData        … 「リスト」を調べる（差し替えたあと、フォームを作る前に実行）
//   3. verifyCreatedForms    … 作成したフォームを抜き取りで調べ、フォルダの中身とも照合する（フォームを作ったあと）
//   結果は「検証結果_○○」シート（対象ごとに別のタブ）の「重大度」で見る。エラー = 直してから進める、警告 = 内容を見て判断、情報 = 参考
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

// 送信先に使わないアドレス（リスト作成.js の KLIST_EXCLUDED_EMAIL_PATTERNS と同じ内容）
const VER_EXCLUDED_EMAIL_PATTERNS = [/^dummy/];

// 「リスト」の列（0始まり）
const VER_COL_STORE = 0;
const VER_COL_EMAIL = 1;
const VER_COL_VENDOR = 2;
const VER_COL_MARK_START = 3;   // D〜H列
const VER_COL_MARK_END = 7;
const VER_COL_URL = 8;
const VER_COL_ID = 9;
const VER_COL_STATUS = 10;
const VER_COL_CUSTOMER = 16;   // Q列: 契約者
const VER_COL_ADDRESS = 17;    // R列: 店舗住所

// 店舗の住所の項目（フィールドコード）
const VER_ADDRESS_CODES = { postal: "契約郵便番号", address: "契約住所連結", prefecture: "契約都道府県", city: "契約市区町村", building: "契約ビル名ほか" };

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
    .concat(emailLabels.map(code))
    .concat(Object.keys(VER_ADDRESS_CODES).map(key => VER_ADDRESS_CODES[key]));
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
      address: verBuildAddress_(record, text),
      emailFields: emailLabels.map(label => ({ label: label, raw: text(record, code(label)) }))
    });
  });
  return { contracts: contracts, notFoundEmailLabels: notFound };
}

// 店舗の住所（「〒郵便番号 住所」）。リスト作成.js とは別に書いた判定
function verBuildAddress_(record, text) {
  const get = key => text(record, VER_ADDRESS_CODES[key]).normalize("NFKC").trim();
  const postal = get("postal");
  let address = get("address");
  if (!address) address = [get("prefecture"), get("city"), get("building")].filter(Boolean).join("");
  if (!address) return "";
  const digits = postal.replace(/\D/g, "");
  const hasPostal = digits !== "" && address.replace(/\D/g, "").indexOf(digits) === 0;
  return postal && !hasPostal ? `〒${postal} ${address}` : address;
}

// 優先順位の順にアドレスを集める。独立した判定（リスト作成.js とは別の書き方）
function verCollectEmails_(emailFields) {
  const valid = [];
  const bad = [];
  emailFields.forEach(f => {
    String(f.raw).normalize("NFKC").toLowerCase().split(/[\s,;、，；]+/).filter(Boolean).forEach(item => {
      const a = item.replace(/^mailto:/, "");
      if (VER_EXCLUDED_EMAIL_PATTERNS.some(pattern => pattern.test(a))) {
        bad.push(`${f.label}: ${a}（ダミーなど）`);
      } else if (/^[a-z0-9._%+\-]+@[a-z0-9\-]+(\.[a-z0-9\-]+)*\.[a-z]{2,}$/.test(a)) {
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
  const formGroups = {};   // アドレスと契約者の組 → 店舗数
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
    const formKey = email + "\u0001" + String(row[VER_COL_CUSTOMER] || "").trim();
    formGroups[formKey] = (formGroups[formKey] || 0) + 1;

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
    if (String(row[VER_COL_CUSTOMER] || "").trim() !== String(expected.contract.customer).trim()) {
      add("エラー", "契約者", store, `契約者がkintoneと違います。リスト: ${row[VER_COL_CUSTOMER]} / kintone: ${expected.contract.customer}（Q列を確認）`, rowLabel);
    }
    const listedAddress = String(row[VER_COL_ADDRESS] || "").trim();
    if (!expected.contract.address) {
      add("警告", "住所", store, "kintoneに店舗の住所がありません。フォームの店舗ページに住所が出ません", rowLabel);
    } else if (listedAddress !== expected.contract.address) {
      add("エラー", "住所", store, `店舗住所がkintoneと違います。リスト: ${listedAddress || "（空）"} / kintone: ${expected.contract.address}（R列を確認）`, rowLabel);
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
    const customers = {};
    stores.forEach(store => {
      const exp = expectedList[store];
      if (exp) customers[exp.contract.customer] = true;
    });
    const customerNames = Object.keys(customers);
    if (customerNames.length > 1) {
      add("情報", "複数の契約者", "", `同じアドレス ${email} に ${customerNames.length} 社の契約者があります。契約者ごとに別のフォーム・別のメールになります`, customerNames.join(" / "));
    }
    const domain = email.split("@")[1] || "";
    if (VER_TYPO_DOMAINS.indexOf(domain) !== -1 || /\.(con|cojp|c0m|comm)$/.test(domain)) {
      add("警告", "アドレス", stores.join(" / "), `アドレスの入力ミスの疑い: ${email}`);
    }
    if (!/^[\x21-\x7e]+$/.test(email)) {
      add("エラー", "アドレス", stores.join(" / "), `アドレスに全角や使えない文字が含まれています: ${email}`);
    }
  });

  Object.keys(formGroups).forEach(key => {
    if (formGroups[key] > VER_MAX_STORES_PER_FORM) {
      add("警告", "フォーム", "", `アドレスと契約者の組 ${key.replace("\u0001", " / ")} に ${formGroups[key]} 店舗が入っています。1つのフォームに入る質問が多くなります`);
    }
  });

  verWriteResult_(ss, sheetName, `kintoneの契約中: ${kintone.contracts.length} 件 / リストの店舗: ${Object.keys(listed).length} 件 / 作成されるフォーム: ${Object.keys(formGroups).length} 件`, results);
}

function verWriteResult_(ss, sheetName, countLine, results) {
  const order = { "エラー": 0, "警告": 1, "情報": 2 };
  results.sort((a, b) => order[a[0]] - order[b[0]]);
  const counts = { "エラー": 0, "警告": 0, "情報": 0 };
  results.forEach(r => { counts[r[0]]++; });

  // 検証の対象ごとに別のタブに出す（「リスト」の検証結果を、フォームの検証結果で上書きしないため）
  const tabName = sheetName.indexOf("作成済みフォーム") === 0 ? VER_RESULT_SHEET_NAME + "_フォーム" : VER_RESULT_SHEET_NAME + "_" + sheetName;
  let out = ss.getSheetByName(tabName);
  if (!out) out = ss.insertSheet(tabName);
  out.clear();
  const summary = [`検証した対象: 「${sheetName}」`, `日時: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss")}`,
    countLine, `エラー ${counts["エラー"]} 件 / 警告 ${counts["警告"]} 件 / 情報 ${counts["情報"]} 件`];
  const header = ["重大度", "区分", "店舗名", "内容", "参考"];
  const rows = summary.map(s => [s, "", "", "", ""]).concat([header], results.length > 0 ? results : [["問題なし", "", "", "検出された食い違いはありません", ""]]);
  out.getRange(1, 1, rows.length, 5).setValues(rows);
  out.getRange(summary.length + 1, 1, 1, 5).setFontWeight("bold");
  console.log(`検証結果（${sheetName}）: エラー ${counts["エラー"]} 件 / 警告 ${counts["警告"]} 件 / 情報 ${counts["情報"]} 件。シート「${tabName}」を確認してください。`);
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
        address: String(row[VER_COL_ADDRESS] || "").trim(),
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
      const items = form.getItems();
      const titles = items.map(item => item.getTitle());
      f.stores.forEach(s => {
        const header = items.filter(item => item.getTitle() === `【${s.name}】の収集希望について`)[0];
        if (header && s.address && header.getHelpText().indexOf(s.address) === -1) {
          add("エラー", "フォーム", s.name, "店舗ページの見出しに、住所が入っていません", f.id);
        }
      });
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

// ===== 全件検証：作成したすべてのフォームの中身を、1件ずつ開いて確認する =====
// verifyCreatedForms は20件の抜き取りだが、こちらは全フォームを対象にする（約900件で、数十分。6分の制限は、時間主導トリガーで自動で続ける）。
// 「リスト」の内容から「あるべきフォームの形」を作り、実際のフォームの項目を、上から1つずつ照合する。何も書き換えない。
//   確認する内容: フォームの題名、説明文（契約者・店舗名・期間・期限）、項目の並びと数、必須の設定、
//                 店舗ごとのページ・見出し・住所・料金、各日の質問と選択肢、受付が止まっていないか
//   最後に、フォルダとの照合（リストのフォームが、フォルダにあるか）と、リスト側の確認（エラー行・未作成）も行う。
// 使い方:
//   1. startFullFormVerification を1回だけ実行（終わるまで自動で続く。途中で止めるときは stopFullFormVerification）
//   2. 「検証結果_フォーム全件」シートの「重大度」を見る。エラー = 直す、警告 = 内容を見て判断、情報 = 参考
//   ※ フォームを作り終えたあとに実行する（作成の途中だと、「まだ作られていない」と出る）

const VER_FULL_SHEET_NAME = "検証結果_フォーム全件";
const VER_FULL_CURSOR_PROPERTY = "VER_FULL_CURSOR";
const VER_FULL_HANDLER = "continueFullFormVerification";
const VER_FULL_TIME_LIMIT_MS = 4.5 * 60 * 1000;
const VER_FULL_RETRY_DELAY_MS = 60 * 1000;
const VER_FULL_SAFETY_DELAY_MS = 7 * 60 * 1000;   // 実行が強制終了した場合に再開するまでの待ち時間
const VER_FULL_HEADER = ["重大度", "区分", "店舗名", "内容", "参考"];

// あるべきフォームの内容（フォーム作成.js とは別に、ここにも書いて照合する）
const VER_EXPECT_TITLE = "【株式会社クリメン】2026年～2027年 年末年始廃棄物回収に関するアンケート";
const VER_EXPECT_PERIOD = "2026年12月30日(水)から2027年1月3日(日)";
const VER_EXPECT_DEADLINE = "2026年11月15日(日)";
const VER_Q_CONTACT_NAME = "ご担当者様のお名前";
const VER_Q_PHONE = "ご連絡可能なお電話番号";
const VER_Q_COMMENT = "ご不明点・ご要望（任意）";
const VER_CHOICES = ["希望する", "希望しない"];
const VER_DEFAULT_FEE = "7,000円";
const VER_VENDOR_FEES = {
  "有限会社長澤商事": "10,000円",
  "株式会社光栄和": "20,000円",
  "鍵本産業株式会社": "10,000円"
};

function startFullFormVerification() {
  verDeleteFullTriggers_();
  PropertiesService.getScriptProperties().deleteProperty(VER_FULL_CURSOR_PROPERTY);
  const ss = verGetSpreadsheet_();
  let sheet = ss.getSheetByName(VER_FULL_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(VER_FULL_SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, VER_FULL_HEADER.length).setValues([VER_FULL_HEADER]);
  continueFullFormVerification();
}

function stopFullFormVerification() {
  verDeleteFullTriggers_();
  PropertiesService.getScriptProperties().deleteProperty(VER_FULL_CURSOR_PROPERTY);
  console.log("全件検証を停止しました。");
}

// 続きの実行を予約する。Googleの一時的なサーバーエラーで失敗することがあるので、少し待って、数回やり直す
function verScheduleFullTrigger_(delayMs) {
  let lastError;
  for (let attempt = 0; attempt < 5; attempt++) {
    try {
      return ScriptApp.newTrigger(VER_FULL_HANDLER).timeBased().after(delayMs).create();
    } catch (e) {
      lastError = e;
      console.warn(`続きの予約に失敗しました（${attempt + 1}回目）: ${e}`);
      Utilities.sleep(2000 * Math.pow(2, attempt));   // 2, 4, 8, 16, 32秒
    }
  }
  throw lastError;
}

function verDeleteFullTriggers_() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === VER_FULL_HANDLER) ScriptApp.deleteTrigger(trigger);
  });
}

// 「リスト」から、作成済みフォームごとの店舗（住所・業者・各日）を取り出す
function verGroupFormsForFull_(data) {
  const headers = data[0];
  const forms = [];
  const byId = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[VER_COL_STATUS] !== "送信済" || !row[VER_COL_ID] || !row[VER_COL_URL]) continue;
    if (!byId[row[VER_COL_ID]]) {
      byId[row[VER_COL_ID]] = { id: row[VER_COL_ID], email: row[VER_COL_EMAIL], customer: String(row[VER_COL_CUSTOMER] || "").trim(), stores: [] };
      forms.push(byId[row[VER_COL_ID]]);
    }
    byId[row[VER_COL_ID]].stores.push({
      name: row[VER_COL_STORE],
      vendor: String(row[VER_COL_VENDOR] || "").trim(),
      address: String(row[VER_COL_ADDRESS] || "").trim(),
      dates: row.slice(VER_COL_MARK_START, VER_COL_MARK_END + 1).map((m, k) => (m === "○" ? headers[VER_COL_MARK_START + k] : "")).filter(Boolean)
    });
  }
  return forms;
}

function continueFullFormVerification() {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) {
    console.log("別の検証が処理中のためスキップしました。");
    return;
  }
  const startTime = Date.now();
  try {
    verDeleteFullTriggers_();
    const ss = verGetSpreadsheet_();
    const sheet = ss.getSheetByName(VER_FULL_SHEET_NAME);
    if (!sheet) throw new Error("startFullFormVerification から実行してください。");
    const listSheet = ss.getSheetByName(VER_LIST_SHEET_NAME);
    if (!listSheet) throw new Error(`シート「${VER_LIST_SHEET_NAME}」が見つかりません。`);
    const data = listSheet.getDataRange().getDisplayValues();
    const forms = verGroupFormsForFull_(data);

    // 強制終了や、続きの予約の失敗に備えた、保険の再開トリガー
    const safetyTrigger = verScheduleFullTrigger_(VER_FULL_SAFETY_DELAY_MS);

    const props = PropertiesService.getScriptProperties();
    let index = Number(props.getProperty(VER_FULL_CURSOR_PROPERTY) || 0);
    if (index > forms.length) index = 0;
    const rows = [];
    while (index < forms.length) {
      if (Date.now() - startTime > VER_FULL_TIME_LIMIT_MS) break;
      verCheckOneForm_(forms[index], rows);
      index++;
    }
    verAppendFullRows_(sheet, rows);

    if (index < forms.length) {
      props.setProperty(VER_FULL_CURSOR_PROPERTY, String(index));
      verScheduleFullTrigger_(VER_FULL_RETRY_DELAY_MS);
      try {
        ScriptApp.deleteTrigger(safetyTrigger);
      } catch (e) {
        console.warn(`保険の予約を消せませんでした（処理には影響しません）: ${e}`);
      }
      console.log(`${index} / ${forms.length} 件のフォームを確認しました。残りは自動で続行します。`);
    } else {
      verDeleteFullTriggers_();
      props.deleteProperty(VER_FULL_CURSOR_PROPERTY);
      verFinishFull_(ss, sheet, data, forms);
    }
  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      console.warn(`ロック解除に失敗しました（処理には影響しません）: ${e}`);
    }
  }
}

function verAppendFullRows_(sheet, rows) {
  if (rows.length === 0) return;
  const start = sheet.getLastRow() + 1;
  const needed = start + rows.length - 1;
  if (needed > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), needed - sheet.getMaxRows());
  sheet.getRange(start, 1, rows.length, VER_FULL_HEADER.length).setValues(rows);
}

// 1つのフォームを、あるべき形と、上から1項目ずつ照合する
function verCheckOneForm_(f, rows) {
  const label = f.stores.map(s => s.name).join("・");
  const add = (level, kind, store, message) => rows.push([level, kind, store || label, message, f.id]);

  let form = null;
  for (let attempt = 0; attempt < 2 && !form; attempt++) {
    try {
      form = FormApp.openById(f.id);
    } catch (e) {
      if (attempt === 1) {
        add("エラー", "フォーム", "", `フォームを開けません: ${e}`);
        return;
      }
      Utilities.sleep(1500);
    }
  }

  try {
    if (form.getTitle() !== VER_EXPECT_TITLE) add("エラー", "題名", "", `フォームの題名が違います: ${form.getTitle()}`);
    if (!form.isAcceptingResponses()) add("警告", "受付", "", "フォームの受付が止まっています");

    // 説明文
    const description = form.getDescription();
    const mustHave = [["期間", VER_EXPECT_PERIOD], ["回答・変更の期限", VER_EXPECT_DEADLINE], ["1月1日の注意", "※1月1日は回収不可"]];
    if (f.customer) mustHave.push(["契約者", `【ご契約者】\n${f.customer}`]);
    f.stores.forEach(s => mustHave.push([`店舗名（${s.name}）`, s.name]));
    mustHave.forEach(pair => {
      if (description.indexOf(pair[1]) === -1) add("エラー", "説明文", "", `説明文に「${pair[0]}」がありません`);
    });

    // あるべき項目の並び
    const T = FormApp.ItemType;
    const expected = [
      { type: T.TEXT, title: VER_Q_CONTACT_NAME, required: true },
      { type: T.TEXT, title: VER_Q_PHONE, required: true },
      { type: T.PARAGRAPH_TEXT, title: VER_Q_COMMENT, required: false }
    ];
    f.stores.forEach((s, i) => {
      const fee = VER_VENDOR_FEES[s.vendor] || VER_DEFAULT_FEE;
      expected.push({ type: T.PAGE_BREAK, store: s.name, title: f.stores.length > 1 ? `希望日の選択（${i + 1}/${f.stores.length}店舗目）` : "希望日の選択" });
      expected.push({ type: T.SECTION_HEADER, store: s.name, title: `【${s.name}】の収集希望について`,
        help: `ご住所：${s.address || "（登録がありません）"}\n特別回収料金：${fee}/1日` });
      s.dates.forEach(d => expected.push({ type: T.MULTIPLE_CHOICE, store: s.name, title: `【${s.name}】${d}の収集を希望しますか？`, required: true, choices: VER_CHOICES }));
    });

    const items = form.getItems();
    if (items.length !== expected.length) add("エラー", "項目の数", "", `項目の数が違います。期待 ${expected.length} / 実際 ${items.length}`);

    let mismatches = 0;
    for (let i = 0; i < Math.min(items.length, expected.length) && mismatches < 3; i++) {
      const item = items[i];
      const exp = expected[i];
      const store = exp.store || "";
      const where = `${i + 1}番目の項目`;
      if (item.getType() !== exp.type) {
        add("エラー", "項目", store, `${where}の種類が違います（期待「${exp.title}」/ 実際「${item.getTitle()}」）`);
        mismatches++;
        continue;
      }
      if (item.getTitle() !== exp.title) {
        add("エラー", "項目", store, `${where}の題名が違います。期待「${exp.title}」/ 実際「${item.getTitle()}」`);
        mismatches++;
        continue;
      }
      if (exp.type === T.SECTION_HEADER && item.getHelpText() !== exp.help) {
        add("エラー", "住所・料金", store, `店舗ページの見出しの説明が違います。期待「${exp.help.replace(/\n/g, " / ")}」/ 実際「${item.getHelpText().replace(/\n/g, " / ")}」`);
        mismatches++;
      } else if (exp.type === T.MULTIPLE_CHOICE) {
        const mc = item.asMultipleChoiceItem();
        const choices = mc.getChoices().map(c => c.getValue());
        if (choices.join("|") !== exp.choices.join("|")) {
          add("エラー", "選択肢", store, `${where}の選択肢が違います: ${choices.join(" / ")}`);
          mismatches++;
        } else if (!mc.isRequired()) {
          add("エラー", "必須", store, `${where}（${exp.title}）が必須になっていません`);
          mismatches++;
        }
      } else if (exp.type === T.TEXT && item.asTextItem().isRequired() !== exp.required) {
        add("エラー", "必須", "", `「${exp.title}」の必須の設定が違います`);
        mismatches++;
      } else if (exp.type === T.PARAGRAPH_TEXT && item.asParagraphTextItem().isRequired() !== exp.required) {
        add("エラー", "必須", "", `「${exp.title}」の必須の設定が違います`);
        mismatches++;
      }
    }
  } catch (e) {
    add("エラー", "フォーム", "", `検査中にエラーが出ました: ${e}`);
  }
}

// 全フォームの確認が終わったあと、フォルダとの照合・リスト側の確認をして、結果をまとめる
function verFinishFull_(ss, sheet, data, forms) {
  const extra = [];
  const add = (level, kind, store, message, ref) => extra.push([level, kind, store || "", message, ref || ""]);

  // リスト側の確認
  let pending = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[VER_COL_STORE]) continue;
    const status = row[VER_COL_STATUS];
    if (status === "エラー") add("エラー", "フォーム", row[VER_COL_STORE], "フォーム作成がエラーになっています", `${i + 1}行目`);
    else if (status === "送信済" && (!row[VER_COL_ID] || !row[VER_COL_URL])) add("エラー", "フォーム", row[VER_COL_STORE], "「送信済」ですが、URLまたはIDが空です", `${i + 1}行目`);
    else if (status !== "送信済") pending++;
  }
  if (pending > 0) add("情報", "進捗", "", `フォームがまだ作られていない店舗が ${pending} 件あります（作成が終わっていない可能性があります）`);

  // 同じフォームIDが、別のお客様（アドレス）の行に使われていないか
  forms.forEach(f => {
    const emails = {};
    data.slice(1).forEach(row => { if (row[VER_COL_ID] === f.id) emails[String(row[VER_COL_EMAIL]).trim()] = true; });
    if (Object.keys(emails).length > 1) add("エラー", "フォームの重複", f.stores.map(s => s.name).join("・"), "同じフォームが、別々のアドレスの行に入っています", f.id);
  });

  // フォルダとの照合
  const folderId = PropertiesService.getScriptProperties().getProperty("FORM_FOLDER_ID");
  let folderCount = "（確認せず）";
  if (folderId) {
    try {
      const inFolder = {};
      const files = DriveApp.getFolderById(folderId.trim()).getFilesByType(MimeType.GOOGLE_FORMS);
      while (files.hasNext()) inFolder[files.next().getId()] = true;
      folderCount = Object.keys(inFolder).length;
      const byId = {};
      forms.forEach(f => {
        byId[f.id] = true;
        if (!inFolder[f.id]) add("エラー", "フォルダ", f.stores.map(s => s.name).join("・"), "フォームが指定のフォルダにありません", f.id);
      });
      Object.keys(inFolder).forEach(id => {
        if (!byId[id]) add("警告", "フォルダ", "", "リストにないフォームがフォルダにあります（テストで作ったものなど）。不要なら削除してください", id);
      });
    } catch (e) {
      add("警告", "フォルダ", "", `フォルダを確認できませんでした: ${e}`);
    }
  }

  // これまでの結果とまとめて、重大度の順に並べて書き直す
  const found = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, VER_FULL_HEADER.length).getValues() : [];
  const results = found.concat(extra);
  const order = { "エラー": 0, "警告": 1, "情報": 2 };
  results.sort((a, b) => order[a[0]] - order[b[0]]);
  const counts = { "エラー": 0, "警告": 0, "情報": 0 };
  results.forEach(r => { counts[r[0]]++; });
  const formsWithError = {};
  results.filter(r => r[0] === "エラー" && r[4]).forEach(r => { formsWithError[r[4]] = true; });

  const summary = [
    "検証した対象: 作成済みのすべてのフォーム（全件）",
    `日時: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss")}`,
    `確認したフォーム: ${forms.length} 件 / フォルダ内のフォーム: ${folderCount} 件 / エラーがあったフォーム: ${Object.keys(formsWithError).length} 件`,
    `エラー ${counts["エラー"]} 件 / 警告 ${counts["警告"]} 件 / 情報 ${counts["情報"]} 件`
  ];
  const body = results.length > 0 ? results : [["問題なし", "", "", "すべてのフォームが、あるべき形になっています", ""]];
  const values = summary.map(s => [s, "", "", "", ""]).concat([VER_FULL_HEADER], body);
  sheet.clear();
  sheet.getRange(1, 1, values.length, VER_FULL_HEADER.length).setValues(values);
  sheet.getRange(summary.length + 1, 1, 1, VER_FULL_HEADER.length).setFontWeight("bold");
  console.log(`全件検証が終わりました（フォーム ${forms.length} 件）: エラー ${counts["エラー"]} 件 / 警告 ${counts["警告"]} 件 / 情報 ${counts["情報"]} 件。シート「${VER_FULL_SHEET_NAME}」を確認してください。`);
}

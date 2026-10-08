// ===== kintone契約管理 → フォーム作成用リストの作成 =====
// kintoneの契約管理アプリから「契約中」の契約を取得し、フォーム作成用の「リスト」と同じ形のシートを作る。
// 作れない契約は「要確認」シートに理由つきで出力する。フォームは作らない（メールも送らない）。
//
// 使い方:
//   1. dumpKeiyakuFields を実行して項目名とフィールドコードの対応を確認（初回のみ）
//   2. buildListFromKintone を実行 → 「リスト_kintone」「要確認」シートができる
//   3. 内容を確認してから、フォーム作成側の「リスト」と差し替える
//
// 必要なスクリプトプロパティ（コード.js と同じ名前）:
//   KINTONE_SUBDOMAIN / KINTONE_KEIYAKU_APP_ID / KINTONE_KEIYAKU_API_TOKEN
//   ※ APIトークンには「レコード閲覧」と、項目名の取得用に「アプリ管理」の権限が必要。
//     「アプリ管理」を付けない場合は、スクリプトプロパティ KLIST_FIELD_CODES に
//     {"契約種別":"フィールドコード", ...} のJSONで項目名→コードを登録する。
//
// スプレッドシートの指定:
//   スクリプトプロパティ FORM_SPREADSHEET_URL があればそのスプレッドシートを使う。
//   無ければスクリプトが紐づいているスプレッドシート（コンテナバインド時）を使う。
// このファイルだけで動く（フォーム作成.js が無くても実行できる）。

const KLIST_SOURCE_SHEET_NAME = "リスト";       // 作成済みフォームの状況を引き継ぐ元
const KLIST_STATUS_SENT = "送信済";
const KLIST_OUTPUT_SHEET_NAME = "リスト_kintone";
const KLIST_REVIEW_SHEET_NAME = "要確認";
const KLIST_TARGET_CONTRACT_TYPE = "契約中";

// 特別回収期間（フォームの列 D〜K に対応する8日分）と、回収不可の日
const KLIST_DATES = [
  [2026, 12, 28], [2026, 12, 29], [2026, 12, 30], [2026, 12, 31],
  [2027, 1, 1], [2027, 1, 2], [2027, 1, 3], [2027, 1, 4]
];
const KLIST_NO_COLLECTION_DATES = ["2027-1-1"];

const KLIST_WEEKDAYS = "日月火水木金土"; // Date#getDay() の順

const KLIST_LABELS = {
  contractType: "契約種別",
  storeName: "契約店舗名称",
  vendor: "収集業者：名称",
  replyEmail: "年末年始　回答アドレス①",
  guideEmail: "案内等送付先メールアドレス:(to)"
};
// 可燃の回収曜日は、チェックボックス1項目（選択肢が曜日）に入っている
const KLIST_BURNABLE_LABEL = "可燃収集日";

const KLIST_EMAIL_PATTERN = /^[a-z0-9._%+\-]+@[a-z0-9\-]+(\.[a-z0-9\-]+)+$/;

// 項目名を比較用にそろえる（全角半角・空白の違いを無視）
function klistNormalizeLabel(label) {
  return String(label).normalize("NFKC").replace(/\s/g, "");
}

function klistGetSpreadsheet() {
  const url = PropertiesService.getScriptProperties().getProperty("FORM_SPREADSHEET_URL");
  if (url) {
    return SpreadsheetApp.openByUrl(url);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("スクリプトプロパティ FORM_SPREADSHEET_URL を設定してください。");
  }
  return ss;
}

function klistGetKintoneConfig() {
  const props = PropertiesService.getScriptProperties();
  const config = {
    subdomain: props.getProperty("KINTONE_SUBDOMAIN"),
    appId: props.getProperty("KINTONE_KEIYAKU_APP_ID"),
    apiToken: props.getProperty("KINTONE_KEIYAKU_API_TOKEN")
  };
  const missing = Object.keys(config).filter(key => !config[key]);
  if (missing.length > 0) {
    throw new Error("スクリプトプロパティが未設定です: KINTONE_SUBDOMAIN / KINTONE_KEIYAKU_APP_ID / KINTONE_KEIYAKU_API_TOKEN");
  }
  return config;
}

function klistFetchJson(url, apiToken) {
  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { "X-Cybozu-API-Token": apiToken },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error(`kintone取得失敗 (HTTP ${response.getResponseCode()})`);
  }
  return JSON.parse(response.getContentText());
}

// 項目の一覧（フィールドコード・項目名・種類）を取得する
function klistFetchFields(config) {
  const url = `https://${config.subdomain}.cybozu.com/k/v1/app/form/fields.json?app=${config.appId}`;
  const properties = klistFetchJson(url, config.apiToken).properties || {};
  return Object.keys(properties).map(code => ({
    code: code,
    label: properties[code].label,
    type: properties[code].type,
    options: properties[code].options ? Object.keys(properties[code].options) : []
  }));
}

// 項目名（正規化済み）→ フィールドコード の対応表を作る。登録済みの対応（スクリプトプロパティ）が優先
function klistResolveFieldCodes(config) {
  const labelToCode = {};

  try {
    klistFetchFields(config).forEach(field => {
      const key = klistNormalizeLabel(field.label);
      if (!(key in labelToCode)) labelToCode[key] = field.code;
    });
  } catch (e) {
    console.warn(`項目名の取得に失敗しました（スクリプトプロパティ KLIST_FIELD_CODES を使います）: ${e}`);
  }

  const override = PropertiesService.getScriptProperties().getProperty("KLIST_FIELD_CODES");
  if (override) {
    const parsed = JSON.parse(override);
    Object.keys(parsed).forEach(label => {
      labelToCode[klistNormalizeLabel(label)] = parsed[label];
    });
  }

  const requiredLabels = Object.keys(KLIST_LABELS).map(key => KLIST_LABELS[key]);
  requiredLabels.push(KLIST_BURNABLE_LABEL);

  const missing = requiredLabels.filter(label => !(klistNormalizeLabel(label) in labelToCode));
  if (missing.length > 0) {
    throw new Error(`次の項目のフィールドコードが分かりません: ${missing.join(" / ")}\n` +
      "dumpKeiyakuFields で項目名を確認するか、KLIST_FIELD_CODES に登録してください。");
  }

  const codeOf = label => labelToCode[klistNormalizeLabel(label)];
  return {
    contractType: codeOf(KLIST_LABELS.contractType),
    storeName: codeOf(KLIST_LABELS.storeName),
    vendor: codeOf(KLIST_LABELS.vendor),
    replyEmail: codeOf(KLIST_LABELS.replyEmail),
    guideEmail: codeOf(KLIST_LABELS.guideEmail),
    burnable: codeOf(KLIST_BURNABLE_LABEL)
  };
}

// 必要な項目だけを、$idベースのページングで全件取得する
function klistFetchRecords(config, fieldCodes) {
  const wanted = ["$id"].concat(
    [fieldCodes.contractType, fieldCodes.storeName, fieldCodes.vendor, fieldCodes.replyEmail,
      fieldCodes.guideEmail, fieldCodes.burnable]
  );
  const fieldsParam = wanted.map((code, i) => `fields%5B${i}%5D=${encodeURIComponent(code)}`).join("&");

  const limit = 500;
  let lastId = 0;
  let allRecords = [];

  while (true) {
    const query = encodeURIComponent(`$id > ${lastId} order by $id asc limit ${limit}`);
    const url = `https://${config.subdomain}.cybozu.com/k/v1/records.json?app=${config.appId}&query=${query}&${fieldsParam}`;
    const records = klistFetchJson(url, config.apiToken).records || [];
    if (records.length === 0) break;

    allRecords = allRecords.concat(records);
    lastId = Number(records[records.length - 1].$id.value);
    if (records.length < limit) break;
  }
  return allRecords;
}

// レコードの項目値を文字列で取り出す（チェックボックス等の配列は「,」でつなぐ）
function klistValue(record, code) {
  const field = record[code];
  if (!field || field.value === null || field.value === undefined) return "";
  const value = field.value;
  if (Array.isArray(value)) {
    return value.map(v => (typeof v === "object" ? (v.name || v.code || "") : String(v))).join(",");
  }
  return String(value);
}

// 可燃収集日（チェックボックス）で選ばれている選択肢の一覧を返す
function klistSelectedOptions(record, code) {
  const field = record[code];
  if (!field || field.value === null || field.value === undefined || field.value === "") return [];
  return Array.isArray(field.value) ? field.value.map(String) : [String(field.value)];
}

// 選択肢が曜日（「月」「月曜」「月曜日」など）なら、日〜土の順の true/false にする
function klistWeekdayFlags(selectedOptions) {
  const initials = selectedOptions.map(option => klistNormalizeLabel(option).charAt(0));
  return KLIST_WEEKDAYS.split("").map(weekday => initials.indexOf(weekday) !== -1);
}

// 送信先を決める。年末年始の回答アドレスを優先し、無ければ案内送付先(to)を使う
function klistPickEmail(replyRaw, guideRaw) {
  const normalize = s => String(s).normalize("NFKC").replace(/\s/g, "").toLowerCase();
  const reply = normalize(replyRaw);
  const guide = normalize(guideRaw);

  if (reply) {
    return KLIST_EMAIL_PATTERN.test(reply)
      ? { email: reply }
      : { reason: "年末年始の回答アドレスがメールアドレスの形式ではありません", raw: replyRaw };
  }
  if (guide) {
    return KLIST_EMAIL_PATTERN.test(guide)
      ? { email: guide }
      : { reason: "案内送付先アドレスがメールアドレスの形式ではありません", raw: guideRaw };
  }
  return { reason: "送信先アドレスがありません", raw: "" };
}

// 期間中の8日分について、見出しと「回収できる曜日か」を作る
function klistBuildDates() {
  return KLIST_DATES.map(([y, m, d]) => {
    const weekdayIndex = new Date(y, m - 1, d).getDay();
    return {
      header: `${m}月${d}日(${KLIST_WEEKDAYS[weekdayIndex]})`,
      weekdayIndex: weekdayIndex,
      collectable: KLIST_NO_COLLECTION_DATES.indexOf(`${y}-${m}-${d}`) === -1
    };
  });
}

// 既存の「リスト」で「送信済」の店舗は、URL・ID・状況を引き継ぐ（作成済みフォームを作り直さないため）
function klistLoadSentStatus(ss) {
  const sentByStore = {};
  const sheet = ss.getSheetByName(KLIST_SOURCE_SHEET_NAME);
  if (!sheet) return sentByStore;

  const data = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[13] === KLIST_STATUS_SENT) {
      sentByStore[row[0]] = [row[11], row[12], row[13]]; // L, M, N列
    }
  }
  return sentByStore;
}

function klistWriteSheet(ss, sheetName, header, rows) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  sheet.clear();
  const values = [header].concat(rows);
  sheet.getRange(1, 1, values.length, header.length).setValues(values);
}

function buildListFromKintone() {
  const config = klistGetKintoneConfig();
  const fieldCodes = klistResolveFieldCodes(config);
  const records = klistFetchRecords(config, fieldCodes);

  const ss = klistGetSpreadsheet();
  const dates = klistBuildDates();
  const sentByStore = klistLoadSentStatus(ss);

  const listRows = [];
  const reviewRows = [];
  let targetCount = 0;
  let carriedCount = 0;
  const seenBurnableOptions = new Set();
  let recognizedWeekdayCount = 0;

  records.forEach(record => {
    if (klistValue(record, fieldCodes.contractType) !== KLIST_TARGET_CONTRACT_TYPE) return;
    targetCount++;

    const recordId = record.$id.value;
    const storeName = klistValue(record, fieldCodes.storeName);
    const vendor = klistValue(record, fieldCodes.vendor);
    const addReview = (reason, raw) => reviewRows.push([recordId, storeName, reason, raw || ""]);

    const picked = klistPickEmail(
      klistValue(record, fieldCodes.replyEmail),
      klistValue(record, fieldCodes.guideEmail)
    );
    if (picked.reason) {
      addReview(picked.reason, picked.raw);
      return;
    }

    const selectedOptions = klistSelectedOptions(record, fieldCodes.burnable);
    selectedOptions.forEach(option => seenBurnableOptions.add(option));
    const burnableOn = klistWeekdayFlags(selectedOptions); // 日〜土の順
    if (burnableOn.some(Boolean)) recognizedWeekdayCount++;
    if (!burnableOn.some(Boolean)) {
      addReview("可燃の回収曜日が登録されていません");
      return;
    }

    const marks = dates.map(date => (date.collectable && burnableOn[date.weekdayIndex] ? "○" : ""));
    if (!marks.some(mark => mark === "○")) {
      addReview("期間中に回収できる日がありません（回収曜日が1月1日の曜日のみ、など）");
      return;
    }

    const carried = sentByStore[storeName] || ["", "", ""];
    if (carried[2]) carriedCount++;
    listRows.push([storeName, picked.email, vendor].concat(marks, carried));
  });

  // 全件で曜日が読み取れなかった場合は、選択肢の名前が想定と違うので、シートを作らずに止める
  if (targetCount > 0 && recognizedWeekdayCount === 0) {
    throw new Error(`「${KLIST_BURNABLE_LABEL}」の選択肢を曜日として読み取れませんでした。選択肢: ${Array.from(seenBurnableOptions).join(" / ") || "（なし）"}`);
  }

  listRows.sort((a, b) => (a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0)); // 同じアドレスが並ぶように

  const listHeader = ["店舗名", "メールアドレス", "業者名"]
    .concat(dates.map(date => date.header), ["フォームURL", "フォームID", "送信状況"]);
  klistWriteSheet(ss, KLIST_OUTPUT_SHEET_NAME, listHeader, listRows);
  klistWriteSheet(ss, KLIST_REVIEW_SHEET_NAME, ["レコード番号", "契約店舗名称", "理由", "アドレス欄の内容"], reviewRows);

  const emailCount = new Set(listRows.map(row => row[1])).size;
  console.log(
    `契約中 ${targetCount} 件 → リスト ${listRows.length} 件（送信先 ${emailCount} 件、うち作成済みを引き継ぎ ${carriedCount} 件）、` +
    `要確認 ${reviewRows.length} 件。シート「${KLIST_OUTPUT_SHEET_NAME}」「${KLIST_REVIEW_SHEET_NAME}」を確認してください。`
  );
}

// 契約管理アプリの項目名とフィールドコードの対応をログに出す（個人情報は出さない）
function dumpKeiyakuFields() {
  const fields = klistFetchFields(klistGetKintoneConfig());
  fields.forEach(field => {
    const optionText = field.type === "CHECK_BOX" && field.options.length > 0 ? `\t選択肢: ${field.options.join("|")}` : "";
    console.log(`${field.code}\t${field.label}\t${field.type}${optionText}`);
  });
  console.log(`項目数: ${fields.length}`);
}

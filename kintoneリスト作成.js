// ===== kintone契約管理 → フォーム作成用リストの作成 =====
// kintoneの契約管理アプリから「契約中」の契約を取得し、次のシートを作る。フォームは作らない（メールも送らない）。
//   ・リスト_kintone … フォーム作成用（フォーム作成側の「リスト」と同じ形）
//   ・要確認       … フォームを作れない契約と、その理由
//   ・契約状況一覧 … 人が見る管理用。契約中の全件について、フォーム作成の状況を一覧にする
//
// 使い方:
//   1. dumpKeiyakuFields を実行して項目名とフィールドコードの対応を確認（初回のみ）
//   2. buildListFromKintone を実行 → 上の3つのシートができる
//   3. 内容を確認してから、フォーム作成側の「リスト」と差し替える
//   4. 状況を最新にしたいときは updateContractStatusList を実行（「契約状況一覧」だけを更新する）
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
const KLIST_STATUS_SHEET_NAME = "契約状況一覧";
const KLIST_RESPONSE_SHEET_NAME = "回答集計";   // 回答集計.js が作る。回答状況の表示に使う
const KLIST_STATUS_ERROR = "エラー";
const KLIST_TARGET_CONTRACT_TYPE = "契約中";

// 特別回収期間（2026年12月30日(水)〜2027年1月3日(日)）の日付と、回収不可の日
const KLIST_DATES = [
  [2026, 12, 30], [2026, 12, 31], [2027, 1, 1], [2027, 1, 2], [2027, 1, 3]
];
const KLIST_NO_COLLECTION_DATES = ["2027-1-1"];
// フォーム作成用リストの日付の列数（D〜H列の5列）。フォーム作成.js・回答集計.js の列の位置と合わせてある。
// 期間の日数を変えるときは、そちらの列の位置も直すこと
const KLIST_DATE_SLOTS = 5;

const KLIST_WEEKDAYS = "日月火水木金土"; // Date#getDay() の順

const KLIST_LABELS = {
  contractType: "契約種別",
  storeName: "契約店舗名称",
  customer: "顧客名称",
  vendor: "収集業者：名称"
};
// 送信先アドレスの項目。上にあるものほど優先する（フォームの送信先は、最初に見つかった正しいアドレス）
// 項目名が実際と違うと見つからないので、dumpKeiyakuFields で確認する。見つからない項目は飛ばす（最低1つは必要）
const KLIST_EMAIL_LABELS = [
  "年末年始　回答アドレス①",
  "年末年始　回答アドレス②",
  "案内等送付先メールアドレス:(to)",
  "送付先別：e-mail",
  "契約時の顧客担当：メールアドレス"
];
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

  const emailCodes = [];
  KLIST_EMAIL_LABELS.forEach(label => {
    const code = labelToCode[klistNormalizeLabel(label)];
    if (code) emailCodes.push({ label: label, code: code });
    else console.warn(`アドレスの項目「${label}」が見つかりません。この項目は使わずに進めます。`);
  });
  if (emailCodes.length === 0) {
    throw new Error("アドレスの項目が1つも見つかりません。dumpKeiyakuFields で項目名を確認してください。");
  }

  const missing = requiredLabels.filter(label => !(klistNormalizeLabel(label) in labelToCode));
  if (missing.length > 0) {
    throw new Error(`次の項目のフィールドコードが分かりません: ${missing.join(" / ")}\n` +
      "dumpKeiyakuFields で項目名を確認するか、KLIST_FIELD_CODES に登録してください。");
  }

  const codeOf = label => labelToCode[klistNormalizeLabel(label)];
  return {
    contractType: codeOf(KLIST_LABELS.contractType),
    storeName: codeOf(KLIST_LABELS.storeName),
    customer: codeOf(KLIST_LABELS.customer),
    vendor: codeOf(KLIST_LABELS.vendor),
    emails: emailCodes,
    burnable: codeOf(KLIST_BURNABLE_LABEL)
  };
}

// 必要な項目だけを、$idベースのページングで全件取得する
function klistFetchRecords(config, fieldCodes) {
  const wanted = ["$id"].concat(
    [fieldCodes.contractType, fieldCodes.storeName, fieldCodes.customer, fieldCodes.vendor, fieldCodes.burnable]
      .concat(fieldCodes.emails.map(e => e.code))
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

// 優先順位に並べた各項目から、メールアドレスを集める。
// 1つの項目に複数のアドレスが入っていても、「,」「;」「、」・空白・改行で分けて取り出す。
// 送信先（email）は、最初に見つかった正しいアドレス。同じアドレスは1回だけ数える。
//   sources: [{label, raw}] 優先順位順
function klistPickEmail(sources) {
  const valid = [];     // {email, label}
  const invalid = [];   // 形式が正しくなかったもの
  sources.forEach(source => {
    const text = String(source.raw).normalize("NFKC").toLowerCase();
    text.split(/[,;、，；\s]+/).filter(Boolean).forEach(candidate => {
      if (KLIST_EMAIL_PATTERN.test(candidate)) {
        if (!valid.some(v => v.email === candidate)) valid.push({ email: candidate, label: source.label });
      } else {
        invalid.push(`${source.label}: ${candidate}`);
      }
    });
  });

  if (valid.length === 0) {
    return invalid.length > 0
      ? { reason: "メールアドレスの形式が正しいものがありません", raw: invalid.join(" / ") }
      : { reason: "送信先アドレスがありません", raw: "" };
  }
  return {
    email: valid[0].email,
    usedLabel: valid[0].label,
    others: valid.slice(1).map(v => `${v.email}（${v.label}）`).join(" / "),
    invalid: invalid.join(" / ")
  };
}

// 期間中の各日について、見出しと「回収できる曜日か」を作る
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

// 既存の「リスト」の状況を読む。「送信済」の店舗はURL・ID・状況を引き継ぎ（作成済みフォームを作り直さないため）、
// 「エラー」の店舗は一覧に表示するために控えておく
function klistLoadSourceStatus(ss) {
  const sentByStore = {};
  const errorStores = new Set();
  const sheet = ss.getSheetByName(KLIST_SOURCE_SHEET_NAME);
  if (!sheet) return { sentByStore: sentByStore, errorStores: errorStores };

  const data = sheet.getDataRange().getDisplayValues();
  // フォームURL・フォームID・送信状況の列は、見出しから探す（見つからなければ I, J, K列）
  const urlCol = data[0].indexOf("フォームURL") >= 0 ? data[0].indexOf("フォームURL") : 8;
  const idCol = urlCol + 1;
  const statusCol = urlCol + 2;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[statusCol] === KLIST_STATUS_SENT) {
      sentByStore[row[0]] = [row[urlCol], row[idCol], row[statusCol]];
    } else if (row[statusCol] === KLIST_STATUS_ERROR) {
      errorStores.add(row[0]);
    }
  }
  return { sentByStore: sentByStore, errorStores: errorStores };
}

// 「回答集計」シートから、回答のある店舗と回答日時を読む（シートが無ければ空）
function klistLoadResponseTimes(ss) {
  const answeredAtByStore = {};
  const sheet = ss.getSheetByName(KLIST_RESPONSE_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return answeredAtByStore;

  const values = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < values.length; i++) {
    answeredAtByStore[values[i][0]] = values[i][1]; // A列: 店舗名, B列: 回答日時
  }
  return answeredAtByStore;
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

// kintoneから契約中の契約を取得し、フォーム作成用リスト・要確認・契約状況一覧の中身を作る（シートには書かない）
function klistCollect(ss) {
  const config = klistGetKintoneConfig();
  const fieldCodes = klistResolveFieldCodes(config);
  const records = klistFetchRecords(config, fieldCodes);

  const dates = klistBuildDates();
  const source = klistLoadSourceStatus(ss);
  const answeredAtByStore = klistLoadResponseTimes(ss);

  const listRows = [];
  const reviewRows = [];
  const statusRows = [];
  let targetCount = 0;
  let carriedCount = 0;
  const seenBurnableOptions = new Set();
  let recognizedWeekdayCount = 0;

  records.forEach(record => {
    if (klistValue(record, fieldCodes.contractType) !== KLIST_TARGET_CONTRACT_TYPE) return;
    targetCount++;

    const recordId = record.$id.value;
    const storeName = klistValue(record, fieldCodes.storeName);
    const customer = klistValue(record, fieldCodes.customer);
    const vendor = klistValue(record, fieldCodes.vendor);
    const recordUrl = `https://${config.subdomain}.cybozu.com/k/${config.appId}/show#record=${recordId}`;

    // 管理用の一覧に1行追加する。状況: フォーム作成済 / 未作成 / エラー / 要確認
    // フォーム作成済の店舗だけ、回答状況（回答済 / 未回答）と回答日時も入れる
    const addStatus = (status, email, reason, pickedInfo) => {
      const answeredAt = answeredAtByStore[storeName];
      const answerState = status === "フォーム作成済" ? (answeredAt ? "回答済" : "未回答") : "";
      statusRows.push([storeName, customer, recordUrl, status, email || "", reason || "",
        answerState, answeredAt || "",
        pickedInfo ? pickedInfo.usedLabel : "", pickedInfo ? pickedInfo.others : "", pickedInfo ? pickedInfo.invalid : ""]);
    };
    const addReview = (reason, raw, pickedInfo) => {
      reviewRows.push([recordId, storeName, reason, raw || ""]);
      addStatus("要確認", pickedInfo && pickedInfo.email, reason, pickedInfo);
    };

    const emailSources = fieldCodes.emails.map(e => ({ label: e.label, raw: klistValue(record, e.code) }));
    const picked = klistPickEmail(emailSources);
    // 各アドレス項目に入っている内容をそのまま並べる（L〜P列。「メール送信用」シートに使う）。見つからない項目の列は空
    const rawEmails = KLIST_EMAIL_LABELS.map(label => {
      const found = emailSources.filter(e => e.label === label)[0];
      return found ? found.raw : "";
    });
    if (picked.reason) {
      addReview(picked.reason, picked.raw);
      return;
    }

    const selectedOptions = klistSelectedOptions(record, fieldCodes.burnable);
    selectedOptions.forEach(option => seenBurnableOptions.add(option));
    const burnableOn = klistWeekdayFlags(selectedOptions); // 日〜土の順
    if (burnableOn.some(Boolean)) recognizedWeekdayCount++;
    if (!burnableOn.some(Boolean)) {
      addReview("可燃の回収曜日が登録されていません", "", picked);
      return;
    }

    const marks = dates.map(date => (date.collectable && burnableOn[date.weekdayIndex] ? "○" : ""));
    if (!marks.some(mark => mark === "○")) {
      addReview("期間中に回収できる日がありません（回収曜日が期間外の曜日のみ、など）", "", picked);
      return;
    }
    while (marks.length < KLIST_DATE_SLOTS) marks.push(""); // 列 D〜H の5列に合わせる

    const carried = source.sentByStore[storeName] || ["", "", ""];
    if (carried[2]) {
      carriedCount++;
      addStatus("フォーム作成済", picked.email, "", picked);
    } else if (source.errorStores.has(storeName)) {
      addStatus(KLIST_STATUS_ERROR, picked.email, "フォーム作成でエラーになりました", picked);
    } else {
      addStatus("未作成", picked.email, "", picked);
    }
    listRows.push([storeName, picked.email, vendor].concat(marks, carried, rawEmails));
  });

  // 全件で曜日が読み取れなかった場合は、選択肢の名前が想定と違うので、シートを作らずに止める
  if (targetCount > 0 && recognizedWeekdayCount === 0) {
    throw new Error(`「${KLIST_BURNABLE_LABEL}」の選択肢を曜日として読み取れませんでした。選択肢: ${Array.from(seenBurnableOptions).join(" / ") || "（なし）"}`);
  }

  listRows.sort((a, b) => (a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0)); // 同じアドレスが並ぶように

  return {
    dates: dates,
    listRows: listRows,
    reviewRows: reviewRows,
    statusRows: statusRows,
    targetCount: targetCount,
    carriedCount: carriedCount
  };
}

function klistWriteStatusSheet(ss, statusRows) {
  klistWriteSheet(ss, KLIST_STATUS_SHEET_NAME,
    ["店舗名", "契約者", "契約管理URL", "状況", "送信先アドレス", "理由", "回答状況", "回答日時",
      "採用したアドレスの項目", "その他のアドレス（優先順）", "形式が正しくないもの"], statusRows);
}

function buildListFromKintone() {
  const ss = klistGetSpreadsheet();
  const result = klistCollect(ss);

  const dateHeaders = result.dates.map(date => date.header);
  while (dateHeaders.length < KLIST_DATE_SLOTS) dateHeaders.push("");
  const listHeader = ["店舗名", "メールアドレス", "業者名"]
    .concat(dateHeaders, ["フォームURL", "フォームID", "送信状況"], KLIST_EMAIL_LABELS);
  klistWriteSheet(ss, KLIST_OUTPUT_SHEET_NAME, listHeader, result.listRows);
  klistWriteSheet(ss, KLIST_REVIEW_SHEET_NAME, ["レコード番号", "契約店舗名称", "理由", "アドレス欄の内容"], result.reviewRows);
  klistWriteStatusSheet(ss, result.statusRows);

  const emailCount = new Set(result.listRows.map(row => row[1])).size;
  console.log(
    `契約中 ${result.targetCount} 件 → リスト ${result.listRows.length} 件（送信先 ${emailCount} 件、うち作成済みを引き継ぎ ${result.carriedCount} 件）、` +
    `要確認 ${result.reviewRows.length} 件。シート「${KLIST_OUTPUT_SHEET_NAME}」「${KLIST_REVIEW_SHEET_NAME}」「${KLIST_STATUS_SHEET_NAME}」を確認してください。`
  );
}

// 「契約状況一覧」だけを最新にする（フォーム作成用のリストや要確認は変更しない）
function updateContractStatusList() {
  const ss = klistGetSpreadsheet();
  const result = klistCollect(ss);
  klistWriteStatusSheet(ss, result.statusRows);

  const counts = {};
  result.statusRows.forEach(row => { counts[row[3]] = (counts[row[3]] || 0) + 1; });
  console.log(`契約状況一覧を更新しました（契約中 ${result.targetCount} 件）: ${JSON.stringify(counts)}`);
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

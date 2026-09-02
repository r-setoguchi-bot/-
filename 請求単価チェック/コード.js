/**
 * ====================================================================
 * 請求単価チェック（契約管理アプリ 見積り添付ファイルとの突合）
 *
 * kintone「契約管理」アプリの各レコードについて、添付された見積りファイル（PDF/Excel/画像）から
 * AIで金額を読み取り、現在の「請求単価」フィールドと一致しているかを確認する。
 *
 * ※このプロジェクトは瀬戸口秘書ボットとは完全に独立した、単独のGASプロジェクトです。
 * ※kintoneへの書き込みは一切行わず、結果をメールとスプレッドシートへ出力するのみです。
 *
 * 【事前に設定が必要なスクリプトプロパティ】（プロジェクトの設定 → スクリプト プロパティ）
 *   KINTONE_SUBDOMAIN        : kintoneのサブドメイン（例: https://xxxx.cybozu.com なら "xxxx"）
 *   KINTONE_KEIYAKU_APP_ID   : 契約管理アプリのアプリID
 *   KINTONE_KEIYAKU_API_TOKEN: 契約管理アプリのAPIトークン（レコード閲覧・アプリ管理の権限が必要）
 *   GEMINI_API_KEY           : Gemini APIキー
 *
 * 【任意】自動判定がうまくいかない場合のみ、以下を追加で手動指定できます
 *   KINTONE_KEIYAKU_MITSUMORI_FIELD : 見積り添付ファイルのフィールドコード
 *   KINTONE_KEIYAKU_TANKA_FIELD     : 請求単価のフィールドコード
 * ====================================================================
 */

const BILLING_RATE_CHECK_CONFIG = {
  appIdProp: "KINTONE_KEIYAKU_APP_ID",
  apiTokenProp: "KINTONE_KEIYAKU_API_TOKEN",
  mitsumoriFieldProp: "KINTONE_KEIYAKU_MITSUMORI_FIELD", // 見積り添付ファイルのフィールドコード（任意・手動指定用）
  tankaFieldProp: "KINTONE_KEIYAKU_TANKA_FIELD",         // 請求単価のフィールドコード（任意・手動指定用）
  resultSheetName: "請求単価チェック結果",
  resultSpreadsheetUrlProp: "RESULT_SPREADSHEET_URL" // 結果シートのURL（初回実行時に自動作成してここへ保存する）
};

/**
 * 契約管理アプリの全レコードを対象に、見積り添付ファイルの金額と請求単価フィールドを突合する関数
 * 時間主導型トリガー、またはGASエディタからの手動実行を想定
 */
function checkBillingRates() {
  const props = PropertiesService.getScriptProperties();
  const subdomain = props.getProperty("KINTONE_SUBDOMAIN");
  const appId = props.getProperty(BILLING_RATE_CHECK_CONFIG.appIdProp);
  const apiToken = props.getProperty(BILLING_RATE_CHECK_CONFIG.apiTokenProp);
  const geminiApiKey = props.getProperty("GEMINI_API_KEY");

  if (!subdomain || !appId || !apiToken || !geminiApiKey) {
    const message = "請求単価チェックに必要なスクリプトプロパティが不足しています（KINTONE_SUBDOMAIN / " +
      BILLING_RATE_CHECK_CONFIG.appIdProp + " / " + BILLING_RATE_CHECK_CONFIG.apiTokenProp + " / GEMINI_API_KEY）。";
    console.error(message);
    notifyByEmail("⚠️ 請求単価チェック：設定エラー", message);
    return;
  }

  let fieldCodes;
  try {
    fieldCodes = resolveKeiyakuFieldCodes(subdomain, appId, apiToken);
  } catch (e) {
    console.error("請求単価チェックを中止しました: " + e.message);
    notifyByEmail("⚠️ 請求単価チェック：実行できませんでした", e.message);
    return;
  }
  const mitsumoriField = fieldCodes.mitsumoriField;
  const tankaField = fieldCodes.tankaField;

  const records = fetchKintoneRecords(subdomain, appId, apiToken);
  const results = [];

  records.forEach(record => {
    const recordId = record.$id.value;
    const displayName = (record["屋号名"] && record["屋号名"].value) ||
                         (record["契約店舗名称"] && record["契約店舗名称"].value) ||
                         (record["会社名"] && record["会社名"].value) ||
                         `レコード#${recordId}`;

    const currentTankaRaw = record[tankaField] ? record[tankaField].value : "";
    const files = (record[mitsumoriField] && record[mitsumoriField].value) || [];

    if (files.length === 0) {
      results.push({ recordId, displayName, status: "見積り未添付", current: currentTankaRaw, extracted: "", note: "" });
      return;
    }

    try {
      const file = files[0]; // 複数添付されている場合は先頭（最新想定）のみをチェック対象にする
      const blob = fetchKintoneFile(subdomain, file.fileKey, apiToken);
      const extraction = extractEstimateAmount(blob, file.contentType, geminiApiKey);

      if (extraction.amount === null) {
        results.push({ recordId, displayName, status: "抽出失敗", current: currentTankaRaw, extracted: "", note: extraction.note || "" });
        return;
      }

      const currentTankaNum = parseAmount(currentTankaRaw);
      if (currentTankaNum === null) {
        results.push({ recordId, displayName, status: "請求単価未入力", current: currentTankaRaw, extracted: extraction.amount, note: extraction.note || "" });
      } else if (currentTankaNum === extraction.amount) {
        results.push({ recordId, displayName, status: "一致", current: currentTankaRaw, extracted: extraction.amount, note: "" });
      } else {
        results.push({ recordId, displayName, status: "不一致", current: currentTankaRaw, extracted: extraction.amount, note: extraction.note || "" });
      }
    } catch (e) {
      console.error(`請求単価チェック中にエラー（レコード#${recordId}）: ` + e.message);
      results.push({ recordId, displayName, status: "エラー", current: currentTankaRaw, extracted: "", note: e.message });
    }
  });

  saveBillingRateCheckResults(results);
  sendBillingRateCheckReport(results);
}

/**
 * 契約管理アプリのフォーム設定（フィールド一覧）をkintone APIから取得し、
 * 「見積り添付ファイル」「請求単価」に該当するフィールドコードをラベル名から自動で特定する。
 * スクリプトプロパティで手動指定されている場合はそちらを優先する。
 * ラベルから一意に特定できない場合は、フィールド一覧を添えてエラーを投げる。
 */
function resolveKeiyakuFieldCodes(subdomain, appId, apiToken) {
  const props = PropertiesService.getScriptProperties();
  const overrideMitsumori = props.getProperty(BILLING_RATE_CHECK_CONFIG.mitsumoriFieldProp);
  const overrideTanka = props.getProperty(BILLING_RATE_CHECK_CONFIG.tankaFieldProp);
  if (overrideMitsumori && overrideTanka) {
    return { mitsumoriField: overrideMitsumori, tankaField: overrideTanka };
  }

  const url = `https://${subdomain}.cybozu.com/k/v1/app/form/fields.json?app=${appId}`;
  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { "X-Cybozu-API-Token": apiToken },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error(`kintoneフィールド情報の取得失敗 (HTTP ${response.getResponseCode()}): ${response.getContentText()}`);
  }

  const properties = JSON.parse(response.getContentText()).properties || {};
  const fields = Object.keys(properties).map(code => ({
    code,
    label: properties[code].label || "",
    type: properties[code].type
  }));

  const mitsumoriField = overrideMitsumori ||
    findUniqueFieldCode(fields, f => f.type === "FILE" && f.label.indexOf("見積") !== -1);
  const tankaField = overrideTanka ||
    findUniqueFieldCode(fields, f => f.label.indexOf("請求単価") !== -1 || (f.label.indexOf("請求") !== -1 && f.label.indexOf("単価") !== -1));

  if (!mitsumoriField || !tankaField) {
    const fieldList = fields.map(f => `${f.code}: ${f.label}（${f.type}）`).join("\n");
    throw new Error(
      "見積り添付ファイル/請求単価のフィールドを自動で特定できませんでした。" +
      "スクリプトプロパティ「" + BILLING_RATE_CHECK_CONFIG.mitsumoriFieldProp + "」「" +
      BILLING_RATE_CHECK_CONFIG.tankaFieldProp + "」で手動指定してください。\n" +
      "契約管理アプリのフィールド一覧:\n" + fieldList
    );
  }

  return { mitsumoriField, tankaField };
}

/**
 * 条件に一致するフィールドが1件だけの場合にそのフィールドコードを返す（0件・複数件はnull＝特定不可）
 */
function findUniqueFieldCode(fields, predicate) {
  const matches = fields.filter(predicate);
  return matches.length === 1 ? matches[0].code : null;
}

/**
 * kintoneアプリの全レコードをREST APIで取得する（$idベースでページング）
 */
function fetchKintoneRecords(subdomain, appId, apiToken) {
  let allRecords = [];
  let lastId = 0;
  const limit = 500;

  while (true) {
    const query = encodeURIComponent(`$id > ${lastId} order by $id asc limit ${limit}`);
    const url = `https://${subdomain}.cybozu.com/k/v1/records.json?app=${appId}&query=${query}`;
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "X-Cybozu-API-Token": apiToken },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error(`kintone取得失敗 (app=${appId}, HTTP ${response.getResponseCode()}): ${response.getContentText()}`);
    }

    const records = JSON.parse(response.getContentText()).records || [];
    if (records.length === 0) break;

    allRecords = allRecords.concat(records);
    lastId = Number(records[records.length - 1].$id.value);
    if (records.length < limit) break;
  }

  return allRecords;
}

/**
 * kintoneの添付ファイルを1件ダウンロードする（file.jsonエンドポイント）
 */
function fetchKintoneFile(subdomain, fileKey, apiToken) {
  const url = `https://${subdomain}.cybozu.com/k/v1/file.json?fileKey=${fileKey}`;
  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { "X-Cybozu-API-Token": apiToken },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error(`kintoneファイル取得失敗 (HTTP ${response.getResponseCode()}): ${response.getContentText()}`);
  }
  return response.getBlob();
}

/**
 * 見積りファイル（PDF/画像/Excel）から請求すべき金額をAIで抽出する
 * PDF・画像はGeminiのマルチモーダル入力でそのまま読み取り、
 * Excel（.xlsx）はセルの文字列を抽出してテキストとしてAIに渡す
 */
function extractEstimateAmount(blob, contentType, apiKey) {
  const promptBase = "これは取引先への見積書です。この見積書に記載されている、実際に請求すべき単価・金額" +
    "（値引き後の最終的な金額。消費税込みの合計金額を優先）を1つだけ特定してください。" +
    "複数の商品/サービスが並んでいる場合は、個別の内訳ではなく合計金額（税込）を優先してください。" +
    "金額がどうしても読み取れない場合はamountをnullにしてください。" +
    "説明文などは一切含めず、次のJSON形式の文字列のみを出力してください:\n" +
    '{"amount": 数値（円、カンマなし。税込金額）またはnull, "note": "抽出時に気になった点があれば一言（無ければ空文字）"}';

  let parts;
  if (contentType === "application/pdf" || contentType.indexOf("image/") === 0) {
    parts = [
      { text: promptBase },
      { inlineData: { mimeType: contentType, data: Utilities.base64Encode(blob.getBytes()) } }
    ];
  } else if (contentType.indexOf("spreadsheetml") !== -1) {
    const sheetText = extractTextFromXlsx(blob);
    if (!sheetText) return { amount: null, note: "Excelファイルの内容を読み取れませんでした。" };
    parts = [{ text: promptBase + "\n\n【見積書の内容（セルの値を抽出したもの）】\n" + sheetText }];
  } else {
    return { amount: null, note: `未対応のファイル形式です（${contentType}）。` };
  }

  return callGeminiForAmount(parts, apiKey);
}

/**
 * Gemini APIへ画像/PDF/テキストを渡し、金額抽出結果のJSONを解析して返す
 */
function callGeminiForAmount(parts, apiKey) {
  const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash-lite:generateContent?key=" + apiKey;
  const payload = {
    "contents": [{ "parts": parts }],
    "generationConfig": { "responseMimeType": "application/json" }
  };
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    const text = json.candidates && json.candidates[0].content.parts[0].text;
    if (!text) return { amount: null, note: "AIから金額を読み取れませんでした。" };

    const cleanJsonStr = text.replace(/```json/g, "").replace(/```/g, "").trim();
    const parsed = JSON.parse(cleanJsonStr);
    const amount = (parsed.amount === null || parsed.amount === undefined) ? null : Number(parsed.amount);
    return { amount: (amount === null || isNaN(amount)) ? null : amount, note: parsed.note || "" };
  } catch (e) {
    console.error("見積り金額抽出のGemini呼び出しでエラー: " + e.message);
    return { amount: null, note: "AI呼び出し中にエラーが発生しました。" };
  }
}

/**
 * カンマ・円マークなどを含む文字列を数値に変換する（変換できない場合はnull）
 */
function parseAmount(raw) {
  if (raw === "" || raw === null || raw === undefined) return null;
  const cleaned = String(raw).replace(/[^\d.]/g, "");
  if (cleaned === "") return null;
  const num = Number(cleaned);
  return isNaN(num) ? null : num;
}

/**
 * xlsx（zip形式）から共有文字列・シートXMLを取り出し、セルの値をタブ区切りテキストへ変換する
 * Drive APIの追加権限を使わずに済むよう、Utilities.unzipで直接パースする
 */
function extractTextFromXlsx(blob) {
  try {
    const entries = Utilities.unzip(blob);
    let sharedStrings = [];
    const sheetTexts = [];

    entries.forEach(entry => {
      if (entry.getName() === "xl/sharedStrings.xml") {
        sharedStrings = parseSharedStringsXml(entry.getDataAsString());
      }
    });

    entries.forEach(entry => {
      if (/^xl\/worksheets\/sheet\d+\.xml$/.test(entry.getName())) {
        sheetTexts.push(parseSheetXml(entry.getDataAsString(), sharedStrings));
      }
    });

    return sheetTexts.join("\n").trim();
  } catch (e) {
    console.error("xlsx解析エラー: " + e.message);
    return "";
  }
}

function parseSharedStringsXml(xml) {
  const result = [];
  const siRegex = /<si>([\s\S]*?)<\/si>/g;
  let siMatch;
  while ((siMatch = siRegex.exec(xml)) !== null) {
    const tRegex = /<t[^>]*>([\s\S]*?)<\/t>/g;
    let text = "";
    let tMatch;
    while ((tMatch = tRegex.exec(siMatch[1])) !== null) {
      text += tMatch[1];
    }
    result.push(decodeXmlEntities(text));
  }
  return result;
}

function parseSheetXml(xml, sharedStrings) {
  const lines = [];
  const rowRegex = /<row\b[^>]*>([\s\S]*?)<\/row>/g;
  let rowMatch;
  while ((rowMatch = rowRegex.exec(xml)) !== null) {
    const cellRegex = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
    const cells = [];
    let cellMatch;
    while ((cellMatch = cellRegex.exec(rowMatch[1])) !== null) {
      const attrs = cellMatch[1];
      const cellBody = cellMatch[2];
      const typeMatch = /\st="([^"]*)"/.exec(attrs);
      const cellType = typeMatch ? typeMatch[1] : null;

      let value = "";
      if (cellType === "s") {
        const vMatch = /<v>([\s\S]*?)<\/v>/.exec(cellBody);
        if (vMatch) value = sharedStrings[Number(vMatch[1])] || "";
      } else if (cellType === "inlineStr") {
        const tMatch = /<t[^>]*>([\s\S]*?)<\/t>/.exec(cellBody);
        if (tMatch) value = decodeXmlEntities(tMatch[1]);
      } else {
        const vMatch = /<v>([\s\S]*?)<\/v>/.exec(cellBody);
        if (vMatch) value = vMatch[1];
      }
      if (value !== "") cells.push(value);
    }
    if (cells.length > 0) lines.push(cells.join("\t"));
  }
  return lines.join("\n");
}

function decodeXmlEntities(str) {
  return str
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

/**
 * 結果保存用のスプレッドシートを取得する。存在しなければ新規作成し、URLをスクリプトプロパティへ保存する
 * （このプロジェクト専用のシートであり、瀬戸口秘書ボットのスプレッドシートとは無関係）
 */
function getOrCreateResultSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const sheetUrl = props.getProperty(BILLING_RATE_CHECK_CONFIG.resultSpreadsheetUrlProp);
  if (sheetUrl) return SpreadsheetApp.openByUrl(sheetUrl);

  const spreadsheet = SpreadsheetApp.create("請求単価チェック結果");
  props.setProperty(BILLING_RATE_CHECK_CONFIG.resultSpreadsheetUrlProp, spreadsheet.getUrl());
  return spreadsheet;
}

/**
 * 指定したシート名のシートを取得し、無ければ新規作成する
 */
function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  return sheet;
}

/**
 * チェック結果をスプレッドシートへ保存する（毎回上書き。詳細確認用）
 */
function saveBillingRateCheckResults(results) {
  const spreadsheet = getOrCreateResultSpreadsheet();
  const sheet = getOrCreateSheet(spreadsheet, BILLING_RATE_CHECK_CONFIG.resultSheetName);
  sheet.clear();

  const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  const rows = [["チェック日時", "レコードID", "契約先", "ステータス", "現在の請求単価", "見積りから読み取った金額", "備考"]];
  results.forEach(r => {
    rows.push([now, r.recordId, r.displayName, r.status, r.current, r.extracted, r.note]);
  });
  sheet.getRange(1, 1, rows.length, 7).setValues(rows);
}

/**
 * チェック結果のうち「一致」以外（要確認）の項目のみをメールで要約送信する
 */
function sendBillingRateCheckReport(results) {
  const needsAttention = results.filter(r => r.status !== "一致");
  const okCount = results.length - needsAttention.length;
  const MAX_LINES = 50;

  let body = `対象 ${results.length}件中、一致 ${okCount}件・要確認 ${needsAttention.length}件でした。\n\n`;

  if (needsAttention.length === 0) {
    body += "すべて一致していました。特に対応は不要です。";
  } else {
    body += needsAttention.slice(0, MAX_LINES).map(r => {
      let line = `【${r.status}】${r.displayName}（現在: ${r.current || "(空)"}`;
      if (r.extracted !== "") line += ` / 見積り: ${r.extracted}`;
      line += "）";
      if (r.note) line += `\n  備考: ${r.note}`;
      return line;
    }).join("\n");

    if (needsAttention.length > MAX_LINES) {
      body += `\n…ほか${needsAttention.length - MAX_LINES}件`;
    }
    body += "\n\n詳細は結果スプレッドシートをご確認ください: " + getOrCreateResultSpreadsheet().getUrl();
  }

  notifyByEmail(`🤖 請求単価チェック結果（要確認 ${needsAttention.length}件）`, body);
}

/**
 * スクリプトの実行者（オーナー）宛にメールを送る。NOTIFY_EMAILが設定されていればそちらを優先する
 */
function notifyByEmail(subject, body) {
  const props = PropertiesService.getScriptProperties();
  const to = props.getProperty("NOTIFY_EMAIL") || Session.getActiveUser().getEmail();
  if (!to) {
    console.error("通知先メールアドレスを特定できませんでした。スクリプトプロパティ「NOTIFY_EMAIL」を設定してください。");
    return;
  }
  MailApp.sendEmail(to, subject, body);
}

/**
 * checkBillingRatesを1日1回（午前8時）自動実行するトリガーを設定する関数
 * 自動チェックを開始したいタイミングでGASエディタからこの関数を1回だけ手動実行してください
 */
function setupBillingRateCheckTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "checkBillingRates") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("checkBillingRates")
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .create();
}

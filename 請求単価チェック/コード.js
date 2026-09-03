/**
 * ====================================================================
 * 請求単価チェック（契約管理アプリ「単価テーブル」と見積書添付ファイルの突合）
 *
 * kintone「契約管理」アプリの各レコードについて、
 * 「単価テーブル」サブテーブルの商品ごとの請求単価と、
 * 「契約書PDF」欄に添付されている見積書（ファイル名に「見積」を含むもの）から
 * AIで読み取った商品ごとの金額を、商品名で突き合わせて確認する。
 *
 * ※このプロジェクトは瀬戸口秘書ボットとは完全に独立した、単独のGASプロジェクトです。
 * ※kintoneへの書き込みは一切行わず、結果をメールとスプレッドシートへ出力するのみです。
 *
 * ※対象レコードが多い（1000件以上など）場合、1回の実行では終わらないため、
 *   数分ごとに自動で続きを実行する「分割処理」方式になっています。
 *   checkBillingRates を実行すると、処理しきれなかった分は1分ごとのトリガーで
 *   自動的に続きが実行され、全件終わったタイミングで完了メールが届きます。
 *
 * 【事前に設定が必要なスクリプトプロパティ】（プロジェクトの設定 → スクリプト プロパティ）
 *   KINTONE_SUBDOMAIN        : kintoneのサブドメイン（例: https://xxxx.cybozu.com なら "xxxx"）
 *   KINTONE_KEIYAKU_APP_ID   : 契約管理アプリのアプリID
 *   KINTONE_KEIYAKU_API_TOKEN: 契約管理アプリのAPIトークン（レコード閲覧・アプリ管理の権限が必要）
 *   GEMINI_API_KEY           : Gemini APIキー
 *
 * 【任意】添付ファイル欄のフィールドコードが異なる場合のみ
 *   KINTONE_KEIYAKU_MITSUMORI_FIELD : 見積書が入っている添付ファイル欄のフィールドコード（未設定時は既定値"契約書PDF"を使う）
 * ====================================================================
 */

const BILLING_RATE_CHECK_CONFIG = {
  appIdProp: "KINTONE_KEIYAKU_APP_ID",
  apiTokenProp: "KINTONE_KEIYAKU_API_TOKEN",
  fileFieldProp: "KINTONE_KEIYAKU_MITSUMORI_FIELD", // 見積書が入っている添付ファイル欄のフィールドコード（任意・手動指定用）
  defaultFileFieldCode: "契約書PDF",                 // 上記が未設定の場合に使う既定のフィールドコード
  subtableFieldCode: "単価テーブル",                  // 請求単価が入っているサブテーブルのフィールドコード
  itemNameFieldCode: "商品名",                        // サブテーブル内：商品名の列
  tankaFieldCode: "請求単価",                         // サブテーブル内：請求単価の列
  estimateFileNameKeyword: "見積",                    // 添付ファイルのうち、これを名前に含むものを見積書とみなす
  resultSheetName: "請求単価チェック結果",
  resultSpreadsheetUrlProp: "RESULT_SPREADSHEET_URL", // 結果シートのURL（初回実行時に自動作成してここへ保存する）
  progressLastIdProp: "BILLING_CHECK_PROGRESS_LAST_ID",
  progressOkCountProp: "BILLING_CHECK_PROGRESS_OK_COUNT",
  progressAttentionCountProp: "BILLING_CHECK_PROGRESS_ATTENTION_COUNT",
  continuationHandlerName: "continueBillingRateCheck"
};

const EXECUTION_TIME_BUDGET_MS = 4.5 * 60 * 1000; // 1回の実行で使ってよい時間（安全のため4分30秒までにしておく）
const KINTONE_PAGE_SIZE = 100; // 1回のkintone取得件数

/**
 * 請求単価チェックを開始する関数。GASエディタからの手動実行を想定
 * 対象レコードが多い場合は1回で終わらないため、進捗をリセットしたうえで1回目のバッチを実行し、
 * 終わらなければ1分ごとに自動で続きが実行されるようにする
 */
function checkBillingRates() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(BILLING_RATE_CHECK_CONFIG.progressLastIdProp);
  props.setProperty(BILLING_RATE_CHECK_CONFIG.progressOkCountProp, "0");
  props.setProperty(BILLING_RATE_CHECK_CONFIG.progressAttentionCountProp, "0");

  const spreadsheet = getOrCreateResultSpreadsheet();
  const sheet = getOrCreateSheet(spreadsheet, BILLING_RATE_CHECK_CONFIG.resultSheetName);
  sheet.clear();
  sheet.appendRow(["チェック日時", "レコードID", "契約先", "商品名", "ステータス", "現在の請求単価", "見積りから読み取った金額"]);

  removeContinuationTrigger();
  runBillingRateCheckBatch();
}

/**
 * 1分ごとのトリガーから呼ばれ、続きのバッチを実行する関数（手動実行はしない）
 */
function continueBillingRateCheck() {
  runBillingRateCheckBatch();
}

/**
 * 実際の1バッチ分の処理。時間切れになったら進捗を保存して抜け、まだ終わっていなければ
 * 続行用トリガーを仕込む。全件終わったらトリガーを消して完了メールを送る
 */
function runBillingRateCheckBatch() {
  const startTime = Date.now();
  const props = PropertiesService.getScriptProperties();
  const subdomain = props.getProperty("KINTONE_SUBDOMAIN");
  const appId = props.getProperty(BILLING_RATE_CHECK_CONFIG.appIdProp);
  const apiToken = props.getProperty(BILLING_RATE_CHECK_CONFIG.apiTokenProp);
  const geminiApiKey = props.getProperty("GEMINI_API_KEY");
  const fileFieldCode = props.getProperty(BILLING_RATE_CHECK_CONFIG.fileFieldProp) || BILLING_RATE_CHECK_CONFIG.defaultFileFieldCode;

  if (!subdomain || !appId || !apiToken || !geminiApiKey) {
    const message = "請求単価チェックに必要なスクリプトプロパティが不足しています（KINTONE_SUBDOMAIN / " +
      BILLING_RATE_CHECK_CONFIG.appIdProp + " / " + BILLING_RATE_CHECK_CONFIG.apiTokenProp + " / GEMINI_API_KEY）。";
    console.error(message);
    notifyByEmail("⚠️ 請求単価チェック：設定エラー", message);
    removeContinuationTrigger();
    return;
  }

  const sheet = getOrCreateSheet(getOrCreateResultSpreadsheet(), BILLING_RATE_CHECK_CONFIG.resultSheetName);
  let lastId = Number(props.getProperty(BILLING_RATE_CHECK_CONFIG.progressLastIdProp) || "0");
  let okCount = Number(props.getProperty(BILLING_RATE_CHECK_CONFIG.progressOkCountProp) || "0");
  let attentionCount = Number(props.getProperty(BILLING_RATE_CHECK_CONFIG.progressAttentionCountProp) || "0");
  let finished = false;

  outer:
  while (true) {
    const query = encodeURIComponent(`$id > ${lastId} order by $id asc limit ${KINTONE_PAGE_SIZE}`);
    const url = `https://${subdomain}.cybozu.com/k/v1/records.json?app=${appId}&query=${query}`;
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "X-Cybozu-API-Token": apiToken },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      const message = `kintone取得失敗 (HTTP ${response.getResponseCode()}): ${response.getContentText()}`;
      console.error(message);
      notifyByEmail("⚠️ 請求単価チェック：実行できませんでした", message);
      removeContinuationTrigger();
      return;
    }

    const records = JSON.parse(response.getContentText()).records || [];
    if (records.length === 0) {
      finished = true;
      break;
    }

    for (let i = 0; i < records.length; i++) {
      if (Date.now() - startTime > EXECUTION_TIME_BUDGET_MS) {
        break outer; // 時間切れ。ここまでの進捗は保存済みなので、続きは次のトリガーで行う
      }

      const record = records[i];
      const rows = buildResultRowsForRecord(record, fileFieldCode, subdomain, apiToken, geminiApiKey);

      if (rows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
        rows.forEach(r => { if (r[4] === "一致") okCount++; else attentionCount++; });
      }

      lastId = Number(record.$id.value);
      props.setProperty(BILLING_RATE_CHECK_CONFIG.progressLastIdProp, String(lastId));
      props.setProperty(BILLING_RATE_CHECK_CONFIG.progressOkCountProp, String(okCount));
      props.setProperty(BILLING_RATE_CHECK_CONFIG.progressAttentionCountProp, String(attentionCount));
    }

    if (records.length < KINTONE_PAGE_SIZE) {
      finished = true;
      break;
    }
  }

  if (finished) {
    removeContinuationTrigger();
    sendBillingRateCheckFinalReport(okCount, attentionCount);
    props.deleteProperty(BILLING_RATE_CHECK_CONFIG.progressLastIdProp);
    props.deleteProperty(BILLING_RATE_CHECK_CONFIG.progressOkCountProp);
    props.deleteProperty(BILLING_RATE_CHECK_CONFIG.progressAttentionCountProp);
  } else {
    ensureContinuationTrigger();
  }
}

/**
 * 続行用トリガー（1分ごと）が無ければ作成する
 */
function ensureContinuationTrigger() {
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === BILLING_RATE_CHECK_CONFIG.continuationHandlerName);
  if (!exists) {
    ScriptApp.newTrigger(BILLING_RATE_CHECK_CONFIG.continuationHandlerName)
      .timeBased()
      .everyMinutes(1)
      .create();
  }
}

/**
 * 続行用トリガーを削除する
 */
function removeContinuationTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === BILLING_RATE_CHECK_CONFIG.continuationHandlerName) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/**
 * レコード1件分を処理し、スプレッドシートに書き込む行（商品ごと）の配列を返す
 * （単価テーブルが無い/見積書が無い/読み取り失敗の場合も、状況が分かる1行を返す）
 */
function buildResultRowsForRecord(record, fileFieldCode, subdomain, apiToken, geminiApiKey) {
  const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
  const recordId = record.$id.value;
  const displayName = (record["屋号名"] && record["屋号名"].value) ||
                       (record["契約店舗名称"] && record["契約店舗名称"].value) ||
                       (record["会社名"] && record["会社名"].value) ||
                       `レコード#${recordId}`;

  const tableRows = (record[BILLING_RATE_CHECK_CONFIG.subtableFieldCode] &&
                      record[BILLING_RATE_CHECK_CONFIG.subtableFieldCode].value) || [];

  if (tableRows.length === 0) {
    return [[now, recordId, displayName, "(全項目)", "単価テーブルなし", "", ""]];
  }

  const files = (record[fileFieldCode] && record[fileFieldCode].value) || [];
  const estimateFiles = files.filter(f => f.name.indexOf(BILLING_RATE_CHECK_CONFIG.estimateFileNameKeyword) !== -1);

  if (estimateFiles.length === 0) {
    return [[now, recordId, displayName, "(全項目)", "見積り未添付", "", ""]];
  }

  let extractedItems;
  try {
    // 複数見積りが添付されている場合は先頭（最新想定）のみをチェック対象にする
    const blob = fetchKintoneFile(subdomain, estimateFiles[0].fileKey, apiToken);
    const extraction = extractEstimateItems(blob, estimateFiles[0].contentType, geminiApiKey);

    if (!extraction.items || extraction.items.length === 0) {
      return [[now, recordId, displayName, "(全項目)", "抽出失敗", "", ""]];
    }
    extractedItems = extraction.items;
  } catch (e) {
    console.error(`見積書の読み取り中にエラー（レコード#${recordId}）: ` + e.message);
    return [[now, recordId, displayName, "(全項目)", "エラー", "", ""]];
  }

  const rows = [];
  tableRows.forEach(row => {
    const itemName = row.value[BILLING_RATE_CHECK_CONFIG.itemNameFieldCode]
      ? row.value[BILLING_RATE_CHECK_CONFIG.itemNameFieldCode].value : "";
    const currentTankaRaw = row.value[BILLING_RATE_CHECK_CONFIG.tankaFieldCode]
      ? row.value[BILLING_RATE_CHECK_CONFIG.tankaFieldCode].value : "";

    if (!itemName) return; // 商品名が空の行はスキップ

    const matched = findMatchingEstimateItem(extractedItems, itemName);
    if (!matched) {
      rows.push([now, recordId, displayName, itemName, "見積りに対応項目なし", currentTankaRaw, ""]);
      return;
    }

    const currentTankaNum = parseAmount(currentTankaRaw);
    if (currentTankaNum === null) {
      rows.push([now, recordId, displayName, itemName, "請求単価未入力", currentTankaRaw, matched.unitPrice]);
    } else if (currentTankaNum === matched.unitPrice) {
      rows.push([now, recordId, displayName, itemName, "一致", currentTankaRaw, matched.unitPrice]);
    } else {
      rows.push([now, recordId, displayName, itemName, "不一致", currentTankaRaw, matched.unitPrice]);
    }
  });

  return rows;
}

/**
 * 見積りから抽出した品目リストの中から、単価テーブルの商品名に対応するものを探す
 * 空白除去のうえ、完全一致または部分一致（どちらかがどちらかを含む）で判定する
 */
function findMatchingEstimateItem(items, itemName) {
  const target = normalizeItemName(itemName);
  if (!target) return null;

  return items.find(it => {
    const n = normalizeItemName(it.itemName);
    if (!n) return false;
    return n === target || n.indexOf(target) !== -1 || target.indexOf(n) !== -1;
  }) || null;
}

function normalizeItemName(name) {
  return String(name || "").replace(/[\s　]/g, "").trim();
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
 * 見積りファイル（PDF/画像/Excel）から、品目名と単価のペアをすべてAIで抽出する
 * PDF・画像はGeminiのマルチモーダル入力でそのまま読み取り、
 * Excel（.xlsx）はセルの文字列を抽出してテキストとしてAIに渡す
 */
function extractEstimateItems(blob, contentType, apiKey) {
  const promptBase = "これは取引先への見積書です。見積書に記載されている品目（商品名・サービス名）と、" +
    "それぞれの単価（金額）のペアを、書かれている行すべてについて抽出してください。" +
    "小計・消費税・合計などの集計行は含めず、個別の品目行だけを対象にしてください。" +
    "説明文などは一切含めず、次のJSON形式の文字列のみを出力してください:\n" +
    '{"items": [{"itemName": "品目名", "unitPrice": 数値（円、カンマなし）}], "note": "抽出時に気になった点があれば一言（無ければ空文字）"}';

  let parts;
  if (contentType === "application/pdf" || contentType.indexOf("image/") === 0) {
    parts = [
      { text: promptBase },
      { inlineData: { mimeType: contentType, data: Utilities.base64Encode(blob.getBytes()) } }
    ];
  } else if (contentType.indexOf("spreadsheetml") !== -1) {
    const sheetText = extractTextFromXlsx(blob);
    if (!sheetText) return { items: [], note: "Excelファイルの内容を読み取れませんでした。" };
    parts = [{ text: promptBase + "\n\n【見積書の内容（セルの値を抽出したもの）】\n" + sheetText }];
  } else {
    return { items: [], note: `未対応のファイル形式です（${contentType}）。` };
  }

  return callGeminiForItems(parts, apiKey);
}

/**
 * Gemini APIへ画像/PDF/テキストを渡し、品目ごとの金額抽出結果のJSONを解析して返す
 */
function callGeminiForItems(parts, apiKey) {
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
    if (!text) return { items: [], note: "AIから内容を読み取れませんでした。" };

    const cleanJsonStr = text.replace(/```json/g, "").replace(/```/g, "").trim();
    const parsed = JSON.parse(cleanJsonStr);
    const items = Array.isArray(parsed.items) ? parsed.items
      .map(it => ({ itemName: String(it.itemName || ""), unitPrice: Number(it.unitPrice) }))
      .filter(it => it.itemName && !isNaN(it.unitPrice)) : [];

    return { items, note: parsed.note || "" };
  } catch (e) {
    console.error("見積り品目抽出のGemini呼び出しでエラー: " + e.message);
    return { items: [], note: "AI呼び出し中にエラーが発生しました。" };
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
 * 全件処理が完了した際に、件数のサマリーをメールで送る（詳細はスプレッドシート参照）
 */
function sendBillingRateCheckFinalReport(okCount, attentionCount) {
  const total = okCount + attentionCount;
  const body = `対象 ${total}件中、一致 ${okCount}件・要確認 ${attentionCount}件でした。\n\n` +
    "商品ごとの詳細（契約先・商品名・現在の請求単価・見積りから読み取った金額）はスプレッドシートをご確認ください:\n" +
    getOrCreateResultSpreadsheet().getUrl();

  notifyByEmail(`🤖 請求単価チェック完了（要確認 ${attentionCount}件）`, body);
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

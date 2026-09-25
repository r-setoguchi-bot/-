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
  costFieldCode: "仕入単価",                          // サブテーブル内：仕入単価の列（赤字チェック用）
  contractorFieldCode: "収集業者名称",                // レコード内：収集業者名のフィールドコード
  estimateFileNameKeyword: "見積",                    // 添付ファイルのうち、これを名前に含むものを見積書とみなす
  resultSheetName: "請求単価チェック結果",
  summarySheetName: "要対応店舗一覧",
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
  removeContinuationTrigger();
  runBillingRateCheckBatch(true);
}

/**
 * 1分ごとのトリガーから呼ばれ、続きのバッチを実行する関数（手動実行はしない）
 */
function continueBillingRateCheck() {
  runBillingRateCheckBatch(false);
}

/**
 * 実際の1バッチ分の処理。時間切れになったら進捗を保存して抜け、まだ終わっていなければ
 * 続行用トリガーを仕込む。全件終わったらトリガーを消して完了メールを送る
 *
 * 手動実行と1分ごとの自動継続が同時に走ってシートを取り合う（行の上書き・重複）事故を防ぐため、
 * スクリプトロックを取得できた場合のみ処理する。ロックが取れない場合は「他の処理が実行中」として
 * 何もせず終了する（自動継続なら次の1分後にまた試みられる）
 */
function runBillingRateCheckBatch(isFreshStart) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    console.error("他の請求単価チェックの処理が実行中のため、今回はスキップしました。手動実行の場合は、実行中の処理が終わってからもう一度お試しください。");
    return;
  }

  try {
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

    if (isFreshStart) {
      props.deleteProperty(BILLING_RATE_CHECK_CONFIG.progressLastIdProp);
      props.setProperty(BILLING_RATE_CHECK_CONFIG.progressOkCountProp, "0");
      props.setProperty(BILLING_RATE_CHECK_CONFIG.progressAttentionCountProp, "0");
      sheet.clear();
      sheet.appendRow(["チェック日時", "レコードID", "契約先", "収集業者名", "商品名", "ステータス", "現在の請求単価", "見積りから読み取った金額", "差額", "差額率(%)", "備考", "赤字チェック"]);
    }

    let lastId = Number(props.getProperty(BILLING_RATE_CHECK_CONFIG.progressLastIdProp) || "0");
    let okCount = Number(props.getProperty(BILLING_RATE_CHECK_CONFIG.progressOkCountProp) || "0");
    let attentionCount = Number(props.getProperty(BILLING_RATE_CHECK_CONFIG.progressAttentionCountProp) || "0");
    let finished = false;

    outer:
    while (true) {
    const query = encodeURIComponent(`$id > ${lastId} order by $id asc limit ${KINTONE_PAGE_SIZE}`);
    const url = `https://${subdomain}.cybozu.com/k/v1/records.json?app=${appId}&query=${query}`;
    console.log(`kintoneからレコード取得開始（$id > ${lastId}）...`);
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
    console.log(`${records.length}件のレコードを取得しました（経過 ${Date.now() - startTime}ms）`);
    if (records.length === 0) {
      finished = true;
      break;
    }

    for (let i = 0; i < records.length; i++) {
      if (Date.now() - startTime > EXECUTION_TIME_BUDGET_MS) {
        console.log(`時間切れのため中断します（経過 ${Date.now() - startTime}ms、このページの${i}/${records.length}件目まで処理済み）`);
        break outer; // 時間切れ。ここまでの進捗は保存済みなので、続きは次のトリガーで行う
      }

      const record = records[i];
      console.log(`[${i + 1}/${records.length}] レコード#${record.$id.value} 処理開始（経過 ${Date.now() - startTime}ms）`);

      let rows;
      try {
        rows = buildResultRowsForRecord(record, fileFieldCode, subdomain, apiToken, geminiApiKey);
      } catch (e) {
        console.error(`レコード#${record.$id.value}の処理中に想定外のエラー: ${e.message}`);
        rows = [[Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss"), record.$id.value, "", "", "(全項目)", "エラー", "", "", "", "", e.message, ""]];
      }

      if (rows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
        SpreadsheetApp.flush();
        rows.forEach(r => { if (r[5] === "一致") okCount++; else attentionCount++; });
      }

      lastId = Number(record.$id.value);
      props.setProperty(BILLING_RATE_CHECK_CONFIG.progressLastIdProp, String(lastId));
      props.setProperty(BILLING_RATE_CHECK_CONFIG.progressOkCountProp, String(okCount));
      props.setProperty(BILLING_RATE_CHECK_CONFIG.progressAttentionCountProp, String(attentionCount));
      console.log(`[${i + 1}/${records.length}] レコード#${record.$id.value} 処理完了（経過 ${Date.now() - startTime}ms）`);
    }

    if (records.length < KINTONE_PAGE_SIZE) {
      finished = true;
      break;
    }
  }

    if (finished) {
      removeContinuationTrigger();
      buildStoreSummarySheet();
      sendBillingRateCheckFinalReport(okCount, attentionCount);
      props.deleteProperty(BILLING_RATE_CHECK_CONFIG.progressLastIdProp);
      props.deleteProperty(BILLING_RATE_CHECK_CONFIG.progressOkCountProp);
      props.deleteProperty(BILLING_RATE_CHECK_CONFIG.progressAttentionCountProp);
    } else {
      ensureContinuationTrigger();
    }
  } finally {
    lock.releaseLock();
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
  const contractorName = (record[BILLING_RATE_CHECK_CONFIG.contractorFieldCode] &&
                           record[BILLING_RATE_CHECK_CONFIG.contractorFieldCode].value) || "";

  const tableRows = (record[BILLING_RATE_CHECK_CONFIG.subtableFieldCode] &&
                      record[BILLING_RATE_CHECK_CONFIG.subtableFieldCode].value) || [];

  if (tableRows.length === 0) {
    return [[now, recordId, displayName, contractorName, "(全項目)", "単価テーブルなし", "", "", "", "", "", ""]];
  }

  const files = (record[fileFieldCode] && record[fileFieldCode].value) || [];
  const estimateFiles = files.filter(f => f.name.indexOf(BILLING_RATE_CHECK_CONFIG.estimateFileNameKeyword) !== -1);

  // 見積りとの照合は「見積りが読み取れた場合」のみ行うが、
  // 仕入単価との赤字チェックは見積りの有無に関係なく全レコードで行う
  let extractedItems = null;
  let extractionFailureStatus = ""; // 見積りはあるが読み取れなかった場合のステータス（"抽出失敗" / "エラー"）
  let extractionNote = "";

  if (estimateFiles.length > 0) {
    const targetFile = estimateFiles[0];
    const fileSizeBytes = Number(targetFile.size) || 0;
    const MAX_FILE_SIZE_BYTES = 15 * 1024 * 1024; // 15MB。大きすぎるファイルはAI呼び出しが極端に遅くなる/失敗するため除外

    if (fileSizeBytes > MAX_FILE_SIZE_BYTES) {
      console.error(`見積りファイルのサイズが大きすぎるためスキップ（レコード#${recordId}、${Math.round(fileSizeBytes / 1024 / 1024)}MB）`);
      extractionFailureStatus = "抽出失敗";
      extractionNote = `見積りファイルのサイズが大きすぎます（${Math.round(fileSizeBytes / 1024 / 1024)}MB）。手動で確認してください。`;
    } else {
      try {
        // 複数見積りが添付されている場合は先頭（最新想定）のみをチェック対象にする
        // Gemini APIのレート制限に引っかかりにくくするため、呼び出し前に少し間隔を空ける
        Utilities.sleep(300);
        const blob = fetchKintoneFile(subdomain, targetFile.fileKey, apiToken);
        const extraction = extractEstimateItems(blob, targetFile.contentType, geminiApiKey);

        if (!extraction.items || extraction.items.length === 0) {
          extractionFailureStatus = "抽出失敗";
          extractionNote = extraction.note || "";
        } else {
          extractedItems = extraction.items;
        }
      } catch (e) {
        console.error(`見積書の読み取り中にエラー（レコード#${recordId}）: ` + e.message);
        extractionFailureStatus = "エラー";
        extractionNote = e.message;
      }
    }
  }

  const rows = [];
  tableRows.forEach(row => {
    const itemName = row.value[BILLING_RATE_CHECK_CONFIG.itemNameFieldCode]
      ? row.value[BILLING_RATE_CHECK_CONFIG.itemNameFieldCode].value : "";
    if (!itemName) return; // 商品名が空の行はスキップ

    const currentTankaRaw = row.value[BILLING_RATE_CHECK_CONFIG.tankaFieldCode]
      ? row.value[BILLING_RATE_CHECK_CONFIG.tankaFieldCode].value : "";
    const costRaw = row.value[BILLING_RATE_CHECK_CONFIG.costFieldCode]
      ? row.value[BILLING_RATE_CHECK_CONFIG.costFieldCode].value : "";

    const currentTankaNum = parseAmount(currentTankaRaw);
    const costNum = parseAmount(costRaw);
    const marginAlert = (currentTankaNum !== null && costNum !== null && currentTankaNum < costNum)
      ? `赤字（仕入${costNum}円 > 請求${currentTankaNum}円）`
      : "";

    let status;
    let extractedAmount = "";
    let diff = "";
    let diffPercent = "";
    let note = "";

    if (extractedItems) {
      const matched = findMatchingEstimateItem(extractedItems, itemName);
      if (!matched) {
        status = "見積りに対応項目なし";
      } else {
        extractedAmount = matched.unitPrice;
        const diffInfo = calcDiff(currentTankaNum, matched.unitPrice);
        diff = diffInfo.diff;
        diffPercent = diffInfo.diffPercent;
        if (currentTankaNum === null) status = "請求単価未入力";
        else if (currentTankaNum === matched.unitPrice) status = "一致";
        else status = "不一致";
      }
    } else if (extractionFailureStatus) {
      status = extractionFailureStatus;
      note = extractionNote;
    } else {
      status = "見積り未添付";
    }

    rows.push([now, recordId, displayName, contractorName, itemName, status, currentTankaRaw, extractedAmount, diff, diffPercent, note, marginAlert]);
  });

  if (rows.length === 0) {
    // 単価テーブルはあるが、商品名がすべて空だった場合など
    return [[now, recordId, displayName, contractorName, "(全項目)", extractionFailureStatus || "見積り未添付", "", "", "", "", extractionNote, ""]];
  }

  return rows;
}

/**
 * 現在の請求単価と見積り金額の差額・差額率(%)を計算する（比較できない場合は空文字を返す）
 */
function calcDiff(currentNum, extractedNum) {
  if (currentNum === null || extractedNum === null || extractedNum === undefined || isNaN(extractedNum)) {
    return { diff: "", diffPercent: "" };
  }
  const diff = extractedNum - currentNum;
  const diffPercent = currentNum !== 0 ? Math.round((diff / currentNum) * 1000) / 10 : "";
  return { diff, diffPercent };
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

  const MAX_ATTEMPTS = 3;
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const code = response.getResponseCode();
      const bodyText = response.getContentText();

      // レート制限（429）は少し待ってから再試行する
      if (code === 429 && attempt < MAX_ATTEMPTS) {
        Utilities.sleep(2000 * attempt);
        continue;
      }

      if (code !== 200) {
        console.error(`Gemini API呼び出し失敗 (HTTP ${code}): ` + bodyText);
        return { items: [], note: `AI呼び出し失敗 (HTTP ${code}): ` + bodyText.substring(0, 200) };
      }

      const json = JSON.parse(bodyText);
      const text = json.candidates && json.candidates[0].content.parts[0].text;
      if (!text) {
        console.error("Geminiの応答に想定した内容が含まれていませんでした: " + bodyText);
        return { items: [], note: "AIから内容を読み取れませんでした（想定外の応答形式）。" };
      }

      const cleanJsonStr = text.replace(/```json/g, "").replace(/```/g, "").trim();
      const parsed = JSON.parse(cleanJsonStr);
      const items = Array.isArray(parsed.items) ? parsed.items
        .map(it => ({ itemName: String(it.itemName || ""), unitPrice: Number(it.unitPrice) }))
        .filter(it => it.itemName && !isNaN(it.unitPrice)) : [];

      return { items, note: parsed.note || "" };
    } catch (e) {
      console.error("見積り品目抽出のGemini呼び出しでエラー: " + e.message);
      if (attempt >= MAX_ATTEMPTS) return { items: [], note: "AI呼び出し中にエラーが発生しました: " + e.message };
      Utilities.sleep(1000 * attempt);
    }
  }
  return { items: [], note: "AI呼び出しがリトライ上限に達しました。" };
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
 * 「請求単価チェック結果」シート（商品ごとの詳細、全件処理後の最終状態）を集計し、
 * 契約先（レコード）ごとに1行の「要対応店舗一覧」シートを作る。
 * 「見積り未添付」は業者ルールでの運用があり得るため、要対応件数には含めない（別列で件数のみ表示）。
 * 差額率(%)が大きい順・要対応の有無の順に並べ替えるので、上から順に見れば「明らかにおかしいもの」から確認できる。
 */
function buildStoreSummarySheet() {
  const spreadsheet = getOrCreateResultSpreadsheet();
  const detailSheet = getOrCreateSheet(spreadsheet, BILLING_RATE_CHECK_CONFIG.resultSheetName);
  const lastRow = detailSheet.getLastRow();
  if (lastRow < 2) return; // 見出しのみ（データなし）

  const data = detailSheet.getRange(2, 1, lastRow - 1, 12).getValues();
  const STATUS_LIST = ["一致", "不一致", "請求単価未入力", "見積りに対応項目なし", "見積り未添付", "単価テーブルなし", "抽出失敗", "エラー"];
  const NEEDS_ATTENTION_STATUSES = ["不一致", "請求単価未入力", "見積りに対応項目なし", "単価テーブルなし", "抽出失敗", "エラー"];

  const summaryByRecord = {};

  data.forEach(r => {
    const recordId = r[1];
    const displayName = r[2];
    const contractorName = r[3];
    const status = r[5];
    const diffPercent = r[9];
    const marginAlert = r[11];

    if (!summaryByRecord[recordId]) {
      const counts = {};
      STATUS_LIST.forEach(s => counts[s] = 0);
      summaryByRecord[recordId] = { recordId, displayName, contractorName, counts, maxAbsDiffPercent: 0, marginAlertCount: 0 };
    }

    const entry = summaryByRecord[recordId];
    if (entry.counts[status] !== undefined) entry.counts[status]++;
    if (typeof diffPercent === "number" && Math.abs(diffPercent) > entry.maxAbsDiffPercent) {
      entry.maxAbsDiffPercent = Math.abs(diffPercent);
    }
    if (marginAlert) entry.marginAlertCount++;
  });

  const header = ["レコードID", "契約先", "収集業者名", "一致", "不一致", "請求単価未入力",
    "見積りに対応項目なし", "見積り未添付", "単価テーブルなし", "抽出失敗", "エラー", "赤字件数", "最大差額率(%)", "要対応"];

  const bodyRows = Object.keys(summaryByRecord).map(recordId => {
    const e = summaryByRecord[recordId];
    const needsAttentionCount = NEEDS_ATTENTION_STATUSES.reduce((sum, s) => sum + e.counts[s], 0) + e.marginAlertCount;
    return [
      e.recordId, e.displayName, e.contractorName,
      e.counts["一致"], e.counts["不一致"], e.counts["請求単価未入力"], e.counts["見積りに対応項目なし"],
      e.counts["見積り未添付"], e.counts["単価テーブルなし"], e.counts["抽出失敗"], e.counts["エラー"],
      e.marginAlertCount, e.maxAbsDiffPercent, needsAttentionCount > 0 ? "要対応" : ""
    ];
  });

  bodyRows.sort((a, b) => {
    const aNeeds = a[13] === "要対応" ? 1 : 0;
    const bNeeds = b[13] === "要対応" ? 1 : 0;
    if (aNeeds !== bNeeds) return bNeeds - aNeeds; // 要対応のものを先に
    if (a[11] !== b[11]) return (b[11] || 0) - (a[11] || 0); // 赤字件数が多い順
    return (b[12] || 0) - (a[12] || 0); // 最大差額率(%)が大きい順
  });

  const summarySheet = getOrCreateSheet(spreadsheet, BILLING_RATE_CHECK_CONFIG.summarySheetName);
  summarySheet.clear();
  summarySheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (bodyRows.length > 0) {
    summarySheet.getRange(2, 1, bodyRows.length, header.length).setValues(bodyRows);
  }
}

/**
 * 全件処理が完了した際に、件数のサマリーをメールで送る（詳細はスプレッドシート参照）
 */
function sendBillingRateCheckFinalReport(okCount, attentionCount) {
  const total = okCount + attentionCount;
  const body = `対象 ${total}件中、一致 ${okCount}件・要確認 ${attentionCount}件でした。\n\n` +
    "・商品ごとの詳細は「請求単価チェック結果」シート\n" +
    "・どの契約先（店舗）を直すべきかは「要対応店舗一覧」シート（差額率が大きい順に並んでいます）\n\n" +
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

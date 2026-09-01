// ==========================================
//  月報自動化システム 【機能完全網羅・エラー根絶版】
// ==========================================

const IS_TEST = false;

const BUDGET_ROW_NEW_COUNT    = 15;
const BUDGET_ROW_CANCEL_COUNT = 22;
const BUDGET_ROW_CANCEL_PROFIT = 34;

/**
 * 「対象月」を1箇所で決める。generateMonthlyReportRequestと
 * receiveGeminiResponseAndWriteToSheetの両方がここを参照することで、
 * 返信メールから対象月を読み取れない場合でも「関数①が対象にした月＝先月」
 * という同じ基準で追従できるようにする。
 */
function getTargetYearMonth() {
  if (IS_TEST) return { year: 2026, month: 4 };
  const today = new Date();
  const lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  return { year: lastMonth.getFullYear(), month: lastMonth.getMonth() + 1 };
}

function generateMonthlyReportRequest() {
  const config = getConfig();
  const { year: targetYear, month: targetMonth } = getTargetYearMonth();

  const baseSubject = targetYear + "年" + targetMonth + "月実績";
  Logger.log("【関数①】本番API連携起動: " + baseSubject);

  const ss = getTargetSpreadsheet(config["月報・予算スプレッドシートURL"]);

  const contractFields = ['収集開始日', '数値_0', '契約店舗名称', '案件番号'];
  const rawContracts = fetchKintoneRecords(config["kintoneサブドメイン"], config["アプリID_契約管理"], config["APIトークン_契約管理"], contractFields);
  writeRawSheet(ss, 'kintone_契約管理', contractFields, rawContracts);

  const cancelFields = ['解約日', '管理費', '案件番号', '契約店舗名称'];
  const rawCancels = fetchKintoneRecords(config["kintoneサブドメイン"], config["アプリID_解約リスト"], config["APIトークン_解約リスト"], cancelFields);
  writeRawSheet(ss, 'kintone_解約リスト', cancelFields, rawCancels);

  const leadFields = ['収集希望開始日', '情報源小分類', '業態詳細', '情報源会社名', '情報源紹介者名', '案件番号', '案件担当者', '顧客ステータス', '問合せ日', '収集開始日_廃棄物', '収集開始月'];
  const rawLeads = fetchKintoneRecords(config["kintoneサブドメイン"], config["アプリID_案件管理"], config["APIトークン_案件管理"], leadFields);
  writeRawSheet(ss, 'kintone_案件管理', leadFields, rawLeads);

  const budgetSheet = ss.getSheetByName("予算");

  const summaryReportData = doMathAggregationFromAPI(rawLeads, rawContracts, rawCancels, budgetSheet, targetYear, targetMonth);
  const rawBusinessTypes = extractBusinessTypesFromAPI(rawContracts, rawLeads, targetYear, targetMonth);
  const rawReferrers = extractReferrersFromAPI(rawLeads, targetYear, targetMonth);
  const rawLeadBusinessTypes = extractLeadBusinessTypesFromAPI(rawLeads, targetYear, targetMonth);

  let emailSubject = "【月報自動生成】" + baseSubject;
  if (IS_TEST) emailSubject += "_テスト";

  let sheetName = baseSubject.trim();
  if (IS_TEST) sheetName += "_テスト";

  let targetSheet = ss.getSheetByName(sheetName);
  if (targetSheet) ss.deleteSheet(targetSheet);

  const templateSheet = ss.getSheetByName("月報フォーマット");
  if (!templateSheet) throw new Error("'月報フォーマット' が見つかりません。");

  targetSheet = templateSheet.copyTo(ss).setName(sheetName);
  targetSheet.showSheet();
  ss.setActiveSheet(targetSheet);
  ss.moveActiveSheet(1);

  targetSheet.getRange("C2").setValue(summaryReportData.months.m1 + "月実績");
  targetSheet.getRange("D2").setValue(summaryReportData.months.m2 + "月実績");
  targetSheet.getRange("E2").setValue(summaryReportData.months.m3 + "月実績");
  targetSheet.getRange("J2").setValue(summaryReportData.months.m1 + "月予算");
  targetSheet.getRange("K2").setValue(summaryReportData.months.m2 + "月予算");
  targetSheet.getRange("L2").setValue(summaryReportData.months.m3 + "月予算");

  targetSheet.getRange("C28").setValue(summaryReportData.months.m1 + "月実績");
  targetSheet.getRange("D28").setValue(summaryReportData.months.m2 + "月実績");
  targetSheet.getRange("E28").setValue(summaryReportData.months.m3 + "月実績");
  targetSheet.getRange("G8").setValue("↓" + summaryReportData.months.m3 + "月実績備考");
  targetSheet.getRange("G28").setValue("↓" + summaryReportData.months.m3 + "月実績備考        ");

  targetSheet.getRange("B58").setValue("成約率（" + summaryReportData.months.m3 + "月）安藤");
  targetSheet.getRange("G58").setValue("成約率（" + summaryReportData.months.m1 + "," + summaryReportData.months.m2 + "," + summaryReportData.months.m3 + "月）安藤");
  targetSheet.getRange("B67").setValue("成約率（" + summaryReportData.months.m3 + "月）五十嵐");
  targetSheet.getRange("G67").setValue("成約率（" + summaryReportData.months.m1 + "," + summaryReportData.months.m2 + "," + summaryReportData.months.m3 + "月）五十嵐");
  targetSheet.getRange("B76").setValue("成約率（" + summaryReportData.months.m3 + "月）瀬戸口");
  targetSheet.getRange("G76").setValue("成約率（" + summaryReportData.months.m1 + "," + summaryReportData.months.m2 + "," + summaryReportData.months.m3 + "月）瀬戸口");
  targetSheet.getRange("B85").setValue("成約率（" + summaryReportData.months.m3 + "月）その他スタッフ");
  targetSheet.getRange("G85").setValue("成約率（" + summaryReportData.months.m1 + "," + summaryReportData.months.m2 + "," + summaryReportData.months.m3 + "月）その他スタッフ");

  for (let rangeKey in summaryReportData.matrix) {
    targetSheet.getRange(rangeKey).setValues(summaryReportData.matrix[rangeKey]);
  }

  const emailBody = buildGeminiPrompt(baseSubject, rawBusinessTypes, rawReferrers, rawLeadBusinessTypes);
  GmailApp.sendEmail(config["通知先メールアドレス"], emailSubject, emailBody);

  Utilities.sleep(1500);
  const sentThreads = GmailApp.search('subject:"' + emailSubject + '"', 0, 1);
  if (sentThreads.length > 0) {
    sentThreads[0].markRead();
    sentThreads[0].moveToArchive();
  }
  Logger.log("【関数①】実行完了。メール送信および送信履歴お片付けに成功しました。");
}

/**
 * AIからの返信メールを解析してシートへ反映する。
 * 未読の「月報」スレッドは全件走査し、実際に書き込みまで完了したスレッドのみ既読化・アーカイブする
 * （旧実装は最新1件のみ処理し、残りの未読スレッドを中身を見ずに一括で既読化していたため、
 * 　処理が1日以上遅れる・テストで複数回実行するなどして未読が複数溜まると2件目以降のデータが
 * 　永久に失われるバグがあった）。
 */
function receiveGeminiResponseAndWriteToSheet() {
  const config = getConfig();
  const threads = GmailApp.search('subject:月報 is:unread', 0, 20);

  if (threads.length === 0) {
    Logger.log("【関数②】未読の月報返信メールが見つかりませんでした。");
    return;
  }

  const userLabel = GmailApp.getUserLabelByName("月報ログ");
  const ss = getTargetSpreadsheet(config["月報・予算スプレッドシートURL"]);

  const extractRealAnswer = (text, tag) => {
    const regex = new RegExp("\\[" + tag + "\\]\\s*([\\s\\S]*?)(?=\\[|$)", "g");
    let match;
    while ((match = regex.exec(text)) !== null) {
      let content = match[1].trim();
      if (content && content.indexOf("（ここに") === -1) {
        if (content === "(なし)" || content === "（なし）" || content === "なし" || content.indexOf("紹介者まとめ）") !== -1 || content.indexOf("業態まとめ）") !== -1) {
          return "";
        }
        content = content.replace(/\n*【[^】]+】\s*$/, "").trim();
        if (content.indexOf("まとめ") !== -1 && content.indexOf("×") === -1 && content.indexOf(",") === -1 && content.length < 20) {
          return "";
        }
        return content;
      }
    }
    return "";
  };

  // 古い返信から順に処理する（届いた順にシートへ反映するため）
  const orderedThreads = threads.slice().reverse();

  orderedThreads.forEach(thread => {
    try {
      const latestMessage = thread.getMessages().pop();
      const mailBody = latestMessage.getPlainBody();
      const mailSubject = latestMessage.getSubject();

      if (mailSubject.indexOf("週報") !== -1) return; // 週報スレッドは対象外（未読のまま残し、週報側の処理に委ねる）
      if (!mailBody.includes("[START_OF_MONTHLY_REPORT]") || !mailBody.includes("[END_OF_MONTHLY_REPORT]")) return; // まだAIの返信が来ていない

      // Workspace Flows経由の完成メール（件名「【月報完成】」など）は年月の文字列を含まないため、
      // その場合は「関数①が対象にした月＝先月」とみなして処理する（generateMonthlyReportRequestと
      // 同じgetTargetYearMonth()を使うことで基準がずれないようにしている）。
      const dateMatch = mailBody.match(/\d{4}年\d{1,2}月実績/) || mailSubject.match(/\d{4}年\d{1,2}月実績/);
      let sheetName;
      if (dateMatch) {
        sheetName = dateMatch[0];
      } else {
        const { year, month } = getTargetYearMonth();
        sheetName = year + "年" + month + "月実績";
        Logger.log("【関数②】件名・本文から対象月を特定できなかったため、直近の対象月（" + sheetName + "）とみなして処理します（件名: " + mailSubject + "）。");
      }
      if (IS_TEST && sheetName.indexOf("_テスト") === -1) sheetName += "_テスト";

      const targetSheet = ss.getSheetByName(sheetName);
      if (!targetSheet) {
        Logger.log("【関数②】シート「" + sheetName + "」が見つからないためスキップしました。");
        return;
      }

      targetSheet.getRange("G10").setValue(extractRealAnswer(mailBody, "CELL_G10_CONTRACT_GOMI_REF"));
      targetSheet.getRange("G11").setValue(extractRealAnswer(mailBody, "CELL_G11_CONTRACT_BUILD_REF"));
      targetSheet.getRange("G12").setValue(extractRealAnswer(mailBody, "CELL_G12_CONTRACT_FUDOSAN_REF"));
      targetSheet.getRange("G13").setValue(extractRealAnswer(mailBody, "CELL_G13_CONTRACT_FIJ_REF"));
      targetSheet.getRange("G14").setValue(extractRealAnswer(mailBody, "CELL_G14_CONTRACT_OSHIBORI_REF"));
      targetSheet.getRange("G15").setValue(extractRealAnswer(mailBody, "CELL_G15_CONTRACT_TENPOS_REF"));
      targetSheet.getRange("G16").setValue(extractRealAnswer(mailBody, "CELL_G16_CONTRACT_ZOTEN_REF"));
      targetSheet.getRange("G17").setValue(extractRealAnswer(mailBody, "CELL_G17_CONTRACT_TOBI_BIZ"));
      targetSheet.getRange("G18").setValue(extractRealAnswer(mailBody, "CELL_G18_CONTRACT_TELE_BIZ"));
      targetSheet.getRange("G19").setValue(extractRealAnswer(mailBody, "CELL_G19_CONTRACT_DM_BIZ"));
      targetSheet.getRange("G20").setValue(extractRealAnswer(mailBody, "CELL_G20_CONTRACT_LP_TEL_BIZ"));
      targetSheet.getRange("G21").setValue(extractRealAnswer(mailBody, "CELL_G21_CONTRACT_LP_MAIL_BIZ"));
      targetSheet.getRange("G22").setValue(extractRealAnswer(mailBody, "CELL_G22_CONTRACT_LP_LINE_BIZ"));
      targetSheet.getRange("G23").setValue(extractRealAnswer(mailBody, "CELL_G23_CONTRACT_HP_TEL_BIZ"));
      targetSheet.getRange("G24").setValue(extractRealAnswer(mailBody, "CELL_G24_CONTRACT_HP_LINE_BIZ")); // 送信側プロンプトのタグ名と一致させた（旧: CELL_G24_HP_LINE_BIZ で常に空欄になっていた）
      targetSheet.getRange("G25").setValue(extractRealAnswer(mailBody, "CELL_G25_HP_MAIL_BIZ"));
      targetSheet.getRange("G26").setValue(extractRealAnswer(mailBody, "CELL_G26_CONTRACT_OTHER_BIZ"));

      targetSheet.getRange("G29").setValue(extractRealAnswer(mailBody, "CELL_G29_LEAD_GOMI_REF"));
      targetSheet.getRange("G30").setValue(extractRealAnswer(mailBody, "CELL_G30_LEAD_BUILD_REF"));
      targetSheet.getRange("G31").setValue(extractRealAnswer(mailBody, "CELL_G31_LEAD_FUDOSAN_REF"));
      targetSheet.getRange("G32").setValue(extractRealAnswer(mailBody, "CELL_G32_LEAD_FIJ_REF"));
      targetSheet.getRange("G33").setValue(extractRealAnswer(mailBody, "CELL_G33_LEAD_OSHIBORI_REF"));
      targetSheet.getRange("G34").setValue(extractRealAnswer(mailBody, "CELL_G34_LEAD_TENPOS_REF"));
      targetSheet.getRange("G35").setValue(extractRealAnswer(mailBody, "CELL_G35_LEAD_ZOTEN_REF"));
      targetSheet.getRange("G36").setValue(extractRealAnswer(mailBody, "CELL_G36_LEAD_TOBI_BIZ"));
      targetSheet.getRange("G37").setValue(extractRealAnswer(mailBody, "CELL_G37_LEAD_TELE_BIZ"));
      targetSheet.getRange("G38").setValue(extractRealAnswer(mailBody, "CELL_G38_LEAD_DM_BIZ"));
      targetSheet.getRange("G39").setValue(extractRealAnswer(mailBody, "CELL_G39_LEAD_LP_TEL_BIZ"));
      targetSheet.getRange("G40").setValue(extractRealAnswer(mailBody, "CELL_G40_LEAD_LP_MAIL_BIZ"));
      targetSheet.getRange("G41").setValue(extractRealAnswer(mailBody, "CELL_G41_LEAD_LP_LINE_BIZ"));
      targetSheet.getRange("G42").setValue(extractRealAnswer(mailBody, "CELL_G42_LEAD_HP_TEL_BIZ"));
      targetSheet.getRange("G43").setValue(extractRealAnswer(mailBody, "CELL_G43_LEAD_HP_LINE_BIZ"));
      targetSheet.getRange("G44").setValue(extractRealAnswer(mailBody, "CELL_G44_LEAD_HP_MAIL_BIZ"));
      targetSheet.getRange("G45").setValue(extractRealAnswer(mailBody, "CELL_G45_LEAD_OTHER_BIZ"));

      targetSheet.getRange("G10:G26").setWrap(true);
      targetSheet.getRange("G29:G45").setWrap(true);
      targetSheet.autoResizeRows(10, 17);
      targetSheet.autoResizeRows(29, 17);

      thread.markRead();
      thread.moveToArchive();
      if (userLabel) thread.addLabel(userLabel);
      Logger.log("【関数②】" + sheetName + " へのマッピング＆お掃除が完了しました。");

    } catch (e) {
      // このスレッドは未読のまま残し、次回実行時に再試行する
      Logger.log("【関数②】スレッド処理中にエラーが発生したためスキップしました: " + e.message);
    }
  });
}

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("config");
  if (!sheet) throw new Error("「config」タブが見つかりません。");
  const data = sheet.getDataRange().getValues();
  const config = {};
  data.forEach(row => { if (row[0]) config[row[0].toString().trim()] = row[1] ? row[1].toString().trim() : ""; });
  return config;
}

function getTargetSpreadsheet(url) {
  if (url && url.indexOf("http") === 0) return SpreadsheetApp.openByUrl(url);
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * kintoneアプリの全レコードをREST APIで取得する（$idベースでページング）。
 * HTTPエラー時は例外を投げて処理を止める（旧実装はbreakするだけで、
 * トークンやフィールドコードの間違いがあってもエラーなしで空データのまま
 * 集計が進んでしまい、原因が分からず「なぜか反映されない」状態になっていた）。
 */
function fetchKintoneRecords(domain, appId, token, fields) {
  if (!appId || !token) return [];
  if (fields.indexOf('$id') === -1) fields.push('$id');
  let allRecords = []; let lastId = 0; const limit = 500;
  while (true) {
    const query = encodeURIComponent(`$id > ${lastId} order by $id asc limit ${limit}`);
    const url = `https://${domain}/k/v1/records.json?app=${appId}&fields=${fields.join(',')}&query=${query}`;
    const response = UrlFetchApp.fetch(url, { method: 'get', headers: { 'X-Cybozu-API-Token': token }, muteHttpExceptions: true });
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

function writeRawSheet(ss, sheetName, fields, records) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  sheet.clear(); if (records.length === 0) return;
  const output = [fields].concat(records.map(rec => fields.map(f => rec[f] ? rec[f].value : '')));
  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  sheet.hideSheet();
}

function doMathAggregationFromAPI(rawLeads, rawContracts, rawCancels, budgetSheet, year, month) {
  const m1 = month - 2 < 1 ? month + 10 : month - 2;
  const m2 = month - 1 < 1 ? month + 11 : month - 1;
  const m3 = month;
  const m1Prefix = (month - 2 < 1 ? year - 1 : year) + "-" + ("0" + m1).slice(-2);
  const m2Prefix = (month - 1 < 1 ? year - 1 : year) + "-" + ("0" + m2).slice(-2);
  const m3Prefix = year + "-" + ("0" + m3).slice(-2);

  const channels = ["ゴミ業者", "ビル管理会社", "不動産", "FIJ", "おしぼり", "てんぽす", "増店", "飛び込み", "テレアポ", "DM", "LP(電話)", "LP(メール)", "LP(ライン)", "HP(電話)", "HP(ライン)", "HP(メール)", "その他"];
  let contractCounts = {}; let leadCounts = {};
  channels.forEach(ch => { contractCounts[ch] = [0, 0, 0]; leadCounts[ch] = [0, 0, 0]; });
  let grossProfitNew = [0, 0, 0]; let grossProfitCancel = [0, 0, 0]; let newContractCount = [0, 0, 0]; let cancelCount = [0, 0, 0];
  const targetRoutes = ["紹介", "増店", "LP", "HP", "その他"];
  const targetStaffs = ["安藤", "五十嵐", "瀬戸口", "その他"];
  let leadsData = {}; let winsData = {};
  targetStaffs.forEach(s => { leadsData[s] = {}; winsData[s] = {}; targetRoutes.forEach(r => { leadsData[s][r] = [0, 0, 0]; winsData[s][r] = [0, 0, 0]; }); });

  function getChannelName(source) { let chName = "その他"; for (let ch of channels) { if (source.includes(ch) || ch.includes(source)) { chName = ch; break; } } return chName; }
  function getRouteKey(chName) { if (chName === "増店") return "増店"; if (chName.indexOf("LP") === 0) return "LP"; if (chName.indexOf("HP") === 0) return "HP"; if (chName === "その他") return "その他"; return "紹介"; }
  function getStaffKeyFromRecord(rec) {
    let staffStr = ""; if (rec['案件担当者'] && rec['案件担当者'].value) { if (Array.isArray(rec['案件担当者'].value)) { staffStr = rec['案件担当者'].value.map(u => u.name || u.code || "").join(","); } else { staffStr = String(rec['案件担当者'].value); } }
    const sLower = staffStr.toLowerCase().trim(); if (!sLower) return "その他"; if (sLower.includes('ando') || sLower.includes('安藤')) return '安藤'; if (sLower.includes('igarasi') || sLower.includes('igarashi') || sLower.includes('五十嵐')) return '五十嵐'; if (sLower.includes('setoguchi') || sLower.includes('setoguti') || sLower.includes('瀬戸口')) return '瀬戸口'; return 'その他';
  }

  rawLeads.forEach(rec => {
    const askDate = (rec['問合せ日'] && rec['問合せ日'].value) ? String(rec['問合せ日'].value) : ""; if (!askDate) return;
    let idx = -1; if (askDate.indexOf(m1Prefix) === 0) idx = 0; else if (askDate.indexOf(m2Prefix) === 0) idx = 1; else if (askDate.indexOf(m3Prefix) === 0) idx = 2;
    if (idx !== -1) { const chName = getChannelName(rec.情報源小分類 ? String(rec.情報源小分類.value) : "その他"); leadCounts[chName][idx]++; const staffKey = getStaffKeyFromRecord(rec); const routeKey = getRouteKey(chName); leadsData[staffKey][routeKey][idx]++; }
  });
  rawLeads.forEach(rec => {
    if (rec['顧客ステータス'] && String(rec['顧客ステータス'].value) === '契約') {
      let contractDate = (rec['収集開始日_廃棄物'] && rec['収集開始日_廃棄物'].value) ? String(rec['収集開始日_廃棄物'].value) : ""; if (!contractDate) { contractDate = (rec['収集開始月'] && rec['収集開始月'].value) ? String(rec['収集開始月'].value) : ""; } if (!contractDate) return;
      let idx = -1; if (contractDate.indexOf(m1Prefix) === 0) idx = 0; else if (contractDate.indexOf(m2Prefix) === 0) idx = 1; else if (contractDate.indexOf(m3Prefix) === 0) idx = 2;
      if (idx !== -1) { const chName = getChannelName(rec.情報源小分類 ? String(rec.情報源小分類.value) : "その他"); contractCounts[chName][idx]++; const staffKey = getStaffKeyFromRecord(rec); const routeKey = getRouteKey(chName); winsData[staffKey][routeKey][idx]++; }
    }
  });
  rawContracts.forEach(rec => {
    const date = (rec.収集開始日 && rec.収集開始日.value) ? String(rec.収集開始日.value) : ""; if (!date) return;
    let idx = -1; if (date.indexOf(m1Prefix) === 0) idx = 0; else if (date.indexOf(m2Prefix) === 0) idx = 1; else if (date.indexOf(m3Prefix) === 0) idx = 2;
    if (idx !== -1) { newContractCount[idx]++; grossProfitNew[idx] += rec.数値_0 ? parseFloat(rec.数値_0.value) || 0 : 0; }
  });
  rawCancels.forEach(rec => {
    const realDate = (rec.解約日 && rec.解約日.value) ? String(rec.解約日.value) : ""; if (!realDate) return;
    let idx = -1; if (realDate.indexOf(m1Prefix) === 0) idx = 0; else if (realDate.indexOf(m2Prefix) === 0) idx = 1; else if (realDate.indexOf(m3Prefix) === 0) idx = 2;
    if (idx !== -1) { cancelCount[idx]++; grossProfitCancel[idx] += rec.管理費 ? parseFloat(rec.管理費.value) || 0 : 0; }
  });

  let grossProfitNet = [0, 0, 0]; let netIncreaseCount = [0, 0, 0];
  for (let i = 0; i < 3; i++) { grossProfitNew[i] = Math.round(grossProfitNew[i]); grossProfitCancel[i] = Math.round(grossProfitCancel[i]); grossProfitNet[i] = grossProfitNew[i] - grossProfitCancel[i]; netIncreaseCount[i] = newContractCount[i] - cancelCount[i]; }
  let contractMatrix = []; let leadMatrix = []; let totalContractsM1 = 0; let totalContractsM2 = 0; let totalContractsM3 = 0; let totalLeadsM1 = 0; let totalLeadsM2 = 0; let totalLeadsM3 = 0;
  channels.forEach(ch => { contractMatrix.push(contractCounts[ch]); leadMatrix.push(leadCounts[ch]); totalContractsM1 += contractCounts[ch][0]; totalContractsM2 += contractCounts[ch][1]; totalContractsM3 += contractCounts[ch][2]; totalLeadsM1 += leadCounts[ch][0]; totalLeadsM2 += leadCounts[ch][1]; totalLeadsM3 += leadCounts[ch][2]; });
  let displayCancelCount = [cancelCount[0], cancelCount[1], cancelCount[2]]; let displayGrossCancel = [grossProfitCancel[0], grossProfitCancel[1], grossProfitCancel[2]];
  let b_net_profit = [0, 0, 0]; let b_net_count = [0, 0, 0]; let b_new_profit = [0, 0, 0]; let b_new_count = [0, 0, 0]; let b_cancel_profit = [0, 0, 0]; let b_cancel_count = [0, 0, 0];
  try {
    if (budgetSheet) {
      const bData = budgetSheet.getDataRange().getValues(); const headerRow = bData[12]; const targetMonths = [m1, m2, m3];
      for (let i = 0; i < 3; i++) {
        const m = targetMonths[i]; let colIdx = -1; for (let c = 0; c < headerRow.length; c++) { if (String(headerRow[c]).indexOf(m + "月") !== -1) { colIdx = c; break; } }
        if (colIdx !== -1) {
          const budgetNewCount = parseFloat(bData[BUDGET_ROW_NEW_COUNT - 1][colIdx]) || 0; const budgetCancelCount = parseFloat(bData[BUDGET_ROW_CANCEL_COUNT - 1][colIdx]) || 0; const budgetCancelProfit = (parseFloat(bData[BUDGET_ROW_CANCEL_PROFIT - 1][colIdx]) || 0) * 1000;
          b_new_count[i] = budgetNewCount; b_cancel_count[i] = Math.abs(budgetCancelCount); b_net_count[i] = budgetNewCount + budgetCancelCount; b_new_profit[i] = budgetNewCount * 7000; b_cancel_profit[i] = Math.abs(budgetCancelProfit); b_net_profit[i] = (budgetNewCount * 7000) + budgetCancelProfit;
        }
      }
    }
  } catch (e) { Logger.log("予算自動取得エラー: " + e.message); }

  function getAvgValuesOnly(matrixData) { let avgMat = []; matrixData.forEach(row => { avgMat.push([Math.round((row[0] + row[1] + row[2]) / 3)]); }); return { avg: avgMat }; }
  let cNetProfitRes = getAvgValuesOnly([grossProfitNet]); let cNetCountRes = getAvgValuesOnly([netIncreaseCount]); let cNewProfitRes = getAvgValuesOnly([grossProfitNew]); let cNewCountRes = getAvgValuesOnly([newContractCount]); let cCancelProfitRes = getAvgValuesOnly([displayGrossCancel]); let cCancelCountRes = getAvgValuesOnly([displayCancelCount]); let contractRes = getAvgValuesOnly(contractMatrix); let contractTotalRes = getAvgValuesOnly([[totalContractsM1, totalContractsM2, totalContractsM3]]); let leadRes = getAvgValuesOnly(leadMatrix); let leadTotalRes = getAvgValuesOnly([[totalLeadsM1, totalLeadsM2, totalLeadsM3]]);
  let totalRows = []; let gLeads = 0; let gWins = 0;
  targetRoutes.forEach(r => { let rLeads = 0; let rWins = 0; targetStaffs.forEach(s => { for(let i=0; i<3; i++) { rLeads += leadsData[s][r][i]; rWins += winsData[s][r][i]; } }); let avgLeads = Math.round(rLeads / 3); let avgWins = Math.round(rWins / 3); totalRows.push([avgLeads, avgWins, avgLeads > 0 ? (avgWins / avgLeads) : 0]); gLeads += avgLeads; gWins += avgWins; });
  let totalMatrix = totalRows.concat([[gLeads, gWins, gLeads > 0 ? (gWins / gLeads) : 0]]);

  function buildStaffMatrix(s) {
    let singleRows = []; let totalRows = []; let leadSingleSum = 0; let winSingleSum = 0; let leadTotalSum = 0; let winTotalSum = 0;
    targetRoutes.forEach(r => {
      let lSingle = leadsData[s][r][2]; let wSingle = winsData[s][r][2]; singleRows.push([lSingle, wSingle, lSingle > 0 ? (wSingle / lSingle) : 0]); leadSingleSum += lSingle; winSingleSum += wSingle;
      let lTotal = leadsData[s][r][0] + leadsData[s][r][1] + leadsData[s][r][2]; let wTotal = winsData[s][r][0] + winsData[s][r][1] + winsData[s][r][2]; totalRows.push([lTotal, wTotal, lTotal > 0 ? (wTotal / lTotal) : 0]); leadTotalSum += lTotal; winTotalSum += wTotal;
    });
    singleRows.push([leadSingleSum, winSingleSum, leadSingleSum > 0 ? (winSingleSum / leadSingleSum) : 0]); totalRows.push([leadTotalSum, winTotalSum, leadTotalSum > 0 ? (winTotalSum / leadTotalSum) : 0]); return { single: singleRows, total: totalRows };
  }

  return {
    "months": { "m1": m1, "m2": m2, "m3": m3 },
    "matrix": {
      "C3:E3": [grossProfitNet], "F3": cNetProfitRes.avg, "C4:E4": [netIncreaseCount], "F4": cNetCountRes.avg, "C5:E5": [grossProfitNew], "F5": cNewProfitRes.avg, "C6:E6": [newContractCount], "F6": cNewCountRes.avg, "C7:E7": [displayGrossCancel], "F7": cCancelProfitRes.avg, "C8:E8": [displayCancelCount], "F8": cCancelCountRes.avg, "C10:E26": contractMatrix, "F10:F26": contractRes.avg, "C27:E27": [[totalContractsM1, totalContractsM2, totalContractsM3]], "F27": contractTotalRes.avg, "C29:E45": leadMatrix, "F29:F45": leadRes.avg, "C46:E46": [[totalLeadsM1, totalLeadsM2, totalLeadsM3]], "F46": leadTotalRes.avg, "J3:L3": [b_net_profit], "J4:L4": [b_net_count], "J5:L5": [b_new_profit], "J6:L6": [b_new_count], "J7:L7": [b_cancel_profit], "J8:L8": [b_cancel_count], "C51:E56": totalMatrix, "C60:E65": buildStaffMatrix("安藤").single, "H60:J65": buildStaffMatrix("安藤").total, "C69:E74": buildStaffMatrix("五十嵐").single, "H69:J74": buildStaffMatrix("五十嵐").total, "C78:E83": buildStaffMatrix("瀬戸口").single, "H78:J83": buildStaffMatrix("瀬戸口").total, "C87:E92": buildStaffMatrix("その他").single, "H87:J92": buildStaffMatrix("その他").total
    }
  };
}

function extractBusinessTypesFromAPI(rawContracts, rawLeads, year, month) {
  let textRows = []; const targetPrefix = year + "-" + ("0" + month).slice(-2);
  rawContracts.forEach(rec => {
    const startDate = (rec.収集開始日 && rec.収集開始日.value) ? String(rec.収集開始日.value) : "";
    if (startDate && startDate.indexOf(targetPrefix) === 0) {
      const shopName = rec.契約店舗名称 ? rec.契約店舗名称.value : "不明"; const code = rec.案件番号 ? rec.案件番号.value : "";
      let bizType = "不明"; let source = "不明"; const matchedLead = rawLeads.find(l => l.案件番号 && l.案件番号.value == code);
      if (matchedLead) { bizType = matchedLead.業態詳細 ? matchedLead.業態詳細.value : "不明"; source = matchedLead.情報源小分類 ? matchedLead.情報源小分類.value : "不明"; }
      textRows.push("- ルート: " + source + " / 業態: " + bizType);
    }
  });
  return textRows.length > 0 ? textRows.join("\n") : "当月の新規契約なし";
}

function extractLeadBusinessTypesFromAPI(rawLeads, year, month) {
  let textRows = []; const targetPrefix = year + "-" + ("0" + month).slice(-2);
  rawLeads.forEach(rec => {
    const askDate = (rec['問合せ日'] && rec['問合せ日'].value) ? String(rec['問合せ日'].value) : "";
    if (askDate && askDate.indexOf(targetPrefix) === 0) {
      const source = rec.情報源小分類 ? String(rec.情報源小分類.value) : "その他";
      const bizType = rec.業態詳細 ? String(rec.業態詳細.value) : "不明";
      textRows.push("- ルート: " + source + " / 業態: " + bizType);
    }
  });
  return textRows.length > 0 ? textRows.join("\n") : "当月の問合せ案件なし";
}

function extractReferrersFromAPI(rawLeads, year, month) {
  let textRows = []; const targetPrefix = year + "-" + ("0" + month).slice(-2);
  rawLeads.forEach(rec => {
    const reqDate = (rec.問合せ日 && rec.問合せ日.value) ? String(rec.問合せ日.value) : "";
    if (reqDate && reqDate.indexOf(targetPrefix) === 0) {
      const source = rec.情報源小分類 ? rec.情報源小分類.value : "";
      const refCompany = rec.情報源会社名 ? rec.情報源会社名.value : "";
      const refName = rec.情報源紹介者名 ? rec.情報源紹介者名.value : "";
      if (refName || refCompany) textRows.push("- ルート: " + source + " / 会社: " + refCompany + " / 紹介者: " + refName);
    }
  });
  return textRows.length > 0 ? textRows.join("\n") : "当月の紹介案件なし";
}

function buildGeminiPrompt(sheetName, rawBusinessTypes, rawReferrers, rawLeadBusinessTypes) {
  return [
    "【月報自動生成指示書】", "[START_OF_MONTHLY_REPORT]",
    "🚨🚨🚨【最重要命題：上段と下段のデータを絶対に混同しないでください】🚨🚨🚨",
    "・生データ①（新規契約店舗）は、[上段]の契約数セクション（G17〜G26）の集計にのみ使用してください。下段のデータは1ミリも混ぜないでください。",
    "・生データ②（問合せ全体）は、[下段]の案件数セクション（G36〜G45）の集計にのみ使用してください。上段のデータは1ミリも混ぜないでください。",
    "",
    "■ 生データ①：新規契約店舗の業態・ルート（上段のG17〜G26の集計にのみ使用！）", rawBusinessTypes,
    "■ 生データ②：問合せ全体の業態・ルート（下段のG36〜G45の集計にのみ使用！）", rawLeadBusinessTypes,
    "■ 生データ③：紹介者の会社名・紹介者名・ルート（G10〜G16、G29〜G35の集計に使用）", rawReferrers,
    "",
    "■ 仕分けの絶対ルール",
    "📊 【ルールA：紹介者名をまとめるセル】（上段10〜16行目、下段29〜35行目）",
    "対象タグ：G10〜G16、G29〜G35（ゴミ業者、ビル管理会社、不動産、FIJ、おしぼり、てんぽす、増店）",
    "・生データ③（紹介者）から, 該当するルートのデータをすべて抽出してください。",
    "・紹介者の表記は, 会社名と紹介者名を合体させて『会社名(紹介者名)』、会社名がなければ『紹介者名』としてください。",
    "・同じ紹介者が複数いる場合は, 必ず『山名さん×2』『FIJ(山崎)×3』のように件数を集計して, カンマ（,）で区切って一列に並べてください。",
    "",
    "📊 【ルールB：お店の業態をまとめるセル】（上段17〜26行目、下段36〜45行目）",
    "対象タグ：G17〜G26、G36〜G45（飛び込み、テレアポ、DM、LP各種、HP各種、その他）",
    "・💡上段（G17〜G26）は『生データ①（契約）』のみを対象に集計してください。",
    "・💡下段（G36〜G45）は『生データ②（問合せ）』のみを対象に集計してください。",
    "・同じ業態が複数ある場合は, 必ず『バー×2』『ラーメン×3』という風に必ず件数を集計して, カンマ（,）で区切って一列に並べてください。重複して同じ名前を単に並べないでください（例：バー, バー は絶対NG）。",
    "",
    "3. 該当するデータが1件もないマスは, 必ず『(なし)』と出力してください。",
    "",
    "■ 出力フォーマット",
    "【契約数セクションの備考（上段）】",
    "[CELL_G10_CONTRACT_GOMI_REF]\n（ゴミ業者の紹介者まとめ）",
    "[CELL_G11_CONTRACT_BUILD_REF]\n（ビル管理の紹介者まとめ）",
    "[CELL_G12_CONTRACT_FUDOSAN_REF]\n（不動産の紹介者まとめ）",
    "[CELL_G13_CONTRACT_FIJ_REF]\n（FIJの紹介者まとめ）",
    "[CELL_G14_CONTRACT_OSHIBORI_REF]\n（おしぼりの紹介者まとめ）",
    "[CELL_G15_CONTRACT_TENPOS_REF]\n（てんぽすの紹介者まとめ）",
    "[CELL_G16_CONTRACT_ZOTEN_REF]\n（増店の紹介者まとめ）",
    "[CELL_G17_CONTRACT_TOBI_BIZ]\n（飛び込みの業態まとめ）",
    "[CELL_G18_CONTRACT_TELE_BIZ]\n（テレアポの業態まとめ）",
    "[CELL_G19_CONTRACT_DM_BIZ]\n（DMの業態まとめ）",
    "[CELL_G20_CONTRACT_LP_TEL_BIZ]\n（LP(電話)の業態まとめ）",
    "[CELL_G21_CONTRACT_LP_MAIL_BIZ]\n（LP(メール)の業態まとめ）",
    "[CELL_G22_CONTRACT_LP_LINE_BIZ]\n（LP(ライン)の業態まとめ）",
    "[CELL_G23_CONTRACT_HP_TEL_BIZ]\n（HP(電話)の業態まとめ）",
    "[CELL_G24_CONTRACT_HP_LINE_BIZ]\n（HP(ライン)の業態まとめ）",
    "[CELL_G25_HP_MAIL_BIZ]\n（HP(メール)の業態まとめ）",
    "[CELL_G26_CONTRACT_OTHER_BIZ]\n（その他の業態まとめ）",
    "",
    "【案件数セクションの備考（下段）】",
    "[CELL_G29_LEAD_GOMI_REF]\n（ゴミ業者の紹介者まとめ）",
    "[CELL_G30_LEAD_BUILD_REF]\n（ビル管理の紹介者まとめ）",
    "[CELL_G31_LEAD_FUDOSAN_REF]\n（不動産の紹介者まとめ）",
    "[CELL_G32_LEAD_FIJ_REF]\n（FIJの紹介者まとめ）",
    "[CELL_G33_LEAD_OSHIBORI_REF]\n（おしぼりの紹介者まとめ）",
    "[CELL_G34_LEAD_TENPOS_REF]\n（てんぽすの紹介者まとめ）",
    "[CELL_G35_LEAD_ZOTEN_REF]\n（増店の紹介者まとめ）",
    "[CELL_G36_LEAD_TOBI_BIZ]\n（飛び込みの業態まとめ）",
    "[CELL_G37_LEAD_TELE_BIZ]\n（テエアポの業態まとめ）",
    "[CELL_G38_LEAD_DM_BIZ]\n（DMの業態まとめ）",
    "[CELL_G39_LEAD_LP_TEL_BIZ]\n（LP(電話)の業態まとめ）",
    "[CELL_G40_LEAD_LP_MAIL_BIZ]\n（LP(メール)の業態まとめ）",
    "[CELL_G41_LEAD_LP_LINE_BIZ]\n（LP(ライン)の業態まとめ）",
    "[CELL_G42_LEAD_HP_TEL_BIZ]\n（HP(電話)の業態まとめ）",
    "[CELL_G43_LEAD_HP_LINE_BIZ]\n（HP(ライン)の業態まとめ）",
    "[CELL_G44_LEAD_HP_MAIL_BIZ]\n（HP(メール)の業態まとめ）",
    "[CELL_G45_LEAD_OTHER_BIZ]\n（その他の業態まとめ）",
    "[END_OF_MONTHLY_REPORT]"
  ].join("\n");
}

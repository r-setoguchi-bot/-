// ===== kintone「2026-2027　年末年始収集　業者案内と顧客送付ﾘｽﾄ」アプリのデータを、シートに書き出す =====
// 読み取りだけで、kintone は変更しない。お客様にメールも送らない。
//
// 必要なスクリプトプロパティ（コード.js と同じ KINTONE_SUBDOMAIN も使う）:
//   KINTONE_GYOSHA_ANNAI_APP_ID    … このアプリのID（kintoneでアプリを開いたURLの /k/ の次の数字）
//   KINTONE_GYOSHA_ANNAI_API_TOKEN … このアプリのAPIトークン（権限は「レコード閲覧」。項目名も見るなら「アプリ管理」も）
//
// 使い方: dumpGyoshaAnnai を実行 → シート「kintone_業者案内と顧客送付リスト」に、全レコード・全項目が書き出される

const GYOSHA_ANNAI_SHEET_NAME = "kintone_業者案内と顧客送付リスト";

function dumpGyoshaAnnai() {
  const props = PropertiesService.getScriptProperties();
  const subdomain = props.getProperty("KINTONE_SUBDOMAIN");
  const appId = props.getProperty("KINTONE_GYOSHA_ANNAI_APP_ID");
  const apiToken = props.getProperty("KINTONE_GYOSHA_ANNAI_API_TOKEN");
  if (!subdomain || !appId || !apiToken) {
    throw new Error("スクリプトプロパティが未設定です: KINTONE_SUBDOMAIN / KINTONE_GYOSHA_ANNAI_APP_ID / KINTONE_GYOSHA_ANNAI_API_TOKEN");
  }

  const records = fetchKintoneRecords(subdomain, appId, apiToken);
  if (records.length === 0) throw new Error("レコードが1件もありません。");

  // フィールドコード → 項目名（取れなければ、フィールドコードだけを見出しにする）
  const labels = {};
  try {
    const response = UrlFetchApp.fetch(`https://${subdomain}.cybozu.com/k/v1/app/form/fields.json?app=${appId}`, {
      method: "get", headers: { "X-Cybozu-API-Token": apiToken }, muteHttpExceptions: true
    });
    if (response.getResponseCode() === 200) {
      const properties = JSON.parse(response.getContentText()).properties || {};
      Object.keys(properties).forEach(code => { labels[code] = properties[code].label; });
    } else {
      console.warn(`項目名を取れませんでした（HTTP ${response.getResponseCode()}）。フィールドコードを見出しにします。`);
    }
  } catch (e) {
    console.warn(`項目名を取れませんでした: ${e}`);
  }

  const codes = [];
  records.forEach(record => Object.keys(record).forEach(code => { if (codes.indexOf(code) < 0) codes.push(code); }));
  const header = codes.map(code => labels[code] ? `${labels[code]}（${code}）` : code);
  const rows = records.map(record => codes.map(code => {
    const field = record[code];
    if (!field || field.value === undefined || field.value === null) return "";
    return typeof field.value === "object" ? JSON.stringify(field.value) : String(field.value);
  }));

  const ss = SpreadsheetApp.getActiveSpreadsheet() || invGetSpreadsheet();
  let sheet = ss.getSheetByName(GYOSHA_ANNAI_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(GYOSHA_ANNAI_SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  console.log(`レコード ${records.length} 件、項目 ${codes.length} 個を、シート「${GYOSHA_ANNAI_SHEET_NAME}」に書き出しました（kintone は変更していません）。`);
}

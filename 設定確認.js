// ===== 設定の確認（テスト用。確認が済んだら、このファイルごと削除してよい） =====
// スクリプトプロパティと接続を確認する。値の中身（APIトークンなど）はログに出さず、○（OK）か×（要確認）だけを出す。
// このファイルだけで動く。ほかのファイルには影響しない（消しても、ほかのファイルは動く）。
//
// 使い方: checkSettings を実行して、実行ログの「○」「×」を見る。「×」のものを直す。

function checkSettings() {
  const props = PropertiesService.getScriptProperties();
  const log = (ok, name, detail) => console.log(`${ok ? "○" : "×"} ${name}: ${detail}`);
  const get = name => String(props.getProperty(name) || "");
  const spaceNote = value => (value !== value.trim() ? "（前後にスペースが入っています）" : "");
  const normalize = s => String(s).normalize("NFKC").replace(/\s/g, "");

  // kintone の3項目
  const sub = get("KINTONE_SUBDOMAIN");
  const app = get("KINTONE_KEIYAKU_APP_ID");
  const token = get("KINTONE_KEIYAKU_API_TOKEN");
  log(!!sub && sub === sub.trim() && !/[\/.:]/.test(sub), "KINTONE_SUBDOMAIN",
    !sub ? "未設定" : /[\/.:]/.test(sub) ? "https:// や .cybozu.com を含めず、先頭の名前だけにしてください" : `設定あり（${sub}）${spaceNote(sub)}`);
  log(/^\d+$/.test(app), "KINTONE_KEIYAKU_APP_ID", !app ? "未設定" : /^\d+$/.test(app) ? `設定あり（${app}）` : "数字だけにしてください" + spaceNote(app));
  log(!!token && token === token.trim(), "KINTONE_KEIYAKU_API_TOKEN", !token ? "未設定" : `設定あり（${token.length}文字）${spaceNote(token)}`);

  if (sub && app && token) {
    try {
      const response = UrlFetchApp.fetch(`https://${sub.trim()}.cybozu.com/k/v1/app/form/fields.json?app=${app.trim()}`,
        { method: "get", headers: { "X-Cybozu-API-Token": token.trim() }, muteHttpExceptions: true });
      if (response.getResponseCode() !== 200) throw new Error(`HTTP ${response.getResponseCode()}`);
      const fields = JSON.parse(response.getContentText()).properties || {};
      log(true, "kintoneへの接続（項目の取得）", `成功（項目 ${Object.keys(fields).length} 件）`);
      const labels = Object.keys(fields).map(code => normalize(fields[code].label));
      const codeOverrides = { "送付先別：e-mail": "送付先別mail1" };   // リスト作成.js の KLIST_EMAIL_CODES と同じ
      [
        "契約種別", "契約店舗名称", "顧客名称", "収集業者：名称", "可燃収集日",
        "年末年始　回答アドレス①", "年末年始　回答アドレス②", "案内等送付先メールアドレス:(to)",
        "送付先別：e-mail", "契約時の顧客担当：メールアドレス"
      ].forEach(label => {
        const byCode = codeOverrides[label] && fields[codeOverrides[label]];
        const found = labels.indexOf(normalize(label)) !== -1 || !!byCode;
        log(found, `項目「${label}」`, found ? (byCode ? `あり（コード ${codeOverrides[label]}）` : "あり") : "見つかりません");
      });
    } catch (e) {
      log(false, "kintoneへの接続", `失敗: ${e}（サブドメイン・アプリID・トークン・「アプリ管理」の権限を確認）`);
    }
  }

  // フォームの保存先フォルダ
  const folderId = get("FORM_FOLDER_ID");
  if (!folderId) {
    log(false, "FORM_FOLDER_ID", "未設定（フォームがマイドライブ直下にできます）");
  } else {
    try {
      log(true, "FORM_FOLDER_ID", `フォルダを開けました（${DriveApp.getFolderById(folderId.trim()).getName()}）${spaceNote(folderId)}`);
    } catch (e) {
      log(false, "FORM_FOLDER_ID", "フォルダを開けません。IDと共有設定を確認してください");
    }
  }

  // スプレッドシート
  try {
    const url = get("FORM_SPREADSHEET_URL");
    const ss = url ? SpreadsheetApp.openByUrl(url) : SpreadsheetApp.getActiveSpreadsheet();
    log(!!ss, "スプレッドシート", ss ? `「${ss.getName()}」を使います` : "見つかりません。FORM_SPREADSHEET_URL を設定してください");
  } catch (e) {
    log(false, "スプレッドシート", String(e));
  }

  // 送信元・テスト用アドレス
  const from = get("MAIL_FROM").trim();
  if (!from) {
    log(true, "MAIL_FROM", "未設定（実行アカウントのアドレスから送ります）");
  } else {
    const own = String(Session.getEffectiveUser().getEmail() || "").toLowerCase();
    const aliases = GmailApp.getAliases().map(a => String(a).toLowerCase());
    const ok = from.toLowerCase() === own || aliases.indexOf(from.toLowerCase()) !== -1;
    log(ok, "MAIL_FROM", ok ? `${from} から送信できます` : `${from} は、このアカウントから送信できるアドレスとして登録されていません`);
  }
  ["INVITE_TEST_TO", "RESP_MAIL_TEST_TO"].forEach(name => {
    const value = get(name).trim();
    log(!value || /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(value), name, !value ? "未設定" : `設定あり（${value}）`);
  });
  console.log(`参考: FORM_TEST_LIMIT=${get("FORM_TEST_LIMIT") || "未設定"} / RESP_MAIL_MODE=${get("RESP_MAIL_MODE") || "未設定（off）"} / INVITE_MODE=${get("INVITE_MODE") || "未設定（送信しない）"} / INV_DAILY_LIMIT=${get("INV_DAILY_LIMIT") || "未設定（制限なし）"}`);
}

// アドレスに関係しそうな項目（名前に「メール」「mail」「アドレス」「送付」を含むもの）を、テーブルの中の項目も含めて一覧にする。
// 項目名と種類だけを出し、個人情報（実際のアドレス）は出さない。
function findEmailFields() {
  const props = PropertiesService.getScriptProperties();
  const sub = String(props.getProperty("KINTONE_SUBDOMAIN") || "").trim();
  const app = String(props.getProperty("KINTONE_KEIYAKU_APP_ID") || "").trim();
  const token = String(props.getProperty("KINTONE_KEIYAKU_API_TOKEN") || "").trim();
  const response = UrlFetchApp.fetch(`https://${sub}.cybozu.com/k/v1/app/form/fields.json?app=${app}`,
    { method: "get", headers: { "X-Cybozu-API-Token": token }, muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) throw new Error(`HTTP ${response.getResponseCode()}`);
  const properties = JSON.parse(response.getContentText()).properties || {};

  const pattern = /メール|mail|アドレス|送付/i;
  let count = 0;
  const show = (code, field, tableLabel) => {
    if (!pattern.test(String(field.label)) && !(tableLabel && pattern.test(tableLabel))) return;
    console.log(`${tableLabel ? "【テーブル「" + tableLabel + "」の中】" : "【通常の項目】"} 項目名=「${field.label}」 コード=${code} 種類=${field.type}`);
    count++;
  };
  Object.keys(properties).forEach(code => {
    const field = properties[code];
    show(code, field, "");
    if (field.type === "SUBTABLE" && field.fields) {
      Object.keys(field.fields).forEach(innerCode => show(innerCode, field.fields[innerCode], field.label || code));
    }
  });
  console.log(`該当した項目: ${count} 件`);
}

// 店舗の住所に関係しそうな項目（名前に「住所」「所在地」「郵便」「都道府県」「市区」「番地」「ビル」を含むもの）を一覧にする。
// 項目名・コード・種類と、最初の契約中の1件の値の例を出す（住所の入っている項目を探すため。アドレスなどは出さない）
function findAddressFields() {
  const props = PropertiesService.getScriptProperties();
  const sub = String(props.getProperty("KINTONE_SUBDOMAIN") || "").trim();
  const app = String(props.getProperty("KINTONE_KEIYAKU_APP_ID") || "").trim();
  const token = String(props.getProperty("KINTONE_KEIYAKU_API_TOKEN") || "").trim();
  const headers = { "X-Cybozu-API-Token": token };
  const response = UrlFetchApp.fetch(`https://${sub}.cybozu.com/k/v1/app/form/fields.json?app=${app}`,
    { method: "get", headers: headers, muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) throw new Error(`HTTP ${response.getResponseCode()}`);
  const properties = JSON.parse(response.getContentText()).properties || {};

  const pattern = /住所|所在地|郵便|都道府県|市区|番地|ビル|地番|アドレス（住所）/;
  const found = [];
  Object.keys(properties).forEach(code => {
    const field = properties[code];
    if (pattern.test(String(field.label))) found.push({ code: code, label: field.label, type: field.type });
  });
  if (found.length === 0) {
    console.log("該当する項目が見つかりませんでした。");
    return;
  }

  // 値の例（契約中の1件目）
  const fieldsParam = found.map((f, i) => `fields%5B${i}%5D=${encodeURIComponent(f.code)}`).join("&");
  const query = encodeURIComponent('契約種別 = "契約中" limit 1');
  const sample = UrlFetchApp.fetch(`https://${sub}.cybozu.com/k/v1/records.json?app=${app}&query=${query}&${fieldsParam}`,
    { method: "get", headers: headers, muteHttpExceptions: true });
  const record = sample.getResponseCode() === 200 ? (JSON.parse(sample.getContentText()).records || [])[0] : null;

  found.forEach(f => {
    const value = record && record[f.code] ? record[f.code].value : "";
    console.log(`項目名=「${f.label}」 コード=${f.code} 種類=${f.type} 値の例=「${Array.isArray(value) ? value.join(",") : value}」`);
  });
  console.log(`該当した項目: ${found.length} 件`);
}

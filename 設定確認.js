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

  // メールの送信（Microsoft Graph）の設定。値は表示しない（シークレットは○か×だけ）
  ["MS_TENANT_ID", "MS_CLIENT_ID", "MS_CLIENT_SECRET", "MS_MAIL_SENDER"].forEach(name => {
    const value = get(name);
    log(!!value && value === value.trim(), name, !value ? "未設定" : `設定あり${spaceNote(value)}`);
  });
  if (get("MS_MAIL_SENDER") && get("MS_MAIL_SENDER").trim().toLowerCase() !== "km-customer@kurimen.co.jp") {
    log(false, "MS_MAIL_SENDER", "km-customer@kurimen.co.jp になっていません（管理者の設定は、このアドレスだけに限定されています）");
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

// ===== Microsoft Graph の動作確認（km-customer からだけ送れることの確認） =====
// 使う前に、スクリプトプロパティに次を入れる: MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET / MS_MAIL_SENDER（km-customer@kurimen.co.jp） / RESP_MAIL_TEST_TO（自分のアドレス）
// 管理者の設定が反映されるまで（30分〜2時間）は 403 になる。202 が出れば反映済み。
//   手順2: testGraphSend()         … km-customer から、RESP_MAIL_TEST_TO 宛てに1通送る。202 なら反映済み
//   手順3: testGraphRestriction()  … r-setoguchi から送ろうとする。403 になれば、km-customer だけに制限できている証拠（手順2が 202 になってから実行する）
// 顧客には送らない。シークレットやトークンの中身は、ログに出さない。

function testGraphSend() {
  const props = PropertiesService.getScriptProperties();
  const sender = String(props.getProperty("MS_MAIL_SENDER") || "").trim();
  const to = String(props.getProperty("RESP_MAIL_TEST_TO") || "").trim();
  if (!sender || !to) throw new Error("MS_MAIL_SENDER と RESP_MAIL_TEST_TO を、スクリプトプロパティに設定してください");
  const code = graphTestSendMail_(sender, to);
  console.log(`km-customer（${sender}）から送信 → HTTP ${code}  ${code === 202 ? "○ 反映済みです。続けて testGraphRestriction を実行してください" : code === 403 ? "× まだ反映待ちです（15〜30分おいて、もう一度）" : "× 想定外の結果です。上のログを確認してください"}`);
}

function testGraphRestriction() {
  const props = PropertiesService.getScriptProperties();
  const to = String(props.getProperty("RESP_MAIL_TEST_TO") || "").trim();
  if (!to) throw new Error("RESP_MAIL_TEST_TO を、スクリプトプロパティに設定してください");
  const other = "r-setoguchi@kurimen.co.jp";
  const code = graphTestSendMail_(other, to);
  console.log(`${other} から送信 → HTTP ${code}  ${code === 403 ? "○ 送れません。km-customer だけに制限できています" : code === 202 ? "× 送れてしまいました。制限が効いていません。管理者に連絡してください（Entra で同意して回避しないでください）" : "× 想定外の結果です。上のログを確認してください"}`);
}

// Graph で、from のメールボックスから to 宛てに、テストメールを1通送る。HTTP ステータスを返す
function graphTestSendMail_(from, to) {
  const props = PropertiesService.getScriptProperties();
  const need = name => {
    const value = String(props.getProperty(name) || "").trim();
    if (!value) throw new Error(`スクリプトプロパティ ${name} を設定してください`);
    return value;
  };
  const tokenResponse = UrlFetchApp.fetch(`https://login.microsoftonline.com/${encodeURIComponent(need("MS_TENANT_ID"))}/oauth2/v2.0/token`, {
    method: "post",
    payload: { grant_type: "client_credentials", client_id: need("MS_CLIENT_ID"), client_secret: need("MS_CLIENT_SECRET"), scope: "https://graph.microsoft.com/.default" },
    muteHttpExceptions: true
  });
  const tokenJson = JSON.parse(tokenResponse.getContentText() || "{}");
  if (tokenResponse.getResponseCode() !== 200 || !tokenJson.access_token) {
    console.error(`トークンを取れませんでした: HTTP ${tokenResponse.getResponseCode()} ${tokenJson.error || ""} ${String(tokenJson.error_description || "").slice(0, 200)}`);
    throw new Error("トークンを取れませんでした（テナントID・クライアントID・シークレットの値と、シークレットの有効期限を確認）");
  }
  console.log("トークンの取得: 成功");
  const response = UrlFetchApp.fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(from)}/sendMail`, {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + tokenJson.access_token },
    payload: JSON.stringify({
      message: {
        subject: "【テスト】Graph の動作確認（顧客には送っていません）",
        body: { contentType: "Text", content: `Graph の動作確認です。送信元: ${from}\n送信時刻: ${new Date()}` },
        toRecipients: [{ emailAddress: { address: to } }]
      },
      saveToSentItems: true
    }),
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  if (code !== 202) console.log(`応答: ${response.getContentText().slice(0, 300)}`);
  return code;
}

// ===== リスト検証のエラー行を調べる（見るだけ。何も書き換えない。確認が済んだら削除してよい） =====
// 検証結果の「参考」に出た行番号（リストシートの行）について、同じフォームに何店舗あるか、送信の状況などを出す。
// 宛先のアドレスは、一部を伏せて表示する。
function inspectErrorRows() {
  const rowNumbers = [43, 416, 634, 682, 870, 1129];   // 検証結果の「参考」の行
  const props = PropertiesService.getScriptProperties();
  const url = String(props.getProperty("FORM_SPREADSHEET_URL") || "").trim();
  const ss = url ? SpreadsheetApp.openByUrl(url) : SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName("リスト").getDataRange().getDisplayValues();
  const mask = email => String(email).replace(/^(.{2})[^@]*@/, "$1***@");

  rowNumbers.forEach(r => {
    const row = data[r - 1];
    if (!row) { console.log(`${r}行目: 見つかりません`); return; }
    const formId = row[9];
    const sameForm = data.map((d, i) => ({ d: d, line: i + 1 })).filter(x => x.line > 1 && x.d[9] === formId && formId);
    console.log(`${r}行目 「${row[0]}」 状況=${row[10] || "（空）"} 宛先=${mask(row[1])} フォームID末尾=${String(formId).slice(-6) || "（なし）"} ` +
      `同じフォームの店舗数=${sameForm.length}（${sameForm.map(x => x.line + "行目:" + x.d[0]).join(" / ")}）`);
  });
}

// ===== 「除外」にした行の確認（見るだけ。何も書き換えない） =====
// 「リスト」のK列が「除外」の行ごとに、同じフォーム（J列）を使う、除外していない行があるかを調べる。
// あると、その店舗を除外しても、同じフォームの他の店舗の分で案内が送られる（除外した店舗も、本文とフォームに載る）。
function checkExcludedRows() {
  const props = PropertiesService.getScriptProperties();
  const url = String(props.getProperty("FORM_SPREADSHEET_URL") || "").trim();
  const ss = url ? SpreadsheetApp.openByUrl(url) : SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName("リスト").getDataRange().getDisplayValues();

  const rowsByForm = {};
  for (let i = 1; i < data.length; i++) {
    const formId = data[i][9];
    if (!formId) continue;
    (rowsByForm[formId] = rowsByForm[formId] || []).push(i);
  }

  let excluded = 0;
  let shared = 0;
  const statusCount = {};
  for (let i = 1; i < data.length; i++) {
    const status = data[i][10] || "（空）";
    statusCount[status] = (statusCount[status] || 0) + 1;
    if (data[i][10] !== "除外") continue;
    excluded++;
    const formId = data[i][9];
    const others = (rowsByForm[formId] || []).filter(j => j !== i && data[j][10] !== "除外");
    if (others.length > 0) {
      shared++;
      console.log(`△ ${i + 1}行目「${data[i][0]}」: 同じフォームに、除外していない店舗が ${others.length} 件あります（${others.map(j => (j + 1) + "行目:" + data[j][0]).join(" / ")}）。案内は、それらの店舗の分で送られます`);
    } else {
      console.log(`○ ${i + 1}行目「${data[i][0]}」: このフォームの店舗は、すべて除外です（案内は送られません）`);
    }
  }
  console.log(`除外の行: ${excluded} 件（うち、同じフォームに除外していない店舗がある行: ${shared} 件）`);
  console.log(`状況の内訳: ${JSON.stringify(statusCount)}`);
}

// ===== フォームの「回答にログインが必要」の確認と解除 =====
// Google Workspace のアカウントで作ったフォームは、既定で「組織内のユーザーだけが回答できる」設定になっていることがある。
// その場合、組織の外のお客様は「アクセス権限がありません」となる。
//   checkFormLoginSample()  … 先頭の数件のフォームの設定を見るだけ（何も変えない）
//   startFixFormLogin()     … 「リスト」の作成済みフォームすべてについて、ログイン必須を解除する（途中で止まっても、自動で続ける）
const FIXLOGIN_HANDLER = "continueFixFormLogin";
const FIXLOGIN_CURSOR = "FIX_LOGIN_CURSOR";

function fixLoginFormIds_() {
  const props = PropertiesService.getScriptProperties();
  const url = String(props.getProperty("FORM_SPREADSHEET_URL") || "").trim();
  const ss = url ? SpreadsheetApp.openByUrl(url) : SpreadsheetApp.getActiveSpreadsheet();
  const data = ss.getSheetByName("リスト").getDataRange().getDisplayValues();
  const ids = [];
  const seen = {};
  for (let i = 1; i < data.length; i++) {
    const id = data[i][9];
    if (id && data[i][10] === "送信済" && !seen[id]) { seen[id] = true; ids.push(id); }
  }
  return ids;
}

function checkFormLoginSample() {
  const ids = fixLoginFormIds_();
  [0, 1, 2, ids.length - 1].forEach(n => {
    const form = FormApp.openById(ids[n]);
    console.log(`${n + 1}件目 末尾=${ids[n].slice(-6)} ログイン必須=${form.requiresLogin()} 回答受付=${form.isAcceptingResponses()}`);
  });
  console.log(`対象のフォーム: ${ids.length} 件`);
}

function startFixFormLogin() {
  PropertiesService.getScriptProperties().deleteProperty(FIXLOGIN_CURSOR);
  fixLoginDeleteTriggers_();
  continueFixFormLogin();
}

function continueFixFormLogin() {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) { console.log("別の処理が動いているためスキップしました。"); return; }
  const startTime = Date.now();
  try {
    fixLoginDeleteTriggers_();
    const safety = ScriptApp.newTrigger(FIXLOGIN_HANDLER).timeBased().after(7 * 60 * 1000).create();
    const ids = fixLoginFormIds_();
    const props = PropertiesService.getScriptProperties();
    let index = Number(props.getProperty(FIXLOGIN_CURSOR) || 0);
    if (index > ids.length) index = 0;
    let fixed = 0, already = 0, errors = 0;
    while (index < ids.length) {
      if (Date.now() - startTime > 4.5 * 60 * 1000) break;
      try {
        const form = FormApp.openById(ids[index]);
        if (form.requiresLogin()) {
          form.setRequireLogin(false);
          if (form.requiresLogin()) throw new Error("解除できませんでした（組織の設定で固定されている可能性があります）");
          fixed++;
        } else {
          already++;
        }
      } catch (e) {
        console.error(`ログイン必須を解除できませんでした: 末尾=${String(ids[index]).slice(-6)} / ${e}`);
        errors++;
      }
      index++;
    }
    if (index < ids.length) {
      props.setProperty(FIXLOGIN_CURSOR, String(index));
      ScriptApp.newTrigger(FIXLOGIN_HANDLER).timeBased().after(60 * 1000).create();
      try { ScriptApp.deleteTrigger(safety); } catch (e) { /* 影響なし */ }
      console.log(`${index} / ${ids.length} 件を確認（今回 解除 ${fixed} 件、もともと不要 ${already} 件、エラー ${errors} 件）。残りは自動で続行します。`);
    } else {
      fixLoginDeleteTriggers_();
      props.deleteProperty(FIXLOGIN_CURSOR);
      console.log(`全 ${ids.length} 件を確認しました（今回 解除 ${fixed} 件、もともと不要 ${already} 件、エラー ${errors} 件）。完了しました。`);
    }
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

function fixLoginDeleteTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === FIXLOGIN_HANDLER) ScriptApp.deleteTrigger(t); });
}

// ===== フォームの「回答者ビュー」（Driveの共有）の確認と変更 =====
// 先に、左の「サービス」に「Drive API（v3）」を追加しておくこと。
//   inspectFormPermissions()  … 先頭の数件のフォームの、回答者ビューの権限を見るだけ（何も変えない）
//   openFormAccessSample()    … 先頭の1件だけ、「リンクを知っている全員が回答できる」に変える（試し）
//   startOpenFormAccess()     … 作成済みのフォームすべてを、同じように変える（途中で止まっても、自動で続ける）
const OPENACCESS_HANDLER = "continueOpenFormAccess";
const OPENACCESS_CURSOR = "OPEN_ACCESS_CURSOR";

function openAccessRetry_(fn) {
  let lastError;
  for (let attempt = 0; attempt < 4; attempt++) {
    try { return fn(); } catch (e) { lastError = e; Utilities.sleep(2000 * (attempt + 1)); }
  }
  throw lastError;
}

function openAccessList_(formId) {
  const result = Drive.Permissions.list(formId, { fields: "permissions(id,type,role,view,domain)", supportsAllDrives: true });
  return result.permissions || [];
}

function openAccessDescribe_(perms) {
  return perms.map(p => `${p.type}/${p.role}/view=${p.view || "-"}${p.domain ? "/" + p.domain : ""}`).join(" , ");
}

function inspectFormPermissions() {
  const ids = fixLoginFormIds_();
  [0, 1, ids.length - 1].forEach(n => {
    try {
      const perms = openAccessRetry_(() => openAccessList_(ids[n]));
      console.log(`${n + 1}件目 末尾=${ids[n].slice(-6)} 権限: ` + openAccessDescribe_(perms));
    } catch (e) {
      console.error(`${n + 1}件目 末尾=${ids[n].slice(-6)} 権限を読めませんでした: ${e}`);
    }
  });
  console.log(`対象のフォーム: ${ids.length} 件`);
}

function openAccessOne_(formId) {
  const perms = openAccessList_(formId);
  if (perms.some(p => p.type === "anyone" && p.view === "published")) return "already";
  Drive.Permissions.create({ type: "anyone", role: "reader", view: "published" }, formId, { supportsAllDrives: true });
  return "changed";
}

function openFormAccessSample() {
  const ids = fixLoginFormIds_();
  const result = openAccessRetry_(() => openAccessOne_(ids[0]));
  const after = openAccessRetry_(() => openAccessList_(ids[0]));
  console.log(`1件目 末尾=${ids[0].slice(-6)} 結果=${result} 変更後の権限: ` + openAccessDescribe_(after));
}

function startOpenFormAccess() {
  PropertiesService.getScriptProperties().deleteProperty(OPENACCESS_CURSOR);
  openAccessDeleteTriggers_();
  continueOpenFormAccess();
}

function continueOpenFormAccess() {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(10 * 1000)) { console.log("別の処理が動いているためスキップしました。"); return; }
  const startTime = Date.now();
  try {
    openAccessDeleteTriggers_();
    const safety = ScriptApp.newTrigger(OPENACCESS_HANDLER).timeBased().after(7 * 60 * 1000).create();
    const ids = fixLoginFormIds_();
    const props = PropertiesService.getScriptProperties();
    let index = Number(props.getProperty(OPENACCESS_CURSOR) || 0);
    if (index > ids.length) index = 0;
    let changed = 0, already = 0, errors = 0;
    while (index < ids.length) {
      if (Date.now() - startTime > 4.5 * 60 * 1000) break;
      try {
        const result = openAccessRetry_(() => openAccessOne_(ids[index]));
        if (result === "changed") changed++; else already++;
      } catch (e) {
        console.error(`変更できませんでした: 末尾=${String(ids[index]).slice(-6)} / ${e}`);
        errors++;
      }
      index++;
    }
    if (index < ids.length) {
      props.setProperty(OPENACCESS_CURSOR, String(index));
      ScriptApp.newTrigger(OPENACCESS_HANDLER).timeBased().after(60 * 1000).create();
      try { ScriptApp.deleteTrigger(safety); } catch (e) { /* 影響なし */ }
      console.log(`${index} / ${ids.length} 件を確認（今回 変更 ${changed} 件、もともと全員可 ${already} 件、エラー ${errors} 件）。残りは自動で続行します。`);
    } else {
      openAccessDeleteTriggers_();
      props.deleteProperty(OPENACCESS_CURSOR);
      console.log(`全 ${ids.length} 件を確認しました（今回 変更 ${changed} 件、もともと全員可 ${already} 件、エラー ${errors} 件）。完了しました。`);
    }
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 影響なし */ }
  }
}

function openAccessDeleteTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === OPENACCESS_HANDLER) ScriptApp.deleteTrigger(t); });
}

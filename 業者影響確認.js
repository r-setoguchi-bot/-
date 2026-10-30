// ===== 業者の年末年始の日程（期間外の日）で影響を受ける店舗の一覧を、シートに書き出す =====
// 読み取りだけで、kintone は変更しない。お客様にメールも送らない。
//   ・光栄和   … 12/29(火)に、小金井市・国分寺市のみ特別回収（8,000円）。火曜が通常の回収日の店舗が、対象
//   ・野口興業 … 1/4(月)に、特別回収（+7,000円）。月曜が通常の回収日の店舗が、対象
// 使い方: listVendorImpact を実行 → シート「業者影響確認」に、2社の全店舗が、影響の有無つきで書き出される
// （kintoneリスト作成.js の関数を使う。契約管理アプリの項目は、そちらの設定と同じ）

const VENDOR_IMPACT_SHEET_NAME = "業者影響確認";
const VENDOR_KOEIWA = "光栄和";
const VENDOR_NOGUCHI = "野口興業";
const VENDOR_KOEIWA_CITIES = ["小金井市", "国分寺市"];

function listVendorImpact() {
  const ss = klistGetSpreadsheet();
  const config = klistGetKintoneConfig();
  const fieldCodes = klistResolveFieldCodes(config);
  const records = klistFetchRecords(config, fieldCodes);
  const dates = klistBuildDates();

  const header = ["影響", "業者", "店舗名", "契約者", "送信先アドレス", "可燃の回収曜日", "市区町村", "住所", "案内の状況", "レコード番号", "契約管理URL"];
  const rows = [];
  records.forEach(record => {
    if (klistValue(record, fieldCodes.contractType) !== KLIST_TARGET_CONTRACT_TYPE) return;
    const vendor = klistValue(record, fieldCodes.vendor);
    const isKoeiwa = vendor.normalize("NFKC").indexOf(VENDOR_KOEIWA) !== -1;
    const isNoguchi = vendor.normalize("NFKC").indexOf(VENDOR_NOGUCHI) !== -1;
    if (!isKoeiwa && !isNoguchi) return;

    const recordId = record.$id.value;
    const options = klistSelectedOptions(record, fieldCodes.burnable);
    const flags = klistWeekdayFlags(options);   // 日〜土
    const weekdays = KLIST_WEEKDAYS.split("").filter((_, i) => flags[i]).join("");
    const address = klistBuildAddress(record);
    const city = klistValue(record, KLIST_ADDRESS_CODES.city).normalize("NFKC").trim();
    const cityText = (city + address).normalize("NFKC");

    const emailSources = fieldCodes.emails.map(e => ({ label: e.label, raw: klistValue(record, e.code) }));
    const picked = klistPickEmail(emailSources);

    // 影響の判定
    let impact = "なし";
    if (isKoeiwa && flags[2] && VENDOR_KOEIWA_CITIES.some(c => cityText.indexOf(c) !== -1)) impact = "12/29が対象（光栄和・小金井市/国分寺市・火曜）";
    if (isNoguchi && flags[1]) impact = "1/4が対象（野口興業・月曜）";

    // 案内の状況
    let status;
    if (picked.reason) status = "要確認（アドレス）";
    else if (!weekdays) status = "要確認（回収曜日なし）";
    else if (!dates.some(date => date.collectable && flags[date.weekdayIndex])) status = "期間外（お知らせメール）";
    else status = "フォームの案内";

    rows.push([impact, isKoeiwa ? VENDOR_KOEIWA : VENDOR_NOGUCHI, klistValue(record, fieldCodes.storeName), klistValue(record, fieldCodes.customer),
      picked.email || "", weekdays, city, address, status, recordId,
      `https://${config.subdomain}.cybozu.com/k/${config.appId}/show#record=${recordId}`]);
  });

  rows.sort((a, b) => (a[0] === "なし") - (b[0] === "なし") || a[1].localeCompare(b[1], "ja") || a[3].localeCompare(b[3], "ja"));
  let sheet = ss.getSheetByName(VENDOR_IMPACT_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(VENDOR_IMPACT_SHEET_NAME);
  sheet.clear();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length > 0) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  const koeiwaAffected = rows.filter(r => r[0].indexOf("12/29") === 0).length;
  const noguchiAffected = rows.filter(r => r[0].indexOf("1/4") === 0).length;
  console.log(`光栄和・野口興業の店舗: ${rows.length} 件。影響あり: 光栄和(12/29) ${koeiwaAffected} 件、野口興業(1/4) ${noguchiAffected} 件。シート「${VENDOR_IMPACT_SHEET_NAME}」に書き出しました（kintone は変更していません）。`);
}

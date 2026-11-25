/**
 * Google Apps Script — วางใน Extensions > Apps Script ของ Google Sheet
 * ชีตที่ต้องมี 2 แท็บ:
 *   Products : sku | name | price | stock | image | desc
 *   Orders   : OrderID | Timestamp | Name | Phone | Address | Postal | Items | Total | Status
 */
const ADMIN_KEY = "choco1234"; // ต้องตรงกับ config.js

function sheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function rows_(name) {
  const v = sheet_(name).getDataRange().getValues();
  const head = v.shift();
  return v.filter(r => r[0] !== "").map(r => {
    const o = {};
    head.forEach((h, i) => (o[h] = r[i] instanceof Date ? r[i].toISOString() : r[i]));
    return o;
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = e.parameter, action = p.action || "products";
  if (action === "products") return json_(rows_("Products"));

  if (p.key !== ADMIN_KEY) return json_({ error: "unauthorized" });

  if (action === "orders") return json_(rows_("Orders"));

  if (action === "mark") {
    const sh = sheet_("Orders"), data = sh.getDataRange().getValues();
    const statusCol = data[0].indexOf("Status") + 1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(p.id)) sh.getRange(i + 1, statusCol).setValue("PACKED");
    }
    return json_({ ok: true });
  }
  return json_({ error: "unknown action" });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const orders = sheet_("Orders");
    const id = "ORD-" + Utilities.formatString("%04d", orders.getLastRow()); // แถวหัวตาราง = 1

    orders.appendRow([id, new Date(), d.name, "'" + d.phone, d.address, "'" + d.postal,
      JSON.stringify(d.items), d.total, "NEW"]);

    // ตัดสต็อก
    const ps = sheet_("Products"), pd = ps.getDataRange().getValues();
    const skuC = pd[0].indexOf("sku"), stockC = pd[0].indexOf("stock");
    d.items.forEach(it => {
      for (let i = 1; i < pd.length; i++) {
        if (pd[i][skuC] === it.sku) ps.getRange(i + 1, stockC + 1).setValue(Math.max(0, pd[i][stockC] - it.qty));
      }
    });
    return json_({ ok: true, id: id });
  } finally {
    lock.releaseLock();
  }
}

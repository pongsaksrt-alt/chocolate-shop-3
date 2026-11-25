/**
 * Cocoa Craft - Google Apps Script (Backend)
 * ใช้เป็น API เชื่อมระหว่างเว็บ (Vercel) กับ Google Sheet
 *
 * วิธีใช้: เปิด Google Sheet > ส่วนขยาย > Apps Script > วางโค้ดนี้
 *  1) แก้ ADMIN_KEY ด้านล่าง
 *  2) เลือกฟังก์ชัน setup แล้วกด Run (อนุญาตสิทธิ์) 1 ครั้ง
 *  3) Deploy > New deployment > Web app
 *       Execute as: Me   |  Who has access: Anyone
 *  4) คัดลอก URL (ลงท้าย /exec) ไปใส่ใน config.js (API_URL)
 */

const ADMIN_KEY = "choco-2026";       // <<< เปลี่ยนเป็นรหัสของคุณ (ใช้เข้าหน้า admin.html)
const SHEET_PRODUCTS = "Products";
const SHEET_ORDERS = "Orders";

const PRODUCT_HEADERS = ["SKU", "Name", "Price", "Image", "Description", "Stock", "Active"];
const ORDER_HEADERS = ["Timestamp", "OrderID", "Name", "Phone", "Address", "Province", "Zip",
                       "Items", "Subtotal", "Shipping", "Total", "Note", "Status"];

// ---------- สร้างชีตและข้อมูลตัวอย่าง (รันครั้งเดียว) ----------
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let p = ss.getSheetByName(SHEET_PRODUCTS) || ss.insertSheet(SHEET_PRODUCTS);
  p.clear();
  p.appendRow(PRODUCT_HEADERS);
  p.appendRow(["CHOC-DARK-70", "ดาร์กช็อกโกแลต 70%", 89, "dark-70.svg",
               "เข้มข้น หอมโกโก้แท้ รสขมนิดๆ ปิดท้ายหวานละมุน", 100, "TRUE"]);
  p.appendRow(["CHOC-MILK-45", "มิลค์ช็อกโกแลต 45%", 79, "milk-45.svg",
               "นุ่ม หวานมัน กลิ่นนมหอม ทานง่ายทุกวัย", 100, "TRUE"]);
  p.appendRow(["CHOC-MATCHA", "ไวท์ช็อกโกแลต มัทฉะ", 99, "matcha.svg",
               "ไวท์ช็อกโกแลตผสมชาเขียวมัทฉะญี่ปุ่น หอมเบาๆ", 100, "TRUE"]);
  p.getRange(1, 1, 1, PRODUCT_HEADERS.length).setFontWeight("bold").setBackground("#f4e6d0");
  p.setFrozenRows(1);

  let o = ss.getSheetByName(SHEET_ORDERS) || ss.insertSheet(SHEET_ORDERS);
  if (o.getLastRow() === 0) {
    o.appendRow(ORDER_HEADERS);
    o.getRange(1, 1, 1, ORDER_HEADERS.length).setFontWeight("bold").setBackground("#f4e6d0");
    o.setFrozenRows(1);
  }
  // เก็บเบอร์โทร/รหัสไปรษณีย์เป็นข้อความ เพื่อไม่ให้เลข 0 นำหน้าหาย
  o.getRange("D:D").setNumberFormat("@");
  o.getRange("G:G").setNumberFormat("@");

  const def = ss.getSheetByName("Sheet1") || ss.getSheetByName("ชีต1");
  if (def && ss.getSheets().length > 2) ss.deleteSheet(def);
}

// ---------- helper ----------
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function table_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  const v = sh.getDataRange().getValues();
  const head = v.shift();
  return v.map((r, i) => {
    const o = { _row: i + 2 };
    head.forEach((h, c) => (o[h] = r[c]));
    return o;
  });
}

// ---------- GET: ดึงสินค้า / ดึงออเดอร์ ----------
function doGet(e) {
  try {
    const action = (e.parameter.action || "products");

    if (action === "products") {
      const products = table_(SHEET_PRODUCTS)
        .filter((r) => String(r.Active).toUpperCase() !== "FALSE" && r.SKU)
        .map((r) => ({
          sku: String(r.SKU), name: r.Name, price: Number(r.Price),
          image: r.Image, desc: r.Description, stock: Number(r.Stock)
        }));
      return json_({ ok: true, products: products });
    }

    if (action === "orders") {
      if (e.parameter.key !== ADMIN_KEY) return json_({ ok: false, error: "รหัสแอดมินไม่ถูกต้อง" });
      const orders = table_(SHEET_ORDERS).filter((r) => r.OrderID).map((r) => {
        let items = [];
        try { items = JSON.parse(r.Items); } catch (err) {}
        return {
          timestamp: String(r.Timestamp), orderId: r.OrderID, name: r.Name,
          phone: String(r.Phone), address: r.Address, province: r.Province,
          zip: String(r.Zip), items: items, subtotal: Number(r.Subtotal),
          shipping: Number(r.Shipping), total: Number(r.Total),
          note: r.Note, status: r.Status
        };
      });
      return json_({ ok: true, orders: orders });
    }

    return json_({ ok: false, error: "unknown action" });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// ---------- POST: สร้างออเดอร์ / เปลี่ยนสถานะ ----------
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);                       // กันออเดอร์ชนกัน / สต็อกผิด
    const d = JSON.parse(e.postData.contents);
    if (d.action === "status") return updateStatus_(d);
    return createOrder_(d);
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function createOrder_(d) {
  // 1) ตรวจข้อมูลผู้รับ
  const clean = (s) => String(s || "").trim();
  const name = clean(d.name), phone = clean(d.phone), address = clean(d.address);
  const province = clean(d.province), zip = clean(d.zip), note = clean(d.note);
  if (name.length < 2) throw new Error("กรุณากรอกชื่อ");
  if (!/^0\d{9}$/.test(phone)) throw new Error("เบอร์โทรไม่ถูกต้อง");
  if (address.length < 10) throw new Error("ที่อยู่สั้นเกินไป");
  if (!/^\d{5}$/.test(zip)) throw new Error("รหัสไปรษณีย์ไม่ถูกต้อง");
  if (!d.items || !d.items.length) throw new Error("ไม่มีสินค้าในตะกร้า");

  // 2) ตรวจสินค้า/สต็อก และคำนวณราคาจากชีต (ไม่เชื่อราคาจากหน้าเว็บ)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ps = ss.getSheetByName(SHEET_PRODUCTS);
  const pv = ps.getDataRange().getValues();
  const items = [];
  let subtotal = 0;

  d.items.forEach((it) => {
    const qty = Math.floor(Number(it.qty));
    if (!(qty > 0 && qty <= 99)) throw new Error("จำนวนสินค้าไม่ถูกต้อง");
    let found = -1;
    for (let r = 1; r < pv.length; r++) if (String(pv[r][0]) === String(it.sku)) { found = r; break; }
    if (found < 0) throw new Error("ไม่พบสินค้า " + it.sku);
    const stock = Number(pv[found][5]);
    if (stock < qty) throw new Error("สินค้า " + pv[found][1] + " เหลือ " + stock + " ชิ้น");
    const price = Number(pv[found][2]);
    items.push({ sku: String(pv[found][0]), name: pv[found][1], price: price, qty: qty, _r: found });
    subtotal += price * qty;
  });

  // 3) ตัดสต็อก
  items.forEach((it) => {
    const newStock = Number(pv[it._r][5]) - it.qty;
    ps.getRange(it._r + 1, 6).setValue(newStock);
    delete it._r;
  });

  // 4) ค่าส่ง + เลขที่ออเดอร์
  const shipping = subtotal >= 300 ? 0 : 40;       // ให้ตรงกับ config.js
  const total = subtotal + shipping;
  const os = ss.getSheetByName(SHEET_ORDERS);
  const seq = Math.max(os.getLastRow(), 1);        // แถวหัวตาราง = 1
  const today = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyMMdd");
  const orderId = "CH" + today + "-" + ("000" + seq).slice(-4);
  const ts = Utilities.formatDate(new Date(), "Asia/Bangkok", "yyyy-MM-dd HH:mm:ss");

  // 5) บันทึกออเดอร์
  // คอลัมน์ Phone/Zip ถูกตั้งเป็นข้อความ (@) ใน setup() เลข 0 นำหน้าจึงไม่หาย
  os.appendRow([ts, orderId, name, phone, address, province, zip,
                JSON.stringify(items), subtotal, shipping, total, note, "Pending"]);

  return json_({ ok: true, orderId: orderId, total: total });
}

function updateStatus_(d) {
  if (d.key !== ADMIN_KEY) throw new Error("รหัสแอดมินไม่ถูกต้อง");
  const os = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ORDERS);
  const ids = os.getRange(1, 2, os.getLastRow(), 1).getValues();
  const ok = ["Pending", "Packed", "Shipped"];
  if (ok.indexOf(d.status) < 0) throw new Error("สถานะไม่ถูกต้อง");
  const targets = [].concat(d.orderIds || []);
  let n = 0;
  for (let r = 1; r < ids.length; r++) {
    if (targets.indexOf(String(ids[r][0])) >= 0) {
      os.getRange(r + 1, 13).setValue(d.status);
      n++;
    }
  }
  return json_({ ok: true, updated: n });
}

/**
 * Chocolate Shop backend (Google Apps Script)
 * ชีต Products : SKU | Name | Price | Image | Stock | Description
 * ชีต Orders   : Timestamp | OrderID | Name | Phone | Address | Province | Zip | Items | Subtotal | Shipping | Total | Note | Status
 */

const SHIPPING_FEE = 40;        // ค่าส่ง (บาท)
const FREE_SHIPPING_OVER = 300; // ซื้อครบเท่านี้ส่งฟรี

// รันครั้งเดียว: สร้างชีต Products (2 SKUs) และชีต Orders ถ้ายังไม่มี
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let p = ss.getSheetByName('Products') || ss.insertSheet('Products');
  p.clear();
  p.appendRow(['SKU', 'Name', 'Price', 'Image', 'Stock', 'Description']);
  p.appendRow(['CHOC-DARK-70', 'ดาร์กช็อกโกแลต 70%', 89, 'images/dark.jpg', 50, 'เข้มข้น หอมโกโก้แท้ ขมกำลังดี']);
  p.appendRow(['CHOC-MILK-35', 'มิลค์ช็อกโกแลต', 79, 'images/milk.jpg', 80, 'หวานนุ่ม ละมุนลิ้น']);
  p.getRange(1, 1, 1, 6).setFontWeight('bold');

  // ถ้าแท็บที่มีหัวตาราง OrderID อยู่แล้ว ให้ใช้ตัวนั้นเป็น Orders
  let o = ss.getSheetByName('Orders');
  if (!o) {
    const first = ss.getSheets()[0];
    if (first.getName() !== 'Products' && first.getRange('B1').getValue() === 'OrderID') {
      first.setName('Orders');
      o = first;
    } else {
      o = ss.insertSheet('Orders');
      o.appendRow(['Timestamp', 'OrderID', 'Name', 'Phone', 'Address', 'Province', 'Zip', 'Items', 'Subtotal', 'Shipping', 'Total', 'Note', 'Status']);
      o.getRange(1, 1, 1, 13).setFontWeight('bold');
    }
  }
}

// ส่งข้อมูลสินค้าให้หน้าเว็บ
function doGet(e) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Products');
  const rows = sh.getDataRange().getValues();
  const head = rows.shift();
  const products = rows
    .filter(r => r[0])
    .map(r => Object.fromEntries(head.map((h, i) => [h.toString().toLowerCase(), r[i]])));
  return json_({ ok: true, products: products, shippingFee: SHIPPING_FEE, freeOver: FREE_SHIPPING_OVER });
}

// รับออเดอร์ บันทึก 1 ออเดอร์ = 1 แถว
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orders = ss.getSheetByName('Orders');
    const products = ss.getSheetByName('Products');
    const pRows = products.getDataRange().getValues();

    const lines = [];
    let subtotal = 0;

    data.items.forEach(item => {
      const qty = Number(item.qty);
      if (!qty || qty < 1) return;
      for (let i = 1; i < pRows.length; i++) {
        if (pRows[i][0] === item.sku) {
          const stock = Number(pRows[i][4]);
          if (qty > stock) throw new Error('สินค้า ' + pRows[i][1] + ' เหลือไม่พอ');
          subtotal += Number(pRows[i][2]) * qty;
          lines.push(pRows[i][1] + ' (' + item.sku + ') x' + qty);
          products.getRange(i + 1, 5).setValue(stock - qty); // ตัดสต็อก
          break;
        }
      }
    });

    if (!lines.length) throw new Error('ไม่มีสินค้าในออเดอร์');

    const shipping = subtotal >= FREE_SHIPPING_OVER ? 0 : SHIPPING_FEE;
    const total = subtotal + shipping;
    const orderId = 'ORD' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMddHHmmss');

    orders.appendRow([
      new Date(), orderId, data.name, "'" + data.phone, data.address,
      data.province, "'" + data.zip, lines.join(', '),
      subtotal, shipping, total, data.note || '', 'รอแพ็ค'
    ]);

    return json_({ ok: true, orderId: orderId, subtotal: subtotal, shipping: shipping, total: total });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

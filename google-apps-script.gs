// Mekos Snackbar POS -> Google Sheets
// 1) CHANGE the secret key below to your own words/numbers (you'll type the same key in the POS settings).
const SECRET_KEY = 'CHANGE-THIS-KEY';

const ORDER_HEADERS = ['Order ID','Date','Time','Type','Payment','Customer','Items','Qty','Subtotal','Delivery fee','Total','Cash received','Change','Status','Paid'];
const ITEM_HEADERS  = ['Order ID','Date','Time','Item','Size','Flavor','Price','Qty','Amount','Status'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);
    if (d.key !== SECRET_KEY) return out({ ok: false, error: 'Wrong key' });
    if (d.ping) return out({ ok: true });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orders = getSheet(ss, 'Orders', ORDER_HEADERS);
    const items = getSheet(ss, 'Items', ITEM_HEADERS);
    // Older sheets: add the new 'Paid' header if it is missing.
    if (orders.getRange(1, 15).getValue() === '') orders.getRange(1, 15).setValue('Paid').setFontWeight('bold').setBackground('#fcbf00');
    const o = d.order;

    // If this order already exists (retry or void), just update its status.
    const found = orders.getRange(1, 1, orders.getLastRow(), 1).createTextFinder(o.id).matchEntireCell(true).findNext();
    if (found) {
      orders.getRange(found.getRow(), 14).setValue(o.status);
      orders.getRange(found.getRow(), 15).setValue(o.paid);
      const last = items.getLastRow();
      if (last > 1) {
        const ids = items.getRange(2, 1, last - 1, 1).getValues();
        ids.forEach((r, i) => { if (r[0] === o.id) items.getRange(i + 2, ITEM_HEADERS.length).setValue(o.status); });
      }
      return out({ ok: true, updated: true });
    }

    orders.appendRow([o.id, o.date, o.time, o.type, o.pay, o.customer, o.items, o.qty, o.subtotal, o.fee, o.total, o.cash, o.change, o.status, o.paid]);
    d.items.forEach(i => items.appendRow([i.id, i.date, i.time, i.item, i.size, i.flavor, i.price, i.qty, i.amount, i.status]));
    return out({ ok: true });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function getSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#fcbf00');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1000, 1).setNumberFormat('@'); // keep Order IDs as text
  }
  return sh;
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

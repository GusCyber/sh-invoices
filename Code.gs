/**
 * Super Headlights Restoration — invoice logger
 * 1. Create a Google Sheet while logged in as superheadlights@gmail.com
 * 2. Extensions → Apps Script → paste this file
 * 3. Deploy → New deployment → Web app
 *    Execute as: Me
 *    Who has access: Anyone
 * 4. Copy the Web app URL into the invoice app Settings
 */

var SHEET_NAME = "Invoices";

var HEADERS = [
  "Timestamp",
  "Date",
  "Invoice",
  "Customer",
  "Phone",
  "Vehicle",
  "Size",
  "Service address",
  "List price",
  "Discount",
  "Total CAD",
  "Payment",
  "Status",
  "Notes"
];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function appendInvoice_(d) {
  getSheet_().appendRow([
    new Date(),
    d.date || "",
    d.invoice || "",
    d.customer || "",
    d.phone || "",
    d.vehicle || "",
    d.size || "",
    d.address || "",
    Number(d.list || 0),
    Number(d.discount || 0),
    Number(d.total || 0),
    d.payment || "",
    d.status || "",
    d.notes || ""
  ]);
}

function doPost(e) {
  try {
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    var data = JSON.parse(raw);
    appendInvoice_(data);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    appendInvoice_((e && e.parameter) ? e.parameter : {});
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

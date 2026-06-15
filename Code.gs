const SHEET_NAME = 'Monthly Shop Profits Tracker';
const SHEET_HEADERS = [
  'Serial No.',
  'Date & Time',
  'Customer Name',
  'Customer Type',
  'Locality',
  'Type of Service',
  'Amount Paid by Customer',
  'Product Cost / Spending',
  'Profit',
  'Unique Identifier Note'
];

// Returns HTML UI or JSON customer history depending on query parameter.
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'customers') {
    return ContentService
      .createTextOutput(JSON.stringify({ customers: getRecentCustomers_() }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService
    .createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Accepts POS entry payload, computes profit, and appends row to the tracker sheet.
function doPost(e) {
  try {
    const payload = parsePayload_(e);
    validatePayload_(payload);

    const amountPaid = toNumber_(payload.amountPaid);
    const productCost = toNumber_(payload.productCost);
    const profit = amountPaid - productCost;

    const sheet = getOrCreateSheet_();
    const serialNo = getNextSerialNo_(sheet);
    const now = new Date();

    sheet.appendRow([
      serialNo,
      now,
      payload.customerName.trim(),
      payload.customerType,
      payload.locality,
      payload.service,
      amountPaid,
      productCost,
      profit,
      (payload.uniqueNote || '').trim()
    ]);

    return jsonResponse_({
      success: true,
      serialNo: serialNo,
      timestamp: now.toISOString(),
      profit: profit
    });
  } catch (error) {
    return jsonResponse_({
      success: false,
      message: error && error.message ? error.message : 'Unable to save entry.'
    });
  }
}

// Ensures the target sheet exists with the exact 10 required header columns.
function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  const firstRow = sheet.getRange(1, 1, 1, SHEET_HEADERS.length).getValues()[0];
  const missingHeaders = SHEET_HEADERS.some((header, idx) => firstRow[idx] !== header);
  if (missingHeaders) {
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setValues([SHEET_HEADERS]);
  }
  return sheet;
}

// Generates the next serial number from the last stored row.
function getNextSerialNo_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;
  const serial = Number(sheet.getRange(lastRow, 1).getValue());
  return Number.isFinite(serial) ? serial + 1 : lastRow;
}

// Builds unique customer options using latest visit for each name + note pair.
function getRecentCustomers_() {
  const sheet = getOrCreateSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, SHEET_HEADERS.length).getValues();
  const latestByKey = {};

  for (let i = values.length - 1; i >= 0; i -= 1) {
    const row = values[i];
    const customerName = String(row[2] || '').trim();
    if (!customerName) continue;
    const uniqueNote = String(row[9] || '').trim();
    const key = `${customerName}__${uniqueNote}`;
    if (latestByKey[key]) continue;

    const visitDate = row[1] instanceof Date ? row[1] : new Date(row[1]);
    latestByKey[key] = {
      customerName: customerName,
      uniqueNote: uniqueNote,
      locality: String(row[4] || ''),
      service: String(row[5] || ''),
      amountPaid: toNumber_(row[6]),
      lastVisitISO: Number.isNaN(visitDate.getTime()) ? '' : visitDate.toISOString()
    };
  }

  return Object.keys(latestByKey)
    .map((key) => latestByKey[key])
    .sort((a, b) => a.customerName.localeCompare(b.customerName));
}

// Supports JSON and form-encoded POST payloads.
function parsePayload_(e) {
  if (!e || !e.postData || !e.postData.contents) return {};
  const contentType = (e.postData.type || '').toLowerCase();
  if (contentType.indexOf('application/json') !== -1) {
    return JSON.parse(e.postData.contents);
  }
  return e.parameter || {};
}

function validatePayload_(payload) {
  if (!payload.customerName) throw new Error('Customer Name is required.');
  if (!payload.customerType) throw new Error('Customer Type is required.');
  if (!payload.locality) throw new Error('Locality is required.');
  if (!payload.service) throw new Error('Type of Service is required.');

  const amount = Number(payload.amountPaid);
  const cost = Number(payload.productCost);
  if (!Number.isFinite(amount) || !Number.isFinite(cost)) {
    throw new Error('Amount Paid and Product Cost must be numeric.');
  }
}

function toNumber_(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

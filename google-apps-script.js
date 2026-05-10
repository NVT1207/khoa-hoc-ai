// =========================================================
// GOOGLE APPS SCRIPT - Quang Thuong AI
// Nhan webhook tu SePay + cung cap API check thanh toan
// =========================================================

// CONFIG: Token SePay cua anh
const SEPAY_TOKEN = 'ZKB8XAVZ05HNUMO73SXFYWGQ3GILNBVQPCJIKFBUE49H2AADBEVKNAJ1PTSWHWLY';

// ===== WEBHOOK TU SEPAY =====
// SePay goi vao day moi khi co giao dich vao tai khoan
// URL webhook: <Deploy URL>?source=sepay
//
// ===== POLLING TU FRONTEND =====
// Frontend goi: <Deploy URL>?action=check&orderCode=QTAI...&amount=4999000
// Tra ve: { paid: true/false }
//
// ===== GHI DATA DANG KY =====
// Frontend POST: { name, phone, email, course, price, orderCode, timestamp, status }

function doGet(e) {
  const action = e.parameter.action;

  // 1. Frontend check trang thai thanh toan
  if (action === 'check') {
    const orderCode = e.parameter.orderCode;
    const amount = parseInt(e.parameter.amount || '0');
    const paid = checkPayment(orderCode, amount);
    return ContentService
      .createTextOutput(JSON.stringify({ paid: paid }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput('OK')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  // Kiem tra xem la webhook SePay hay form dang ky
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput('ERROR');
  }

  // ===== WEBHOOK TU SEPAY =====
  // SePay gui format: { gateway, transactionDate, transferType, transferAmount, content, referenceCode, ... }
  if (data.transferAmount !== undefined && data.content !== undefined) {
    return handleSePayWebhook(data);
  }

  // ===== FORM DANG KY / UPDATE STATUS =====
  return handleRegistration(data);
}

// ==========================================
// WEBHOOK SEPAY
// ==========================================
function handleSePayWebhook(data) {
  // Kiem tra Authorization tu SePay (co trong header)
  // Luu giao dich vao sheet "Transactions"

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Transactions');
  if (!sheet) {
    sheet = ss.insertSheet('Transactions');
    sheet.appendRow(['Time', 'Amount', 'Content', 'Reference', 'Type']);
  }

  sheet.appendRow([
    new Date(),
    data.transferAmount,
    data.content || '',
    data.referenceCode || '',
    data.transferType || 'in'
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// CHECK THANH TOAN (goi tu frontend)
// ==========================================
function checkPayment(orderCode, expectedAmount) {
  if (!orderCode) return false;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transactions');
  if (!sheet) return false;

  const data = sheet.getDataRange().getValues();
  // Chi check 100 giao dich gan nhat (toi uu)
  const startRow = Math.max(1, data.length - 100);

  for (let i = startRow; i < data.length; i++) {
    const row = data[i];
    const amount = row[1];
    const content = (row[2] || '').toString().toUpperCase();
    const type = (row[4] || '').toString().toLowerCase();

    // Chi tinh giao dich tien vao, dung so tien, dung orderCode
    if (type === 'in' &&
        amount >= expectedAmount &&
        content.includes(orderCode.toUpperCase())) {
      return true;
    }
  }

  return false;
}

// ==========================================
// LUU THONG TIN DANG KY
// ==========================================
function handleRegistration(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Registrations');
  if (!sheet) {
    sheet = ss.insertSheet('Registrations');
    sheet.appendRow(['Time', 'Name', 'Phone', 'Email', 'Package', 'Amount', 'OrderCode', 'Status']);
  }

  sheet.appendRow([
    data.timestamp || new Date(),
    data.name || '',
    data.phone || '',
    data.email || '',
    data.course || '',
    data.price || 0,
    data.orderCode || '',
    data.status || 'pending'
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

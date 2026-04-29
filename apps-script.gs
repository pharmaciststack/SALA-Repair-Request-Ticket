// =====================================================
// Google Apps Script — วางโค้ดทั้งหมดนี้ใน Apps Script
// แล้ว Deploy > New Deployment > Web App
// Execute as: Me | Who has access: Anyone
// =====================================================

const SHEET_NAME = 'Tickets';

function doGet(e) {
  // ?action=getAdmins — return dynamic admin email list
  if (e && e.parameter && e.parameter.action === 'getAdmins') {
    const sheet = getAdminSheet();
    const data  = sheet.getDataRange().getValues();
    const admins = data.slice(1).map(r => String(r[0]).trim()).filter(Boolean);
    return json({ ok: true, admins });
  }

  // default — return tickets
  const sheet = getSheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return json({ ok: true, data: [] });

  const headers = data[0];
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
  return json({ ok: true, data: rows });
}

function doPost(e) {
  const body  = JSON.parse(e.postData.contents);
  const sheet = getSheet();

  if (body.action === 'add') {
    const t = body.ticket;
    sheet.appendRow([
      t.id, t.name, t.email, t.branch, t.phone,
      t.equip, t.deviceId || '', t.category, t.desc, t.urgency,
      t.status, t.note, t.createdAt, t.updatedAt,
      JSON.stringify(t.images || [])
    ]);

  } else if (body.action === 'uploadImage') {
    const url = uploadImageToDrive(body.base64, body.filename, body.mimeType);
    return json({ ok: true, url });

  } else if (body.action === 'update') {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const col = (name) => headers.indexOf(name) + 1;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(body.id)) {
        const c = col(body.field);
        if (c > 0) sheet.getRange(i + 1, c).setValue(body.value);
        const cu = col('updatedAt');
        if (cu > 0) sheet.getRange(i + 1, cu).setValue(new Date().toISOString());
        break;
      }
    }

  } else if (body.action === 'delete') {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(body.id)) {
        sheet.deleteRow(i + 1);
        break;
      }
    }

  } else if (body.action === 'clear') {
    const last = sheet.getLastRow();
    if (last > 1) sheet.deleteRows(2, last - 1);

  // ── Admin Management ──────────────────────────────
  } else if (body.action === 'addAdmin') {
    const aSheet = getAdminSheet();
    const data   = aSheet.getDataRange().getValues();
    const existing = data.slice(1).map(r => String(r[0]).toLowerCase().trim());
    const email = String(body.email).toLowerCase().trim();
    if (email && !existing.includes(email)) {
      aSheet.appendRow([body.email.trim(), new Date().toISOString()]);
    }

  } else if (body.action === 'removeAdmin') {
    const aSheet = getAdminSheet();
    const data   = aSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === String(body.email).toLowerCase().trim()) {
        aSheet.deleteRow(i + 1);
        break;
      }
    }
  }

  return json({ ok: true });
}

// ── Image Upload ──────────────────────────────────────
function uploadImageToDrive(base64, filename, mimeType) {
  const folder = getOrCreateFolder('RepairTicketImages');
  const blob   = Utilities.newBlob(
    Utilities.base64Decode(base64.split(',')[1]),
    mimeType,
    filename
  );
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

// ── Sheet Setup ───────────────────────────────────────
function getSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'id','name','email','branch','phone',
      'equip','deviceId','category','desc','urgency',
      'status','note','createdAt','updatedAt','images'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getAdminSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName('AdminEmails');
  if (!sheet) {
    sheet = ss.insertSheet('AdminEmails');
    sheet.appendRow(['email', 'addedAt']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

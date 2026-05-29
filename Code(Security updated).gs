// =====================================================
// Google Apps Script — Security Updated Version
// วางโค้ดทั้งหมดนี้ใน Apps Script แทนโค้ดเดิม
// Deploy > Manage Deployments > Edit > New version > Deploy
// Execute as: Me | Who has access: Anyone
// =====================================================

const SHEET_NAME         = 'Tickets';
const COMPANY            = 'บริษัท ศาลาโอสถรีเทล จำกัด';
const SYSTEM_NAME        = 'ระบบแจ้งซ่อมบำรุง';
const SYSTEM_URL         = 'https://goodyearzph.github.io/salaosot-repair-system/';
const GOOGLE_CLIENT_ID   = '854838901494-cuhmkrl29oj80i12apt7no01k763r3o2.apps.googleusercontent.com';
const SUPER_ADMIN_EMAILS = ['pharmacist@salaosot.com', 'goodyearzph@gmail.com'];
const ALLOWED_UPDATE_FIELDS = ['status', 'note'];

// ── Auth helpers ──────────────────────────────────────
// Decodes a Google ID token locally (no external network call).
// Checks expiry and audience; returns verified email (lowercase) or throws.
function verifyToken(idToken) {
  if (!idToken) throw new Error('Missing token');
  const parts = idToken.split('.');
  if (parts.length !== 3) throw new Error('Invalid JWT');
  // Convert base64url → base64, add padding, decode bytes → string → JSON
  const b64 = parts[1].replace(/-/g, '+').replace(/_/g, '/');
  const padded = b64 + '==='.slice((b64.length + 3) % 4);
  const bytes = Utilities.base64Decode(padded);
  const jsonStr = bytes.map(function(b) { return String.fromCharCode(b); }).join('');
  const payload = JSON.parse(jsonStr);
  if (!payload.email) throw new Error('No email in token');
  if (payload.exp * 1000 < Date.now()) throw new Error('Token expired');
  if (payload.aud !== GOOGLE_CLIENT_ID) throw new Error('Wrong audience');
  return payload.email.toLowerCase();
}

function isAdminEmail(email) {
  const lower = String(email).toLowerCase();
  if (SUPER_ADMIN_EMAILS.map(e => e.toLowerCase()).includes(lower)) return true;
  const rows = getAdminSheet().getDataRange().getValues().slice(1);
  return rows.some(r => String(r[0]).toLowerCase().trim() === lower);
}

// ── doGet ─────────────────────────────────────────────
function doGet(e) {
  const action  = e && e.parameter && e.parameter.action;
  const idToken = e && e.parameter && e.parameter.idToken;

  // No-auth version check — open this URL in browser to confirm deployment
  if (action === 'ping') {
    return json({ ok: true, version: '3.0', time: new Date().toISOString() });
  }

  let callerEmail;
  try {
    callerEmail = verifyToken(idToken);
  } catch(err) {
    return json({ ok: false, error: 'Unauthorized', debug: String(err) });
  }

  if (action === 'getAdmins') {
    const admins = getAdminSheet().getDataRange().getValues()
      .slice(1).map(r => String(r[0]).trim()).filter(Boolean);
    return json({ ok: true, admins });
  }

  if (action === 'getTechs') {
    const techs = getTechSheet().getDataRange().getValues()
      .slice(1).map(r => String(r[0]).trim()).filter(Boolean);
    return json({ ok: true, techs });
  }

  if (action === 'getUser') {
    const rows = getUsersSheet().getDataRange().getValues().slice(1);
    const user = rows.find(r => String(r[0]).toLowerCase().trim() === callerEmail);
    if (user) return json({ ok: true, registered: true, branch: String(user[1] || ''), role: String(user[3] || 'user') });
    return json({ ok: true, registered: false });
  }

  if (action === 'getUsers') {
    if (!isAdminEmail(callerEmail)) return json({ ok: false, error: 'Forbidden' });
    const rows = getUsersSheet().getDataRange().getValues().slice(1);
    const users = rows.map(r => ({ email: r[0], branch: r[1], registeredAt: r[2], role: r[3] || 'user' }));
    return json({ ok: true, users });
  }

  // default — return tickets
  // Admins see all tickets; regular users see only their own
  const sheet = getSheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return json({ ok: true, data: [] });
  const headers = data[0];
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
  const filtered = isAdminEmail(callerEmail)
    ? rows
    : rows.filter(r => String(r.email || '').toLowerCase() === callerEmail);
  return json({ ok: true, data: filtered });
}

// ── doPost ────────────────────────────────────────────
function doPost(e) {
  const body  = JSON.parse(e.postData.contents);
  const sheet = getSheet();

  let callerEmail;
  try {
    callerEmail = verifyToken(body.idToken);
  } catch(err) {
    return json({ ok: false, error: 'Unauthorized' });
  }

  // Destructive / management actions are admin-only
  const adminOnly = ['update', 'delete', 'clear', 'addAdmin', 'removeAdmin', 'addTech', 'removeTech'];
  if (adminOnly.includes(body.action) && !isAdminEmail(callerEmail)) {
    return json({ ok: false, error: 'Forbidden' });
  }

  if (body.action === 'add') {
    const t = body.ticket;
    // Use the server-verified email — ignore whatever the client sent
    t.email = callerEmail;
    sheet.appendRow([
      t.id, t.name, t.email, t.branch, t.phone,
      t.equip, t.deviceId || '', t.category, t.desc, t.urgency,
      t.status, t.note, t.createdAt, t.updatedAt,
      JSON.stringify(t.images || [])
    ]);
    try { sendNewTicketNotification(t); } catch(err) {}

  } else if (body.action === 'uploadImage') {
    const url = uploadImageToDrive(body.base64, body.filename, body.mimeType);
    return json({ ok: true, url });

  } else if (body.action === 'update') {
    // Whitelist allowed fields — prevents arbitrary column overwrites
    if (!ALLOWED_UPDATE_FIELDS.includes(body.field)) {
      return json({ ok: false, error: 'Invalid field' });
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const col = (name) => headers.indexOf(name) + 1;
    const rows = sheet.getDataRange().getValues();
    let ticketRow = null;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(body.id)) {
        const c = col(body.field);
        if (c > 0) sheet.getRange(i + 1, c).setValue(body.value);
        const cu = col('updatedAt');
        if (cu > 0) sheet.getRange(i + 1, cu).setValue(new Date().toISOString());
        ticketRow = rows[i];
        break;
      }
    }
    // Send completion email when status → done
    if (body.field === 'status' && body.value === 'done' && ticketRow) {
      try {
        const h = headers;
        const tObj = {};
        h.forEach((k, i) => { tObj[k] = ticketRow[i]; });
        sendCompletionNotification(tObj);
      } catch(err) {}
    }

  } else if (body.action === 'delete') {
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(body.id)) { sheet.deleteRow(i + 1); break; }
    }

  } else if (body.action === 'clear') {
    const last = sheet.getLastRow();
    if (last > 1) sheet.deleteRows(2, last - 1);

  // ── User Registration & Role Management ──────────
  } else if (body.action === 'registerUser') {
    const uSheet = getUsersSheet();
    const existing = uSheet.getDataRange().getValues().slice(1);
    const alreadyExists = existing.some(r => String(r[0]).toLowerCase().trim() === callerEmail);
    if (!alreadyExists) {
      uSheet.appendRow([callerEmail, String(body.branch || '').trim(), new Date().toISOString(), 'user']);
    }

  } else if (body.action === 'setRole') {
    if (!isAdminEmail(callerEmail)) return json({ ok: false, error: 'Forbidden' });
    const targetEmail = String(body.email || '').toLowerCase().trim();
    if (SUPER_ADMIN_EMAILS.map(e => e.toLowerCase()).includes(targetEmail)) {
      return json({ ok: false, error: 'Cannot modify super admin' });
    }
    // Update role in Users sheet
    const uSheet = getUsersSheet();
    const uData = uSheet.getDataRange().getValues();
    for (let i = 1; i < uData.length; i++) {
      if (String(uData[i][0]).toLowerCase().trim() === targetEmail) {
        uSheet.getRange(i + 1, 4).setValue(body.role || 'user');
        break;
      }
    }
    // Sync with AdminEmails sheet
    const aSheet = getAdminSheet();
    const aData = aSheet.getDataRange().getValues();
    if (body.role === 'admin') {
      const exists = aData.slice(1).some(r => String(r[0]).toLowerCase().trim() === targetEmail);
      if (!exists) aSheet.appendRow([body.email.trim(), new Date().toISOString()]);
    } else {
      for (let i = 1; i < aData.length; i++) {
        if (String(aData[i][0]).toLowerCase().trim() === targetEmail) { aSheet.deleteRow(i + 1); break; }
      }
    }

  // ── Admin Management ─────────────────────────────
  } else if (body.action === 'addAdmin') {
    const aSheet = getAdminSheet();
    const existing = aSheet.getDataRange().getValues().slice(1).map(r => String(r[0]).toLowerCase().trim());
    const email = String(body.email).toLowerCase().trim();
    if (email && !existing.includes(email)) aSheet.appendRow([body.email.trim(), new Date().toISOString()]);

  } else if (body.action === 'removeAdmin') {
    const aSheet = getAdminSheet();
    const data = aSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === String(body.email).toLowerCase().trim()) {
        aSheet.deleteRow(i + 1); break;
      }
    }

  // ── Tech Email Management ─────────────────────────
  } else if (body.action === 'addTech') {
    const tSheet = getTechSheet();
    const existing = tSheet.getDataRange().getValues().slice(1).map(r => String(r[0]).toLowerCase().trim());
    const email = String(body.email).toLowerCase().trim();
    if (email && !existing.includes(email)) tSheet.appendRow([body.email.trim(), new Date().toISOString()]);

  } else if (body.action === 'removeTech') {
    const tSheet = getTechSheet();
    const data = tSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase().trim() === String(body.email).toLowerCase().trim()) {
        tSheet.deleteRow(i + 1); break;
      }
    }
  }

  return json({ ok: true });
}

// ── Email Notifications ───────────────────────────────
function getTechEmails() {
  return getTechSheet().getDataRange().getValues()
    .slice(1).map(r => String(r[0]).trim()).filter(Boolean);
}

function sendNewTicketNotification(t) {
  const urgencyLabel = { low: '🟢 ต่ำ (7-30 วัน)', medium: '🟡 ปานกลาง (2-7 วัน)', high: '🔴 ด่วนมาก (1-2 วัน)' }[t.urgency] || t.urgency;
  const categoryLabel = { electrical: '⚡ ไฟฟ้า', plumbing: '🚿 ประปา', ac: '❄️ แอร์', computer: '💻 IT', furniture: '🪑 เฟอร์นิเจอร์', other: '🔩 อื่นๆ' }[t.category] || t.category;
  const subject = `🔧 [แจ้งซ่อมใหม่] ${t.equip} — สาขา ${t.branch} (${urgencyLabel})`;
  const body = `<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
  <div style="background:#f59e0b;color:#fff;padding:16px 24px;border-radius:8px 8px 0 0">
    <h2 style="margin:0;font-size:18px">🔧 มีการแจ้งซ่อมใหม่</h2>
    <p style="margin:4px 0 0;font-size:12px;opacity:.85">${COMPANY} · ${SYSTEM_NAME}</p>
  </div>
  <div style="background:#fff;border:1px solid #e5e7eb;border-top:none;padding:24px;border-radius:0 0 8px 8px">
    <table style="width:100%;border-collapse:collapse;font-size:14px">
      <tr><td style="padding:6px 0;color:#6b7280;width:130px">เลขที่ Ticket</td><td style="padding:6px 0;font-weight:600">#${String(t.id).slice(-8)}</td></tr>
      <tr><td style="padding:6px 0;color:#6b7280">ผู้แจ้ง</td><td style="padding:6px 0;font-weight:600">${t.name} (${t.email})</td></tr>
      <tr><td style="padding:6px 0;color:#6b7280">สาขา</td><td style="padding:6px 0;font-weight:600">${t.branch}</td></tr>
      ${t.phone ? `<tr><td style="padding:6px 0;color:#6b7280">เบอร์ติดต่อ</td><td style="padding:6px 0">${t.phone}</td></tr>` : ''}
      <tr><td style="padding:6px 0;color:#6b7280">อุปกรณ์</td><td style="padding:6px 0;font-weight:600">${t.equip}</td></tr>
      ${t.deviceId ? `<tr><td style="padding:6px 0;color:#6b7280">รหัสอุปกรณ์</td><td style="padding:6px 0">${t.deviceId}</td></tr>` : ''}
      <tr><td style="padding:6px 0;color:#6b7280">หมวดหมู่</td><td style="padding:6px 0">${categoryLabel}</td></tr>
      <tr><td style="padding:6px 0;color:#6b7280">ความเร่งด่วน</td><td style="padding:6px 0"><strong style="color:${t.urgency==='high'?'#dc2626':t.urgency==='medium'?'#d97706':'#16a34a'}">${urgencyLabel}</strong></td></tr>
      <tr><td style="padding:6px 0;color:#6b7280">วันที่แจ้ง</td><td style="padding:6px 0">${new Date(t.createdAt).toLocaleString('th-TH')}</td></tr>
    </table>
    <div style="background:#f9fafb;border:1px solid #e5e7eb;border-radius:6px;padding:12px;margin:16px 0">
      <p style="margin:0 0 4px;font-size:12px;color:#6b7280;font-weight:600">รายละเอียดปัญหา</p>
      <p style="margin:0;font-size:14px;color:#374151">${t.desc}</p>
    </div>
    <a href="${SYSTEM_URL}" style="display:inline-block;background:#f59e0b;color:#fff;padding:10px 24px;border-radius:6px;text-decoration:none;font-weight:600;font-size:14px">เปิดระบบจัดการ →</a>
  </div>
  <p style="text-align:center;font-size:11px;color:#9ca3af;margin-top:12px">อีเมลนี้ส่งอัตโนมัติจากระบบแจ้งซ่อมบำรุง ${COMPANY}</p>
</div>`;

  const techEmails = getTechEmails();
  if (techEmails.length) {
    MailApp.sendEmail({ to: techEmails.join(','), subject, htmlBody: body });
  }

  if (t.email) {
    const reqSubject = `✅ รับแจ้งซ่อมแล้ว #${String(t.id).slice(-8)} — ${t.equip}`;
    const reqBody = `<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
  <div style="background:#10b981;color:#fff;padding:16px 24px;border-radius:8px 8px 0 0">
    <h2 style="margin:0;font-size:18px">✅ ได้รับการแจ้งซ่อมแล้ว</h2>
    <p style="margin:4px 0 0;font-size:12px;opacity:.85">${COMPANY} · ${SYSTEM_NAME}</p>
  </div>
  <div style="background:#fff;border:1px solid #e5e7eb;border-top:none;padding:24px;border-radius:0 0 8px 8px">
    <p style="font-size:14px;color:#374151">เรียน คุณ${t.name}</p>
    <p style="font-size:14px;color:#374151">ระบบได้รับการแจ้งซ่อมของคุณเรียบร้อยแล้ว ทีมช่างเทคนิคจะดำเนินการโดยเร็วที่สุด</p>
    <table style="width:100%;border-collapse:collapse;font-size:14px;margin:16px 0">
      <tr><td style="padding:6px 0;color:#6b7280;width:130px">เลขที่ Ticket</td><td style="padding:6px 0;font-weight:600">#${String(t.id).slice(-8)}</td></tr>
      <tr><td style="padding:6px 0;color:#6b7280">อุปกรณ์</td><td style="padding:6px 0;font-weight:600">${t.equip}</td></tr>
      <tr><td style="padding:6px 0;color:#6b7280">สาขา</td><td style="padding:6px 0">${t.branch}</td></tr>
      <tr><td style="padding:6px 0;color:#6b7280">วันที่แจ้ง</td><td style="padding:6px 0">${new Date(t.createdAt).toLocaleString('th-TH')}</td></tr>
    </table>
    <p style="font-size:13px;color:#6b7280">หากมีข้อสงสัยกรุณาติดต่อฝ่าย IT</p>
  </div>
  <p style="text-align:center;font-size:11px;color:#9ca3af;margin-top:12px">อีเมลนี้ส่งอัตโนมัติจากระบบแจ้งซ่อมบำรุง ${COMPANY}</p>
</div>`;
    MailApp.sendEmail({ to: t.email, subject: reqSubject, htmlBody: reqBody });
  }
}

function sendCompletionNotification(t) {
  const reqEmail = String(t.email || '').trim();
  const ticketId = '#' + String(t.id).slice(-8);
  const subject  = `🎉 ซ่อมเสร็จแล้ว ${ticketId} — ${t.equip} สาขา ${t.branch}`;
  const body = `<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">
  <div style="background:#3b82f6;color:#fff;padding:16px 24px;border-radius:8px 8px 0 0">
    <h2 style="margin:0;font-size:18px">🎉 ดำเนินการซ่อมเสร็จสิ้น</h2>
    <p style="margin:4px 0 0;font-size:12px;opacity:.85">${COMPANY} · ${SYSTEM_NAME}</p>
  </div>
  <div style="background:#fff;border:1px solid #e5e7eb;border-top:none;padding:24px;border-radius:0 0 8px 8px">
    <p style="font-size:14px;color:#374151">เรียน คุณ${t.name || ''}</p>
    <p style="font-size:14px;color:#374151">การแจ้งซ่อมของคุณได้รับการดำเนินการเสร็จสิ้นแล้ว</p>
    <table style="width:100%;border-collapse:collapse;font-size:14px;margin:16px 0">
      <tr><td style="padding:6px 0;color:#6b7280;width:130px">เลขที่ Ticket</td><td style="padding:6px 0;font-weight:600">${ticketId}</td></tr>
      <tr><td style="padding:6px 0;color:#6b7280">อุปกรณ์</td><td style="padding:6px 0;font-weight:600">${t.equip || ''}</td></tr>
      <tr><td style="padding:6px 0;color:#6b7280">สาขา</td><td style="padding:6px 0">${t.branch || ''}</td></tr>
      <tr><td style="padding:6px 0;color:#6b7280">วันที่แจ้ง</td><td style="padding:6px 0">${t.createdAt ? new Date(t.createdAt).toLocaleString('th-TH') : '-'}</td></tr>
      <tr><td style="padding:6px 0;color:#6b7280">วันที่เสร็จ</td><td style="padding:6px 0">${new Date().toLocaleString('th-TH')}</td></tr>
    </table>
    ${t.note ? `<div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:6px;padding:12px;margin:16px 0">
      <p style="margin:0 0 4px;font-size:12px;color:#1d4ed8;font-weight:600">💬 หมายเหตุจากช่าง</p>
      <p style="margin:0;font-size:14px;color:#1e3a8a">${t.note}</p>
    </div>` : ''}
    <p style="font-size:13px;color:#6b7280">ขอบคุณที่ใช้บริการระบบแจ้งซ่อมบำรุง หากพบปัญหาอีกครั้งสามารถแจ้งซ่อมได้ที่ลิงก์ด้านล่าง</p>
    <a href="${SYSTEM_URL}" style="display:inline-block;background:#3b82f6;color:#fff;padding:10px 24px;border-radius:6px;text-decoration:none;font-weight:600;font-size:14px">แจ้งซ่อมอีกครั้ง →</a>
  </div>
  <p style="text-align:center;font-size:11px;color:#9ca3af;margin-top:12px">อีเมลนี้ส่งอัตโนมัติจากระบบแจ้งซ่อมบำรุง ${COMPANY}</p>
</div>`;

  if (reqEmail) MailApp.sendEmail({ to: reqEmail, subject, htmlBody: body });

  const techEmails = getTechEmails();
  if (techEmails.length) {
    const techSubject = `✅ [ซ่อมเสร็จ] ${ticketId} ${t.equip || ''} — สาขา ${t.branch || ''}`;
    MailApp.sendEmail({ to: techEmails.join(','), subject: techSubject, htmlBody: body });
  }
}

// ── Image Upload ──────────────────────────────────────
function uploadImageToDrive(base64, filename, mimeType) {
  const folder = getOrCreateFolder('RepairTicketImages');
  const blob   = Utilities.newBlob(Utilities.base64Decode(base64.split(',')[1]), mimeType, filename);
  const file   = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w800';
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

// ── Sheet Setup ───────────────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id','name','email','branch','phone','equip','deviceId','category','desc','urgency','status','note','createdAt','updatedAt','images']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getAdminSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('AdminEmails');
  if (!sheet) { sheet = ss.insertSheet('AdminEmails'); sheet.appendRow(['email','addedAt']); sheet.setFrozenRows(1); }
  return sheet;
}

function getTechSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('TechEmails');
  if (!sheet) { sheet = ss.insertSheet('TechEmails'); sheet.appendRow(['email','addedAt']); sheet.setFrozenRows(1); }
  return sheet;
}

function getUsersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(['email', 'branch', 'registeredAt', 'role']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

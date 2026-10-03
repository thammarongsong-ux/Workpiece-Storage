/*************************************************************
 * Workpiece Storage - Backend (Google Apps Script)
 * เชื่อม Google Drive + Google Sheet + โหมด Admin
 *
 * Sheet ID  : 1IndHBNkd_PrciT7itRAVDjWuQzL3x0xpOrSZhUDIQ00
 * Drive ID  : 1NkxlMHyKXB7Y6c8RmC4B7Bxy1-jYKz8e
 *
 * วิธีใช้:
 * 1. วางไฟล์นี้ทับของเดิม > Save
 * 2. รันฟังก์ชัน setupSheet() 1 ครั้ง (สร้างชีต ผลงาน/วิชา/นักเรียน)
 * 3. Deploy > Manage deployments > Edit > Version: New version > Deploy
 *    URL ยังเป็นตัวเดิม ไม่ต้องแก้ Frontend
 *
 * ความปลอดภัย:
 * - ตั้งค่า Share ของ Sheet และ Drive โฟลเดอร์เป็น Restricted
 *   (เฉพาะเจ้าของ/ครู) เพราะ Apps Script ใช้สิทธิ์เจ้าของเข้าถึงได้อยู่แล้ว
 *   ไม่ต้องเปิด Anyone with link ไฟล์นักเรียนก็ยังอัปโหลด/แสดง thumbnail ได้
 *   (thumbnail ของไฟล์ที่อัปโหลดใหม่โค้ดจะเปิด Anyone ให้เฉพาะไฟล์นั้น)
 * - ลิงก์ตรงไป Sheet/Drive จะแสดงเฉพาะตอน Admin ล็อกอินแล้วเท่านั้น
 *************************************************************/

const CONFIG = {
  SHEET_ID: '1IndHBNkd_PrciT7itRAVDjWuQzL3x0xpOrSZhUDIQ00',
  ROOT_FOLDER_ID: '1NkxlMHyKXB7Y6c8RmC4B7Bxy1-jYKz8e',
  SHEET_NAME: 'ผลงาน',
  SHEET_SUBJECTS: 'วิชา',
  SHEET_STUDENTS: 'นักเรียน',
  ADMIN_USER: 'FirstStar',
  ADMIN_PASS: '574001',
  HEADERS: [
    'Timestamp','ชื่อ-สกุล','ชั้น/ห้อง','วิชา','ประเภทงาน','ชื่อชิ้นงาน',
    'คำอธิบาย','ชนิดไฟล์','ชื่อไฟล์','File URL','Drive File ID','สถานะตรวจ','ความเห็นครู'
  ],
  DEFAULT_SUBJECTS: ['ภาษาไทย','คณิตศาสตร์','วิทยาศาสตร์','ภาษาอังกฤษ','สังคมศึกษา','ศิลปะ','การงานอาชีพ','สุขศึกษา/พละ','คอมพิวเตอร์']
};

/** รันครั้งเดียวเพื่อสร้างหัวตารางทุกชีต */
function setupSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  // 1) ชีตผลงาน
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_NAME, CONFIG.HEADERS);
  sh.setFrozenRows(1);
  // 2) ชีตวิชา: คอลัมน์ B = ห้องที่กำหนด (คั่นด้วย comma, ว่าง = ทุกห้อง)
  const SUBJECT_HEADERS = ['ชื่อวิชา', 'ชั้น/ห้อง (comma คั่น, ว่าง=ทุกห้อง)'];
  const shS = getOrCreateSheet_(ss, CONFIG.SHEET_SUBJECTS, SUBJECT_HEADERS);
  if (!shS.getRange(1, 2).getValue()) shS.getRange(1, 2).setValue(SUBJECT_HEADERS[1]); // เติมหัวคอลัมน์ B ให้ชีตเก่า
  if (shS.getLastRow() < 2) {
    shS.getRange(2, 1, CONFIG.DEFAULT_SUBJECTS.length, 1).setValues(CONFIG.DEFAULT_SUBJECTS.map(s => [s]));
  }
  // 3) ชีตนักเรียน
  getOrCreateSheet_(ss, CONFIG.SHEET_STUDENTS, ['ชื่อ-สกุล', 'ชั้น/ห้อง']);
  Logger.log('setupSheet OK');
}

function getOrCreateSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  return getOrCreateSheet_(ss, CONFIG.SHEET_NAME, CONFIG.HEADERS);
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function isAdmin_(d) {
  return String(d && d.adminUser || '') === CONFIG.ADMIN_USER && String(d && d.adminPass || '') === CONFIG.ADMIN_PASS;
}

/** แปลง "ป.5/1, ป.5/2" -> ["ป.5/1","ป.5/2"] (รับ string หรือ array) */
function parseClasses_(v) {
  if (!v) return [];
  const arr = (v instanceof Array) ? v : String(v).split(',');
  const out = [];
  arr.forEach(x => { const t = String(x || '').trim(); if (t && out.indexOf(t) < 0) out.push(t); });
  return out;
}

// GET ?action=list | ?action=listConfig | ?action=ping
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'ping';
    if (action === 'list') return listWorks_();
    if (action === 'listConfig') return listConfig_();
    return jsonOut_({ ok: true, message: 'Student Workpiece Storage API ready', time: new Date() });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// POST {action:'submit'|'updateStatus'|'login'|'addSubject'|'editSubject'|'deleteSubject'|'addStudent'|'editStudent'|'deleteStudent'|'deleteWork', ...}
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'submit') return submitWork_(data);
    if (data.action === 'login') {
      return jsonOut_({ ok: isAdmin_(data), message: isAdmin_(data) ? 'login ok' : 'user/pass ไม่ถูกต้อง' });
    }
    // --- เมนู Admin ต้องแนบ adminUser/adminPass มาด้วย ---
    if (data.action === 'updateStatus') return updateStatus_(data);
    if (data.action === 'deleteWork') return deleteWork_(data);
    if (data.action === 'addSubject') return addSubject_(data);
    if (data.action === 'editSubject') return editSubject_(data);
    if (data.action === 'deleteSubject') return deleteSubject_(data);
    if (data.action === 'addStudent') return addStudent_(data);
    if (data.action === 'editStudent') return editStudent_(data);
    if (data.action === 'deleteStudent') return deleteStudent_(data);
    return jsonOut_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.stack || err) });
  }
}

/** ดึงรายชื่อวิชา + นักเรียน จาก Sheet (สาธารณะ อ่านได้อย่างเดียว) */
function listConfig_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const shS = getOrCreateSheet_(ss, CONFIG.SHEET_SUBJECTS, ['ชื่อวิชา', 'ชั้น/ห้อง (comma คั่น, ว่าง=ทุกห้อง)']);
  if (!shS.getRange(1, 2).getValue()) shS.getRange(1, 2).setValue('ชั้น/ห้อง (comma คั่น, ว่าง=ทุกห้อง)');
  const shT = getOrCreateSheet_(ss, CONFIG.SHEET_STUDENTS, ['ชื่อ-สกุล', 'ชั้น/ห้อง']);
  let subjectRows = [];
  if (shS.getLastRow() >= 2) {
    const sv = shS.getRange(2, 1, shS.getLastRow() - 1, 2).getValues();
    sv.forEach((r, i) => {
      const nm = String(r[0] || '').trim();
      if (nm) subjectRows.push({ row: i + 2, name: nm, classrooms: parseClasses_(r[1]) }); // ว่าง = ทุกห้อง
    });
  }
  let subjects = subjectRows.map(o => o.name);
  if (!subjects.length) {
    subjects = CONFIG.DEFAULT_SUBJECTS.slice();
    subjectRows = subjects.map((s, i) => ({ row: i + 2, name: s, classrooms: [] }));
  }
  let students = [];
  if (shT.getLastRow() >= 2) {
    students = shT.getRange(2, 1, shT.getLastRow() - 1, 2).getValues()
      .map((r, i) => ({ row: i + 2, fullName: String(r[0] || '').trim(), classroom: String(r[1] || '').trim() }))
      .filter(s => s.fullName);
  }
  return jsonOut_({ ok: true, subjects: subjects, subjectRows: subjectRows, students: students });
}

/** ดึงข้อมูลผลงานทั้งหมด */
function listWorks_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return jsonOut_({ ok: true, rows: [] });
  const vals = sh.getRange(2, 1, last - 1, CONFIG.HEADERS.length).getValues();
  const rows = vals
    .map((r, i) => ({
      row: i + 2,
      timestamp: r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Bangkok', 'yyyy-MM-dd HH:mm') : String(r[0] || ''),
      fullName: String(r[1] || ''),
      classroom: String(r[2] || ''),
      subject: String(r[3] || ''),
      category: String(r[4] || ''),
      title: String(r[5] || ''),
      description: String(r[6] || ''),
      fileKind: String(r[7] || 'image'),
      fileName: String(r[8] || ''),
      fileUrl: String(r[9] || ''),
      driveFileId: String(r[10] || ''),
      status: String(r[11] || 'รอตรวจ'),
      comment: String(r[12] || '')
    }))
    .filter(r => r.fullName || r.title);
  rows.reverse();
  return jsonOut_({ ok: true, rows: rows, total: rows.length });
}

/** นักเรียนส่งงาน (ไม่ต้องล็อกอิน) */
function submitWork_(d) {
  const fullName = String(d.fullName || '').trim();
  const subject  = String(d.subject || '').trim();
  const category = String(d.category || '').trim();
  const title    = String(d.title || '').trim();
  const classroom = String(d.classroom || '').trim();
  if (!fullName || !classroom || !subject || !category || !title) {
    return jsonOut_({ ok: false, error: 'กรุณาเลือก ชื่อ-สกุล / ห้อง / วิชา / ประเภทงาน / ชื่อชิ้นงาน ให้ครบ' });
  }
  // กันเลือกวิชาผิดห้อง: ถ้าวิชานี้กำกับห้องไว้ ห้องที่ส่งต้องอยู่ในลิสต์
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const shS = ss.getSheetByName(CONFIG.SHEET_SUBJECTS);
    if (shS && shS.getLastRow() >= 2) {
      const sv = shS.getRange(2, 1, shS.getLastRow() - 1, 2).getValues();
      // ชื่อวิชาอาจซ้ำกันต่างห้อง (เช่น "HB เทอม 1" มีทั้ง ม.2 และ ม.3) -> ผ่านถ้ามีแถวใดแถวหนึ่งอนุญาตห้องนี้
      let found = false, allowed = false, allowList = [];
      for (let i = 0; i < sv.length; i++) {
        if (String(sv[i][0] || '').trim() === subject) {
          found = true;
          const allow = parseClasses_(sv[i][1]);
          if (!allow.length || allow.indexOf(classroom) >= 0) { allowed = true; break; }
          allowList = allowList.concat(allow.filter(x => allowList.indexOf(x) < 0));
        }
      }
      if (found && !allowed) {
        return jsonOut_({ ok: false, error: 'วิชา "' + subject + '" ไม่ได้กำหนดให้ห้อง ' + classroom + ' (กำหนดไว้: ' + allowList.join(', ') + ')' });
      }
    }
  } catch (e) { /* ถ้าอ่าน mapping ไม่ได้ ให้ผ่านไปก่อน ไม่บล็อกการส่ง */ }
  const fileKind = (d.fileKind || 'image');
  let fileUrl = '';
  let driveFileId = '';
  let fileName = String(d.fileName || '');
  if (fileKind === 'link') {
    fileUrl = String(d.linkUrl || '').trim();
    if (!fileUrl) return jsonOut_({ ok: false, error: 'กรุณาแนบลิงก์ชิ้นงาน' });
    if (!/^https?:\/\//i.test(fileUrl)) fileUrl = 'https://' + fileUrl;
    fileName = fileUrl;
  } else {
    if (!d.fileBase64) return jsonOut_({ ok: false, error: 'ไม่พบไฟล์แนบ' });
    const bytes = Utilities.base64Decode(String(d.fileBase64).split(',').pop());
    const mime = String(d.mimeType || 'application/octet-stream');
    if (!fileName) fileName = title + '_' + Date.now();
    const blob = Utilities.newBlob(bytes, mime, fileName);
    const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
    const sub1 = getOrCreateFolder_(root, subject);
    const sub2 = getOrCreateFolder_(sub1, category);
    const safeName = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss') + '_' + fullName + '_' + fileName;
    const file = sub2.createFile(blob.setName(safeName));
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    driveFileId = file.getId();
    fileUrl = 'https://drive.google.com/file/d/' + driveFileId + '/view';
  }
  getSheet_().appendRow([new Date(), fullName, classroom, subject, category, title,
    String(d.description || ''), fileKind, fileName, fileUrl, driveFileId, 'รอตรวจ', '']);
  return jsonOut_({ ok: true, fileUrl: fileUrl, driveFileId: driveFileId });
}

/** ครูตรวจงาน (Admin เท่านั้น) */
function updateStatus_(d) {
  if (!isAdmin_(d)) return jsonOut_({ ok: false, error: 'ต้องล็อกอิน Admin ก่อน' });
  const row = Number(d.row);
  const status = String(d.status || '');
  const comment = String(d.comment || '');
  if (!row || row < 2) return jsonOut_({ ok: false, error: 'row ไม่ถูกต้อง' });
  if (['รอตรวจ', 'ผ่าน', 'ต้องแก้ไข'].indexOf(status) < 0) return jsonOut_({ ok: false, error: 'status ไม่ถูกต้อง' });
  const sh = getSheet_();
  sh.getRange(row, 12).setValue(status);
  sh.getRange(row, 13).setValue(comment);
  return jsonOut_({ ok: true, row: row, status: status });
}

/** ลบงาน (Admin เท่านั้น) */
function deleteWork_(d) {
  if (!isAdmin_(d)) return jsonOut_({ ok: false, error: 'ต้องล็อกอิน Admin ก่อน' });
  const row = Number(d.row);
  if (!row || row < 2) return jsonOut_({ ok: false, error: 'row ไม่ถูกต้อง' });
  const sh = getSheet_();
  // ลบซ้ำได้อย่างปลอดภัย: ถ้าแถวหายไปแล้วถือว่าสำเร็จ (กันเคสลบสำเร็จแต่เน็ตหลุดตอนตอบกลับ)
  if (row > sh.getLastRow()) return jsonOut_({ ok: true, alreadyDeleted: true, row: row });
  const fileId = String(sh.getRange(row, 11).getValue() || '').trim();
  if (fileId) {
    try { DriveApp.getFileById(fileId).setTrashed(true); }
    catch (e) { /* ไฟล์ใน Drive อาจถูกลบไปก่อนแล้ว — ลบแค่แถวใน Sheet ต่อ */ }
  }
  sh.deleteRow(row);
  return jsonOut_({ ok: true, row: row });
}

/** เพิ่มวิชา (Admin) */
function addSubject_(d) {
  if (!isAdmin_(d)) return jsonOut_({ ok: false, error: 'ต้องล็อกอิน Admin ก่อน' });
  const name = String(d.subject || '').trim();
  if (!name) return jsonOut_({ ok: false, error: 'กรุณาระบุชื่อวิชา' });
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_SUBJECTS, ['ชื่อวิชา', 'ชั้น/ห้อง (comma คั่น, ว่าง=ทุกห้อง)']);
  const rooms = parseClasses_(d.classrooms);
  // กันเพิ่มซ้ำแถวเดิมเป๊ะๆ (ชื่อเดียวกันแต่คนละห้อง อนุญาตได้ เช่น HB เทอม 1 ของ ม.2/ม.3)
  if (sh.getLastRow() >= 2) {
    const sv = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    const dupe = sv.some(r => String(r[0] || '').trim() === name && parseClasses_(r[1]).join('|') === rooms.join('|'));
    if (dupe) return jsonOut_({ ok: false, error: 'มีวิชานี้ + ห้องชุดนี้แล้ว' });
  }
  sh.appendRow([name, rooms.join(', ')]);
  return jsonOut_({ ok: true, subject: name, classrooms: rooms });
}

/** แก้ไขชื่อวิชา (Admin) + อัปเดตงานเดิมที่ใช้ชื่อเก่า */
function editSubject_(d) {
  if (!isAdmin_(d)) return jsonOut_({ ok: false, error: 'ต้องล็อกอิน Admin ก่อน' });
  const row = Number(d.row);
  const name = String(d.subject || '').trim();
  if (!row || row < 2) return jsonOut_({ ok: false, error: 'row ไม่ถูกต้อง' });
  if (!name) return jsonOut_({ ok: false, error: 'กรุณาระบุชื่อวิชาใหม่' });
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_SUBJECTS);
  if (!sh) return jsonOut_({ ok: false, error: 'ไม่พบชีตวิชา' });
  const oldName = String(sh.getRange(row, 1).getValue() || '').trim();
  // กันชื่อซ้ำกับวิชาอื่น (ถ้าไม่ได้เปลี่ยนชื่อ ข้ามเช็คนี้ — กันเคสชื่อซ้ำต่างห้องที่มีอยู่แล้วเช่น HB เทอม 1)
  if (name !== oldName && sh.getLastRow() >= 2) {
    const cur = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim());
    if (cur.indexOf(name) >= 0) return jsonOut_({ ok: false, error: 'มีวิชาชื่อนี้แล้ว' });
  }
  sh.getRange(row, 1).setValue(name);
  // อัปเดต mapping ห้อง (ถ้าส่งมา) — ส่ง classrooms มาเป็น string/array, ไม่ส่ง = คงเดิม
  let rooms = parseClasses_(sh.getRange(row, 2).getValue());
  if (d.hasOwnProperty('classrooms')) {
    rooms = parseClasses_(d.classrooms);
    sh.getRange(row, 2).setValue(rooms.join(', '));
  }
  // อัปเดตงานเดิมที่ใช้ชื่อวิชาเก่าให้เป็นชื่อใหม่ (คอลัมน์ D ของชีตผลงาน)
  let updated = 0;
  if (oldName && oldName !== name && String(d.updateWorks || '1') !== '0') {
    const ws = getSheet_();
    if (ws.getLastRow() >= 2) {
      const n = ws.getLastRow() - 1;
      const col = ws.getRange(2, 4, n, 1).getValues();
      for (let i = 0; i < n; i++) {
        if (String(col[i][0]).trim() === oldName) { ws.getRange(i + 2, 4).setValue(name); updated++; }
      }
    }
  }
  return jsonOut_({ ok: true, oldName: oldName, subject: name, classrooms: rooms, updatedWorks: updated });
}

function deleteSubject_(d) {
  if (!isAdmin_(d)) return jsonOut_({ ok: false, error: 'ต้องล็อกอิน Admin ก่อน' });
  const row = Number(d.row);
  if (!row || row < 2) return jsonOut_({ ok: false, error: 'row ไม่ถูกต้อง' });
  SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_SUBJECTS).deleteRow(row);
  return jsonOut_({ ok: true });
}

/** เพิ่มนักเรียน (Admin) */
function addStudent_(d) {
  if (!isAdmin_(d)) return jsonOut_({ ok: false, error: 'ต้องล็อกอิน Admin ก่อน' });
  const name = String(d.fullName || '').trim();
  const classroom = String(d.classroom || '').trim();
  if (!name) return jsonOut_({ ok: false, error: 'กรุณาระบุชื่อ-สกุล' });
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_STUDENTS, ['ชื่อ-สกุล', 'ชั้น/ห้อง']);
  // กันเพิ่มชื่อซ้ำในห้องเดิม (ชื่อเดียวกันแต่คนละห้อง อนุญาตได้)
  if (sh.getLastRow() >= 2) {
    const sv = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    const dupe = sv.some(r => String(r[0] || '').trim() === name && String(r[1] || '').trim() === classroom);
    if (dupe) return jsonOut_({ ok: false, error: 'มีชื่อนี้ในห้องนี้แล้ว' });
  }
  sh.appendRow([name, classroom]);
  return jsonOut_({ ok: true });
}

/** แก้ไขนักเรียน (Admin) + อัปเดตชื่อ/ห้องในงานเดิม */
function editStudent_(d) {
  if (!isAdmin_(d)) return jsonOut_({ ok: false, error: 'ต้องล็อกอิน Admin ก่อน' });
  const row = Number(d.row);
  const name = String(d.fullName || '').trim();
  const classroom = String(d.classroom || '').trim();
  if (!row || row < 2) return jsonOut_({ ok: false, error: 'row ไม่ถูกต้อง' });
  if (!name) return jsonOut_({ ok: false, error: 'กรุณาระบุชื่อ-สกุล' });
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_STUDENTS);
  if (!sh) return jsonOut_({ ok: false, error: 'ไม่พบชีตนักเรียน' });
  const oldName = String(sh.getRange(row, 1).getValue() || '').trim();
  const oldClass = String(sh.getRange(row, 2).getValue() || '').trim();
  sh.getRange(row, 1).setValue(name);
  sh.getRange(row, 2).setValue(classroom);
  // อัปเดตงานเดิมของชื่อเก่า (คอลัมน์ B ชื่อ / C ห้อง ของชีตผลงาน)
  let updated = 0;
  if (oldName && (oldName !== name || oldClass !== classroom) && String(d.updateWorks || '1') !== '0') {
    const ws = getSheet_();
    if (ws.getLastRow() >= 2) {
      const n = ws.getLastRow() - 1;
      const cols = ws.getRange(2, 2, n, 2).getValues();
      for (let i = 0; i < n; i++) {
        if (String(cols[i][0]).trim() === oldName && String(cols[i][1]).trim() === oldClass) {
          if (oldName !== name) ws.getRange(i + 2, 2).setValue(name);
          if (oldClass !== classroom) ws.getRange(i + 2, 3).setValue(classroom);
          updated++;
        }
      }
    }
  }
  return jsonOut_({ ok: true, oldName: oldName, fullName: name, classroom: classroom, updatedWorks: updated });
}

function deleteStudent_(d) {
  if (!isAdmin_(d)) return jsonOut_({ ok: false, error: 'ต้องล็อกอิน Admin ก่อน' });
  const row = Number(d.row);
  if (!row || row < 2) return jsonOut_({ ok: false, error: 'row ไม่ถูกต้อง' });
  SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_STUDENTS).deleteRow(row);
  return jsonOut_({ ok: true });
}

function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

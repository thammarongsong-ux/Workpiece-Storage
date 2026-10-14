/*************************************************************
 * Workpiece Storage - Backend (Google Apps Script)
 * เชื่อม Google Drive + Google Sheet + โหมด Admin + แยกปีการศึกษา/ภาคเรียน
 *
 * แนวคิดการแยกเทอม:
 * - ชีตผลงานแยกตามเทอม: "ผลงาน_2568_1" = ปีการศึกษา 2568 ภาคเรียนที่ 1
 *   ชีตเดิมชื่อ "ผลงาน" เก็บไว้เป็นคลังเดิม (อ่านรวมได้ ไม่ลบ)
 * - โฟลเดอร์ Drive: ROOT / "ปีการศึกษา 2568" / "ภาคเรียนที่ 1" / วิชา / ประเภทงาน / ไฟล์
 * - ชีต "วิชา" และ "นักเรียน" ใช้ร่วมกันทุกเทอม (master กลาง)
 *
 * วิธีอัปเกรดจากเวอร์ชันเดิม:
 * 1. วางไฟล์นี้ทับของเดิม > Save
 * 2. รันฟังก์ชัน setupSheet() 1 ครั้ง (สร้างชีตเทอมปัจจุบัน + อัปเกรดหัวตาราง)
 * 3. Deploy > Manage deployments > Edit > Version: New version > Deploy
 *    URL ยังเป็นตัวเดิม ไม่ต้องแก้ Frontend (แต่ควรอัป index.html ตัวใหม่ที่มีตัวเลือกปี/เทอม)
 *************************************************************/

const CONFIG = {
  SHEET_ID: '1IndHBNkd_PrciT7itRAVDjWuQzL3x0xpOrSZhUDIQ00',
  ROOT_FOLDER_ID: '1NkxlMHyKXB7Y6c8RmC4B7Bxy1-jYKz8e',
  SHEET_NAME: 'ผลงาน',            // ชีตเดิม (legacy) - เก็บไว้ไม่ลบ
  SHEET_BASE: 'ผลงาน',            // prefix ของชีตรายเทอม เช่น ผลงาน_2568_1
  SHEET_SUBJECTS: 'วิชา',
  SHEET_STUDENTS: 'นักเรียน',
  SHEET_TEACHERS: 'ครู',
  TEACHER_HEADERS: ['USER', 'PASS', 'ชื่อครู', 'วิชาที่สอน (comma คั่น)'],
  ADMIN_USER: 'FirstStar',
  ADMIN_PASS: '574001',
  VERSION: 'term-8', // ต้องตรงกับ BACKEND_VERSION ใน index.html (ไว้เช็คว่า Deploy ตัวล่าสุดแล้วหรือยัง)
  HEADERS: [
    'Timestamp','ชื่อ-สกุล','ชั้น/ห้อง','วิชา','ประเภทงาน','ชื่อชิ้นงาน',
    'คำอธิบาย','ชนิดไฟล์','ชื่อไฟล์','File URL','Drive File ID','สถานะตรวจ','ความเห็นครู',
    'ปีการศึกษา','ภาคเรียน','รหัสกลุ่ม','ผู้ตรวจ'
  ],
  HEADERS_LEGACY_LEN: 13,
  DEFAULT_SUBJECTS: ['ภาษาไทย','คณิตศาสตร์','วิทยาศาสตร์','ภาษาอังกฤษ','สังคมศึกษา','ศิลปะ','การงานอาชีพ','สุขศึกษา/พละ','คอมพิวเตอร์']
};

/** ---------- เทอม: ปีการศึกษา + ภาคเรียน ---------- */

/** เทอมปัจจุบัน (ปีการศึกษาไทยเริ่ม พ.ค.): พ.ค.-ต.ค. = เทอม 1, พ.ย.-เม.ย. = เทอม 2 */
function currentTerm_() {
  const now = new Date();
  const m = now.getMonth() + 1;
  const gY = now.getFullYear();
  if (m >= 5 && m <= 10) return { year: String(gY + 543), semester: '1' };
  if (m >= 11) return { year: String(gY + 543), semester: '2' };
  return { year: String(gY + 543 - 1), semester: '2' }; // ม.ค.-เม.ย.
}

function normalizeTerm_(year, sem) {
  const y = String(year || '').trim();
  const s = String(sem || '').trim();
  if (!/^\d{4}$/.test(y)) return null;
  if (s !== '1' && s !== '2') return null;
  const yi = Number(y);
  if (yi < 2500 || yi > 2600) return null; // กันพิมพ์ปี ค.ศ. ผิด
  return { year: y, semester: s, termId: y + '_' + s, sheetName: CONFIG.SHEET_BASE + '_' + y + '_' + s };
}

function termFromSheetName_(name) {
  const m = /^ผลงาน_(\d{4})_([12])$/.exec(String(name || ''));
  if (!m) return null;
  return { year: m[1], semester: m[2], termId: m[1] + '_' + m[2], sheetName: String(name) };
}

function termLabel_(year, sem) {
  if (!year || !sem) return 'คลังเดิม (ก่อนแยกเทอม)';
  return 'ปี ' + year + ' เทอม ' + sem;
}

function worksSheetName_(year, sem) {
  const t = normalizeTerm_(year, sem);
  return t ? t.sheetName : null;
}

/** รันครั้งเดียวเพื่อสร้างหัวตารางทุกชีต + ชีตเทอมปัจจุบัน */
function setupSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  // ชีตเดิม (เผื่อยังไม่มี)
  ensureHeaders_(getOrCreateSheet_(ss, CONFIG.SHEET_NAME, CONFIG.HEADERS));
  // ชีตเทอมปัจจุบัน เช่น ผลงาน_2568_1
  const cur = currentTerm_();
  ensureHeaders_(getOrCreateSheet_(ss, worksSheetName_(cur.year, cur.semester), CONFIG.HEADERS));
  // ชีตวิชา
  const SUBJECT_HEADERS = ['ชื่อวิชา', 'ชั้น/ห้อง (comma คั่น, ว่าง=ทุกห้อง)'];
  const shS = getOrCreateSheet_(ss, CONFIG.SHEET_SUBJECTS, SUBJECT_HEADERS);
  if (!shS.getRange(1, 2).getValue()) shS.getRange(1, 2).setValue(SUBJECT_HEADERS[1]);
  if (shS.getLastRow() < 2) {
    shS.getRange(2, 1, CONFIG.DEFAULT_SUBJECTS.length, 1).setValues(CONFIG.DEFAULT_SUBJECTS.map(function(s) { return [s]; }));
  }
  // ชีตนักเรียน
  getOrCreateSheet_(ss, CONFIG.SHEET_STUDENTS, ['ชื่อ-สกุล', 'ชั้น/ห้อง']);
  // ชีตบัญชีครู (USER/PASS/ชื่อ/วิชาที่สอน) — สร้างเปล่าไว้ก่อน Admin มากรอกเอง
  getOrCreateSheet_(ss, CONFIG.SHEET_TEACHERS, CONFIG.TEACHER_HEADERS);
  // อัปเกรดหัวตารางชีตรายเทอมที่มีอยู่แล้ว (ถ้ามี)
  listTermSheets_().forEach(function(t) {
    const sh = ss.getSheetByName(t.sheetName);
    if (sh) ensureHeaders_(sh);
  });
  Logger.log('setupSheet OK, current term: ' + cur.year + '/' + cur.semester);
}

/** เติม/อัปเกรดหัวตารางเป็น 17 คอลัมน์ (รองรับชีตเก่า 13/15/16 คอลัมน์โดยไม่ลบข้อมูล) */
function ensureHeaders_(sh) {
  if (!sh) return sh;
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
    sh.setFrozenRows(1);
    return sh;
  }
  const width = sh.getLastColumn();
  if (width < CONFIG.HEADERS.length) {
    sh.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
  } else {
    // เติมเฉพาะหัวที่ขาด (คอลัมน์ 14-17) กันทับหัวเดิมที่ผู้ใช้อาจแก้
    const h14 = String(sh.getRange(1, 14).getValue() || '').trim();
    const h15 = String(sh.getRange(1, 15).getValue() || '').trim();
    const h16 = String(sh.getRange(1, 16).getValue() || '').trim();
    if (!h14) sh.getRange(1, 14).setValue(CONFIG.HEADERS[13]);
    if (!h15) sh.getRange(1, 15).setValue(CONFIG.HEADERS[14]);
    if (!h16) sh.getRange(1, 16).setValue(CONFIG.HEADERS[15]);
    try { if (!String(sh.getRange(1, 17).getValue() || '').trim()) sh.getRange(1, 17).setValue(CONFIG.HEADERS[16]); } catch (e) {}
  }
  return sh;
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

/** ชีตงานรายเทอม (สร้างใหม่ถ้ายังไม่มี) */
function getWorksSheet_(year, sem) {
  const t = normalizeTerm_(year, sem);
  if (!t) throw new Error('ปีการศึกษา/ภาคเรียนไม่ถูกต้อง');
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  return ensureHeaders_(getOrCreateSheet_(ss, t.sheetName, CONFIG.HEADERS));
}

function getLegacySheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  return ensureHeaders_(getOrCreateSheet_(ss, CONFIG.SHEET_NAME, CONFIG.HEADERS));
}

/** รายชื่อชีตรายเทอมทั้งหมดที่มีอยู่จริง */
function listTermSheets_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const out = [];
  ss.getSheets().forEach(function(sh) {
    const t = termFromSheetName_(sh.getName());
    if (t) out.push(t);
  });
  out.sort(function(a, b) {
    if (a.year !== b.year) return a.year < b.year ? -1 : 1;
    return a.semester < b.semester ? -1 : 1;
  });
  return out;
}

/** หาชีตงานจากพารามิเตอร์ที่ Frontend ส่งมา (รองรับของเก่าที่ส่งแค่ row) */
function resolveWorksSheet_(d) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  if (d && d.sheetName) {
    const sh = ss.getSheetByName(String(d.sheetName));
    if (sh) return sh;
  }
  if (d && d.termId) {
    const m = /^(\d{4})_([12])$/.exec(String(d.termId));
    if (m) {
      const sh = ss.getSheetByName(CONFIG.SHEET_BASE + '_' + m[1] + '_' + m[2]);
      if (sh) return sh;
    }
  }
  const y = d && (d.academicYear || d.year);
  const s = d && (d.semester || d.term);
  if (y && s) {
    const t = normalizeTerm_(y, s);
    if (t) {
      const sh = ss.getSheetByName(t.sheetName);
      if (sh) return sh;
    }
  }
  return getLegacySheet_();
}

/** ชีตงานทั้งหมด (legacy + รายเทอม) สำหรับค้น/อัปเดตข้ามเทอม */
function allWorksSheets_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const arr = [];
  const legacy = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (legacy) arr.push({ sh: legacy, term: null });
  listTermSheets_().forEach(function(t) {
    const sh = ss.getSheetByName(t.sheetName);
    if (sh) arr.push({ sh: sh, term: t });
  });
  return arr;
}

function jsonOut_(obj) {
  try { if (obj && typeof obj === 'object' && !obj.version) obj.version = CONFIG.VERSION; } catch (e) {}
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function isAdmin_(d) {
  return String(d && d.adminUser || '') === CONFIG.ADMIN_USER && String(d && d.adminPass || '') === CONFIG.ADMIN_PASS;
}

/** ---------- บัญชีครู (ตรวจได้เฉพาะวิชาที่สอน, จัดการระบบไม่ได้) ---------- */

function getTeachersSheet_(ss) {
  const s = ss || SpreadsheetApp.openById(CONFIG.SHEET_ID);
  return getOrCreateSheet_(s, CONFIG.SHEET_TEACHERS, CONFIG.TEACHER_HEADERS);
}

/** ค้นครูจาก USER+PASS (trim ตรงตัว, เทียบ case-sensitive) คืน {row,user,name,subjects} หรือ null */
function findTeacher_(user, pass) {
  const u = String(user || '').trim();
  const p = String(pass || '').trim();
  if (!u || !p) return null;
  const sh = getTeachersSheet_(null);
  if (!sh || sh.getLastRow() < 2) return null;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim() === u && String(vals[i][1] || '').trim() === p) {
      return { row: i + 2, user: u, name: String(vals[i][2] || '').trim() || u, subjects: parseClasses_(vals[i][3]) };
    }
  }
  return null;
}

/** ล็อกอิน: admin ก่อน, ไม่ใช่ค่อยหาชีตครู — ไม่เปิดเผยว่า user ผิดหรือ pass ผิด */
function login_(d) {
  if (isAdmin_(d)) {
    return jsonOut_({ ok: true, role: 'admin', user: CONFIG.ADMIN_USER, name: CONFIG.ADMIN_USER, subjects: [] });
  }
  const t = findTeacher_(d && d.adminUser, d && d.adminPass);
  if (t) {
    return jsonOut_({ ok: true, role: 'teacher', user: t.user, name: t.name, subjects: t.subjects });
  }
  return jsonOut_({ ok: false, error: 'USER หรือ PASS ไม่ถูกต้อง' });
}

/** ตรวจได้ไหม: admin ได้ทุกวิชา, ครูได้เฉพาะวิชาในลิสต์ตัวเอง */
function canGradeSubject_(d, subject) {
  if (isAdmin_(d)) return { ok: true, role: 'admin' };
  const t = findTeacher_(d && d.adminUser, d && d.adminPass);
  if (!t) return { ok: false, error: 'ต้องล็อกอินก่อน' };
  const subj = String(subject || '').trim();
  if (t.subjects.indexOf(subj) < 0) {
    return { ok: false, error: 'วิชา "' + subj + '" ไม่ใช่ของคุณ (คุณตรวจได้: ' + (t.subjects.join(', ') || '—') + ')' };
  }
  return { ok: true, role: 'teacher', teacher: t };
}

/** แปลง "ป.5/1, ป.5/2" -> ["ป.5/1","ป.5/2"] (รับ string หรือ array) */
function parseClasses_(v) {
  if (!v) return [];
  const arr = (v instanceof Array) ? v : String(v).split(',');
  const out = [];
  arr.forEach(function(x) { const t = String(x || '').trim(); if (t && out.indexOf(t) < 0) out.push(t); });
  return out;
}

/** ---------- กันส่งงานซ้ำ (กดเบิ้ล/เน็ต timeout แล้วกดใหม่) ---------- */
// Frontend สร้าง submitId ใหม่ทุกครั้งที่กรอกฟอร์มใหม่ แต่ใช้รหัสเดิมเมื่อกด "ส่งอีกครั้ง"
// backend จำผลลัพธ์สำเร็จไว้ใน ScriptCache 6 ชม. เจอรหัสเดิม = คืนผลเดิม ไม่สร้างไฟล์/แถวใหม่
function submitCacheGet_(sid) {
  if (!/^[A-Za-z0-9-]{8,64}$/.test(String(sid || ''))) return null;
  try {
    const s = CacheService.getScriptCache().get('sid_' + sid);
    return s ? JSON.parse(s) : null;
  } catch (e) { return null; }
}

function submitCachePut_(sid, obj) {
  if (!/^[A-Za-z0-9-]{8,64}$/.test(String(sid || ''))) return;
  try { CacheService.getScriptCache().put('sid_' + sid, JSON.stringify(obj), 21600); } catch (e) {}
}

// GET ?action=init (รวม list+config ในครั้งเดียว เร็วสุด) | ?action=list[&year=&semester=] | ?action=listConfig | ?action=ping | ?action=terms
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'ping';
    if (action === 'init') {
      const p = (e && e.parameter) || {};
      return getBootstrap_(p.year || p.academicYear || '', p.semester || p.term || '');
    }
    if (action === 'list') {
      const p = (e && e.parameter) || {};
      return listWorks_(p.year || p.academicYear || '', p.semester || p.term || '');
    }
    if (action === 'terms') return jsonOut_({ ok: true, terms: listTermSheets_(), current: currentTerm_() });
    if (action === 'listConfig') return listConfig_();
    return jsonOut_({ ok: true, message: 'Student Workpiece Storage API ready', time: new Date(), current: currentTerm_() });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

/** รวมทุกอย่างที่หน้าเว็บต้องใช้ตอนเปิดครั้งแรกใน response เดียว (ลด cold-start จาก 2 รอบเหลือ 1 รอบ) */
function getBootstrap_(yearFilter, semFilter) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const cfg = readConfigObj_(ss);
  const data = readWorksObj_(ss, yearFilter, semFilter);
  return jsonOut_({
    ok: true,
    rows: data.rows, total: data.total,
    terms: data.terms, current: currentTerm_(),
    subjects: cfg.subjects, subjectRows: cfg.subjectRows, students: cfg.students
  });
}

// POST {action:'submit'|'updateStatus'|'login'|'addSubject'|'editSubject'|'deleteSubject'|'addStudent'|'editStudent'|'deleteStudent'|'deleteWork'|'listTeachers'|'addTeacher'|'editTeacher'|'deleteTeacher', ..., academicYear, semester}
// งานเขียน (submit/ตรวจ/ลบ/จัดการ) ต้องผ่าน ScriptLock ใน doPost — รายชื่อนี้บอกว่า action ไหนต้องเข้าคิว
var WRITE_ACTIONS = {submit:1,updateStatus:1,deleteWork:1,backfillTerm:1,addSubject:1,editSubject:1,deleteSubject:1,addStudent:1,editStudent:1,deleteStudent:1,addTeacher:1,editTeacher:1,deleteTeacher:1};
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    // งานเขียนทุกชนิดต้องเข้าคิวด้วย ScriptLock (กันหลายคนกดพร้อมกันแล้วข้อมูลพันกัน:
    // โฟลเดอร์ซ้ำ, แถวซ้อน, เช็คซ้ำแล้วเขียนทับกัน) — งานอ่าน (login/list/init) ไม่ล็อก เร็วเหมือนเดิม
    if (WRITE_ACTIONS[data.action]) {
      const lock = LockService.getScriptLock();
      const got = lock.tryLock(10000);
      if (!got) return jsonOut_({ ok: false, busy: true, error: 'มีผู้ใช้งานพร้อมกันเยอะ ระบบให้คิวไม่ทัน กรุณากดลองใหม่อีกครั้ง' });
      try {
        return dispatchPost_(data);
      } finally {
        try { lock.releaseLock(); } catch (x) {}
      }
    }
    return dispatchPost_(data);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err && err.stack || err) });
  }
}

function dispatchPost_(data) {
    if (data.action === 'submit') return submitWork_(data);
    if (data.action === 'login') {
      return login_(data);
    }
    if (data.action === 'updateStatus') return updateStatus_(data);
    if (data.action === 'deleteWork') return deleteWork_(data);
    if (data.action === 'backfillTerm') return backfillTerm_(data);
    if (data.action === 'addSubject') return addSubject_(data);
    if (data.action === 'editSubject') return editSubject_(data);
    if (data.action === 'deleteSubject') return deleteSubject_(data);
    if (data.action === 'addStudent') return addStudent_(data);
    if (data.action === 'editStudent') return editStudent_(data);
    if (data.action === 'deleteStudent') return deleteStudent_(data);
    if (data.action === 'listTeachers') return listTeachers_(data);
    if (data.action === 'addTeacher') return addTeacher_(data);
    if (data.action === 'editTeacher') return editTeacher_(data);
    if (data.action === 'deleteTeacher') return deleteTeacher_(data);
    return jsonOut_({ ok: false, error: 'unknown action' });
}

/** ดึงรายชื่อวิชา + นักเรียน จาก Sheet (สาธารณะ อ่านได้อย่างเดียว) */
function listConfig_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const cfg = readConfigObj_(ss);
  cfg.ok = true;
  cfg.current = currentTerm_();
  cfg.terms = listTermSheets_();
  return jsonOut_(cfg);
}

/** อ่าน config เป็น object (ใช้ร่วมกับ getBootstrap_ เปิด SS ครั้งเดียว) */
function readConfigObj_(ss) {
  const shS = getOrCreateSheet_(ss, CONFIG.SHEET_SUBJECTS, ['ชื่อวิชา', 'ชั้น/ห้อง (comma คั่น, ว่าง=ทุกห้อง)']);
  if (!shS.getRange(1, 2).getValue()) shS.getRange(1, 2).setValue('ชั้น/ห้อง (comma คั่น, ว่าง=ทุกห้อง)');
  const shT = getOrCreateSheet_(ss, CONFIG.SHEET_STUDENTS, ['ชื่อ-สกุล', 'ชั้น/ห้อง']);
  let subjectRows = [];
  if (shS.getLastRow() >= 2) {
    const sv = shS.getRange(2, 1, shS.getLastRow() - 1, 2).getValues();
    sv.forEach(function(r, i) {
      const nm = String(r[0] || '').trim();
      if (nm) subjectRows.push({ row: i + 2, name: nm, classrooms: parseClasses_(r[1]) });
    });
  }
  let subjects = subjectRows.map(function(o) { return o.name; });
  if (!subjects.length) {
    subjects = CONFIG.DEFAULT_SUBJECTS.slice();
    subjectRows = subjects.map(function(s, i) { return { row: i + 2, name: s, classrooms: [] }; });
  }
  let students = [];
  if (shT.getLastRow() >= 2) {
    students = shT.getRange(2, 1, shT.getLastRow() - 1, 2).getValues()
      .map(function(r, i) { return { row: i + 2, fullName: String(r[0] || '').trim(), classroom: String(r[1] || '').trim() }; })
      .filter(function(s) { return !!s.fullName; });
  }
  return { subjects: subjects, subjectRows: subjectRows, students: students };
}

/** ลิสต์ชีตงานทั้งหมดจาก ss ที่เปิดไว้แล้ว (ไม่เปิดใหม่ซ้ำ) */
function allWorksSheetsOf_(ss) {
  const arr = [];
  const legacy = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (legacy) arr.push({ sh: legacy, term: null });
  ss.getSheets().forEach(function(sh) {
    const t = termFromSheetName_(sh.getName());
    if (t) arr.push({ sh: sh, term: t });
  });
  arr.sort(function(a, b) {
    const ya = a.term ? a.term.year : '', yb = b.term ? b.term.year : '';
    if (ya !== yb) return ya < yb ? -1 : 1;
    const sa = a.term ? a.term.semester : '', sb = b.term ? b.term.semester : '';
    return sa < sb ? -1 : (sa > sb ? 1 : 0);
  });
  return arr;
}

/** อ่านแถวจากชีตหนึ่งให้เป็น object (รองรับทั้ง 13 และ 15 คอลัมน์) */
function readWorksSheet_(sh, term) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const width = Math.max(sh.getLastColumn(), CONFIG.HEADERS_LEGACY_LEN);
  const ncols = Math.min(width, CONFIG.HEADERS.length);
  const vals = sh.getRange(2, 1, last - 1, ncols).getValues();
  const sheetName = sh.getName();
  const year = term ? term.year : (ncols >= 14 ? '' : '');
  const sem = term ? term.semester : '';
  return vals.map(function(r, i) {
    // ถ้าเป็นชีตรายเทอม ใช้ปี/เทอมจากชื่อชีตเป็นหลัก; ถ้า legacy ลองอ่านจากคอลัมน์ 14-15 (ถ้ามี)
    let y = term ? term.year : String(r[13] || '').trim();
    let s = term ? term.semester : String(r[14] || '').trim();
    return {
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
      comment: String(r[12] || ''),
      academicYear: y,
      semester: s,
      groupId: String(r[15] || '').trim(),
      gradedBy: String(r[16] || '').trim(),
      termId: term ? term.termId : (y && (s === '1' || s === '2') ? y + '_' + s : ''),
      sheetName: sheetName,
      termLabel: term ? termLabel_(term.year, term.semester) : (y && s ? termLabel_(y, s) : 'คลังเดิม (ก่อนแยกเทอม)')
    };
  }).filter(function(r) { return r.fullName || r.title; });
}

/** ดึงข้อมูลผลงาน (รวมทุกเทอม + คลังเดิม) หรือกรองเฉพาะเทอม — กรองระดับแถว ข้อมูลจึงตรง */
function listWorks_(yearFilter, semFilter) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  return jsonOut_(readWorksObj_(ss, yearFilter, semFilter));
}

/** อ่านงานเป็น object (เปิด SS ครั้งเดียว, กรองที่แถว ไม่ทิ้งชีต legacy ที่มีปีในคอลัมน์) */
function readWorksObj_(ss, yearFilter, semFilter) {
  const yf = String(yearFilter || '').trim();
  const sf = String(semFilter || '').trim();
  const terms = [];
  ss.getSheets().forEach(function(sh) {
    const t = termFromSheetName_(sh.getName());
    if (t) terms.push(t);
  });
  terms.sort(function(a, b) {
    if (a.year !== b.year) return a.year < b.year ? -1 : 1;
    return a.semester < b.semester ? -1 : 1;
  });
  let rows = [];
  allWorksSheetsOf_(ss).forEach(function(o) { rows = rows.concat(readWorksSheet_(o.sh, o.term)); });
  // กรองระดับแถว: ถูกต้องทั้งชีตรายเทอมและคลังเดิมที่มีคอลัมน์ปี/เทอม
  if (yf) rows = rows.filter(function(r) { return String(r.academicYear || '') === yf; });
  if (sf === '1' || sf === '2') rows = rows.filter(function(r) { return String(r.semester || '') === sf; });
  // เรียงใหม่สุดก่อน (timestamp รูปแบบ yyyy-MM-dd HH:mm เรียง string ได้)
  rows.sort(function(a, b) {
    if (a.timestamp < b.timestamp) return 1;
    if (a.timestamp > b.timestamp) return -1;
    return 0;
  });
  return { rows: rows, total: rows.length, terms: terms };
}

/** นักเรียนส่งงาน (ไม่ต้องล็อกอิน) — ต้องระบุ academicYear + semester */
function submitWork_(d) {
  const subject  = String(d.subject || '').trim();
  const category = String(d.category || '').trim();
  const title    = String(d.title || '').trim();
  let term = normalizeTerm_(d.academicYear || d.year, d.semester || d.term);
  if (!term) {
    // รองรับ Frontend เก่าที่ยังไม่ส่งปี/เทอมมา: ลงเทอมปัจจุบันอัตโนมัติ
    const cur = currentTerm_();
    term = normalizeTerm_(cur.year, cur.semester);
  }
  if (!term) {
    return jsonOut_({ ok: false, error: 'กรุณาเลือก ปีการศึกษา / ภาคเรียน (1 หรือ 2) ให้ถูกต้อง' });
  }
  // สมาชิก: งานเดี่ยวส่ง 1 คน, งานกลุ่มส่ง members[] หลายคน (อัปโหลดไฟล์ครั้งเดียว แตกแถวรายคน)
  let members = [];
  if (d.members instanceof Array && d.members.length) {
    const seen = {};
    d.members.forEach(function(m) {
      const nm = String((m && (m.fullName || m.name)) || '').trim();
      const cl = String((m && (m.classroom || m.class)) || '').trim();
      if (!nm || !cl) return;
      const k = nm + '||' + cl;
      if (!seen[k]) { seen[k] = 1; members.push({ fullName: nm, classroom: cl }); }
    });
  } else {
    const nm = String(d.fullName || '').trim();
    const cl = String(d.classroom || '').trim();
    if (nm && cl) members.push({ fullName: nm, classroom: cl });
  }
  if (!members.length) {
    return jsonOut_({ ok: false, error: 'กรุณาเลือก ชื่อ-สกุล / ห้อง อย่างน้อย 1 คน' });
  }
  if (members.length > 10) {
    return jsonOut_({ ok: false, error: 'กลุ่มใหญ่สุด 10 คนต่อชิ้นงาน' });
  }
  if (!subject || !category || !title) {
    return jsonOut_({ ok: false, error: 'กรุณาเลือก วิชา / ประเภทงาน / ชื่อชิ้นงาน ให้ครบ' });
  }
  // ถ้าล็อกอินเป็นครู (ส่ง adminUser/adminPass มาด้วย) จำกัดวิชาที่ส่งได้เฉพาะวิชาตัวเอง
  // นักเรียนทั่วไปไม่ล็อกอิน ส่งได้ทุกวิชาเหมือนเดิม
  if (!isAdmin_(d) && String(d.adminUser || '').trim()) {
    const perm = canGradeSubject_(d, subject);
    if (!perm.ok) return jsonOut_({ ok: false, error: perm.error || 'บัญชีครูไม่ถูกต้อง' });
  }
  // เคยสำเร็จด้วยรหัสนี้แล้ว (กดเบิ้ล/timeout แล้วกดซ้ำ) = คืนผลเดิม ไม่สร้างไฟล์+แถวใหม่
  const sid = String(d.submitId || '');
  const dup = submitCacheGet_(sid);
  if (dup) { dup.duplicate = true; return jsonOut_(dup); }
  // สุ่มท้ายรหัสกลุ่ม กัน 2 กลุ่มส่งวินาทีเดียวกันแล้วรหัสชนกัน (สมาชิกจะปนกันในหน้า Padlet)
  const groupId = members.length > 1
    ? ('G' + Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMddHHmmss') + Math.random().toString(36).slice(2, 5))
    : '';
  // กันเลือกวิชาผิดห้อง (เช็คทุกคนในกลุ่ม — ชื่อวิชาซ้ำต่างห้องผ่านถ้าห้องใดห้องหนึ่งตรง)
  // เปิด SS ครั้งเดียวแล้วใช้ซ้ำถึงตอน appendRow (ลด cold-start ~1-2 วิ กัน timeout ตอนส่ง)
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  try {
    const shS = ss.getSheetByName(CONFIG.SHEET_SUBJECTS);
    if (shS && shS.getLastRow() >= 2) {
      const sv = shS.getRange(2, 1, shS.getLastRow() - 1, 2).getValues();
      let found = false, allowList = [];
      for (let i = 0; i < sv.length; i++) {
        if (String(sv[i][0] || '').trim() === subject) {
          found = true;
          const allow = parseClasses_(sv[i][1]);
          if (!allow.length) { allowList = []; break; } // วิชานี้ใช้ได้ทุกห้อง
          allowList = allowList.concat(allow.filter(function(x) { return allowList.indexOf(x) < 0; }));
        }
      }
      if (found && allowList.length) {
        const bad = members.filter(function(m) { return allowList.indexOf(m.classroom) < 0; });
        if (bad.length) {
          return jsonOut_({ ok: false, error: 'วิชา "' + subject + '" ไม่ได้กำหนดให้ห้อง ' + bad.map(function(m) { return m.classroom; }).join(', ') + ' (กำหนดไว้: ' + allowList.join(', ') + ')' });
        }
      }
    }
  } catch (e) { }
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
    // โฟลเดอร์แยกเทอม: ROOT / ปีการศึกษา XXXX / ภาคเรียนที่ X / วิชา / ประเภทงาน
    const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
    const fYear = getOrCreateFolder_(root, 'ปีการศึกษา ' + term.year);
    const fTerm = getOrCreateFolder_(fYear, 'ภาคเรียนที่ ' + term.semester);
    const sub1 = getOrCreateFolder_(fTerm, subject);
    const sub2 = getOrCreateFolder_(sub1, category);
    const safeName = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd_HHmmss') + '_' + members[0].fullName + (members.length > 1 ? '_กลุ่ม' + members.length + 'คน' : '') + '_' + fileName;
    const file = sub2.createFile(blob.setName(safeName));
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    driveFileId = file.getId();
    fileUrl = 'https://drive.google.com/file/d/' + driveFileId + '/view';
  }
  const now = new Date();
  const ws = ensureHeaders_(getOrCreateSheet_(ss, term.sheetName, CONFIG.HEADERS));
  members.forEach(function(m) {
    ws.appendRow([now, m.fullName, m.classroom, subject, category, title,
      String(d.description || ''), fileKind, fileName, fileUrl, driveFileId, 'รอตรวจ', '', term.year, term.semester, groupId]);
  });
  const resp = { ok: true, fileUrl: fileUrl, driveFileId: driveFileId, academicYear: term.year, semester: term.semester, sheetName: term.sheetName, groupId: groupId, rows: members.length, isGroup: members.length > 1 };
  submitCachePut_(sid, resp); // จำผลไว้กันกดซ้ำ
  return jsonOut_(resp);
}

/** ครูตรวจงาน (Admin ทุกวิชา / ครูเฉพาะวิชาที่สอน) — ต้องส่ง sheetName/termId/year+semester มาด้วย (ของเก่าส่งแค่ row = คลังเดิม) */
function updateStatus_(d) {
  const row = Number(d.row);
  const status = String(d.status || '');
  const comment = String(d.comment || '');
  if (!row || row < 2) return jsonOut_({ ok: false, error: 'row ไม่ถูกต้อง' });
  if (['รอตรวจ', 'ผ่าน', 'ต้องแก้ไข'].indexOf(status) < 0) return jsonOut_({ ok: false, error: 'status ไม่ถูกต้อง' });
  const sh = resolveWorksSheet_(d);
  if (row > sh.getLastRow()) return jsonOut_({ ok: false, error: 'ไม่พบแถวนี้ในชีต ' + sh.getName() });
  // อ่านตัวตนของแถว (B..F) ครั้งเดียว: ใช้ทั้งเช็คสิทธิ์ + เช็คว่าแถวเลื่อนเพราะมีคนลบพร้อมกันหรือไม่
  const idv = sh.getRange(row, 2, 1, 5).getValues()[0];
  const expN = String(d.expectFullName || ''), expT = String(d.expectTitle || '');
  if ((expN && String(idv[0] || '') !== expN) || (expT && String(idv[4] || '') !== expT)) {
    return jsonOut_({ ok: false, stale: true, error: 'ข้อมูลแถวนี้เปลี่ยนไปแล้ว (มีคนลบ/แก้ไขพร้อมกัน) กรุณากดรีเฟรชแล้วทำใหม่' });
  }
  // เช็คสิทธิ์จากวิชาจริงในแถว (ไม่ใช้ค่าที่ client ส่งมา กันปลอม)
  const subject = String(idv[2] || '').trim();
  const perm = canGradeSubject_(d, subject);
  if (!perm.ok) return jsonOut_({ ok: false, error: perm.error || 'ต้องล็อกอินก่อน' });
  const grader = perm.role === 'admin' ? 'Admin' : (perm.teacher ? perm.teacher.name : '');
  const stamp = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm');
  const gradedBy = (grader ? grader + ' ' : '') + stamp;
  try { if (!String(sh.getRange(1, 17).getValue() || '').trim()) sh.getRange(1, 17).setValue(CONFIG.HEADERS[16]); } catch (e) {}
  sh.getRange(row, 12).setValue(status);
  sh.getRange(row, 13).setValue(comment);
  sh.getRange(row, 17).setValue(gradedBy);
  return jsonOut_({ ok: true, row: row, status: status, sheetName: sh.getName(), role: perm.role, gradedBy: gradedBy });
}

/** ลบงาน (Admin เท่านั้น) */
function deleteWork_(d) {
  if (!isAdmin_(d)) return jsonOut_({ ok: false, error: 'ต้องล็อกอิน Admin ก่อน' });
  const row = Number(d.row);
  if (!row || row < 2) return jsonOut_({ ok: false, error: 'row ไม่ถูกต้อง' });
  const sh = resolveWorksSheet_(d);
  if (row > sh.getLastRow()) return jsonOut_({ ok: true, alreadyDeleted: true, row: row, sheetName: sh.getName() });
  // กันลบผิดแถว: ถ้าชื่อ/ชื่องานไม่ตรงกับที่เห็นบนจอ = แถวเลื่อนเพราะมีคนลบพร้อมกัน ให้หยุดก่อน
  const expN = String(d.expectFullName || ''), expT = String(d.expectTitle || '');
  if (expN || expT) {
    const idv = sh.getRange(row, 2, 1, 5).getValues()[0];
    if ((expN && String(idv[0] || '') !== expN) || (expT && String(idv[4] || '') !== expT)) {
      return jsonOut_({ ok: false, stale: true, error: 'ข้อมูลแถวนี้เปลี่ยนไปแล้ว (มีคนลบพร้อมกัน) กรุณากดรีเฟรชแล้วทำใหม่ — ยังไม่ได้ลบอะไร' });
    }
  }
  const fileId = String(sh.getRange(row, 11).getValue() || '').trim();
  // งานกลุ่มแชร์ไฟล์เดียวกัน: ลบไฟล์ใน Drive ก็ต่อเมื่อไม่มีแถวอื่นอ้างถึงแล้ว
  const canTrash = fileId && !fileStillUsed_(sh.getName(), row, fileId);
  if (canTrash) {
    try { DriveApp.getFileById(fileId).setTrashed(true); }
    catch (e) { }
  }
  sh.deleteRow(row);
  return jsonOut_({ ok: true, row: row, sheetName: sh.getName() });
}

/** มีแถวอื่น (ที่ไม่ใช่แถวกำลังลบ) อ้าง driveFileId นี้อยู่หรือไม่ */
function fileStillUsed_(excludeSheetName, excludeRow, fileId) {
  const all = allWorksSheets_();
  for (let k = 0; k < all.length; k++) {
    const ws = all[k].sh;
    if (ws.getLastRow() < 2 || ws.getLastColumn() < 11) continue;
    const n = ws.getLastRow() - 1;
    const col = ws.getRange(2, 11, n, 1).getValues();
    for (let i = 0; i < n; i++) {
      if (String(col[i][0] || '').trim() === fileId) {
        if (ws.getName() === excludeSheetName && (i + 2) === Number(excludeRow)) continue;
        return true;
      }
    }
  }
  return false;
}

/** ย้อนหลังใส่ปี/เทอมให้แถวที่ยังว่าง (Admin เท่านั้น) — แตะเฉพาะคอลัมน์ N-O ที่ว่าง ไม่แตะแถวที่มีปีแล้ว */
function backfillTerm_(d) {
  if (!isAdmin_(d)) return jsonOut_({ ok: false, error: 'ต้องล็อกอิน Admin ก่อน' });
  const term = normalizeTerm_(d.academicYear || d.year, d.semester || d.term);
  if (!term) return jsonOut_({ ok: false, error: 'กรุณาระบุ ปีการศึกษา / ภาคเรียน ให้ถูกต้อง' });
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let updated = 0, sheets = 0, scanned = 0;
  allWorksSheetsOf_(ss).forEach(function(o) {
    const ws = o.sh;
    ensureHeaders_(ws);
    if (ws.getLastRow() < 2) return;
    sheets++;
    const n = ws.getLastRow() - 1;
    // ชีตรายเทอมใช้ปี/เทอมของชีตนั้น, ชีตคลังเดิมใช้ค่าที่ Admin ระบุ
    const y = o.term ? o.term.year : term.year;
    const s = o.term ? o.term.semester : term.semester;
    const range = ws.getRange(2, 14, n, 2);
    const vals = range.getValues();
    let dirty = false;
    for (let i = 0; i < n; i++) {
      scanned++;
      if (!String(vals[i][0] || '').trim()) { vals[i][0] = y; vals[i][1] = s; updated++; dirty = true; }
    }
    if (dirty) range.setValues(vals); // เขียนครั้งเดียวต่อชีต
  });
  return jsonOut_({ ok: true, updated: updated, scanned: scanned, sheets: sheets, academicYear: term.year, semester: term.semester });
}

/** เพิ่มวิชา (Admin) */
function addSubject_(d) {
  if (!isAdmin_(d)) return jsonOut_({ ok: false, error: 'ต้องล็อกอิน Admin ก่อน' });
  const name = String(d.subject || '').trim();
  if (!name) return jsonOut_({ ok: false, error: 'กรุณาระบุชื่อวิชา' });
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_SUBJECTS, ['ชื่อวิชา', 'ชั้น/ห้อง (comma คั่น, ว่าง=ทุกห้อง)']);
  const rooms = parseClasses_(d.classrooms);
  if (sh.getLastRow() >= 2) {
    const sv = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    const dupe = sv.some(function(r) { return String(r[0] || '').trim() === name && parseClasses_(r[1]).join('|') === rooms.join('|'); });
    if (dupe) return jsonOut_({ ok: false, error: 'มีวิชานี้ + ห้องชุดนี้แล้ว' });
  }
  sh.appendRow([name, rooms.join(', ')]);
  return jsonOut_({ ok: true, subject: name, classrooms: rooms });
}

/** แก้ไขชื่อวิชา (Admin) + อัปเดตงานเดิมทุกเทอมที่ใช้ชื่อเก่า */
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
  if (name !== oldName && sh.getLastRow() >= 2) {
    const cur = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function(r) { return String(r[0]).trim(); });
    if (cur.indexOf(name) >= 0) return jsonOut_({ ok: false, error: 'มีวิชาชื่อนี้แล้ว' });
  }
  sh.getRange(row, 1).setValue(name);
  let rooms = parseClasses_(sh.getRange(row, 2).getValue());
  if (d.hasOwnProperty('classrooms')) {
    rooms = parseClasses_(d.classrooms);
    sh.getRange(row, 2).setValue(rooms.join(', '));
  }
  let updated = 0;
  if (oldName && oldName !== name && String(d.updateWorks || '1') !== '0') {
    allWorksSheets_().forEach(function(o) {
      const ws = o.sh;
      if (ws.getLastRow() >= 2) {
        const n = ws.getLastRow() - 1;
        const range = ws.getRange(2, 4, n, 1);
        const col = range.getValues();
        let dirty = false;
        for (let i = 0; i < n; i++) {
          if (String(col[i][0]).trim() === oldName) { col[i][0] = name; updated++; dirty = true; }
        }
        if (dirty) range.setValues(col); // เขียนครั้งเดียวต่อชีต เร็ว/ไม่ค้าง
      }
    });
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
  if (sh.getLastRow() >= 2) {
    const sv = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    const dupe = sv.some(function(r) { return String(r[0] || '').trim() === name && String(r[1] || '').trim() === classroom; });
    if (dupe) return jsonOut_({ ok: false, error: 'มีชื่อนี้ในห้องนี้แล้ว' });
  }
  sh.appendRow([name, classroom]);
  return jsonOut_({ ok: true });
}

/** แก้ไขนักเรียน (Admin) + อัปเดตชื่องานเดิมทุกเทอม */
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
  let updated = 0;
  if (oldName && (oldName !== name || oldClass !== classroom) && String(d.updateWorks || '1') !== '0') {
    allWorksSheets_().forEach(function(o) {
      const ws = o.sh;
      if (ws.getLastRow() >= 2) {
        const n = ws.getLastRow() - 1;
        const range = ws.getRange(2, 2, n, 2);
        const cols = range.getValues();
        let dirty = false;
        for (let i = 0; i < n; i++) {
          if (String(cols[i][0]).trim() === oldName) {
            if (oldName !== name) cols[i][0] = name;
            if (oldClass !== classroom) cols[i][1] = classroom;
            updated++; dirty = true;
          }
        }
        if (dirty) range.setValues(cols); // เขียนครั้งเดียวต่อชีต
      }
    });
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

/** ---------- จัดการบัญชีครู (Admin เท่านั้น) ---------- */

/** ลิสต์บัญชีครูให้ Admin (รวม PASS ไว้รีเซ็ต — ไม่เคยส่งให้สาธารณะ) */
function listTeachers_(d) {
  if (!isAdmin_(d)) return jsonOut_({ ok: false, error: 'เฉพาะ Admin เท่านั้น' });
  const sh = getTeachersSheet_(null);
  const out = [];
  if (sh && sh.getLastRow() >= 2) {
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    vals.forEach(function(r, i) {
      const user = String(r[0] || '').trim();
      if (user) out.push({ row: i + 2, user: user, pass: String(r[1] || ''), name: String(r[2] || '').trim() || user, subjects: parseClasses_(r[3]) });
    });
  }
  return jsonOut_({ ok: true, teachers: out });
}

function addTeacher_(d) {
  if (!isAdmin_(d)) return jsonOut_({ ok: false, error: 'เฉพาะ Admin เท่านั้น' });
  const user = String(d.user || d.username || '').trim();
  const pass = String(d.pass || d.password || '').trim();
  const name = String(d.name || d.teacherName || '').trim() || user;
  const subjects = parseClasses_(d.subjects);
  if (!user || !pass || !name) return jsonOut_({ ok: false, error: 'กรุณากรอก USER / PASS / ชื่อครู ให้ครบ' });
  if (user === CONFIG.ADMIN_USER) return jsonOut_({ ok: false, error: 'USER นี้ซ้ำกับ Admin' });
  if (!subjects.length) return jsonOut_({ ok: false, error: 'กรุณาระบุวิชาที่สอนอย่างน้อย 1 วิชา' });
  const sh = getTeachersSheet_(null);
  if (sh.getLastRow() >= 2) {
    const col = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < col.length; i++) {
      if (String(col[i][0] || '').trim() === user) return jsonOut_({ ok: false, error: 'มี USER นี้แล้ว' });
    }
  }
  const clash = teacherSubjectClash_(sh, subjects, 0);
  if (clash) return jsonOut_({ ok: false, error: 'วิชา "' + clash.subject + '" มีครูใช้แล้ว (' + clash.name + ' @' + clash.user + ') 1 วิชา = 1 ครูเท่านั้น' });
  sh.appendRow([user, pass, name, subjects.join(', ')]);
  return jsonOut_({ ok: true, user: user, name: name, subjects: subjects });
}

function editTeacher_(d) {
  if (!isAdmin_(d)) return jsonOut_({ ok: false, error: 'เฉพาะ Admin เท่านั้น' });
  const row = Number(d.row);
  if (!row || row < 2) return jsonOut_({ ok: false, error: 'row ไม่ถูกต้อง' });
  const sh = getTeachersSheet_(null);
  if (!sh || row > sh.getLastRow()) return jsonOut_({ ok: false, error: 'ไม่พบบัญชีนี้' });
  const user = String(d.user || d.username || '').trim();
  const pass = String(d.pass || d.password || '').trim();
  const name = String(d.name || d.teacherName || '').trim();
  const subjects = parseClasses_(d.subjects);
  if (!user || !pass || !name) return jsonOut_({ ok: false, error: 'กรุณากรอก USER / PASS / ชื่อครู ให้ครบ' });
  if (user === CONFIG.ADMIN_USER) return jsonOut_({ ok: false, error: 'USER นี้ซ้ำกับ Admin' });
  if (!subjects.length) return jsonOut_({ ok: false, error: 'กรุณาระบุวิชาที่สอนอย่างน้อย 1 วิชา' });
  const n = sh.getLastRow() - 1;
  const col = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < n; i++) {
    if ((i + 2) !== row && String(col[i][0] || '').trim() === user) {
      return jsonOut_({ ok: false, error: 'มี USER นี้แล้ว (แถว ' + (i + 2) + ')' });
    }
  }
  const clash = teacherSubjectClash_(sh, subjects, row);
  if (clash) return jsonOut_({ ok: false, error: 'วิชา "' + clash.subject + '" มีครูใช้แล้ว (' + clash.name + ' @' + clash.user + ') 1 วิชา = 1 ครูเท่านั้น' });
  sh.getRange(row, 1, 1, 4).setValues([[user, pass, name, subjects.join(', ')]]);
  return jsonOut_({ ok: true, row: row, user: user, name: name, subjects: subjects });
}

function deleteTeacher_(d) {
  if (!isAdmin_(d)) return jsonOut_({ ok: false, error: 'เฉพาะ Admin เท่านั้น' });
  const row = Number(d.row);
  if (!row || row < 2) return jsonOut_({ ok: false, error: 'row ไม่ถูกต้อง' });
  const sh = getTeachersSheet_(null);
  if (!sh || row > sh.getLastRow()) return jsonOut_({ ok: true, alreadyDeleted: true, row: row });
  sh.deleteRow(row);
  return jsonOut_({ ok: true, row: row });
}

/** วิชาซ้ำข้ามบัญชีหรือไม่ (1 วิชา = 1 ครู) — คืน {subject,user,name} ที่ชน หรือ null */
function teacherSubjectClash_(sh, subjects, excludeRow) {
  if (!sh || sh.getLastRow() < 2 || !subjects || !subjects.length) return null;
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  for (let i = 0; i < vals.length; i++) {
    if ((i + 2) === Number(excludeRow)) continue;
    const user = String(vals[i][0] || '').trim();
    if (!user) continue;
    const name = String(vals[i][2] || '').trim() || user;
    const has = parseClasses_(vals[i][3]);
    for (let k = 0; k < subjects.length; k++) {
      if (has.indexOf(subjects[k]) >= 0) return { subject: subjects[k], user: user, name: name };
    }
  }
  return null;
}

function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

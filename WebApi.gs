/**
 * WebApi.gs — เพิ่มไฟล์นี้ในโปรเจกต์ Apps Script เดิม (ไฟล์เดียวกับ Code.gs) ห้ามแก้ Code.gs
 *
 * ทำหน้าที่ให้หน้าเว็บบน GitHub Pages คุยกับ Google Sheets ชุดเดิมผ่าน doPost
 * โดยเรียกใช้ฟังก์ชันที่มีอยู่แล้วใน Code.gs (adminLogin, getStudents, addTransaction ฯลฯ)
 * ข้อมูลและชีตเดิมทั้งหมดยังอยู่ครบ ไม่มีการลบหรือย้าย
 *
 * ฝากเงินผ่าน QR (ไม่ต้องเข้าสู่ระบบ)
 * - QR ของนักเรียนแต่ละคนมี "กุญแจ" สุ่ม 24 ตัวอักษร เก็บใน Script Properties (QR_KEYS)
 *   ใครไม่มี QR ใบนั้นจะเดาลิงก์ฝากเงินของนักเรียนไม่ได้
 * - หน้า QR ทำได้อย่างเดียวคือ "ฝาก" ถอนไม่ได้ และฝากได้ไม่เกิน QR_MAX_AMOUNT บาทต่อครั้ง
 * - ทุกรายการจาก QR บันทึกผู้บันทึกว่า "สแกน QR" ตรวจย้อนหลังได้
 * - ถ้า QR ใบไหนหาย ให้กด "สร้าง QR ใหม่" ในหน้าครู กุญแจเดิมจะใช้ไม่ได้ทันที
 *
 * หลังวางไฟล์นี้: Deploy > Manage deployments > แก้ไข (รูปดินสอ) > Version: New version > Deploy
 * URL /exec เดิมใช้ต่อได้เหมือนเดิม
 */

const WEBAPI = {
  QR_MAX_AMOUNT: 500,
  QR_KEYS_PROP: 'QR_KEYS',
  QR_RECORDER: 'สแกน QR',
  YEAR_KEY: 'ปีการศึกษา'
};

function doPost(e) {
  let req = {};
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return webJson_({ ok: false, error: 'คำขอไม่ถูกต้อง' });
  }
  try {
    return webJson_({ ok: true, data: webRoute_(req) });
  } catch (err) {
    const msg = String((err && err.message) || err);
    const auth = /เข้าสู่ระบบใหม่|ไม่มีสิทธิ์/.test(msg);
    return webJson_({ ok: false, error: auth ? 'AUTH' : msg });
  }
}

function webRoute_(req) {
  const t = req.token;
  switch (req.action) {
    /* ---- ไม่ต้องเข้าสู่ระบบ ---- */
    case 'login': {
      const r = adminLogin(req.password);
      if (!r.success) throw new Error(r.message || 'รหัสผ่านไม่ถูกต้อง');
      return { token: r.token };
    }
    case 'qrInfo': return webQrInfo_(req.key);
    case 'qrDeposit': return webQrDeposit_(req.key, req.amount, req.note);

    /* ---- ต้องเข้าสู่ระบบ ---- */
    case 'getAll': checkAdmin_(t); return webGetAll_();
    case 'addTransaction': {
      const tx = req.tx || {};
      const r = addTransaction({
        studentId: tx.studentId, type: tx.type, amount: tx.amount,
        transactionDate: tx.date, recorder: tx.by, note: tx.note
      }, t);
      return { balance: r.balance };
    }
    case 'addStudent': {
      const s = req.student || {};
      const r = addStudent({ code: s.code, name: s.name, number: s.number, goal: s.goal, note: s.note }, t);
      return { id: r.id };
    }
    case 'updateStudent': {
      const s = req.student || {};
      updateStudent({ id: s.id, code: s.code, name: s.name, number: s.number, goal: s.goal, note: s.note }, t);
      return true;
    }
    case 'deleteStudent': deleteStudent(req.id, t); return true;
    case 'saveSettings': checkAdmin_(t); return webSaveSettings_(req.settings || {});
    case 'uploadLogo': {
      uploadSchoolLogo(req.dataUrl, 'school-logo', t);
      return { url: getSystemInfo().logoDataUrl || req.dataUrl };
    }
    case 'removeLogo': removeSchoolLogo(t); return true;
    case 'changePassword': changeAdminPassword(req.oldPassword, req.newPassword, t); return true;
    case 'syncSheets': syncStudentSheets(t); return true;
    case 'getQrKey': checkAdmin_(t); return { key: webQrKeyFor_(req.id, !!req.renew) };
    case 'getQrKeys': {
      checkAdmin_(t);
      const out = {};
      getStudents().forEach(s => { out[s.id] = webQrKeyFor_(s.id, false); });
      return out;
    }
  }
  throw new Error('ไม่รู้จักคำสั่ง ' + req.action);
}

/* -------------------- ข้อมูลทั้งหมดสำหรับหน้าครู -------------------- */

function webGetAll_() {
  const students = getStudents();
  const byCodeName = {};
  students.forEach(s => { byCodeName[s.code + '|' + s.name] = s.id; });

  const sh = getOrCreateSheet_(getSpreadsheet(), CONFIG.TRANSACTION_SHEET);
  const values = sh.getDataRange().getValues().slice(1);
  const transactions = values.filter(r => r[0] || r[1]).map(r => {
    const code = String(r[2] || '').trim();
    const name = String(r[3] || '').trim();
    return {
      id: String(r[0] || ''),
      date: toIsoString_(r[1]),
      studentId: String(r[9] || '').trim() || byCodeName[code + '|' + name] || ('code:' + code),
      type: String(r[4] || ''),
      amount: Number(r[5] || 0),
      by: String(r[7] || ''),
      note: String(r[8] || '')
    };
  });

  const info = getSystemInfo();
  const map = getSettingsMap_();
  return {
    students: students.map(s => ({ id: s.id, code: s.code, number: s.number, name: s.name, goal: s.goal, note: s.note })),
    transactions: transactions,
    settings: {
      school: info.school, system: info.system, className: info.className,
      teacher: info.homeroomTeacher, teacherPos: info.homeroomPosition,
      academic: info.academicTeacher, academicPos: info.academicPosition,
      director: info.directorName, directorPos: info.directorPosition,
      year: map[WEBAPI.YEAR_KEY] || '',
      logo: info.logoDataUrl || info.logoUrl || ''
    },
    qrMax: WEBAPI.QR_MAX_AMOUNT
  };
}

/* upsertSettings_ ใน Code.gs ไม่เขียนทับค่าที่มีอยู่แล้ว จึงใช้ฟังก์ชันนี้แทนเมื่อกดบันทึกจากหน้าเว็บ */
function webSaveSettings_(s) {
  const pairs = {
    'ชื่อโรงเรียน': s.school, 'ชื่อระบบ': s.system, 'ชั้นเรียน': s.className,
    'ชื่อครูประจำชั้น': s.teacher, 'ตำแหน่งครูประจำชั้น': s.teacherPos,
    'ชื่อครูวิชาการ': s.academic, 'ตำแหน่งครูวิชาการ': s.academicPos,
    'ชื่อผู้อำนวยการ': s.director, 'ตำแหน่งผู้อำนวยการ': s.directorPos
  };
  pairs[WEBAPI.YEAR_KEY] = s.year;
  const sh = getOrCreateSheet_(getSpreadsheet(), CONFIG.SETTINGS_SHEET);
  ensureHeader_(sh, SETTINGS_HEADERS);
  const vals = sh.getDataRange().getValues();
  const rowOf = {};
  vals.forEach((r, i) => { if (i > 0 && r[0] !== '' && r[0] != null) rowOf[String(r[0]).trim()] = i + 1; });
  Object.keys(pairs).forEach(k => {
    const v = pairs[k] == null ? '' : String(pairs[k]);
    if (rowOf[k]) sh.getRange(rowOf[k], 2).setValue(v);
    else sh.appendRow([k, v]);
  });
  return true;
}

/* -------------------- QR ฝากเงิน -------------------- */

function webQrKeys_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(WEBAPI.QR_KEYS_PROP) || '{}'); }
  catch (err) { return {}; }
}

function webQrKeyFor_(studentId, renew) {
  studentId = String(studentId || '');
  if (!getStudents().some(s => s.id === studentId)) throw new Error('ไม่พบนักเรียน');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const keys = webQrKeys_();
    if (!keys[studentId] || renew) {
      keys[studentId] = Utilities.getUuid().replace(/-/g, '').substring(0, 24);
      PropertiesService.getScriptProperties().setProperty(WEBAPI.QR_KEYS_PROP, JSON.stringify(keys));
    }
    return keys[studentId];
  } finally {
    lock.releaseLock();
  }
}

function webStudentByKey_(key) {
  key = String(key || '');
  if (!/^[a-f0-9]{24}$/.test(key)) throw new Error('QR นี้ไม่ถูกต้อง');
  const keys = webQrKeys_();
  const id = Object.keys(keys).find(k => keys[k] === key);
  const s = id && getStudents().find(x => x.id === id);
  if (!s) throw new Error('QR นี้ใช้ไม่ได้แล้ว กรุณาขอ QR ใบใหม่จากครู');
  return s;
}

function webQrInfo_(key) {
  const s = webStudentByKey_(key);
  const info = getSystemInfo();
  return {
    name: s.name, code: s.code, number: s.number, goal: s.goal,
    balance: getBalance(s.code, s.id),
    school: info.school, className: info.className, max: WEBAPI.QR_MAX_AMOUNT
  };
}

function webQrDeposit_(key, amount, note) {
  const s = webStudentByKey_(key);
  amount = Number(amount);
  if (!Number.isFinite(amount) || amount <= 0) throw new Error('จำนวนเงินต้องมากกว่า 0 บาท');
  if (amount > WEBAPI.QR_MAX_AMOUNT) throw new Error('ฝากผ่าน QR ได้ไม่เกิน ' + WEBAPI.QR_MAX_AMOUNT + ' บาทต่อครั้ง');
  if (Math.round(amount * 100) !== amount * 100) throw new Error('จำนวนเงินมีทศนิยมได้ไม่เกิน 2 ตำแหน่ง');

  const cache = CacheService.getScriptCache();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // กันกดซ้ำ: ยอดเดียวกันของคนเดียวกันภายใน 8 วินาทีจะไม่ถูกบันทึกซ้ำ
    const dupKey = 'QRDUP_' + s.id + '_' + amount;
    if (cache.get(dupKey)) throw new Error('เพิ่งบันทึกยอดนี้ไปเมื่อสักครู่ ถ้าต้องการฝากซ้ำให้รอสักครู่แล้วลองใหม่');

    const current = getBalance(s.code, s.id);
    const newBalance = roundMoney_(current + amount);
    const sh = getOrCreateSheet_(getSpreadsheet(), CONFIG.TRANSACTION_SHEET);
    ensureTransactionHeaders_(sh);
    sh.appendRow([
      Utilities.getUuid(), new Date(), s.code, s.name, 'ฝาก',
      amount, newBalance, WEBAPI.QR_RECORDER, String(note || '').substring(0, 100), s.id
    ]);
    const row = sh.getLastRow();
    sh.getRange(row, 2).setNumberFormat('dd/MM/yyyy HH:mm');
    sh.getRange(row, 6, 1, 2).setNumberFormat('#,##0.00');
    cache.put(dupKey, '1', 8);
    rebuildStudentSheet_(s);
    return { name: s.name, amount: amount, balance: newBalance };
  } finally {
    lock.releaseLock();
  }
}

function webJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

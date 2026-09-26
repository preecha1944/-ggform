/**
 * ระบบสอบออนไลน์จับเวลา (Google Apps Script Web App)
 * - ผู้สอบกรอกชื่อ-สกุล และรหัสนักศึกษา แล้วเริ่มทำข้อสอบ
 * - จับเวลาตามที่กำหนด หมดเวลาแล้วส่งคำตอบอัตโนมัติ ตอบได้กี่ข้อก็คิดคะแนนเท่านั้น
 * - เวลาเริ่มและกำหนดส่งเก็บฝั่งเซิร์ฟเวอร์ รีเฟรชหน้าก็ไม่ได้เวลาเพิ่ม และรหัสหนึ่งสอบได้ครั้งเดียว
 * - ออกจากหน้าสอบ (สลับแท็บ/แอป, ปิดแล้วเปิดใหม่) ครั้งแรกเตือน เกินกำหนดส่งคำตอบทันทีและตัดสิทธิ์
 * - ผลสอบบันทึกลง Google Sheet ผู้สอบเห็นคะแนนและเฉลยทันทีหลังส่ง
 *
 * ติดตั้ง: ดู README.md
 */

const EXAM = {
  title: 'แบบทดสอบความรู้ก่อนเรียน',
  subtitle: 'รายวิชา 220104 จริยธรรมสำหรับผู้บริหารการศึกษา (Professional Ethics for Educational Administrators)',
  durationMinutes: 30,
  // เผื่อเวลาให้คำตอบที่ส่งอัตโนมัติตอนหมดเวลาเดินทางถึงเซิร์ฟเวอร์ (เน็ตช้า)
  graceSeconds: 60,
  showAnswerKey: true,
  // จำนวนครั้งที่ออกจากหน้าสอบได้โดยแค่ถูกเตือน ออกครั้งถัดไปจะถูกส่งคำตอบทันที (0 = เด้งออกตั้งแต่ครั้งแรก)
  allowedLeaves: 1,
  // รหัสนักศึกษา: ตัวเลขอย่างน้อย 5 หลัก (อนุญาต - คั่นได้)
  studentIdPattern: '^[0-9][0-9-]{4,}$',
};

const LETTERS = ['ก', 'ข', 'ค', 'ง', 'จ'];
const SHEET_NAME = 'ผลสอบ';
const COL = { id: 1, name: 2, start: 3, deadline: 4, submitted: 5, status: 6, score: 7, leaves: 8, leaveLog: 9, firstAnswer: 10 };
const SAVED_COL = COL.firstAnswer + QUESTIONS.length;
const STATUS = {
  active: 'กำลังทำ',
  done: 'ส่งแล้ว',
  timeout: 'หมดเวลา (ส่งอัตโนมัติ)',
  kicked: 'ตัดสิทธิ์ (ออกจากหน้าสอบเกินกำหนด)',
};

/** รันครั้งเดียวก่อน Deploy: สร้าง Google Sheet สำหรับเก็บผลสอบ */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SHEET_ID')) {
    Logger.log('มีชีตผลสอบอยู่แล้ว: ' + SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getUrl());
    return;
  }
  const ss = SpreadsheetApp.create('ผลสอบ - ' + EXAM.title + ' 220104');
  const sheet = ss.getSheets()[0].setName(SHEET_NAME);
  const headers = ['รหัสนักศึกษา', 'ชื่อ-สกุล', 'เวลาเริ่ม', 'กำหนดส่ง', 'เวลาส่ง', 'สถานะ', 'คะแนน (เต็ม ' + QUESTIONS.length + ')',
    'ออกจากหน้าสอบ (ครั้ง)', 'บันทึกการออกจากหน้าสอบ'];
  QUESTIONS.forEach(function (_, i) { headers.push('ข้อ ' + (i + 1)); });
  headers.push('คำตอบที่บันทึกระหว่างสอบ');
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.getRange(1, COL.start, sheet.getMaxRows(), 3).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  props.setProperty('SHEET_ID', ss.getId());
  Logger.log('สร้างชีตผลสอบแล้ว: ' + ss.getUrl());
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(EXAM.title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getExamInfo() {
  return {
    title: EXAM.title,
    subtitle: EXAM.subtitle,
    durationMinutes: EXAM.durationMinutes,
    total: QUESTIONS.length,
    studentIdPattern: EXAM.studentIdPattern,
    allowedLeaves: EXAM.allowedLeaves,
  };
}

/**
 * เริ่มสอบ หรือกลับเข้ามาทำต่อ (เวลายังคงนับจากตอนเริ่มครั้งแรก)
 * การกลับเข้ามาใหม่ (ปิดแท็บ/รีเฟรช) นับเป็นการออกจากหน้าสอบหนึ่งครั้ง
 */
function startExam(name, studentId) {
  name = String(name || '').trim();
  studentId = String(studentId || '').trim();
  if (!name) throw new Error('กรุณากรอกชื่อ-สกุล');
  if (!new RegExp(EXAM.studentIdPattern).test(studentId)) throw new Error('รหัสนักศึกษาไม่ถูกต้อง');

  return withLock_(function (sheet) {
    const now = Date.now();
    let row = findRow_(sheet, studentId);
    const isResume = row > 0;
    if (!row) {
      const deadline = now + EXAM.durationMinutes * 60 * 1000;
      const values = [studentId, name, new Date(now), new Date(deadline), '', STATUS.active, '', 0, ''];
      QUESTIONS.forEach(function () { values.push(''); });
      values.push('[]');
      sheet.appendRow(values);
      row = sheet.getLastRow();
    }
    const rec = readRow_(sheet, row);
    if (rec.status !== STATUS.active) return { state: 'finished', result: buildResult_(rec) };
    if (now > rec.deadline + EXAM.graceSeconds * 1000) {
      return { state: 'finished', result: finalize_(sheet, row, rec, rec.saved, STATUS.timeout) };
    }
    if (isResume) {
      const leaves = addLeave_(sheet, row, rec, 'เปิดหน้าสอบใหม่');
      if (leaves > EXAM.allowedLeaves) {
        return { state: 'finished', result: finalize_(sheet, row, rec, rec.saved, STATUS.kicked) };
      }
    }
    return {
      state: 'active',
      name: rec.name,
      studentId: rec.studentId,
      remainingMs: Math.max(0, rec.deadline - now),
      leaves: rec.leaves,
      allowedLeaves: EXAM.allowedLeaves,
      saved: rec.saved,
      questions: QUESTIONS.map(function (item) { return { q: item.q, choices: item.choices }; }),
    };
  });
}

/** บันทึกคำตอบระหว่างสอบ (รับเฉพาะก่อนหมดเวลา) */
function saveAnswers(studentId, answers) {
  return withLock_(function (sheet) {
    const row = findRow_(sheet, String(studentId).trim());
    if (!row) throw new Error('ไม่พบข้อมูลผู้สอบ');
    const rec = readRow_(sheet, row);
    if (rec.status !== STATUS.active || Date.now() > rec.deadline + EXAM.graceSeconds * 1000) return { ok: false };
    sheet.getRange(row, SAVED_COL).setValue(JSON.stringify(normalize_(answers)));
    return { ok: true };
  });
}

/**
 * ผู้สอบออกจากหน้าสอบ (สลับแท็บ/แอป/ย่อหน้าต่าง) ระหว่างสอบ
 * ถ้าเกินจำนวนที่อนุญาต ส่งคำตอบที่มีอยู่ทันทีและตัดสิทธิ์
 */
function reportLeave(studentId, answers) {
  return withLock_(function (sheet) {
    const row = findRow_(sheet, String(studentId).trim());
    if (!row) throw new Error('ไม่พบข้อมูลผู้สอบ');
    const rec = readRow_(sheet, row);
    if (rec.status !== STATUS.active) return { leaves: rec.leaves, result: buildResult_(rec) };
    const leaves = addLeave_(sheet, row, rec, 'ออกจากหน้าสอบ');
    if (leaves > EXAM.allowedLeaves) {
      const inTime = Date.now() <= rec.deadline + EXAM.graceSeconds * 1000;
      return { leaves: leaves, result: finalize_(sheet, row, rec, inTime ? normalize_(answers) : rec.saved, STATUS.kicked) };
    }
    if (Date.now() <= rec.deadline + EXAM.graceSeconds * 1000) {
      sheet.getRange(row, SAVED_COL).setValue(JSON.stringify(normalize_(answers)));
    }
    return { leaves: leaves, result: null };
  });
}

/** ผู้สอบกลับมาที่หน้าสอบ: บันทึกว่าหายไปนานเท่าไร */
function reportReturn(studentId, awaySeconds) {
  return withLock_(function (sheet) {
    const row = findRow_(sheet, String(studentId).trim());
    if (!row) return;
    appendLog_(sheet, row, 'กลับมา (หายไป ' + Math.round(Number(awaySeconds) || 0) + ' วินาที)');
  });
}

/** ส่งคำตอบ (กดส่งเอง หรือส่งอัตโนมัติเมื่อหมดเวลา) */
function submitExam(studentId, answers, isAuto) {
  return withLock_(function (sheet) {
    const row = findRow_(sheet, String(studentId).trim());
    if (!row) throw new Error('ไม่พบข้อมูลผู้สอบ');
    const rec = readRow_(sheet, row);
    if (rec.status !== STATUS.active) return buildResult_(rec);
    const inTime = Date.now() <= rec.deadline + EXAM.graceSeconds * 1000;
    // ส่งช้าเกินเวลาเผื่อ: ใช้เฉพาะคำตอบที่บันทึกไว้ก่อนหมดเวลา
    const final = inTime ? normalize_(answers) : rec.saved;
    return finalize_(sheet, row, rec, final, isAuto || !inTime ? STATUS.timeout : STATUS.done);
  });
}

function finalize_(sheet, row, rec, answers, status) {
  let score = 0;
  const letters = answers.map(function (a, i) {
    if (a === null) return '';
    if (a === LETTERS.indexOf(QUESTIONS[i].answer)) score++;
    return LETTERS[a];
  });
  sheet.getRange(row, COL.submitted, 1, 3).setValues([[new Date(), status, score]]);
  sheet.getRange(row, COL.firstAnswer, 1, letters.length).setValues([letters]);
  sheet.getRange(row, SAVED_COL).setValue(JSON.stringify(answers));
  rec.status = status;
  rec.saved = answers;
  return buildResult_(rec);
}

function addLeave_(sheet, row, rec, what) {
  rec.leaves += 1;
  sheet.getRange(row, COL.leaves).setValue(rec.leaves);
  appendLog_(sheet, row, what + ' ครั้งที่ ' + rec.leaves);
  return rec.leaves;
}

function appendLog_(sheet, row, text) {
  const cell = sheet.getRange(row, COL.leaveLog);
  const time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss');
  const prev = String(cell.getValue() || '');
  cell.setValue((prev ? prev + '\n' : '') + time + ' ' + text);
}

function buildResult_(rec) {
  const answers = rec.saved;
  let score = 0;
  const review = QUESTIONS.map(function (item, i) {
    const correct = LETTERS.indexOf(item.answer);
    if (answers[i] === correct) score++;
    return { q: item.q, choices: item.choices, chosen: answers[i], correct: correct };
  });
  return {
    name: rec.name,
    studentId: rec.studentId,
    status: rec.status,
    score: score,
    total: QUESTIONS.length,
    answered: answers.filter(function (a) { return a !== null; }).length,
    leaves: rec.leaves,
    review: EXAM.showAnswerKey ? review : null,
  };
}

function normalize_(answers) {
  return QUESTIONS.map(function (_, i) {
    const a = answers && answers[i];
    return typeof a === 'number' && a >= 0 && a < LETTERS.length && a % 1 === 0 ? a : null;
  });
}

function readRow_(sheet, row) {
  const v = sheet.getRange(row, 1, 1, SAVED_COL).getValues()[0];
  let saved = [];
  try { saved = JSON.parse(v[SAVED_COL - 1] || '[]'); } catch (e) { saved = []; }
  return {
    studentId: String(v[COL.id - 1]),
    name: String(v[COL.name - 1]),
    deadline: new Date(v[COL.deadline - 1]).getTime(),
    status: String(v[COL.status - 1]),
    leaves: Number(v[COL.leaves - 1]) || 0,
    saved: normalize_(saved),
  };
}

function findRow_(sheet, studentId) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const ids = sheet.getRange(2, COL.id, last - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === studentId) return i + 2;
  }
  return 0;
}

function withLock_(fn) {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('ระบบยังไม่ได้ตั้งค่า (ผู้สอนต้องรันฟังก์ชัน setup ก่อน)');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn(SpreadsheetApp.openById(id).getSheetByName(SHEET_NAME));
  } finally {
    lock.releaseLock();
  }
}

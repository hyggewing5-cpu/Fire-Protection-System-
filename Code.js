/**
 * ระบบตรวจป้องกันอัคคีภัยก่อนปิดสำนักงาน
 * Google Apps Script (Container-bound กับ Google Sheet)
 *
 * ชีตที่ใช้:
 *   Config          A1 = ชื่อหน่วยงาน, A3 = รหัสผ่านผู้ดูแล (จำเป็นสำหรับเพิ่ม/ลบห้องและตั้งค่ารายงาน)
 *                   A2 = URL โลโก้แบบเดิม (ใช้เป็นค่าสำรอง; เมื่อบันทึกการตั้งค่าในหน้าเว็บ ระบบจะย้ายโลโก้ไปเก็บใน ReportSettings)
 *   Rooms           รหัสห้อง | ชื่อห้อง | ผู้รับผิดชอบ
 *   ChecklistLogs   วันที่-เวลา | รหัสห้อง | ชื่อห้อง | ผู้ตรวจสอบ | ข้อ 1..19 | หมายเหตุ
 *   ReportSettings  key | value | คำอธิบาย  (สร้างอัตโนมัติ เก็บหัวข้อรายงาน โลโก้ ผู้ลงชื่อ ฯลฯ)
 *   Staff           รหัสเจ้าหน้าที่ | ชื่อ-สกุล | รหัสห้อง  (หนึ่งคนต่อหนึ่งแถว สร้างอัตโนมัติ และย้ายรายชื่อจากคอลัมน์ "ผู้รับผิดชอบ" ของ Rooms ให้ครั้งแรก)
 *   ChecklistItems  ลำดับ | ข้อความรายการตรวจ  (สร้างอัตโนมัติ แก้ไขได้จากหน้าเว็บ แท็บ "ตั้งค่า > รายการตรวจ")
 *
 * โครงสร้างคอลัมน์ ChecklistLogs (เข้ากันได้กับข้อมูลเดิม 24 คอลัมน์):
 *   1-4 = วันที่-เวลา, รหัสห้อง, ชื่อห้อง, ผู้ตรวจสอบ | 5-23 = ข้อ 1-19 | 24 = หมายเหตุ | 25-35 = ข้อ 20-30
 */

var SHEET_ROOMS = 'Rooms';
var SHEET_LOGS = 'ChecklistLogs';
var SHEET_CONFIG = 'Config';
var SHEET_SETTINGS = 'ReportSettings';
var SHEET_ITEMS = 'ChecklistItems';
var SHEET_STAFF = 'Staff';
var STAFF_NAME_MAX = 100;
var BASE_ITEMS = 19;                 // ข้อ 1-19 อยู่คอลัมน์ 5-23 เหมือนเดิม
var MAX_ITEMS = 30;                  // จำนวนรายการตรวจสูงสุดที่เพิ่มได้ (ข้อ 20-30 อยู่คอลัมน์ 25-35)
var NOTE_COL = 4 + BASE_ITEMS + 1;   // 24 คอลัมน์หมายเหตุ (คงตำแหน่งเดิม)
var LOG_COLS = 4 + MAX_ITEMS + 1;    // 35
var ITEM_MAX_LEN = 300;
var RESULT_OK = 'ปกติ';
var RESULT_BAD = 'ไม่ปกติ';
var RESULT_NA = 'ไม่มี';
var LOGO_MAX_LEN = 45000; // ช่องเซลล์ใน Google Sheet เก็บได้ไม่เกิน 50,000 ตัวอักษร

var DEFAULT_ITEMS = [
  "1. ถอดปลั๊กเครื่องใช้ไฟฟ้าทุกชนิด (ยกเว้นตู้เย็นและตู้แช่ที่ติดตั้งสวิตช์ตัดตอนไฟฟ้าอัตโนมัติเป็นการเฉพาะ)",
  "2. ปิดสวิตช์เครื่องปรับอากาศ (ถ้ามี)",
  "3. ดับธูปเทียนบูชาพระ (ถ้ามี)",
  "4. กำจัดขยะ น้ำมันเชื้อเพลิง และสิ่งที่เป็นเชื้อเพลิง",
  "5. ปิดวาล์วถังก๊าซหุงต้ม (ถ้ามี)",
  "6. ปิดก๊อกน้ำประปาเรียบร้อย ไม่มีน้ำรั่วไหล",
  "7. ตรวจสอบสภาพสายไฟฟ้าไม่ชำรุดและมีฉนวนหุ้มเรียบร้อย",
  "8. เครื่องใช้ไฟฟ้าประเภทความร้อน เช่น กระติกต้มน้ำร้อนไฟฟ้า ต้องมีฉนวนกันความร้อนหรือวัสดุทนไฟรองรับ",
  "9. สายไฟฟ้าไม่แตะพาดกับรางน้ำหรือโลหะ",
  "10. ตรวจสอบสภาพไฟฟ้า จุดต่อสายไฟฟ้า ภายในอาคารไม่มีรอยไหม้เกรียมหรือสภาพผิดปกติ",
  "11. ไม่มีการใช้สายไฟฟ้าเปลือยเสียบแทนปลั๊กไฟฟ้า",
  "12. ไม่มีการใช้ลวดทองแดงแทนฟิวส์ตะกั่ว",
  "13. ไม่วางพัสดุและสิ่งของต่าง ๆ ปิดบัง พิง พาด ทับอุปกรณ์ไฟฟ้าและอุปกรณ์ดับเพลิงประจำอาคาร",
  "14. ไม่มีน้ำรั่วไหลลงมาสัมผัสอุปกรณ์ไฟฟ้าในอาคาร",
  "15. ไม่มีการวางตั้งพัสดุสิ่งของใด ๆ กีดขวางทางหนีไฟ",
  "16. อุปกรณ์และเครื่องดับเพลิงมีสภาพพร้อมใช้งาน",
  "17. ปิดสวิตช์ไฟฟ้าแสงสว่างทุกจุดเรียบร้อย",
  "18. ปิดและล็อกหน้าต่าง ประตู เรียบร้อย",
  "19. ปลดสวิตช์ตัดตอนกระแสไฟฟ้าประจำอาคารลง (เฉพาะบางอาคาร)"
];

// ค่าเริ่มต้นของหน้ารายงาน (แก้ได้จากหน้าเว็บ แท็บ "ตั้งค่า")
var DEFAULT_SETTINGS = {
  orgLogo: '',
  reportTitle: 'รายงานการตรวจอัคคีภัย (ก่อนปิดอาคารสำนักงาน)',
  reportSubtitle: 'หน่วยงาน {org} ประจำเดือน {month}',
  footerNote: '',
  signer1Label: 'ผู้สรุปรายงาน',
  signer1Name: '',
  signer1Position: 'เจ้าหน้าที่ความปลอดภัย (จป.วิชาชีพ)',
  signer2Label: '',
  signer2Name: '',
  signer2Position: '',
  showLogo: true,
  logoHeight: 60,
  headerAlign: 'center',
  tableFontSize: 14,
  orientation: 'landscape',
  showLegend: true,
  showSignature: true
};
var SETTING_KEYS = Object.keys(DEFAULT_SETTINGS);

// ---------------------------------------------------------------- Web app

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบตรวจป้องกันอัคคีภัยก่อนปิดสำนักงาน')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------------------------------------------------------------- API สำหรับหน้าเว็บที่โฮสต์นอก Apps Script
// หน้าเว็บภายนอก (เช่น Netlify) ส่ง POST แบบ text/plain เป็น JSON {"fn":"ชื่อฟังก์ชัน","args":[...]}
// ตอบกลับ {"ok":true,"data":...} หรือ {"ok":false,"error":"..."}  (เรียกได้เฉพาะฟังก์ชันในรายการด้านล่าง)
// ต้อง Deploy เป็น Web app โดยตั้ง Execute as = Me และ Who has access = Anyone

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'คำขอไม่ถูกต้อง' });
  }
  var fn = String((req && req.fn) || '');
  var args = Object.prototype.toString.call(req && req.args) === '[object Array]' ? req.args : [];

  var table = {
    getAppData: getAppData,
    saveChecklist: saveChecklist,
    saveChecklistItems: saveChecklistItems,
    saveReportSettings: saveReportSettings,
    checkAdmin: checkAdmin,
    manageRoom: manageRoom,
    manageStaff: manageStaff,
    getDailyUncheckedReport: getDailyUncheckedReport,
    getMonthlySummary: getMonthlySummary
  };
  if (!Object.prototype.hasOwnProperty.call(table, fn)) {
    return json_({ ok: false, error: 'ไม่รู้จักคำสั่ง: ' + fn });
  }
  try {
    return json_({ ok: true, data: table[fn].apply(null, args) });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

// ---------------------------------------------------------------- Helpers

function tz_() {
  return Session.getScriptTimeZone() || 'Asia/Bangkok';
}

function getSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function getConfigSheet_() {
  var sheet = getSheet_(SHEET_CONFIG);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 3, 2).setValues([
      ['สำนักงานของท่าน', 'ชื่อหน่วยงาน'],
      ['', 'URL โลโก้ (ไม่บังคับ ตั้งค่าได้จากหน้าเว็บ)'],
      ['', 'รหัสผ่านผู้ดูแล (ต้องตั้งก่อนจัดการห้อง/ตั้งค่ารายงาน)']
    ]);
  }
  return sheet;
}

function getRoomsSheet_() {
  return getSheet_(SHEET_ROOMS, ['รหัสห้อง', 'ชื่อห้อง', 'ผู้รับผิดชอบ']);
}

function logHeaders_() {
  var h = ['วันที่-เวลา', 'รหัสห้อง', 'ชื่อห้อง', 'ผู้ตรวจสอบ'];
  for (var i = 1; i <= BASE_ITEMS; i++) h.push('ข้อ ' + i);
  h.push('หมายเหตุ');
  for (var j = BASE_ITEMS + 1; j <= MAX_ITEMS; j++) h.push('ข้อ ' + j);
  return h;
}

// ตำแหน่งคอลัมน์ (เริ่มที่ 0) ของรายการตรวจลำดับที่ i (เริ่มที่ 0)
function itemColIdx_(i) {
  return i < BASE_ITEMS ? 4 + i : NOTE_COL + (i - BASE_ITEMS);
}

function itemResults_(row, count) {
  var out = [];
  for (var i = 0; i < count; i++) {
    var v = row[itemColIdx_(i)];
    out.push(v === undefined || v === null ? '' : String(v));
  }
  return out;
}

// สร้างชีตบันทึกผล และแก้หัวตารางให้ครบ 24 คอลัมน์เสมอ
function getLogSheet_() {
  var sheet = getSheet_(SHEET_LOGS);
  if (sheet.getMaxColumns() < LOG_COLS) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), LOG_COLS - sheet.getMaxColumns());
  }
  if (sheet.getRange(1, LOG_COLS).getValue() !== 'ข้อ ' + MAX_ITEMS) {
    sheet.getRange(1, 1, 1, LOG_COLS).setValues([logHeaders_()]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSettingsSheet_() {
  var existed = !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  var sheet = getSheet_(SHEET_SETTINGS, ['key', 'value', 'คำอธิบาย']);
  if (!existed) {
    sheet.getRange('B:B').setNumberFormat('@'); // เก็บเป็นข้อความ กันชีตแปลงตัวเลข/ค่าจริงเอง
  }
  return sheet;
}

// อ่านชีต Rooms ตามที่เก็บไว้ (คอลัมน์ C เป็นข้อมูลแบบเดิม ใช้เฉพาะตอนย้ายไปชีต Staff)
function readRoomsRaw_() {
  var sheet = getRoomsSheet_();
  if (sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var rooms = [];
  values.forEach(function (r) {
    if (r[0] !== '' && r[0] !== null) {
      rooms.push({ id: String(r[0]), name: String(r[1]), owner: String(r[2] || '') });
    }
  });
  return rooms;
}

function splitNames_(s) {
  return String(s || '').split(/[,;\n]+/).map(function (x) { return x.trim(); }).filter(Boolean);
}

// ชีตเจ้าหน้าที่: สร้างครั้งแรกพร้อมย้ายรายชื่อเดิมจากคอลัมน์ "ผู้รับผิดชอบ" ของชีต Rooms (คั่นด้วย , ; หรือขึ้นบรรทัดใหม่)
function getStaffSheet_() {
  var existed = !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STAFF);
  var sheet = getSheet_(SHEET_STAFF, ['รหัสเจ้าหน้าที่', 'ชื่อ-สกุล', 'รหัสห้อง']);
  if (!existed) {
    var rows = [];
    readRoomsRaw_().forEach(function (r) {
      splitNames_(r.owner).forEach(function (nm) {
        rows.push(['S' + (rows.length + 1), nm.slice(0, STAFF_NAME_MAX), r.id]);
      });
    });
    if (rows.length) {
      sheet.getRange(2, 2, rows.length, 1).setNumberFormat('@'); // กันชื่อที่ขึ้นต้นด้วย = ถูกตีเป็นสูตร
      sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    }
  }
  return sheet;
}

function readStaff_() {
  var sheet = getStaffSheet_();
  if (sheet.getLastRow() < 2) return [];
  var out = [];
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().forEach(function (r) {
    var id = String(r[0] || '').trim();
    var name = String(r[1] || '').trim();
    if (id && name) out.push({ id: id, name: name, roomId: String(r[2] || '').trim() });
  });
  return out;
}

// ห้อง + รายชื่อเจ้าหน้าที่ประจำห้อง (staff = รายชื่อ, owner = รายชื่อคั่นด้วย , สำหรับแสดงในรายงาน)
function readRooms_() {
  var staff = readStaff_();
  return readRoomsRaw_().map(function (r) {
    var names = staff.filter(function (s) { return s.roomId === r.id; }).map(function (s) { return s.name; });
    return { id: r.id, name: r.name, owner: names.join(', '), staff: names };
  });
}

// แปลงค่าในคอลัมน์วันที่ (Date หรือข้อความ yyyy-MM-dd HH:mm:ss) เป็น Date
function toDate_(v) {
  if (v instanceof Date) return v;
  var m = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?/.exec(String(v));
  if (m) {
    return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  }
  var d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

// อ่านบันทึกทั้งหมดเป็นโครงสร้างที่ใช้งานง่าย
function readLogs_() {
  var sheet = getLogSheet_();
  if (sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, LOG_COLS).getValues();
  var out = [];
  var itemCount = getItems_().length;
  values.forEach(function (row) {
    var d = toDate_(row[0]);
    if (!d) return;
    out.push({
      date: Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'),
      ym: Utilities.formatDate(d, tz_(), 'yyyy-MM'),
      day: parseInt(Utilities.formatDate(d, tz_(), 'd'), 10),
      roomId: String(row[1]),
      roomName: String(row[2]),
      inspector: String(row[3]),
      results: itemResults_(row, itemCount),
      note: String(row[NOTE_COL - 1] || '')
    });
  });
  return out;
}

function getAdminPassword_() {
  return String(getConfigSheet_().getRange(3, 1).getValue() || '').trim();
}

function assertAdmin_(password) {
  var real = getAdminPassword_();
  if (!real) {
    throw new Error('ยังไม่ได้ตั้งรหัสผ่านผู้ดูแล กรุณาพิมพ์รหัสผ่านที่ชีต Config เซลล์ A3 ก่อน');
  }
  if (String(password || '') !== real) {
    throw new Error('รหัสผ่านไม่ถูกต้อง');
  }
}

// ---------------------------------------------------------------- Checklist items

// ตัดเลขลำดับนำหน้า ("5. ", "5) ") ออก เพราะระบบใส่เลขให้เองตามลำดับ
function cleanItem_(s) {
  return String(s == null ? '' : s).replace(/^\s*\d+\s*[.)]\s*/, '').trim();
}

function getItemsSheet_() {
  var existed = !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ITEMS);
  var sheet = getSheet_(SHEET_ITEMS, ['ลำดับ', 'ข้อความรายการตรวจ']);
  if (!existed) {
    var rows = DEFAULT_ITEMS.map(function (t, i) { return [i + 1, cleanItem_(t)]; });
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('@');
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
  return sheet;
}

// รายการตรวจปัจจุบัน (ข้อความล้วน ไม่มีเลขนำหน้า)
function getItems_() {
  var sheet = getItemsSheet_();
  var items = [];
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().forEach(function (r) {
      var t = cleanItem_(r[0]);
      if (t) items.push(t);
    });
  }
  if (!items.length) {
    items = DEFAULT_ITEMS.map(cleanItem_);
  }
  return items.slice(0, MAX_ITEMS);
}

function numberedItems_(items) {
  return items.map(function (t, i) { return (i + 1) + '. ' + t; });
}

// ---------------------------------------------------------------- Report settings

function toBool_(v, def) {
  if (v === true || v === 'true') return true;
  if (v === false || v === 'false') return false;
  return def;
}

// แปลงค่าที่รับมา (จากชีตหรือจากหน้าเว็บ) ให้เป็นชนิดที่ถูกต้องและอยู่ในช่วงที่กำหนด
function normalizeSettings_(input) {
  input = input || {};
  var has = function (k) { return Object.prototype.hasOwnProperty.call(input, k); };
  var D = DEFAULT_SETTINGS;

  function str(k, max) {
    var v = has(k) ? input[k] : D[k];
    return String(v == null ? '' : v).trim().slice(0, max);
  }
  function num(k, min, max) {
    var n = parseInt(has(k) ? input[k] : D[k], 10);
    if (isNaN(n)) n = D[k];
    return Math.min(max, Math.max(min, n));
  }
  function oneOf(k, list) {
    var v = has(k) ? String(input[k]) : D[k];
    return list.indexOf(v) === -1 ? D[k] : v;
  }
  function bool(k) {
    return toBool_(has(k) ? input[k] : D[k], D[k]);
  }

  return {
    orgLogo: has('orgLogo') ? String(input.orgLogo || '').trim() : D.orgLogo,
    reportTitle: str('reportTitle', 200),
    reportSubtitle: str('reportSubtitle', 250),
    footerNote: str('footerNote', 500),
    signer1Label: str('signer1Label', 100),
    signer1Name: str('signer1Name', 100),
    signer1Position: str('signer1Position', 150),
    signer2Label: str('signer2Label', 100),
    signer2Name: str('signer2Name', 100),
    signer2Position: str('signer2Position', 150),
    showLogo: bool('showLogo'),
    logoHeight: num('logoHeight', 20, 200),
    headerAlign: oneOf('headerAlign', ['left', 'center', 'right']),
    tableFontSize: num('tableFontSize', 10, 20),
    orientation: oneOf('orientation', ['landscape', 'portrait']),
    showLegend: bool('showLegend'),
    showSignature: bool('showSignature')
  };
}

// โลโก้ต้องเป็น URL (http/https) หรือรูป data URL (png/jpeg/gif/webp) ที่ไม่ใหญ่เกินกำหนด; ไม่รับ SVG
function validateLogo_(logo) {
  if (!logo) return '';
  if (/^https?:\/\/\S+$/i.test(logo)) {
    if (logo.length > 2000) throw new Error('URL โลโก้ยาวเกินไป');
    return logo;
  }
  if (/^data:image\/(png|jpeg|gif|webp);base64,[A-Za-z0-9+\/=]+$/.test(logo)) {
    if (logo.length > LOGO_MAX_LEN) throw new Error('ไฟล์โลโก้ใหญ่เกินไป กรุณาใช้รูปที่เล็กลง');
    return logo;
  }
  throw new Error('รูปแบบโลโก้ไม่ถูกต้อง (รองรับ PNG, JPG, GIF, WEBP หรือลิงก์ http/https)');
}

function readSettings_() {
  var sheet = getSettingsSheet_();
  var raw = {};
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var k = String(r[0]);
      if (SETTING_KEYS.indexOf(k) !== -1) raw[k] = r[1];
    });
  }
  return normalizeSettings_(raw);
}

// ---------------------------------------------------------------- API

// ข้อมูลพื้นฐาน: หน่วยงาน, ห้อง, รายการตรวจ, การตั้งค่ารายงาน
function getAppData() {
  var cfg = getConfigSheet_();
  var orgName = String(cfg.getRange(1, 1).getValue() || '').trim() || 'สำนักงานของท่าน';
  var settings = readSettings_();
  var items = getItems_();
  var logo = settings.orgLogo || String(cfg.getRange(2, 1).getValue() || '').trim();
  settings.orgLogo = logo;
  return {
    orgName: orgName,
    orgLogo: logo,
    report: settings,
    rooms: readRooms_(),
    staff: readStaff_(),
    items: numberedItems_(items),
    itemTexts: items,
    maxItems: MAX_ITEMS,
    adminReady: getAdminPassword_() !== ''
  };
}

// บันทึกการตั้งค่ารายงาน (ชื่อหน่วยงาน โลโก้ หัวข้อ ผู้ลงชื่อ การจัดหน้า) ต้องใช้รหัสผ่านผู้ดูแล
function saveReportSettings(settings, password) {
  try {
    assertAdmin_(password);
  } catch (e) {
    return { status: 'error', message: e.message };
  }
  settings = settings || {};

  var clean = normalizeSettings_(settings);
  try {
    clean.orgLogo = validateLogo_(clean.orgLogo);
  } catch (e) {
    return { status: 'error', message: e.message };
  }
  var orgName = String(settings.orgName || '').trim().slice(0, 150);
  if (!orgName) return { status: 'error', message: 'กรุณากรอกชื่อหน่วยงาน' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var cfg = getConfigSheet_();
    cfg.getRange(1, 1).setValue(orgName);
    cfg.getRange(2, 1).setValue(''); // โลโก้ย้ายมาเก็บที่ ReportSettings แล้ว

    var sheet = getSettingsSheet_();
    var rows = SETTING_KEYS.map(function (k) { return [k, String(clean[k]), '']; });
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
    }
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('@');
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  } finally {
    lock.releaseLock();
  }
  return { status: 'success', message: 'บันทึกการตั้งค่าสำเร็จ', orgName: orgName, report: clean };
}

// บันทึกรายการตรวจ (เพิ่ม/แก้ข้อความ/ลบรายการท้ายสุด) ต้องใช้รหัสผ่านผู้ดูแล
// หมายเหตุ: ผลตรวจเก่าเก็บตามลำดับข้อ การแก้ข้อความของข้อเดิมจึงมีผลกับความหมายของข้อมูลย้อนหลังในข้อนั้น
function saveChecklistItems(items, password) {
  try {
    assertAdmin_(password);
  } catch (e) {
    return { status: 'error', message: e.message };
  }
  if (Object.prototype.toString.call(items) !== '[object Array]') {
    return { status: 'error', message: 'ข้อมูลรายการตรวจไม่ถูกต้อง' };
  }
  var clean = items.map(cleanItem_);
  if (clean.some(function (t) { return !t; })) {
    return { status: 'error', message: 'มีรายการที่ว่างอยู่ กรุณากรอกให้ครบหรือลบรายการนั้น' };
  }
  if (clean.length < 1) return { status: 'error', message: 'ต้องมีรายการตรวจอย่างน้อย 1 ข้อ' };
  if (clean.length > MAX_ITEMS) return { status: 'error', message: 'เพิ่มได้สูงสุด ' + MAX_ITEMS + ' ข้อ' };
  for (var i = 0; i < clean.length; i++) {
    if (clean[i].length > ITEM_MAX_LEN) {
      return { status: 'error', message: 'ข้อ ' + (i + 1) + ' ยาวเกิน ' + ITEM_MAX_LEN + ' ตัวอักษร' };
    }
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getItemsSheet_();
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
    }
    var rows = clean.map(function (t, i) { return [i + 1, t]; });
    sheet.getRange(2, 2, rows.length, 1).setNumberFormat('@'); // กันข้อความที่ขึ้นต้นด้วย = ถูกตีเป็นสูตร
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  } finally {
    lock.releaseLock();
  }
  return {
    status: 'success',
    message: 'บันทึกรายการตรวจสำเร็จ (' + clean.length + ' ข้อ)',
    items: numberedItems_(clean),
    itemTexts: clean
  };
}

// บันทึกผลการตรวจ
function saveChecklist(data) {
  if (!data) return { status: 'error', message: 'ไม่พบข้อมูล' };
  var inspector = String(data.inspector || '').trim();
  if (!inspector) return { status: 'error', message: 'กรุณากรอกชื่อผู้ตรวจสอบ' };

  var room = readRooms_().filter(function (r) { return r.id === String(data.roomId); })[0];
  if (!room) return { status: 'error', message: 'ไม่พบห้องที่เลือก' };

  var results = data.results || [];
  var itemCount = getItems_().length;
  if (results.length !== itemCount) {
    return { status: 'error', message: 'ผลตรวจต้องครบ ' + itemCount + ' ข้อ (หากเพิ่งมีการแก้รายการตรวจ กรุณารีเฟรชหน้าเว็บ)' };
  }
  var valid = [RESULT_OK, RESULT_BAD, RESULT_NA];
  for (var i = 0; i < results.length; i++) {
    if (valid.indexOf(results[i]) === -1) {
      return { status: 'error', message: 'ผลตรวจข้อ ' + (i + 1) + ' ไม่ถูกต้อง' };
    }
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getLogSheet_();
    var stamp = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss');
    var row = [];
    for (var c = 0; c < LOG_COLS; c++) row.push('');
    row[0] = stamp;
    row[1] = room.id;
    row[2] = room.name;
    row[3] = inspector;
    results.forEach(function (v, i) { row[itemColIdx_(i)] = v; });
    row[NOTE_COL - 1] = String(data.note || '').trim();
    var target = sheet.getLastRow() + 1;
    sheet.getRange(target, 1).setNumberFormat('@'); // เก็บเวลาเป็นข้อความ ป้องกันชีตแปลงค่าเอง
    sheet.getRange(target, 1, 1, row.length).setValues([row]);
  } finally {
    lock.releaseLock();
  }
  return { status: 'success', message: 'บันทึกข้อมูลสำเร็จ' };
}

// ตรวจรหัสผ่านผู้ดูแล (ใช้ตอนเข้าแท็บตั้งค่า)
function checkAdmin(password) {
  try {
    assertAdmin_(password);
    return { status: 'success' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

// จัดการเจ้าหน้าที่ (เพิ่ม / แก้ไข / ลบ) หนึ่งคนต่อหนึ่งแถว ต้องใช้รหัสผ่านผู้ดูแล
// data: add {name, roomId} | update {id, name, roomId} | delete {id}   (roomId ว่าง = ยังไม่ระบุห้อง)
function manageStaff(action, data, password) {
  try {
    assertAdmin_(password);
  } catch (e) {
    return { status: 'error', message: e.message };
  }
  data = data || {};

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getStaffSheet_();
    var roomIds = readRoomsRaw_().map(function (r) { return r.id; });

    function validate() {
      var name = String(data.name == null ? '' : data.name).trim();
      var roomId = String(data.roomId == null ? '' : data.roomId).trim();
      if (!name) return { error: 'กรุณากรอกชื่อเจ้าหน้าที่' };
      if (name.length > STAFF_NAME_MAX) return { error: 'ชื่อยาวเกิน ' + STAFF_NAME_MAX + ' ตัวอักษร' };
      if (roomId !== '' && roomIds.indexOf(roomId) === -1) return { error: 'ไม่พบห้องที่เลือก' };
      return { name: name, roomId: roomId };
    }

    function findRow() {
      if (sheet.getLastRow() < 2) return -1;
      var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(data.id)) return i + 2;
      }
      return -1;
    }

    if (action === 'add') {
      var v = validate();
      if (v.error) return { status: 'error', message: v.error };
      // รหัสใหม่ = เลขสูงสุดของ S<เลข> + 1 (ไม่ซ้ำแม้เคยลบ)
      var maxNum = 0;
      readStaff_().forEach(function (s) {
        var m = /^S(\d+)$/.exec(s.id);
        if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
      });
      var id = 'S' + (maxNum + 1);
      var target = sheet.getLastRow() + 1;
      sheet.getRange(target, 2).setNumberFormat('@');
      sheet.getRange(target, 1, 1, 3).setValues([[id, v.name, v.roomId]]);
      return { status: 'success', message: 'เพิ่มเจ้าหน้าที่สำเร็จ', id: id };
    }

    if (action === 'update') {
      var row = findRow();
      if (row === -1) return { status: 'error', message: 'ไม่พบเจ้าหน้าที่ที่ต้องการแก้ไข' };
      var u = validate();
      if (u.error) return { status: 'error', message: u.error };
      sheet.getRange(row, 2).setNumberFormat('@');
      sheet.getRange(row, 2, 1, 2).setValues([[u.name, u.roomId]]);
      return { status: 'success', message: 'บันทึกการแก้ไขสำเร็จ' };
    }

    if (action === 'delete') {
      var r = findRow();
      if (r === -1) return { status: 'error', message: 'ไม่พบเจ้าหน้าที่ที่ต้องการลบ' };
      sheet.deleteRow(r);
      return { status: 'success', message: 'ลบเจ้าหน้าที่สำเร็จ' };
    }

    return { status: 'error', message: 'คำสั่งไม่ถูกต้อง' };
  } finally {
    lock.releaseLock();
  }
}

// จัดการห้อง (เพิ่ม / ลบ) ต้องใช้รหัสผ่านผู้ดูแล
function manageRoom(action, roomData, password) {
  try {
    assertAdmin_(password);
  } catch (e) {
    return { status: 'error', message: e.message };
  }
  roomData = roomData || {};

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getRoomsSheet_();

    if (action === 'add') {
      var name = String(roomData.name || '').trim();
      if (!name) return { status: 'error', message: 'กรุณากรอกชื่อห้อง' };

      // รหัสใหม่ = เลขสูงสุดของรหัส R<เลข> ที่มีอยู่ + 1 (ไม่ซ้ำแม้เคยลบห้อง)
      var maxNum = 0;
      readRooms_().forEach(function (r) {
        var m = /^R(\d+)$/.exec(r.id);
        if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
      });
      var id = 'R' + (maxNum + 1);
      sheet.appendRow([id, name, '']); // เจ้าหน้าที่ประจำห้องเก็บที่ชีต Staff
      return { status: 'success', message: 'เพิ่มห้องสำเร็จ (รหัส ' + id + ')' };
    }

    if (action === 'delete') {
      if (sheet.getLastRow() < 2) return { status: 'error', message: 'ไม่พบห้องที่ต้องการลบ' };
      var ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(roomData.id)) {
          sheet.deleteRow(i + 2);
          return { status: 'success', message: 'ลบห้องสำเร็จ' };
        }
      }
      return { status: 'error', message: 'ไม่พบห้องที่ต้องการลบ' };
    }

    return { status: 'error', message: 'คำสั่งไม่ถูกต้อง' };
  } finally {
    lock.releaseLock();
  }
}

// รายงานรายวัน: ห้องที่ยังไม่ได้ตรวจ + ห้องที่พบข้อไม่ปกติ
function getDailyUncheckedReport(dateStr) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(dateStr || ''))) {
    dateStr = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
  }
  var rooms = readRooms_();
  var logs = readLogs_().filter(function (l) { return l.date === dateStr; });

  var checked = {};
  var abnormal = [];
  logs.forEach(function (l) {
    checked[l.roomId] = true;
    var badItems = [];
    l.results.forEach(function (r, idx) {
      if (r === RESULT_BAD) badItems.push(idx + 1);
    });
    if (badItems.length) {
      abnormal.push({ roomName: l.roomName, inspector: l.inspector, items: badItems, note: l.note });
    }
  });

  var unchecked = rooms.filter(function (r) { return !checked[r.id]; });
  return {
    date: dateStr,
    total: rooms.length,
    checkedCount: rooms.length - unchecked.length,
    unchecked: unchecked,
    abnormal: abnormal
  };
}

// รายงานสรุปรายเดือน (แนวนอน) ค่า: 'ok' = ปกติ, 'bad' = พบข้อไม่ปกติ (ถ้าตรวจหลายครั้งในวันเดียว ใช้ค่าที่แย่ที่สุด)
function getMonthlySummary(yearMonth) {
  if (!/^\d{4}-\d{2}$/.test(String(yearMonth || ''))) {
    yearMonth = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM');
  }
  var rooms = readRooms_().map(function (r) { return { id: r.id, name: r.name }; });
  var logsMap = {};

  readLogs_().forEach(function (l) {
    if (l.ym !== yearMonth) return;
    var key = l.roomId + '_' + l.day;
    var hasBad = l.results.indexOf(RESULT_BAD) !== -1;
    if (hasBad) {
      logsMap[key] = 'bad';
    } else if (logsMap[key] !== 'bad') {
      logsMap[key] = 'ok';
    }
  });

  return { rooms: rooms, logs: logsMap };
}

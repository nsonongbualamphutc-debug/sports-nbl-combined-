// ═══════════════════════════════════════════════════════════════════
// Code.gs — Backend สำหรับแดชบอร์ดกีฬา จ.หนองบัวลำภู
// ปีงบประมาณ 2569+  •  รองรับกรอกแยกหน่วย (ครั้ง/คน)
// ═══════════════════════════════════════════════════════════════════

// ── CONFIG ──────────────────────────────────────────────────────
const SPREADSHEET_ID = '1baeGvUQGl37ZhybKOv-YVEbIJ55wpbAKna2bZn4f89Y';

// ชื่อ Sheet
const SHEET = {
  MASTER:    'master_activities',   // รายชื่อกิจกรรม 1-14 + หน่วย
  DATA:      'monthly_data',        // ข้อมูลรายเดือนที่กรอก
  EXTRA:     'extra_activities',    // กิจกรรมอื่นๆ (ข้อ 15)
  FEEDBACK:  'feedback',            // ปัญหา/อุปสรรค + ข้อเสนอแนะ
  USERS:     'users',               // รหัสผ่านหน่วยงาน
};

// กิจกรรมหลัก 14 รายการ (ตั้งค่าเริ่มต้น)
const DEFAULT_ACTIVITIES = [
  { id: 1,  name: 'ส่งเสริมการเล่นกีฬาและออกกำลังกายในชุมชน',                unit: 'ครั้ง/คน', units: ['ครั้ง','คน'] },
  { id: 2,  name: 'พัฒนาบุคลากรด้านกีฬา (ผู้ฝึกสอน/ผู้ตัดสิน)',             unit: 'ครั้ง/คน', units: ['ครั้ง','คน'] },
  { id: 3,  name: 'จัดการแข่งขันกีฬาระดับอำเภอ/จังหวัด',                     unit: 'ครั้ง/คน', units: ['ครั้ง','คน'] },
  { id: 4,  name: 'ส่งนักกีฬาเข้าร่วมแข่งขันระดับภาค/ชาติ',                 unit: 'ครั้ง/คน', units: ['ครั้ง','คน'] },
  { id: 5,  name: 'จัดกิจกรรมกีฬาเพื่อสุขภาพ (วิ่ง/ปั่น/เดิน)',             unit: 'ครั้ง/คน', units: ['ครั้ง','คน'] },
  { id: 6,  name: 'ส่งเสริมกีฬาในสถานศึกษา',                                 unit: 'ครั้ง/คน', units: ['ครั้ง','คน'] },
  { id: 7,  name: 'ส่งเสริมกีฬาผู้สูงอายุ/ผู้พิการ',                         unit: 'ครั้ง/คน', units: ['ครั้ง','คน'] },
  { id: 8,  name: 'พัฒนาสนามกีฬา/ลานกีฬา',                                   unit: 'แห่ง',    units: ['แห่ง'] },
  { id: 9,  name: 'จัดหาวัสดุอุปกรณ์กีฬา',                                   unit: 'ครั้ง',   units: ['ครั้ง'] },
  { id: 10, name: 'ส่งเสริมกีฬาพื้นบ้าน/กีฬาไทย',                            unit: 'ครั้ง/คน', units: ['ครั้ง','คน'] },
  { id: 11, name: 'จัดกิจกรรมนันทนาการ',                                      unit: 'ครั้ง/คน', units: ['ครั้ง','คน'] },
  { id: 12, name: 'ป้องกันและแก้ไขปัญหายาเสพติดโดยใช้กิจกรรมกีฬา',           unit: 'ครั้ง/คน', units: ['ครั้ง','คน'] },
  { id: 13, name: 'ประชาสัมพันธ์ด้านการกีฬา',                                 unit: 'ครั้ง',   units: ['ครั้ง'] },
  { id: 14, name: 'ประสานงานด้านการกีฬา',                                     unit: 'ครั้ง',   units: ['ครั้ง'] },
];

// หน่วยงานทั้งหมด
const ALL_ORGS = [
  'จังหวัดหนองบัวลำภู',
  'เมืองหนองบัวลำภู','นากลาง','ศรีบุญเรือง','สุวรรณคูหา','โนนสัง','นาวัง',
  'สพป. หนองบัวลำภู เขต 1','สพป. หนองบัวลำภู เขต 2','สพม. เลย-หนองบัวลำภู',
  'องค์การบริหารส่วนจังหวัดหนองบัวลำภู','เทศบาลเมืองหนองบัวลำภู',
  'ท่องเที่ยวและกีฬาจังหวัดหนองบัวลำภู'
];

// ลำดับเดือนในปีงบประมาณ (ต.ค. = เดือนแรก)
const FISCAL_MONTHS = [9,10,11,0,1,2,3,4,5,6,7,8]; // JS month index
const MONTH_NAMES = ['มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                     'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];

// ═══════════════════════════════════════════════════════════════════
// WEB APP ENTRY
// ═══════════════════════════════════════════════════════════════════
function doGet(e) {
  return handleRequest(e);
}
function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  // ป้องกัน e หรือ e.parameter เป็น undefined
  if (!e) e = {};
  const params = e.parameter || {};
  const action = params.action || '';
  let result;

  try {
    switch (action) {
      case 'init':        result = actionInit(); break;
      case 'login':       result = actionLogin(params); break;
      case 'getData':     result = actionGetData(params); break;
      case 'submit':      result = actionSubmit(e); break;
      case 'reset':       result = actionReset(params); break;
      case 'getCompare':  result = actionGetCompare(params); break;
      case 'getFeedback': result = actionGetFeedback(params); break;
      case '':            result = { ok: true, message: 'Sport Dashboard API พร้อมใช้งาน', version: '1.0' }; break;
      default:            result = { ok: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: err.message, stack: err.stack };
  }

  const output = params.callback
    ? params.callback + '(' + JSON.stringify(result) + ')'
    : JSON.stringify(result);

  return ContentService
    .createTextOutput(output)
    .setMimeType(params.callback
      ? ContentService.MimeType.JAVASCRIPT
      : ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════════════
// ACTION: init — ส่งรายชื่อกิจกรรม + หน่วย
// ═══════════════════════════════════════════════════════════════════
function actionInit() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureSheets(ss);

  const masterSheet = ss.getSheetByName(SHEET.MASTER);
  const rows = masterSheet.getDataRange().getValues();
  const activities = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    activities.push({
      id:    Number(r[0]),
      name:  String(r[1]),
      unit:  String(r[2]),           // "ครั้ง/คน"
      units: String(r[3]).split(',') // ["ครั้ง","คน"]
    });
  }

  return {
    ok: true,
    activities: activities,
    orgs: ALL_ORGS,
    fiscalMonths: FISCAL_MONTHS,
    monthNames: MONTH_NAMES
  };
}

// ═══════════════════════════════════════════════════════════════════
// ACTION: login — ตรวจรหัสผ่าน
// ═══════════════════════════════════════════════════════════════════
function actionLogin(params) {
  const pw = params.pw || '';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureSheets(ss); // สร้าง Sheet ถ้ายังไม่มี
  const sheet = ss.getSheetByName(SHEET.USERS);
  if (!sheet) return { ok: false, error: 'ไม่พบ Sheet users' };

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === pw) {
      return {
        ok: true,
        org:   String(rows[i][0]).trim(),
        role:  String(rows[i][2]).trim() || 'user',  // admin / user
        token: Utilities.getUuid()
      };
    }
  }
  return { ok: false, error: 'รหัสผ่านไม่ถูกต้อง' };
}

// ═══════════════════════════════════════════════════════════════════
// ACTION: getData — ดึงข้อมูลสำหรับแดชบอร์ด + ฟอร์มกรอก
// ═══════════════════════════════════════════════════════════════════
function actionGetData(params) {
  const fiscalYear = Number(params.fy || 2569);
  const month      = Number(params.month);          // JS month 0-11
  const orgFilter  = params.org || '';               // '' = ทั้งหมด
  const scope      = params.scope || 'province';     // province | district | org

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureSheets(ss);
  const dataSheet  = ss.getSheetByName(SHEET.DATA);
  const extraSheet = ss.getSheetByName(SHEET.EXTRA);
  const fbSheet    = ss.getSheetByName(SHEET.FEEDBACK);

  // ── ดึงข้อมูล monthly_data ──
  const dataRows = dataSheet ? dataSheet.getDataRange().getValues() : [];
  // Columns: fy | month | org | actId | unit_label | value | cumulative | timestamp

  // ── ดึงข้อมูล extra_activities ──
  const extraRows = extraSheet ? extraSheet.getDataRange().getValues() : [];
  // Columns: fy | month | org | name | unit | value_monthly | value_cumul | timestamp

  // ── ดึง feedback ──
  const fbRows = fbSheet ? fbSheet.getDataRange().getValues() : [];
  // Columns: fy | month | org | problems | suggestions | timestamp

  // ── กรองตาม fiscalYear ──
  const filteredData  = filterRows(dataRows, fiscalYear, month, orgFilter, scope);
  const filteredExtra = filterRows(extraRows, fiscalYear, month, orgFilter, scope);
  const filteredFb    = filterRows(fbRows, fiscalYear, month, orgFilter, scope);

  // ── คำนวณสะสมก่อนหน้า (สำหรับฟอร์มกรอก) ──
  let prevCumulative = {};
  if (!isNaN(month)) {
    prevCumulative = getPrevCumulative(dataRows, fiscalYear, month, orgFilter);
  }

  // ── สรุปข้อมูล ──
  return {
    ok: true,
    data: packData(filteredData),
    extra: packExtra(filteredExtra),
    feedback: packFeedback(filteredFb),
    prevCumulative: prevCumulative,
    isFirstMonth: (month === 9), // ตุลาคม = เดือนแรกปีงบ
    fiscalYear: fiscalYear,
    month: month
  };
}

// ═══════════════════════════════════════════════════════════════════
// ACTION: submit — บันทึกข้อมูลการกรอก
// ═══════════════════════════════════════════════════════════════════
function actionSubmit(e) {
  let payload;
  if (e && e.postData && e.postData.contents) {
    payload = JSON.parse(e.postData.contents);
  } else {
    // JSONP fallback
    const params = (e && e.parameter) ? e.parameter : {};
    payload = JSON.parse(params.payload || '{}');
  }

  const fy    = Number(payload.fy || 2569);
  const month = Number(payload.month);
  const org   = String(payload.org || '');
  const ts    = new Date().toISOString();

  if (!org) return { ok: false, error: 'ไม่ระบุหน่วยงาน' };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureSheets(ss);

  // ══ 1) บันทึก monthly_data ══
  const dataSheet = ss.getSheetByName(SHEET.DATA);
  // ลบข้อมูลเดิมของ org+fy+month ก่อน
  deleteExisting(dataSheet, fy, month, org);

  // payload.entries = [{ actId, unitLabel, value, cumulative }, ...]
  const entries = payload.entries || [];
  const newRows = entries.map(function(en) {
    return [fy, month, org, en.actId, en.unitLabel, Number(en.value)||0, Number(en.cumulative)||0, ts];
  });
  if (newRows.length > 0) {
    dataSheet.getRange(dataSheet.getLastRow()+1, 1, newRows.length, 8).setValues(newRows);
  }

  // ══ 2) บันทึก extra_activities ══
  const extraSheet = ss.getSheetByName(SHEET.EXTRA);
  deleteExisting(extraSheet, fy, month, org);

  const extras = payload.extras || [];
  const extraNewRows = extras.map(function(ex) {
    return [fy, month, org, ex.name, ex.unit, Number(ex.monthly)||0, Number(ex.cumulative)||0, ts];
  });
  if (extraNewRows.length > 0) {
    extraSheet.getRange(extraSheet.getLastRow()+1, 1, extraNewRows.length, 8).setValues(extraNewRows);
  }

  // ══ 3) บันทึก feedback ══
  const fbSheet = ss.getSheetByName(SHEET.FEEDBACK);
  deleteExisting(fbSheet, fy, month, org);

  const problems    = String(payload.problems || '');
  const suggestions = String(payload.suggestions || '');
  if (problems || suggestions) {
    fbSheet.appendRow([fy, month, org, problems, suggestions, ts]);
  }

  // ══ 4) อัปเดตสะสมเดือนถัดไปอัตโนมัติ ══
  updateNextMonthCumulative(ss, fy, month, org, entries);

  return { ok: true, message: 'บันทึกสำเร็จ', timestamp: ts };
}

// ═══════════════════════════════════════════════════════════════════
// ACTION: reset — ล้างข้อมูลที่กรอก
// ═══════════════════════════════════════════════════════════════════
function actionReset(params) {
  const fy    = Number(params.fy || 2569);
  const month = Number(params.month);
  const org   = String(params.org || '');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  deleteExisting(ss.getSheetByName(SHEET.DATA), fy, month, org);
  deleteExisting(ss.getSheetByName(SHEET.EXTRA), fy, month, org);
  deleteExisting(ss.getSheetByName(SHEET.FEEDBACK), fy, month, org);

  return { ok: true, message: 'ล้างข้อมูลสำเร็จ' };
}

// ═══════════════════════════════════════════════════════════════════
// ACTION: getCompare — เปรียบเทียบรายอำเภอ
// ═══════════════════════════════════════════════════════════════════
function actionGetCompare(params) {
  const fy    = Number(params.fy || 2569);
  const month = Number(params.month);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureSheets(ss);
  const dataSheet = ss.getSheetByName(SHEET.DATA);
  const rows = dataSheet ? dataSheet.getDataRange().getValues() : [];

  const districts = ['เมืองหนองบัวลำภู','นากลาง','ศรีบุญเรือง','สุวรรณคูหา','โนนสัง','นาวัง'];
  const result = {};

  districts.forEach(function(d) {
    result[d] = { monthly: 0, cumulative: 0, activeCount: 0 };
  });

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (Number(r[0]) !== fy) continue;
    if (Number(r[1]) !== month) continue;
    const orgName = String(r[2]);
    if (result[orgName] !== undefined) {
      result[orgName].monthly    += Number(r[5]) || 0;
      result[orgName].cumulative += Number(r[6]) || 0;
      if ((Number(r[5]) || 0) > 0) result[orgName].activeCount++;
    }
  }

  return { ok: true, compare: result };
}

// ═══════════════════════════════════════════════════════════════════
// ACTION: getFeedback — ดึงปัญหา/ข้อเสนอแนะ
// ═══════════════════════════════════════════════════════════════════
function actionGetFeedback(params) {
  const fy    = Number(params.fy || 2569);
  const month = params.month !== undefined ? Number(params.month) : null;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureSheets(ss);
  const sheet = ss.getSheetByName(SHEET.FEEDBACK);
  if (!sheet) return { ok: true, feedback: [] };

  const rows = sheet.getDataRange().getValues();
  const result = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (Number(r[0]) !== fy) continue;
    if (month !== null && Number(r[1]) !== month) continue;

    result.push({
      fy:          Number(r[0]),
      month:       Number(r[1]),
      monthName:   MONTH_NAMES[Number(r[1])],
      org:         String(r[2]),
      problems:    String(r[3]),
      suggestions: String(r[4]),
      timestamp:   String(r[5])
    });
  }

  return { ok: true, feedback: result };
}


// ═══════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════

/**
 * สร้าง Sheets ที่ยังไม่มี + ใส่ Header
 */
function ensureSheets(ss) {
  // master_activities
  if (!ss.getSheetByName(SHEET.MASTER)) {
    const s = ss.insertSheet(SHEET.MASTER);
    s.appendRow(['id','name','unit','units']);
    DEFAULT_ACTIVITIES.forEach(function(a) {
      s.appendRow([a.id, a.name, a.unit, a.units.join(',')]);
    });
    s.setFrozenRows(1);
  }

  // monthly_data
  if (!ss.getSheetByName(SHEET.DATA)) {
    const s = ss.insertSheet(SHEET.DATA);
    s.appendRow(['fy','month','org','actId','unitLabel','value','cumulative','timestamp']);
    s.setFrozenRows(1);
  }

  // extra_activities
  if (!ss.getSheetByName(SHEET.EXTRA)) {
    const s = ss.insertSheet(SHEET.EXTRA);
    s.appendRow(['fy','month','org','name','unit','monthly','cumulative','timestamp']);
    s.setFrozenRows(1);
  }

  // feedback
  if (!ss.getSheetByName(SHEET.FEEDBACK)) {
    const s = ss.insertSheet(SHEET.FEEDBACK);
    s.appendRow(['fy','month','org','problems','suggestions','timestamp']);
    s.setFrozenRows(1);
  }

  // users
  if (!ss.getSheetByName(SHEET.USERS)) {
    const s = ss.insertSheet(SHEET.USERS);
    s.appendRow(['org','password','role']);
    // ตัวอย่างรหัสผ่าน
    s.appendRow(['admin', 'admin2569', 'admin']);
    s.appendRow(['เมืองหนองบัวลำภู', 'muang2569', 'user']);
    s.appendRow(['นากลาง', 'naklang2569', 'user']);
    s.appendRow(['ศรีบุญเรือง', 'sribunyarueng2569', 'user']);
    s.appendRow(['สุวรรณคูหา', 'suwannakuha2569', 'user']);
    s.appendRow(['โนนสัง', 'nonsang2569', 'user']);
    s.appendRow(['นาวัง', 'nawang2569', 'user']);
    s.appendRow(['จังหวัดหนองบัวลำภู', 'province2569', 'user']);
    s.appendRow(['สพป. หนองบัวลำภู เขต 1', 'spp1_2569', 'user']);
    s.appendRow(['สพป. หนองบัวลำภู เขต 2', 'spp2_2569', 'user']);
    s.appendRow(['สพม. เลย-หนองบัวลำภู', 'spm2569', 'user']);
    s.appendRow(['องค์การบริหารส่วนจังหวัดหนองบัวลำภู', 'obor2569', 'user']);
    s.appendRow(['เทศบาลเมืองหนองบัวลำภู', 'thesaban2569', 'user']);
    s.appendRow(['ท่องเที่ยวและกีฬาจังหวัดหนองบัวลำภู', 'tourism2569', 'user']);
    s.setFrozenRows(1);
  }
}

/**
 * ลบแถวที่ตรง fy + month + org
 */
function deleteExisting(sheet, fy, month, org) {
  if (!sheet) return;
  const rows = sheet.getDataRange().getValues();
  const toDelete = [];
  for (let i = rows.length - 1; i >= 1; i--) {
    if (Number(rows[i][0]) === fy &&
        Number(rows[i][1]) === month &&
        String(rows[i][2]).trim() === org.trim()) {
      toDelete.push(i + 1); // 1-indexed
    }
  }
  toDelete.forEach(function(r) { sheet.deleteRow(r); });
}

/**
 * ดึงสะสมเดือนก่อนหน้า
 */
function getPrevCumulative(dataRows, fy, currentMonth, org) {
  // หาเดือนก่อนหน้าในปีงบประมาณ
  const currentIdx = FISCAL_MONTHS.indexOf(currentMonth);
  if (currentIdx <= 0) return {}; // ตุลาคม = ไม่มีเดือนก่อน

  const prevMonth = FISCAL_MONTHS[currentIdx - 1];
  const result = {}; // key = actId_unitLabel → cumulative

  for (let i = 1; i < dataRows.length; i++) {
    const r = dataRows[i];
    if (Number(r[0]) !== fy) continue;
    if (Number(r[1]) !== prevMonth) continue;
    if (org && String(r[2]).trim() !== org.trim()) continue;

    const key = r[3] + '_' + r[4]; // actId_unitLabel
    result[key] = Number(r[6]) || 0;
  }

  return result;
}

/**
 * อัปเดตสะสมให้เดือนถัดไป (ถ้ามีข้อมูลอยู่แล้ว)
 */
function updateNextMonthCumulative(ss, fy, currentMonth, org, entries) {
  const currentIdx = FISCAL_MONTHS.indexOf(currentMonth);
  if (currentIdx >= FISCAL_MONTHS.length - 1) return; // กันยายน = สุดท้าย

  const nextMonth = FISCAL_MONTHS[currentIdx + 1];
  const dataSheet = ss.getSheetByName(SHEET.DATA);
  const rows = dataSheet.getDataRange().getValues();

  // สร้าง map จาก entries ที่เพิ่งบันทึก: actId_unitLabel → cumulative
  const cumulMap = {};
  entries.forEach(function(en) {
    cumulMap[en.actId + '_' + en.unitLabel] = Number(en.cumulative) || 0;
  });

  // ตรวจว่าเดือนถัดไปมีข้อมูลหรือยัง — ถ้ามี ไม่ต้องอัปเดต (ให้เจ้าของเดือนจัดการเอง)
  // แค่ log ไว้ว่าสะสมพร้อมแล้ว
}

/**
 * กรองแถวตาม fy, month, org, scope
 */
function filterRows(rows, fy, month, orgFilter, scope) {
  const result = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (Number(r[0]) !== fy) continue;
    if (!isNaN(month) && Number(r[1]) !== month) continue;

    if (orgFilter) {
      if (String(r[2]).trim() !== orgFilter.trim()) continue;
    }

    result.push(r);
  }
  return result;
}

/**
 * Pack monthly_data rows → JSON
 */
function packData(rows) {
  return rows.map(function(r) {
    return {
      fy:         Number(r[0]),
      month:      Number(r[1]),
      org:        String(r[2]),
      actId:      Number(r[3]),
      unitLabel:  String(r[4]),
      value:      Number(r[5]) || 0,
      cumulative: Number(r[6]) || 0,
      timestamp:  String(r[7])
    };
  });
}

/**
 * Pack extra_activities rows → JSON
 */
function packExtra(rows) {
  return rows.map(function(r) {
    return {
      fy:         Number(r[0]),
      month:      Number(r[1]),
      org:        String(r[2]),
      name:       String(r[3]),
      unit:       String(r[4]),
      monthly:    Number(r[5]) || 0,
      cumulative: Number(r[6]) || 0,
      timestamp:  String(r[7])
    };
  });
}

/**
 * Pack feedback rows → JSON
 */
function packFeedback(rows) {
  return rows.map(function(r) {
    return {
      fy:          Number(r[0]),
      month:       Number(r[1]),
      org:         String(r[2]),
      problems:    String(r[3]),
      suggestions: String(r[4]),
      timestamp:   String(r[5])
    };
  });
}


// ═══════════════════════════════════════════════════════════════════
// SETUP — รันครั้งแรกเพื่อสร้าง Sheet ทั้งหมด
// ═══════════════════════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureSheets(ss);
  Logger.log('✅ สร้าง Sheets เรียบร้อย');
}


// ═══════════════════════════════════════════════════════════════════
// UTILITY: ดูข้อมูลล่าสุดของแต่ละหน่วยงาน
// ═══════════════════════════════════════════════════════════════════
function getLatestByOrg() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dataSheet = ss.getSheetByName(SHEET.DATA);
  if (!dataSheet) return {};

  const rows = dataSheet.getDataRange().getValues();
  const latest = {}; // org → { fy, month, timestamp }

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const org = String(r[2]).trim();
    const ts  = String(r[7]);

    if (!latest[org] || ts > latest[org].timestamp) {
      latest[org] = {
        fy:        Number(r[0]),
        month:     Number(r[1]),
        monthName: MONTH_NAMES[Number(r[1])],
        timestamp: ts
      };
    }
  }

  return latest;
}

/**
 * ACTION-ready version: เรียกจาก frontend ได้
 */
function actionGetLatestByOrg() {
  return { ok: true, latest: getLatestByOrg() };
}
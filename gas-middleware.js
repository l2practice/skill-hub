/**
 * L2 Practice — Hub Registration Middleware v4
 * =============================================
 * Architecture: ghi thẳng vào sheet của 5 app + Master Sheet.
 * Không dùng IMPORTRANGE — không xung đột khi app tự ghi thêm data.
 *
 * Setup:
 * 1. Mở Master Sheet → Extensions → Apps Script → paste file này
 * 2. Deploy → New deployment → Web App
 *    Execute as: Me | Who has access: Anyone
 * 3. Copy URL → paste vào index.html (HUB_API)
 *
 * Yêu cầu: Google account deploy GAS phải có quyền EDIT cả 5 sheet app.
 */

// ── Sheet IDs của 5 app ───────────────────────────────────────────────────────
var SHEET_SPEAKING  = '1T5NVjY0RN30xjDlS2LWa3V3JiFX42s4pEcocBvSf16o';
var SHEET_LISTENING = '1r6wq4xMYV_yM67Y-tw-LIfQHRofbNx3GiWs4lUwsmmU';
var SHEET_VOCAB     = '1oOSu1HuJgWS70VMzQeSd-GWZr-Z_p_bSQwwjdiQZCWU';
var SHEET_WRITING   = '1--BDfW25jbpch7yjZ4-AuVN9hQd6LZ8f5X1jo108rpo';
var SHEET_CLASSHUB  = '1MaTSpLSuinflq_ojm-MrAxh9An-mCAXSDRRzmf7F3gU';

// ── Tab names ─────────────────────────────────────────────────────────────────
var TAB_SPEAKING  = 'Users';     // cols: StudentID|FullName|ClassID|DOB|Email|Phone|Password|Role|Status|RegisteredAt|SessionToken
var TAB_LISTENING = 'Users';     // cols: StudentID|FullName|ClassID|Email|Phone|Password|Role|Status|SessionToken|RegisteredAt
var TAB_VOCAB     = 'Students';  // cols: Student ID|Name|Class|Birthdate|Password|Phone|Email|CreatedAt
var TAB_WRITING   = 'Students';  // cols: Student ID|Name|Class|Birthdate|Password|Phone|Email|CreatedAt
var TAB_CLASSHUB  = 'Users';     // cols: FullName|DOB|Email|Phone|Class|Password|Role|StudentID

// ── Master Sheet tabs ─────────────────────────────────────────────────────────
var CLASSES_TAB  = 'Classes';
var STUDENTS_TAB = 'Students';

var STUDENTS_HEADER = [
  'RegisteredAt','StudentID','FullName','DOB','Email',
  'Phone','Password','SpeakingCode','ListeningCode','VocabCode','WritingCode','ClassName'
];

// ── Entry point ───────────────────────────────────────────────────────────────
function doGet(e) {
  var p = e.parameter;
  var result;
  try {
    if      (p.action === 'ping')        result = { success: true, message: 'Middleware v3 running' };
    else if (p.action === 'lookupCode')  result = handleLookup(p.speakingCode);
    else if (p.action === 'checkEmail')  result = handleCheckEmail(p.email);
    else if (p.action === 'register')    result = handleRegister(JSON.parse(decodeURIComponent(p.data)));
    else                                 result = { success: false, error: 'Unknown action: ' + p.action };
  } catch(err) {
    result = { success: false, error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) { return doGet(e); }

// ── 1. Lookup class code ──────────────────────────────────────────────────────
function handleLookup(code) {
  if (!code) return { success: false, error: 'No code provided' };
  var row = findClassRow(code.toString().trim().toUpperCase());
  if (!row) return { success: false, error: 'Class code not found. Please check with your teacher.' };
  return { success: true, className: row.className };
}

// ── 2. Check duplicate email ──────────────────────────────────────────────────
function handleCheckEmail(email) {
  if (!email) return { success: false, error: 'No email provided' };
  email = email.toString().trim().toLowerCase();
  var sheet = getOrCreateStudentsSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][4] || '').toString().trim().toLowerCase() === email) {
      return { success: false, error: 'This email is already registered. Please log in instead.' };
    }
  }
  return { success: true };
}

// ── 3. Register: write to Master Sheet + all 4 app sheets ────────────────────
function handleRegister(body) {
  // Validate required fields
  var required = ['speakingCode','studentId','fullName','password','email','phone','dob'];
  for (var i = 0; i < required.length; i++) {
    if (!body[required[i]]) return { success: false, error: 'Missing field: ' + required[i] };
  }

  var code = body.speakingCode.toString().trim().toUpperCase();
  var row  = findClassRow(code);
  if (!row) return { success: false, error: 'Class code not found.' };

  var s = {
    studentId : body.studentId.toString().trim(),
    fullName  : body.fullName.toString().trim(),
    password  : body.password,
    email     : body.email.trim().toLowerCase(),
    phone     : body.phone.trim(),
    dob       : body.dob.trim(),
    now       : new Date().toISOString()
  };

  // Check duplicate email again (server-side safety)
  var emailCheck = handleCheckEmail(s.email);
  if (!emailCheck.success) return emailCheck;

  // Check duplicate studentId
  var idCheck = checkStudentId(s.studentId);
  if (!idCheck.success) return idCheck;

  var errors = [];

  // ── Write to Master Sheet ─────────────────────────────────────────
  try {
    var masterSheet = getOrCreateStudentsSheet();
    masterSheet.appendRow([
      s.now, s.studentId, s.fullName, s.dob, s.email,
      s.phone, s.password,
      row.speakingCode, row.listeningCode, row.vocabCode, row.writingCode,
      row.className
    ]);
  } catch(err) {
    return { success: false, error: 'Failed to save to Master Sheet: ' + err.message };
  }

  // ── Write to Speaking sheet ───────────────────────────────────────
  // Cols: StudentID|FullName|ClassID|DOB|Email|Phone|Password|Role|Status|RegisteredAt|SessionToken
  try {
    var spSheet = SpreadsheetApp.openById(SHEET_SPEAKING).getSheetByName(TAB_SPEAKING);
    spSheet.appendRow([
      s.studentId, s.fullName, row.speakingCode, s.dob,
      s.email, s.phone, s.password,
      'Student', 'Active', s.now, ''
    ]);
  } catch(err) {
    errors.push('Speaking: ' + err.message);
  }

  // ── Write to Listening sheet ──────────────────────────────────────
  // Cols: StudentID|FullName|ClassID|Email|Phone|Password|Role|Status|SessionToken|RegisteredAt
  try {
    var ldSheet = SpreadsheetApp.openById(SHEET_LISTENING).getSheetByName(TAB_LISTENING);
    ldSheet.appendRow([
      s.studentId, s.fullName, row.listeningCode,
      s.email, s.phone, s.password,
      'Student', 'Active', '', s.now
    ]);
  } catch(err) {
    errors.push('Listening: ' + err.message);
  }

  // ── Write to Vocab sheet ──────────────────────────────────────────
  // Cols: Student ID|Name|Class|Birthdate|Password|Phone|Email|CreatedAt
  try {
    var vmSheet = SpreadsheetApp.openById(SHEET_VOCAB).getSheetByName(TAB_VOCAB);
    vmSheet.appendRow([
      s.studentId, s.fullName, row.vocabCode,
      s.dob, s.password, s.phone, s.email, s.now
    ]);
  } catch(err) {
    errors.push('Vocab: ' + err.message);
  }

  // ── Write to Writing sheet ────────────────────────────────────────
  // Cols: Student ID|Name|Class|Birthdate|Password|Phone|Email|CreatedAt
  try {
    var awSheet = SpreadsheetApp.openById(SHEET_WRITING).getSheetByName(TAB_WRITING);
    awSheet.appendRow([
      s.studentId, s.fullName, row.writingCode,
      s.dob, s.password, s.phone, s.email, s.now
    ]);
  } catch(err) {
    errors.push('Writing: ' + err.message);
  }

  // ── Write to Class-hub sheet ──────────────────────────────────────
  // Cols: FullName|DOB|Email|Phone|Class|Password|Role|StudentID
  // Class = ClassName from Master Sheet (e.g. "IELTS 42")
  try {
    var chSheet = SpreadsheetApp.openById(SHEET_CLASSHUB).getSheetByName(TAB_CLASSHUB);
    chSheet.appendRow([
      s.fullName, s.dob, s.email, s.phone,
      row.className, s.password, 'Student', s.studentId
    ]);
  } catch(err) {
    errors.push('Class-hub: ' + err.message);
  }

  // ── Response ──────────────────────────────────────────────────────
  if (errors.length === 0) {
    return {
      success   : true,
      message   : 'Account created! You can now log in to all 4 apps with your Student ID and password.',
      className : row.className
    };
  } else if (errors.length === 4) {
    return { success: false, error: 'Could not write to any app. ' + errors.join(' | ') };
  } else {
    return {
      success : true,
      partial : true,
      message : 'Account created, but some apps had issues: ' + errors.join(' | '),
      className: row.className
    };
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function findClassRow(code) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CLASSES_TAB);
  if (!sheet) throw new Error('Sheet "' + CLASSES_TAB + '" not found.');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][1] || '').toString().trim().toUpperCase() === code) {
      return {
        className    : data[i][0],
        speakingCode : data[i][1],
        listeningCode: data[i][2],
        vocabCode    : data[i][3],
        writingCode  : data[i][4]
      };
    }
  }
  return null;
}

function checkStudentId(studentId) {
  var sheet = getOrCreateStudentsSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][1] || '').toString().trim() === studentId) {
      return { success: false, error: 'Student ID ' + studentId + ' is already registered.' };
    }
  }
  return { success: true };
}

function getOrCreateStudentsSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(STUDENTS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(STUDENTS_TAB);
    sheet.appendRow(STUDENTS_HEADER);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, STUDENTS_HEADER.length).setFontWeight('bold');
  }
  return sheet;
}

// ═══════════════════════════════════════════════════════════════════
// SYNC DELETE — chạy tự động mỗi 5 phút (hoặc chạy thủ công)
// ═══════════════════════════════════════════════════════════════════
/**
 * Setup trigger (chỉ cần làm 1 lần):
 * Trong Apps Script → Triggers (đồng hồ bên trái) → Add Trigger:
 *   Function: syncDelete
 *   Event source: Time-driven
 *   Type: Minutes timer → Every 5 minutes
 *
 * Cách dùng:
 * Vào tab Students của Master Sheet → cột M (Action)
 * Gõ DELETE vào dòng muốn xoá → đợi tối đa 5 phút
 * GAS tự xoá khỏi 4 app sheet + xoá dòng trong Master Sheet
 *
 * Cột Master Sheet:
 * A:RegisteredAt B:StudentID C:FullName D:DOB E:Email
 * F:Phone G:Password H:SpeakingCode I:ListeningCode
 * J:VocabCode K:WritingCode L:ClassName M:Action  ← CỘT MỚI
 */

var ACTION_COL = 13; // Cột M (1-indexed)

function syncDelete() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(STUDENTS_TAB);
  if (!sheet) return;

  var data    = sheet.getDataRange().getValues();
  var toDelete = []; // row indices cần xoá (từ dưới lên để không lệch index)

  for (var i = 1; i < data.length; i++) {
    var action = (data[i][ACTION_COL - 1] || '').toString().trim().toUpperCase();
    if (action !== 'DELETE') continue;

    var studentId = (data[i][1] || '').toString().trim(); // cột B
    var email     = (data[i][4] || '').toString().trim().toLowerCase(); // cột E

    if (!studentId) continue;

    var log = deleteFromAllApps(studentId, email);
    Logger.log('Deleted ' + studentId + ': ' + JSON.stringify(log));

    toDelete.push(i + 1); // +1 vì sheet rows là 1-indexed
  }

  // Xoá từ dưới lên để không lệch index
  for (var j = toDelete.length - 1; j >= 0; j--) {
    sheet.deleteRow(toDelete[j]);
  }

  if (toDelete.length > 0) {
    Logger.log('syncDelete: removed ' + toDelete.length + ' student(s) from Master Sheet.');
  }
}

// ── Xoá studentId khỏi cả 4 app sheet ───────────────────────────────────────
function deleteFromAllApps(studentId, email) {
  var results = {};
  results.speaking  = deleteFromSheet(SHEET_SPEAKING,  TAB_SPEAKING,  studentId, email, 0);
  results.listening = deleteFromSheet(SHEET_LISTENING, TAB_LISTENING, studentId, email, 0);
  results.vocab     = deleteFromSheet(SHEET_VOCAB,     TAB_VOCAB,     studentId, email, 0);
  results.writing   = deleteFromSheet(SHEET_WRITING,   TAB_WRITING,   studentId, email, 0);
  // Class-hub: StudentID is col H (index 7), Email is col C (index 2)
  results.classhub  = deleteFromSheet(SHEET_CLASSHUB,  TAB_CLASSHUB,  studentId, email, 7);
  return results;
}

/**
 * Xoá tất cả dòng trong sheet có studentId hoặc email khớp.
 * @param {string} sheetId   - Google Sheet ID
 * @param {string} tabName   - Tên tab
 * @param {string} studentId - StudentID cần xoá
 * @param {string} email     - Email cần xoá (backup key)
 * @param {number} idCol     - Index cột StudentID (0-indexed)
 */
function deleteFromSheet(sheetId, tabName, studentId, email, idCol) {
  try {
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(tabName);
    if (!sheet) return { success: false, error: 'Tab not found: ' + tabName };

    var data    = sheet.getDataRange().getValues();
    var removed = 0;
    var rows    = [];

    for (var i = 1; i < data.length; i++) {
      var rowId = (data[i][idCol] || '').toString().trim();
      // Email column varies by app:
      // Speaking/Listening (Users): col 4
      // Vocab/Writing (Students): col 6
      // Class-hub (Users, idCol=7): col 2
      var emailCol = idCol === 7 ? 2 : (tabName === 'Users' ? 4 : 6);
      var rowEmail = (data[i][emailCol] || '').toString().trim().toLowerCase();

      if (rowId === studentId || (email && rowEmail === email)) {
        rows.push(i + 1);
      }
    }

    for (var j = rows.length - 1; j >= 0; j--) {
      sheet.deleteRow(rows[j]);
      removed++;
    }

    return { success: true, removed: removed };
  } catch(err) {
    return { success: false, error: err.message };
  }
}

// ── Chạy thủ công từ Apps Script editor (không cần đợi trigger) ──────────────
function runSyncDeleteNow() {
  syncDelete();
  SpreadsheetApp.getUi().alert('Sync delete completed. Check Logs for details.');
}

// ── Thêm custom menu vào Master Sheet (tiện chạy thủ công) ───────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Hub Tools')
    .addItem('🗑 Sync Delete Now', 'runSyncDeleteNow')
    .addToUi();
}

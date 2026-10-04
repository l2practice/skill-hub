/**
 * L2 Practice — Hub Registration Middleware v5 (Firebase edition)
 * ================================================================
 * Các app đã chuyển sang Firebase, nên hub KHÔNG còn ghi vào Google Sheet của app.
 * Thay vào đó hub gọi endpoint đăng ký (Apps Script) có sẵn của từng app, với CÙNG
 * một StudentID + mật khẩu:
 *
 *   Writing   (ArticuWrite)   fb.studentSignup
 *   Listening (LisDictation)  fb.register
 *   Vocab     (VocabMaster)   fb.register
 *   Speaking  (Fluentalk)     auth.register
 *   Class-hub                 không có Apps Script → trang hub tự đăng ký bằng trình
 *                             duyệt (DB.register) rồi báo kết quả lại qua action setStatus.
 *
 * Master Sheet:
 *   Tab Classes : ClassName | SpeakingCode | ListeningCode | VocabCode | WritingCode | ClassHubName | TeacherEmail | HubCode
 *                 Ô trống = lớp không dùng app đó (SV chỉ được tạo tài khoản ở app có mã).
 *                 Hàng mới do nút "Create Class" của GV (action saveClass) ghi; HubCode là mã GV phát cho SV.
 *   Tab Students: A RegisteredAt | B StudentID | C FullName | D DOB | E Email | F Phone
 *                 G (Password — không còn ghi, để trống) | H..K mã 4 app | L ClassName
 *                 M Action | N AppStatus (JSON trạng thái từng app)
 *
 * Setup:
 * 1. Mở Master Sheet → Extensions → Apps Script → paste file này.
 * 2. Thêm cột F "ClassHubName" vào tab Classes (điền tên lớp trong Class-hub nếu lớp dùng Class-hub).
 * 3. Deploy → New deployment → Web App. Execute as: Me | Access: Anyone → copy URL vào index.html (HUB_API).
 */

// ── Endpoint Apps Script (/exec) của các app ───────────────────────────────────
var APP_URL = {
  speaking : 'https://script.google.com/macros/s/AKfycbzxCevaShEohE2PkyB71xyunBfF6YqpW9ESwvnEVVKfYTIc2FGUDXW4bIM_jUnXuYrWmA/exec',
  listening: 'https://script.google.com/macros/s/AKfycbyadq7DEYYcTNKILHotdXw7cCElBwggj4JGHJ3JD6tM07agn1CQq6aSklIwii5G0iiQ/exec',
  vocab    : 'https://script.google.com/macros/s/AKfycbwj-XE8zxBifrn7BgcbIGegqeeoKAPnYIBUPX7dOuCQozNQvkOgmS9bT3tC92W3kwoM/exec',
  writing  : 'https://script.google.com/macros/s/AKfycbxgVhsy3WKU-hL7rW7GZaNsn0B-z6zt6iH2Q-UlpbJVqP9koAE49P175m0tR3ISGp-m/exec'
};
// Web API key của từng Firebase project (công khai) — dùng để xác thực ID token của GV khi tạo lớp
var APP_API_KEY = {
  speaking : 'AIzaSyAeB-tcXD9QOkppW4oshpFI4aXe9T33kws',
  listening: 'AIzaSyABj5BoT_Bz8aGJ6bys8LWCLAFhut5VJL8',
  vocab    : 'AIzaSyAhoWEygnchnXPf1BaG2T6ZvpV0VY7oeeY',
  writing  : 'AIzaSyCj8WTr6eaqMGhqKltiZ9444LELV-7ZDIw'
};
var APP_LABEL = { speaking:'Speaking', listening:'Listening', vocab:'Vocab', writing:'Writing', classhub:'Class Hub' };

var CLASSES_TAB  = 'Classes';
var STUDENTS_TAB = 'Students';
var STUDENTS_HEADER = [
  'RegisteredAt','StudentID','FullName','DOB','Email',
  'Phone','Password','SpeakingCode','ListeningCode','VocabCode','WritingCode','ClassName','Action','AppStatus'
];
var ACTION_COL = 13; // M
var STATUS_COL = 14; // N

// ── Entry points ───────────────────────────────────────────────────────────────
function doGet(e)  { return respond(route(e.parameter || {})); }
function doPost(e) {
  var p = e.parameter || {};
  try {
    if (e.postData && e.postData.contents) p = JSON.parse(e.postData.contents);
  } catch (err) { return respond({ success: false, error: 'Bad request body' }); }
  return respond(route(p));
}

function respond(result) {
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function route(p) {
  try {
    switch (p.action) {
      case 'ping':       return { success: true, message: 'Middleware v5 running' };
      case 'lookupCode': return handleLookup(p.code || p.speakingCode);
      case 'checkEmail': return handleCheckEmail(p.email);
      case 'register':   return handleRegister(typeof p.data === 'string' ? JSON.parse(p.data) : p.data);
      case 'setStatus':  return handleSetStatus(p);
      case 'saveClass':  return handleSaveClass(p);
      default:           return { success: false, error: 'Unknown action: ' + p.action };
    }
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ── 1. Lookup class code ───────────────────────────────────────────────────────
// Chấp nhận HubCode hoặc mã của BẤT KỲ app nào có trong lớp (lớp không dạy Speaking vẫn dùng mã Listening...).
function handleLookup(code) {
  if (!code) return { success: false, error: 'No code provided' };
  var row = findClassRow(code.toString().trim().toUpperCase());
  if (!row) return { success: false, error: 'Class code not found. Please check with your teacher.' };
  return { success: true, className: row.className, apps: appsOf(row) };
}

// ── 2. Check duplicate email ───────────────────────────────────────────────────
function handleCheckEmail(email) {
  if (!email) return { success: false, error: 'No email provided' };
  email = email.toString().trim().toLowerCase();
  var data = getOrCreateStudentsSheet().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][4] || '').toString().trim().toLowerCase() === email) {
      return { success: false, error: 'This email is already registered. Please log in instead.' };
    }
  }
  return { success: true };
}

// ── 3. Register ────────────────────────────────────────────────────────────────
// Một StudentID + một mật khẩu cho mọi app của lớp. Gọi lại với cùng StudentID + email
// = thử lại: chỉ đăng ký lại các app chưa thành công.
function handleRegister(body) {
  body = body || {};
  body.code = body.code || body.speakingCode;
  var required = ['code','studentId','fullName','password','email','phone','dob'];
  for (var i = 0; i < required.length; i++) {
    if (!body[required[i]]) return { success: false, error: 'Missing field: ' + required[i] };
  }

  var row = findClassRow(body.code.toString().trim().toUpperCase());
  if (!row) return { success: false, error: 'Class code not found.' };

  var s = {
    studentId: body.studentId.toString().trim(),
    fullName : body.fullName.toString().trim(),
    password : body.password.toString(),
    email    : body.email.toString().trim().toLowerCase(),
    phone    : body.phone.toString().trim(),
    dob      : body.dob.toString().trim()
  };
  if (s.password.length < 6) return { success: false, error: 'Password must be at least 6 characters.' };

  // Trùng StudentID / email với một SV khác → từ chối. Trùng cả hai = thử lại.
  var sheet = getOrCreateStudentsSheet();
  var data  = sheet.getDataRange().getValues();
  var existingRow = 0, prev = {};
  for (var r = 1; r < data.length; r++) {
    var sameId    = (data[r][1] || '').toString().trim() === s.studentId;
    var sameEmail = (data[r][4] || '').toString().trim().toLowerCase() === s.email;
    if (sameId && sameEmail) {
      existingRow = r + 1;
      try { prev = JSON.parse(data[r][STATUS_COL - 1] || '{}'); } catch (e) {}
      break;
    }
    if (sameId)    return { success: false, error: 'Student ID ' + s.studentId + ' is already registered.' };
    if (sameEmail) return { success: false, error: 'This email is already registered. Please log in instead.' };
  }

  // Chỉ các app lớp này dùng, và chưa thành công ở lần trước
  var APPS = ['speaking','listening','vocab','writing'];
  var todo = APPS.filter(function (k) {
    return row[k + 'Code'] && !isOk(prev[k]);
  });

  var resps = todo.length ? UrlFetchApp.fetchAll(todo.map(function (k) { return buildRequest(k, row[k + 'Code'], s); })) : [];

  var status = {};
  APPS.forEach(function (k) { if (row[k + 'Code']) status[k] = prev[k] || ''; });
  todo.forEach(function (k, idx) { status[k] = classify(resps[idx]); });

  // Class-hub: trình duyệt sẽ đăng ký rồi gọi setStatus. Trạng thái ban đầu "pending".
  var needClassHub = !!row.classHubName && !isOk(prev.classhub);
  if (row.classHubName) status.classhub = needClassHub ? 'pending' : prev.classhub;

  // Ghi Master Sheet (không lưu mật khẩu)
  var line = [
    new Date().toISOString(), s.studentId, s.fullName, s.dob, s.email, s.phone, '',
    row.speakingCode, row.listeningCode, row.vocabCode, row.writingCode, row.className, '', JSON.stringify(status)
  ];
  if (existingRow) {
    // giữ nguyên thời điểm đăng ký ban đầu và cột Action
    line[0] = data[existingRow - 1][0];
    line[ACTION_COL - 1] = data[existingRow - 1][ACTION_COL - 1];
    sheet.getRange(existingRow, 1, 1, line.length).setValues([line]);
  } else {
    sheet.appendRow(line);
  }

  var failed = APPS.filter(function (k) { return status[k] !== undefined && !isOk(status[k]); });
  if (todo.length && !needClassHub && failed.length === todo.length) {
    return { success: false, error: 'Could not create any account. ' + describe(status), status: status };
  }
  return {
    success: true, className: row.className, status: status,
    classHubName: needClassHub ? row.classHubName : '',
    partial: failed.length > 0,
    message: failed.length ? 'Some apps need a retry: ' + describe(status) : 'Account created.'
  };
}

// Class-hub báo kết quả đăng ký (từ trình duyệt) → ghi vào cột AppStatus
function handleSetStatus(p) {
  var sid = (p.studentId || '').toString().trim(), email = (p.email || '').toString().trim().toLowerCase();
  if (!sid || !email) return { success: false, error: 'Missing studentId/email' };
  var st = (p.classhub === 'ok' || p.classhub === 'exists') ? p.classhub : ('error: ' + String(p.classhub || '').slice(0, 200));
  var sheet = getOrCreateStudentsSheet(), data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][1] || '').toString().trim() === sid && (data[i][4] || '').toString().trim().toLowerCase() === email) {
      var cur = {};
      try { cur = JSON.parse(data[i][STATUS_COL - 1] || '{}'); } catch (e) {}
      cur.classhub = st;
      sheet.getRange(i + 1, STATUS_COL).setValue(JSON.stringify(cur));
      return { success: true };
    }
  }
  return { success: false, error: 'Student not found in Master Sheet.' };
}

// ── 4. Create Class (GV) ───────────────────────────────────────────────────────
// Trang hub đã tạo lớp ở từng app bằng tài khoản GV; ở đây chỉ lưu bản đồ mã lớp và sinh HubCode.
// Chỉ nhận khi ID token là của một tài khoản GV (claim role=teacher) của một trong 4 app.
function handleSaveClass(p) {
  var v = verifyTeacher(p.idToken, p.project);
  if (!v.ok) return { success: false, error: v.error };
  var className = (p.className || '').toString().trim();
  if (!className) return { success: false, error: 'Missing class name.' };
  var codes = p.codes || {}, clean = {};
  ['speaking','listening','vocab','writing'].forEach(function (k) {
    var c = (codes[k] || '').toString().trim().toUpperCase();
    if (c && !/^[A-Z0-9-]{3,20}$/.test(c)) throw new Error('Bad ' + APP_LABEL[k] + ' code.');
    clean[k] = c;
  });
  if (!clean.speaking && !clean.listening && !clean.vocab && !clean.writing)
    return { success: false, error: 'No app code to save.' };
  var classHubName = (p.classHubName || '').toString().trim();

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CLASSES_TAB);
    if (!sheet) throw new Error('Sheet "' + CLASSES_TAB + '" not found.');
    var head = sheet.getRange(1, 1, 1, 8).getValues()[0];
    ['ClassName','SpeakingCode','ListeningCode','VocabCode','WritingCode','ClassHubName','TeacherEmail','HubCode']
      .forEach(function (h, i) { if (!head[i]) sheet.getRange(1, i + 1).setValue(h).setFontWeight('bold'); });

    var data = sheet.getDataRange().getValues(), used = {};

    // Thêm app còn thiếu vào lớp đã có (HubCode): chỉ điền ô đang TRỐNG, không ghi đè
    var existing = (p.hubCode || '').toString().trim().toUpperCase();
    if (existing) {
      for (var e = 1; e < data.length; e++) {
        if ((data[e][7] || '').toString().trim().toUpperCase() !== existing) continue;
        var cols = { speaking: 2, listening: 3, vocab: 4, writing: 5 }, added = [];
        Object.keys(cols).forEach(function (k) {
          if (clean[k] && !(data[e][cols[k] - 1] || '').toString().trim()) { sheet.getRange(e + 1, cols[k]).setValue(clean[k]); added.push(k); }
        });
        if (classHubName && !(data[e][5] || '').toString().trim()) sheet.getRange(e + 1, 6).setValue(classHubName);
        return { success: true, hubCode: existing, added: added };
      }
      return { success: false, error: 'Class code ' + existing + ' not found.' };
    }
    for (var i = 1; i < data.length; i++) for (var c = 1; c <= 4; c++) used[(data[i][c] || '').toString().trim().toUpperCase()] = 1;
    for (var j = 1; j < data.length; j++) used[(data[j][7] || '').toString().trim().toUpperCase()] = 1;
    var hubCode = '', CH = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
    for (var t = 0; t < 50 && !hubCode; t++) {
      var x = 'HUB-'; for (var n = 0; n < 6; n++) x += CH.charAt(Math.floor(Math.random() * CH.length));
      if (!used[x]) hubCode = x;
    }
    if (!hubCode) throw new Error('Could not generate a hub code.');
    sheet.appendRow([className, clean.speaking, clean.listening, clean.vocab, clean.writing, classHubName, v.email, hubCode]);
    return { success: true, hubCode: hubCode };
  } finally { lock.releaseLock(); }
}

// ID token → tài khoản GV hợp lệ của project đó? (Firebase Auth REST, bằng web API key)
function verifyTeacher(idToken, project) {
  if (!idToken || !APP_API_KEY[project]) return { ok: false, error: 'Missing teacher login.' };
  var r = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + APP_API_KEY[project], {
    method: 'post', contentType: 'application/json', payload: JSON.stringify({ idToken: idToken }), muteHttpExceptions: true
  });
  var j; try { j = JSON.parse(r.getContentText()); } catch (e) { return { ok: false, error: 'Could not verify teacher login.' }; }
  var u = (j.users || [])[0];
  if (!u) return { ok: false, error: 'Teacher login expired or invalid.' };
  var claims = {}; try { claims = JSON.parse(u.customAttributes || '{}'); } catch (e) {}
  // Speaking (Fluentalk) legacy teacher has uid 'teacher'
  if (claims.role !== 'teacher' && !(project === 'speaking' && u.localId === 'teacher')) return { ok: false, error: 'This is not a teacher account.' };
  return { ok: true, email: (u.email || '').toLowerCase() };
}

// ── Gọi endpoint đăng ký của từng app ──────────────────────────────────────────
function buildRequest(app, classCode, s) {
  var action, payload;
  if (app === 'writing') {
    action = 'fb.studentSignup';
    payload = { studentId: s.studentId, name: s.fullName, class: classCode, email: s.email,
                birthdate: s.dob, phone: s.phone, password: s.password };
  } else if (app === 'listening' || app === 'vocab') {
    action = 'fb.register';
    payload = { studentId: s.studentId, fullName: s.fullName, classId: classCode, email: s.email,
                phone: s.phone, birthdate: s.dob, password: s.password };
  } else { // speaking
    action = 'auth.register';
    payload = { studentId: s.studentId, fullName: s.fullName, classId: classCode, email: s.email,
                dob: s.dob, phone: s.phone, password: s.password };
  }
  return {
    url: APP_URL[app], method: 'post', contentType: 'text/plain;charset=utf-8',
    payload: JSON.stringify({ action: action, payload: payload }),
    muteHttpExceptions: true, followRedirects: true
  };
}

// 'ok' | 'exists' (Student ID đã đăng ký rồi — coi như xong) | 'error: ...'
function classify(resp) {
  var j;
  try { j = JSON.parse(resp.getContentText()); }
  catch (e) { return 'error: HTTP ' + resp.getResponseCode() + ' (not JSON)'; }
  if (j && j.success) return 'ok';
  var msg = String((j && (j.error || j.message)) || 'unknown error');
  if (/student id|mã sv/i.test(msg) && /đã được đăng ký|đã tồn tại|already/i.test(msg)) return 'exists';
  return 'error: ' + msg.slice(0, 200);
}
function isOk(v) { return v === 'ok' || v === 'exists'; }
function describe(status) {
  return Object.keys(status).filter(function (k) { return !isOk(status[k]) && status[k] !== 'pending'; })
    .map(function (k) { return APP_LABEL[k] + ' (' + String(status[k]).replace(/^error: /, '') + ')'; }).join(' | ');
}

// ── Helpers ────────────────────────────────────────────────────────────────────
// Tab Classes: A ClassName | B Speaking | C Listening | D Vocab | E Writing | F ClassHubName | G TeacherEmail | H HubCode
function findClassRow(code) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CLASSES_TAB);
  if (!sheet) throw new Error('Sheet "' + CLASSES_TAB + '" not found.');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var codes = [1, 2, 3, 4].map(function (c) { return (data[i][c] || '').toString().trim(); });
    var hubCode = (data[i][7] || '').toString().trim();
    if (codes.some(function (c) { return c && c.toUpperCase() === code; }) || (hubCode && hubCode.toUpperCase() === code)) {
      return {
        className    : data[i][0],
        speakingCode : codes[0], listeningCode: codes[1], vocabCode: codes[2], writingCode: codes[3],
        classHubName : (data[i][5] || '').toString().trim(),
        hubCode      : hubCode
      };
    }
  }
  return null;
}

function appsOf(row) {
  var out = [];
  if (row.speakingCode)  out.push('speaking');
  if (row.listeningCode) out.push('listening');
  if (row.vocabCode)     out.push('vocab');
  if (row.writingCode)   out.push('writing');
  if (row.classHubName)  out.push('classhub');
  return out;
}

function getOrCreateStudentsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(STUDENTS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(STUDENTS_TAB);
    sheet.appendRow(STUDENTS_HEADER);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, STUDENTS_HEADER.length).setFontWeight('bold');
  } else if (sheet.getRange(1, STATUS_COL).getValue() === '') {
    sheet.getRange(1, STATUS_COL).setValue('AppStatus').setFontWeight('bold');
  }
  return sheet;
}

// ═══════════════════════════════════════════════════════════════════
// SYNC DELETE — chạy tự động mỗi 5 phút (Triggers → syncDelete)
// ═══════════════════════════════════════════════════════════════════
// Gõ DELETE vào cột M (Action) của tab Students → dòng đó bị xoá khỏi Master Sheet.
// LƯU Ý: từ v5 hàm này CHỈ xoá dòng trong Master Sheet. Tài khoản trong Firebase của
// từng app (và dữ liệu học) phải xoá/khoá trong app tương ứng — các app đã chuyển
// sang Firebase nên không còn Google Sheet để xoá.
function syncDelete() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STUDENTS_TAB);
  if (!sheet) return;
  var data = sheet.getDataRange().getValues(), rows = [];
  for (var i = 1; i < data.length; i++) {
    if ((data[i][ACTION_COL - 1] || '').toString().trim().toUpperCase() === 'DELETE') rows.push(i + 1);
  }
  for (var j = rows.length - 1; j >= 0; j--) sheet.deleteRow(rows[j]);
  if (rows.length) Logger.log('syncDelete: removed ' + rows.length + ' row(s) from Master Sheet.');
}

function runSyncDeleteNow() {
  syncDelete();
  SpreadsheetApp.getUi().alert('Done. Only the Master Sheet rows were removed; Firebase accounts are unchanged.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Hub Tools').addItem('🗑 Sync Delete Now', 'runSyncDeleteNow').addToUi();
}

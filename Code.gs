/**
 * ============================================================
 *  Madlouh Medical Group — Nursing Assessment System
 *  Backend: Google Apps Script (Web App / API)
 *  Sheets:  Users , Submissions
 * ============================================================
 *  Deploy as Web App:
 *    Deploy > New deployment > Type: Web app
 *    Execute as: Me
 *    Who has access: Anyone
 *  Copy the /exec URL into API_URL inside index.html
 * ============================================================
 */

// ---------- CONFIG ----------
const SPREADSHEET_ID   = '1yo5KzIn_nqZ50YIJX_LobWQiT0fl1d1vYn_1cdlDwRU';
const SHEET_USERS       = 'Users';
const SHEET_SUBMISSIONS = 'Submissions';
const OTP_TTL_MINUTES    = 5;
const SESSION_TTL_HOURS  = 12;
const ADMIN_EMAIL        = 'it.madlouh@gmail.com';
const SENDER_NAME        = 'Madlouh Medical Complex';

// Leave blank to auto-create/find a Drive folder named below the first time
// a signature is saved. Or paste a specific Drive folder ID to force it.
const SIGNATURE_FOLDER_ID   = '';
const SIGNATURE_FOLDER_NAME = 'Madlouh Nursing Assessment Signatures';

// ---------- ENTRY POINTS ----------
function doGet(e) {
  return handle(e);
}

function doPost(e) {
  return handle(e);
}

function handle(e) {
  var params = (e && e.parameter) || {};
  var action = params.action || '';
  var body = {};

  // submitForm is the only action sent as POST (it carries a large signature
  // image, which does not fit in a URL). Every other action is sent as a
  // plain GET with query parameters — GET requests are not affected by the
  // "POST becomes GET on redirect, body is lost" issue that some browsers /
  // proxies apply to Apps Script's internal redirect, so this is the most
  // reliable way to call this web app for small payloads.
  if (e && e.postData && e.postData.contents) {
    try {
      body = JSON.parse(e.postData.contents);
    } catch (err) {
      return json({ success: false, code: 'bad_request', message: 'Invalid JSON body' });
    }
    if (!action) action = body.action || '';
  } else {
    body = params;
  }

  // Read-only actions do not need the script lock — locking them only adds
  // latency and can make them wait behind unrelated write operations.
  var READ_ONLY_ACTIONS = { checkSession: true, listUsers: true, exportSubmissions: true, getStats: true };
  var needsLock = !READ_ONLY_ACTIONS[action];

  var lock = null;
  if (needsLock) {
    lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
    } catch (err) {
      return json({ success: false, code: 'busy', message: 'النظام مشغول، حاول مرة أخرى.' });
    }
  }

  try {
    return json(route(action, body));
  } catch (err) {
    // Never let an uncaught exception fall through — Apps Script would then
    // return an HTML error page instead of JSON, which breaks the client.
    Logger.log('Unhandled error in action "' + action + '": ' + err + (err && err.stack ? '\n' + err.stack : ''));
    return json({ success: false, code: 'server_error', message: 'حدث خطأ في الخادم. حاول مرة أخرى.' });
  } finally {
    if (lock) lock.releaseLock();
  }
}

function route(action, body) {
  switch (action) {
    case 'register':
      return registerUser(body.fullName, body.email);
    case 'requestOtp':
      return requestOtp(body.email);
    case 'verifyOtp':
      return verifyOtp(body.email, body.otp);
    case 'checkSession':
      return checkSession(body.token);
    case 'submitForm':
      return submitForm(body.token, body.formData, body.signature);
    case 'listUsers':
      return listUsers(body.token);
    case 'setUserStatus':
      return setUserStatus(body.token, body.targetEmail, body.status);
    case 'exportSubmissions':
      return exportSubmissions(body.token);
    case 'getStats':
      return getStats(body.token);
    default:
      return { success: false, code: 'unknown_action', message: 'Unknown action: ' + action };
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- HELPERS ----------
function ss() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function usersSheet() {
  return ss().getSheetByName(SHEET_USERS);
}

function submissionsSheet() {
  return ss().getSheetByName(SHEET_SUBMISSIONS);
}

// Users columns: Timestamp | FullName | Email | Role | Status | OTP | OTPExpiry | SessionToken | SessionExpiry | LastLogin
function findUserRowByEmail(email) {
  var sheet = usersSheet();
  var data = sheet.getDataRange().getValues();
  email = String(email || '').trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2] || '').trim().toLowerCase() === email) {
      return { rowIndex: i + 1, row: data[i], headers: data[0] };
    }
  }
  return null;
}

function colIndex(headers, name) {
  return headers.indexOf(name);
}

function generateOtp() {
  return String(Math.floor(100000 + Math.random() * 900000));
}

function generateToken() {
  return Utilities.getUuid() + '-' + new Date().getTime();
}

// ---------- REGISTRATION ----------
function registerUser(fullName, email) {
  fullName = String(fullName || '').trim();
  email = String(email || '').trim().toLowerCase();
  if (!fullName || !email) {
    return { success: false, code: 'invalid_fields', message: 'يرجى إدخال الاسم والبريد الإلكتروني' };
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { success: false, code: 'invalid_email', message: 'صيغة البريد الإلكتروني غير صحيحة' };
  }

  var existing = findUserRowByEmail(email);
  if (existing) {
    return { success: false, code: 'already_registered', message: 'هذا البريد مسجل بالفعل. يرجى تسجيل الدخول أو انتظار موافقة الإدارة.' };
  }

  var sheet = usersSheet();
  sheet.appendRow([
    new Date(),   // Timestamp
    fullName,     // FullName
    email,        // Email
    'Nurse',      // Role
    'Pending',    // Status
    '', '',       // OTP, OTPExpiry
    '', '',       // SessionToken, SessionExpiry
    ''            // LastLogin
  ]);

  // Notify admin
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      name: SENDER_NAME,
      subject: 'طلب تسجيل جديد - نظام التمريض',
      body: 'يوجد طلب تسجيل جديد بانتظار الموافقة:\n\nالاسم: ' + fullName + '\nالبريد الإلكتروني: ' + email +
            '\n\nيرجى الدخول إلى شيت Users وتغيير الحالة (Status) إلى Approved لتفعيل الحساب.'
    });
  } catch (err) {
    // Email failure should not block registration
  }

  return { success: true, code: 'registered_ok', message: 'تم إرسال طلب التسجيل. سيتم تفعيل حسابك من قبل الإدارة قريباً.' };
}

// ---------- LOGIN: STEP 1 (REQUEST OTP) ----------
function requestOtp(email) {
  email = String(email || '').trim().toLowerCase();
  var found = findUserRowByEmail(email);
  if (!found) {
    return { success: false, code: 'email_not_found', message: 'هذا البريد غير مسجل. يرجى إنشاء حساب جديد.' };
  }

  var headers = found.headers;
  var status = found.row[colIndex(headers, 'Status')];
  if (status === 'Disabled') {
    return { success: false, code: 'account_disabled', message: 'تم تعطيل هذا الحساب من قبل الإدارة.' };
  }
  if (status !== 'Approved') {
    return { success: false, code: 'pending_approval', message: 'حسابك بانتظار موافقة الإدارة.' };
  }

  var otp = generateOtp();
  var expiry = new Date(new Date().getTime() + OTP_TTL_MINUTES * 60000);

  var sheet = usersSheet();
  sheet.getRange(found.rowIndex, colIndex(headers, 'OTP') + 1).setValue(otp);
  sheet.getRange(found.rowIndex, colIndex(headers, 'OTPExpiry') + 1).setValue(expiry);

  try {
    MailApp.sendEmail({
      to: email,
      name: SENDER_NAME,
      subject: 'رمز الدخول (OTP) - نظام التمريض',
      body: 'رمز الدخول الخاص بك هو: ' + otp + '\n\nصالح لمدة ' + OTP_TTL_MINUTES + ' دقائق فقط.\n\nإذا لم تطلب هذا الرمز، تجاهل هذه الرسالة.'
    });
  } catch (err) {
    return { success: false, code: 'email_send_failed', message: 'تعذر إرسال البريد الإلكتروني. تواصل مع الإدارة.' };
  }

  return { success: true, code: 'otp_sent', message: 'تم إرسال رمز الدخول إلى بريدك الإلكتروني.' };
}

// ---------- LOGIN: STEP 2 (VERIFY OTP) ----------
function verifyOtp(email, otp) {
  email = String(email || '').trim().toLowerCase();
  otp = String(otp || '').trim();
  var found = findUserRowByEmail(email);
  if (!found) {
    return { success: false, code: 'email_not_found', message: 'هذا البريد غير مسجل.' };
  }

  var headers = found.headers;
  var savedOtp = String(found.row[colIndex(headers, 'OTP')] || '');
  var otpExpiry = found.row[colIndex(headers, 'OTPExpiry')];

  if (!savedOtp || savedOtp !== otp) {
    return { success: false, code: 'otp_invalid', message: 'رمز الدخول غير صحيح.' };
  }
  if (!otpExpiry || new Date(otpExpiry).getTime() < new Date().getTime()) {
    return { success: false, code: 'otp_expired', message: 'انتهت صلاحية الرمز. اطلب رمزاً جديداً.' };
  }

  var token = generateToken();
  var sessionExpiry = new Date(new Date().getTime() + SESSION_TTL_HOURS * 3600000);
  var sheet = usersSheet();
  sheet.getRange(found.rowIndex, colIndex(headers, 'SessionToken') + 1).setValue(token);
  sheet.getRange(found.rowIndex, colIndex(headers, 'SessionExpiry') + 1).setValue(sessionExpiry);
  sheet.getRange(found.rowIndex, colIndex(headers, 'LastLogin') + 1).setValue(new Date());
  // clear OTP after use
  sheet.getRange(found.rowIndex, colIndex(headers, 'OTP') + 1).setValue('');
  sheet.getRange(found.rowIndex, colIndex(headers, 'OTPExpiry') + 1).setValue('');

  return {
    success: true,
    code: 'login_ok',
    token: token,
    fullName: found.row[colIndex(headers, 'FullName')],
    role: found.row[colIndex(headers, 'Role')],
    email: email
  };
}

// ---------- SESSION CHECK ----------
function checkSession(token) {
  if (!token) return { success: false };
  var sheet = usersSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var tCol = colIndex(headers, 'SessionToken');
  var eCol = colIndex(headers, 'SessionExpiry');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][tCol]) === String(token)) {
      var expiry = data[i][eCol];
      if (expiry && new Date(expiry).getTime() > new Date().getTime()) {
        return {
          success: true,
          code: 'session_ok',
          fullName: data[i][colIndex(headers, 'FullName')],
          role: data[i][colIndex(headers, 'Role')],
          email: data[i][colIndex(headers, 'Email')]
        };
      }
      return { success: false, code: 'session_expired', message: 'انتهت الجلسة. سجل الدخول مرة أخرى.' };
    }
  }
  return { success: false, code: 'session_invalid', message: 'جلسة غير صالحة.' };
}

// ---------- ADMIN: LIST / APPROVE / DISABLE USERS ----------
function requireAdmin(token) {
  var session = checkSession(token);
  if (!session.success) {
    return { ok: false, response: { success: false, code: 'must_login', message: 'يجب تسجيل الدخول أولاً.' } };
  }
  if (session.role !== 'Admin') {
    return { ok: false, response: { success: false, code: 'not_admin', message: 'غير مصرح لك بالوصول لهذه الصفحة.' } };
  }
  return { ok: true, session: session };
}

function listUsers(token) {
  var check = requireAdmin(token);
  if (!check.ok) return check.response;

  var sheet = usersSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var users = [];
  for (var i = 1; i < data.length; i++) {
    users.push({
      fullName: data[i][colIndex(headers, 'FullName')],
      email: data[i][colIndex(headers, 'Email')],
      role: data[i][colIndex(headers, 'Role')],
      status: data[i][colIndex(headers, 'Status')],
      lastLogin: data[i][colIndex(headers, 'LastLogin')] ? String(data[i][colIndex(headers, 'LastLogin')]) : ''
    });
  }
  return { success: true, users: users };
}

// newStatus must be one of: Approved, Pending, Disabled
function setUserStatus(token, targetEmail, newStatus) {
  var check = requireAdmin(token);
  if (!check.ok) return check.response;

  var allowed = ['Approved', 'Pending', 'Disabled'];
  if (allowed.indexOf(newStatus) === -1) {
    return { success: false, message: 'حالة غير صحيحة.' };
  }

  var found = findUserRowByEmail(targetEmail);
  if (!found) {
    return { success: false, message: 'المستخدم غير موجود.' };
  }
  if (String(found.row[colIndex(found.headers, 'Role')]) === 'Admin' && newStatus !== 'Approved') {
    return { success: false, message: 'لا يمكن تعطيل حساب أدمن.' };
  }

  var headers = found.headers;
  var sheet = usersSheet();
  sheet.getRange(found.rowIndex, colIndex(headers, 'Status') + 1).setValue(newStatus);
  if (newStatus !== 'Approved') {
    // Force logout of any active session for this user
    sheet.getRange(found.rowIndex, colIndex(headers, 'SessionToken') + 1).setValue('');
    sheet.getRange(found.rowIndex, colIndex(headers, 'SessionExpiry') + 1).setValue('');
  }
  return { success: true, code: 'status_updated', message: 'تم تحديث حالة الحساب.' };
}

// ---------- DOWNLOAD: EXPORT SUBMISSIONS AS EXCEL ----------
// Any logged-in user can download the recorded data (not admin-only), per
// how this system is used day to day.
function exportSubmissions(token) {
  var session = checkSession(token);
  if (!session.success) {
    return { success: false, code: 'must_login', message: 'يجب تسجيل الدخول أولاً.' };
  }

  try {
    var sheetId = submissionsSheet().getSheetId();
    var url = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/export?format=xlsx&gid=' + sheetId;
    var oauthToken = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + oauthToken },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      Logger.log('exportSubmissions HTTP ' + response.getResponseCode() + ': ' + response.getContentText());
      return { success: false, code: 'export_failed', message: 'تعذر تجهيز ملف الإكسيل.' };
    }

    var base64 = Utilities.base64Encode(response.getBlob().getBytes());
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT', 'yyyyMMdd_HHmm');
    return { success: true, base64: base64, fileName: 'Nursing_Assessment_Data_' + stamp + '.xlsx' };
  } catch (err) {
    Logger.log('exportSubmissions error: ' + err);
    return { success: false, code: 'export_failed', message: 'تعذر تجهيز ملف الإكسيل.' };
  }
}

// ---------- DASHBOARD: SUBMISSION STATISTICS ----------
function getStats(token) {
  var session = checkSession(token);
  if (!session.success) {
    return { success: false, code: 'must_login', message: 'يجب تسجيل الدخول أولاً.' };
  }

  var sheet = submissionsSheet();
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return { success: true, total: 0, charts: {} };
  }

  var headers = data[0];
  var rows = data.slice(1);

  var FIELDS = ['Fever', 'Rash', 'CoughSOB', 'FallRiskScore', 'LanguageSpoken', 'ReasonForVisit', 'ModeOfAccess', 'Allergies'];
  var charts = {};
  FIELDS.forEach(function (field) {
    var idx = colIndex(headers, field);
    var counts = {};
    rows.forEach(function (row) {
      var value = String(idx >= 0 ? row[idx] : '').trim();
      if (!value) value = 'N/A';
      counts[value] = (counts[value] || 0) + 1;
    });
    charts[field] = counts;
  });

  return { success: true, total: rows.length, charts: charts };
}

// ---------- FORM SUBMISSION ----------
// formData is a flat object whose keys MATCH the Submissions sheet header names.
function submitForm(token, formData, signatureBase64) {
  var session = checkSession(token);
  if (!session.success) {
    return { success: false, code: 'must_login', message: 'يجب تسجيل الدخول أولاً.' };
  }

  var submissionId = Utilities.getUuid();
  var signatureUrl = '';
  if (signatureBase64) {
    try {
      signatureUrl = saveSignatureToDrive(signatureBase64, submissionId, formData);
    } catch (err) {
      // If Drive save fails, keep going but leave the field blank rather than
      // losing the whole submission; log for troubleshooting.
      Logger.log('Signature save failed: ' + err);
    }
  }

  var sheet = submissionsSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = new Array(headers.length).fill('');

  for (var i = 0; i < headers.length; i++) {
    var h = headers[i];
    if (h === 'SubmissionID') row[i] = submissionId;
    else if (h === 'Timestamp') row[i] = new Date();
    else if (h === 'SubmittedBy') row[i] = session.fullName;
    else if (h === 'SubmittedByEmail') row[i] = session.email;
    else if (h === 'PatientSignature') row[i] = signatureUrl;
    else if (formData && Object.prototype.hasOwnProperty.call(formData, h)) row[i] = formData[h];
  }

  sheet.appendRow(row);
  return { success: true, code: 'submit_ok', message: 'تم حفظ النموذج بنجاح.' };
}

// Decodes a "data:image/png;base64,...." string, saves it as a PNG in the
// signatures Drive folder, sets it viewable by link, and returns the file URL.
function saveSignatureToDrive(dataUrl, submissionId, formData) {
  var match = String(dataUrl).match(/^data:(image\/\w+);base64,(.*)$/);
  if (!match) return '';
  var mimeType = match[1];
  var base64 = match[2];
  var bytes = Utilities.base64Decode(base64);

  var fileNo = (formData && formData.FileNo) ? String(formData.FileNo) : 'unknown';
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT', 'yyyyMMdd_HHmmss');
  var fileName = 'signature_' + fileNo + '_' + stamp + '_' + submissionId.substring(0, 8) + '.png';

  var blob = Utilities.newBlob(bytes, mimeType, fileName);
  var folder = getOrCreateSignatureFolder();
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getOrCreateSignatureFolder() {
  if (SIGNATURE_FOLDER_ID) {
    return DriveApp.getFolderById(SIGNATURE_FOLDER_ID);
  }
  var folders = DriveApp.getFoldersByName(SIGNATURE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(SIGNATURE_FOLDER_NAME);
}

// ---------- ONE-TIME SETUP HELPER ----------
// Run this once manually from the Apps Script editor (select function > Run)
// to create/reset the two sheets with correct headers, if you did not upload
// the pre-built Excel file.
function setupSheets() {
  var spreadsheet = ss();

  var usersHeaders = ['Timestamp', 'FullName', 'Email', 'Role', 'Status', 'OTP', 'OTPExpiry', 'SessionToken', 'SessionExpiry', 'LastLogin'];
  var users = spreadsheet.getSheetByName(SHEET_USERS) || spreadsheet.insertSheet(SHEET_USERS);
  users.clear();
  users.appendRow(usersHeaders);
  users.appendRow([new Date(), 'IT Admin', ADMIN_EMAIL, 'Admin', 'Approved', '', '', '', '', '']);

  var subsHeaders = [
    'SubmissionID','Timestamp','FileNo','PatientName',
    'Fever','Rash','CoughSOB','TraveledOutsideKSA','TravelWhenDetails','HealthcareWorkerExposure','CloseContactSimilarSymptoms',
    'ActionSurgicalMask','ActionIsolationRoom','ActionWaitingAreaSeparation','ActionHandHygiene','ActionPrivateRoom',
    'ReceivedFrom','FirstVisit','LanguageSpoken','HistoryTakenFrom','Allergies',
    'ModeOfAccess','ReasonForVisit','Accompanied','InterpreterNeed','Diet','Transportation',
    'ChiefComplaint','Temp','TempMethod','BPSystolic','BPDiastolic','Pulse','Height','Skin','RR','Weight',
    'WeightChangePast6Months','FunctionalStatus','RecentCareChanges',
    'PsychologicalStatus','LiveAlone','CareGiverAvailable','PainScreening',
    'FallPast3Months','DifficultyWalking','UseAssistiveDevice','ReceiveFallRiskMedication','FallRiskMedications',
    'FallRiskScore','InterventionYellowWristband','InterventionDesignatedClinic','InterventionWheelChair','InterventionFallsPamphlet',
    'SubmittedBy','SubmittedByEmail','PatientSignature'
  ];
  var subs = spreadsheet.getSheetByName(SHEET_SUBMISSIONS) || spreadsheet.insertSheet(SHEET_SUBMISSIONS);
  subs.clear();
  subs.appendRow(subsHeaders);

  SpreadsheetApp.flush();
}

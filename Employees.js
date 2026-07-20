function getMasterSheet_() {
  var sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!sheetId) throw new Error('SHEET_ID not set. Run runFullSetup() first.');
  return SpreadsheetApp.openById(sheetId);
}

function getEmployeesSheet_() { return getMasterSheet_().getSheetByName('Employees'); }
function getClaimsSheet_() { return getMasterSheet_().getSheetByName('Claims'); }
function getPolicySheet_() { return getMasterSheet_().getSheetByName('Policy'); }
function getLogsSheet_() { return getMasterSheet_().getSheetByName('Logs'); }

function splitRoles_(rolesStr) {
  return String(rolesStr || '')
    .split(',')
    .map(function (r) { return r.trim(); })
    .filter(Boolean);
}

function hasRole_(rolesArr, role) {
  return rolesArr.indexOf(role) !== -1;
}

/** Reads Employees tab into an array of row objects keyed by header. */
function readEmployeeRows_() {
  var sheet = getEmployeesSheet_();
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = {};
    for (var c = 0; c < headers.length; c++) row[headers[c]] = values[i][c];
    row._rowIndex = i + 1;
    rows.push(row);
  }
  return rows;
}

function findEmployeeByEmail_(email) {
  if (!email) return null;
  var lower = email.toLowerCase();
  var rows = readEmployeeRows_();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['Email']).toLowerCase() === lower) return rows[i];
  }
  return null;
}

function getEmployeesWithRole_(role) {
  return readEmployeeRows_()
    .filter(function (r) { return r['Status'] === 'Active' && hasRole_(splitRoles_(r['Role']), role); })
    .map(function (r) { return r['Email']; });
}

// ---------- Auth token (non-Google login fallback) ----------

var TOKEN_TTL_MS_ = 30 * 24 * 60 * 60 * 1000; // 30 days

function getOtpSecret_() {
  var secret = PropertiesService.getScriptProperties().getProperty('OTP_SECRET');
  if (!secret) throw new Error('OTP_SECRET not set. Run migrateSchemaV2() first.');
  return secret;
}

function signPayload_(payload) {
  var bytes = Utilities.computeHmacSha256Signature(payload, getOtpSecret_());
  return Utilities.base64EncodeWebSafe(bytes);
}

function generateToken_(email) {
  var expiry = Date.now() + TOKEN_TTL_MS_;
  var payload = email + '|' + expiry;
  var sig = signPayload_(payload);
  return Utilities.base64EncodeWebSafe(payload) + '.' + sig;
}

function verifyToken_(token) {
  if (!token || token.indexOf('.') === -1) return null;
  var parts = token.split('.');
  var payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
  var expectedSig = signPayload_(payload);
  if (expectedSig !== parts[1]) return null;
  var pieces = payload.split('|');
  var email = pieces[0];
  var expiry = Number(pieces[1]);
  if (Date.now() > expiry) return null;
  return email;
}

/** Resolves the caller's email: Google session first, then a verified OTP token. */
function resolveIdentity_(token) {
  var email = Session.getActiveUser().getEmail();
  if (email) return email;
  if (token) return verifyToken_(token);
  return null;
}

function requireActiveEmployee_(token) {
  var email = resolveIdentity_(token);
  if (!email) throw new Error('Not signed in.');
  var employee = findEmployeeByEmail_(email);
  if (!employee || employee['Status'] !== 'Active') throw new Error('Account not recognized or inactive: ' + email);
  employee._roles = splitRoles_(employee['Role']);
  employee._email = email;
  return employee;
}

function buildIdentityPayload_(email) {
  if (!email) return { success: true, signedIn: false };
  var employee = findEmployeeByEmail_(email);
  if (!employee) return { success: true, signedIn: true, known: false, email: email };
  if (employee['Status'] !== 'Active') return { success: true, signedIn: true, known: true, active: false, email: email };
  return {
    success: true,
    signedIn: true,
    known: true,
    active: true,
    email: email,
    employeeId: employee['Employee ID'],
    name: employee['Name'],
    managerEmail: employee['Manager Email'],
    brandName: employee['Brand Name'] || '',
    department: employee['Department'] || '',
    roles: splitRoles_(employee['Role'])
  };
}

function handleWhoAmI(token) {
  return buildIdentityPayload_(resolveIdentity_(token));
}

function handleRequestOtp(email) {
  if (!email) throw new Error('Email is required.');
  var employee = findEmployeeByEmail_(email);
  if (!employee || employee['Status'] !== 'Active') throw new Error('This email is not a recognized active account. Contact your Admin.');
  var otp = String(Math.floor(100000 + Math.random() * 900000));
  CacheService.getScriptCache().put('otp_' + email.toLowerCase(), otp, 300);
  MailApp.sendEmail({
    to: email,
    subject: 'Your Curefoods Reimbursement Portal login code',
    body: 'Your login code is ' + otp + '. It expires in 5 minutes.\n\nIf you did not request this, you can ignore this email.'
  });
  return { success: true };
}

function handleVerifyOtp(email, otp) {
  if (!email || !otp) throw new Error('Email and code are required.');
  var cached = CacheService.getScriptCache().get('otp_' + email.toLowerCase());
  if (!cached || cached !== String(otp).trim()) throw new Error('Incorrect or expired code.');
  CacheService.getScriptCache().remove('otp_' + email.toLowerCase());
  var identity = buildIdentityPayload_(email);
  identity.token = generateToken_(email);
  return identity;
}

function getMasterSheet_() {
  var sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!sheetId) throw new Error('SHEET_ID not set. Run runFullSetup() first.');
  return SpreadsheetApp.openById(sheetId);
}

function getEmployeesSheet_() { return getMasterSheet_().getSheetByName('Employees Data'); }
function getClaimsSheet_() { return getMasterSheet_().getSheetByName('Claims'); }
function getPolicySheet_() { return getMasterSheet_().getSheetByName('Policy'); }
function getTravelPolicySheet_() { return getMasterSheet_().getSheetByName('Travel policy'); }
function getLogsSheet_() { return getMasterSheet_().getSheetByName('Logs'); }

function splitCommaList_(str) {
  return String(str || '')
    .split(',')
    .map(function (s) { return s.trim(); })
    .filter(Boolean);
}

function splitRoles_(rolesStr) {
  return splitCommaList_(rolesStr);
}

function hasRole_(rolesArr, role) {
  return rolesArr.indexOf(role) !== -1;
}

/**
 * Reads the "Employees Data" tab (the real HRMS export, ~3600 rows across
 * Curefoods/Cakezone/Sharief Bhai/Frozen Bottle) into row objects normalized
 * to the canonical keys the rest of the app uses (Employee ID, Name, Email,
 * Manager Email, Brand Name, Department, Grade, Role, Status), regardless of
 * that sheet's own column headers (Emp Code, Emp Name, Official E-Mail ID,
 * Manager email, Company, ...). Grade is new — used by the Travel policy
 * engine to resolve an employee's Hotel/Food DA tier.
 */
function readEmployeeRows_() {
  var sheet = getEmployeesSheet_();
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var idx = {};
  headers.forEach(function (h, c) { idx[h] = c; });
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var raw = values[i];
    if (!raw[idx['Emp Code']]) continue; // skip blank rows
    var row = {
      'Employee ID': raw[idx['Emp Code']],
      'Name': raw[idx['Emp Name']],
      'Email': raw[idx['Official E-Mail ID']],
      'Manager Email': raw[idx['Manager email']],
      'Brand Name': raw[idx['Company']] || '',
      'Department': raw[idx['Department']] || '',
      'Grade': raw[idx['Grade']] || '',
      'Role': raw[idx['Role']] || '',
      'Status': raw[idx['Status']] || '',
      'Finance Company': raw[idx['Finance Company']] || ''
    };
    row._rowIndex = i + 1;
    rows.push(row);
  }
  return rows;
}

/**
 * requireActiveEmployee_ (and therefore this) runs on every single authenticated
 * server call, and readEmployeeRows_ is a full ~3600-row sheet scan — the single
 * most repeated, most expensive operation in the app. Caching the per-email result
 * cuts that to near-zero for the common case of one user firing several calls in
 * quick succession (loading a tab, switching subtabs, submitting). Kept short
 * (seconds, not the 5-30 min typical for this kind of cache) so a manual
 * deactivation/edit in Employees Data takes effect almost immediately rather than
 * being masked for a long window. Misses are cached too (as '') so a bad/stale
 * token doesn't keep re-scanning the sheet either.
 */
var EMPLOYEE_LOOKUP_CACHE_TTL_ = 60;

function findEmployeeByEmail_(email) {
  if (!email) return null;
  var lower = email.toLowerCase();
  var cache = CacheService.getScriptCache();
  var cacheKey = 'emp_' + lower;
  var cached = cache.get(cacheKey);
  if (cached !== null) return cached === '' ? null : JSON.parse(cached);
  var rows = readEmployeeRows_();
  var found = null;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]['Email']).toLowerCase() === lower) { found = rows[i]; break; }
  }
  cache.put(cacheKey, found ? JSON.stringify(found) : '', EMPLOYEE_LOOKUP_CACHE_TTL_);
  return found;
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

/**
 * Resolves the caller's email: a verified OTP token first, then the Google
 * session. Token takes priority because it's an explicit choice the person
 * made (e.g. someone signed into a @curefoods.in Google account that isn't
 * their real employee record - a Cakezone/Sharief Bhai person with a
 * curefoods.in email too - used the "sign in with a different email" OTP
 * flow instead). Without this, the ambient Google session would win again
 * on every later visit and silently undo their choice.
 */
function resolveIdentity_(token) {
  if (token) {
    var tokenEmail = verifyToken_(token);
    if (tokenEmail) return tokenEmail;
  }
  return Session.getActiveUser().getEmail() || null;
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
    grade: employee['Grade'] || '',
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

var CLAIM_COLS_ = {
  claimId: 1, employeeId: 2, employeeName: 3, employeeEmail: 4, managerEmail: 5,
  category: 6, amount: 7, description: 8, expenseDate: 9, status: 10,
  submittedAt: 11, managerRemarks: 12, managerActionAt: 13, adminActionAt: 14,
  transactionRef: 15, paidAt: 16, receiptUrl: 17, overPolicy: 18, batchId: 19,
  resubmittedFrom: 20, brandName: 21, department: 22,
  returnReason: 23, returnedBy: 24, returnedAt: 25
};

/**
 * Claim IDs must be unique even when two different people submit at nearly the same
 * moment. Deriving the ID from sheet.getLastRow() (the old approach) is not atomic —
 * two concurrent submissions can both read the same row count before either has
 * appended, producing the same ID (confirmed live: two real claims both got
 * CFI-2026-00008, which then broke manager approve/reject for one of them since
 * claim lookup-by-ID always finds the first matching row). A persistent counter in
 * Script Properties, incremented under a script-wide lock, is atomic regardless of
 * how many people submit at once.
 */
function nextClaimId_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var year = new Date().getFullYear();
    var key = 'CLAIM_COUNTER_' + year;
    var props = PropertiesService.getScriptProperties();
    var next = Number(props.getProperty(key) || '0') + 1;
    props.setProperty(key, String(next));
    return 'CFI-' + year + '-' + String(next).padStart(5, '0');
  } finally {
    lock.releaseLock();
  }
}

/**
 * google.script.run's client bridge can silently deliver `null` instead of the real
 * response when the payload contains raw Date objects nested inside an array of
 * objects (confirmed: direct Apps Script calls always returned valid data; only the
 * real web-app round-trip failed). Converting dates to ISO strings avoids that.
 */
function sanitizeForClient_(value) {
  if (value instanceof Date) return value.toISOString();
  return value;
}

function readClaimRows_() {
  var sheet = getClaimsSheet_();
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = {};
    for (var c = 0; c < headers.length; c++) row[headers[c]] = sanitizeForClient_(values[i][c]);
    row._rowIndex = i + 1;
    rows.push(row);
  }
  return rows;
}

/** Finds one claim by ID within an already-fetched row array — avoids a fresh full-sheet readClaimRows_() for every claim ID in a bulk loop. */
function findClaimInRows_(rows, claimId) {
  for (var i = 0; i < rows.length; i++) {
    if (rows[i]['Claim ID'] === claimId) return rows[i];
  }
  return null;
}

function findClaimRow_(claimId) {
  return findClaimInRows_(readClaimRows_(), claimId);
}

/* ===================== TRAVEL POLICY ENGINE ===================== */
/**
 * Reads the "Travel policy" sheet (brand-columns × grade-tier rows) into
 * structured rows: { category, byBrand: {companyName: rawCellText}, grades: [...] }.
 * Some category labels repeat 2-3 times (e.g. "Hotel" appears once per grade
 * tier) — those rows are disambiguated ONLY by which grade codes are listed
 * in that row's Grade column, never by the category text itself, so this
 * stays correct even if Abhilash reorders or relabels rows later.
 */
function readTravelPolicyRows_() {
  var values = getTravelPolicySheet_().getDataRange().getValues();
  var headers = values[0];
  var gradeCol = headers.indexOf('Grade');
  var brandCols = [];
  for (var c = 1; c < headers.length; c++) {
    if (c === gradeCol) continue;
    brandCols.push({ name: headers[c], index: c });
  }
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i][0]) continue;
    var byBrand = {};
    brandCols.forEach(function (b) { byBrand[b.name] = values[i][b.index]; });
    var gradesCell = gradeCol > -1 ? values[i][gradeCol] : '';
    var grades = String(gradesCell || '').split(/[\n,]+/).map(function (g) { return g.trim(); }).filter(Boolean);
    // Wrapped cells (e.g. "Local Conv.\n2-Wheeler", "Bills Required\n(Hotel)") store a literal
    // newline where the sheet just visually wraps the text — collapse all whitespace runs to a
    // single space so category lookups elsewhere can use the plain single-line label.
    var category = String(values[i][0]).trim().replace(/\s+/g, ' ');
    rows.push({ category: category, byBrand: byBrand, grades: grades });
  }
  return rows;
}

/** Pulls the first number out of a policy cell like "₹5.5/km", "₹2,500/month", "(>500 km)". Null for blank/"Not defined". */
function parsePolicyNumber_(cellText) {
  if (cellText === null || cellText === undefined || cellText === '') return null;
  var s = String(cellText);
  if (/not defined/i.test(s)) return null;
  var m = s.match(/[\d,]+(\.\d+)?/);
  return m ? Number(m[0].replace(/,/g, '')) : null;
}

function findPolicyRowByCategory_(rows, category) {
  for (var i = 0; i < rows.length; i++) if (rows[i].category === category) return rows[i];
  return null;
}

function findPolicyRowByCategoryAndGrade_(rows, category, grade) {
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].category !== category) continue;
    if (rows[i].grades.indexOf(grade) !== -1) return rows[i];
  }
  return null;
}

/** Whether Hotel claims need a receipt for this brand ("Yes — compulsory" vs "No — flat allowance"). Defaults to required if the policy row is missing. */
function getBillsRequiredForHotel_(policyRows, brand) {
  var row = findPolicyRowByCategory_(policyRows, 'Bills Required (Hotel)');
  if (!row) return true;
  return !/no\s*—?\s*flat allowance/i.test(String(row.byBrand[brand] || ''));
}

function getMonthlyCap_(policyRows, brand) {
  var row = findPolicyRowByCategory_(policyRows, 'Monthly Cap');
  return row ? parsePolicyNumber_(row.byBrand[brand]) : null;
}

/**
 * Only "Local Conveyance" has an actual cumulative cap in the new policy
 * (the Monthly Cap row) — Hotel/Food DA are rate×quantity so a correctly
 * computed amount is definitionally at-policy, and Flight/Outstation Travel
 * has no defined ceiling. Kept as a (category, brand) function, not folded
 * into computeTravelAmount_, since handleApproveClaim also needs to
 * re-check this when a manager edits a claim's amount after submission.
 */
function getPolicyMax_(category, brand) {
  if (category !== 'Local Conveyance') return null;
  return getMonthlyCap_(readTravelPolicyRows_(), brand);
}

/** Travel-focused category list for the New Claim form — non-travel categories (Office Supplies, Medical, etc.) were dropped per Abhilash's ask. */
function getCategoryList_() {
  return ['Local Conveyance', 'Hotel', 'Food DA', 'Flight / Outstation Travel'];
}

/** Local Conveyance vehicle types with no per-km policy rate — employee's entered fare is trusted directly, like Flight/Outstation Travel. */
var FARE_VEHICLE_TYPES_ = ['Auto', 'Cab', 'Bus', 'Metro'];

/**
 * Computes amount + an audit-trail description for the 3 rate-based
 * categories from the employee's Company (brand) and Grade, using whatever
 * quantity field the client sent (km / nights / days). Returns null when
 * the caller already supplied a flat payload.amount directly instead (bulk
 * upload, and Flight/Outstation Travel, which has no rate table at all) —
 * validateClaimPayload_ falls back to trusting that amount as-is in that case.
 */
function computeTravelAmount_(employee, payload) {
  if (payload.amount !== undefined && payload.amount !== null && String(payload.amount).trim() !== '') {
    return null;
  }
  var policyRows = readTravelPolicyRows_();
  var brand = employee['Brand Name'];
  var grade = employee['Grade'];
  var category = payload.category;

  if (category === 'Local Conveyance') {
    var vehicle = payload.vehicleType;
    if (FARE_VEHICLE_TYPES_.indexOf(vehicle) !== -1) {
      var fare = Number(payload.fare);
      if (!(fare > 0)) throw new Error('Enter the fare amount.');
      return { amount: Math.round(fare * 100) / 100, description: vehicle + ' fare' };
    }
    if (vehicle !== '2-Wheeler' && vehicle !== '4-Wheeler') throw new Error('Select a vehicle type.');
    var km = Number(payload.km);
    if (!(km > 0)) throw new Error('Enter the distance travelled in km.');
    var cRow = findPolicyRowByCategory_(policyRows, 'Local Conv. ' + vehicle);
    var cRate = cRow ? parsePolicyNumber_(cRow.byBrand[brand]) : null;
    if (cRate === null) throw new Error('No Local Conveyance rate defined for ' + (brand || 'your company') + '. Contact Admin.');
    return { amount: Math.round(km * cRate * 100) / 100, description: km + ' km (' + vehicle + ') @ ₹' + cRate + '/km' };
  }

  if (category === 'Hotel') {
    var nights = Number(payload.nights);
    if (!(nights > 0)) throw new Error('Enter the number of nights.');
    var hRow = findPolicyRowByCategoryAndGrade_(policyRows, 'Hotel', grade);
    var hRate = hRow ? parsePolicyNumber_(hRow.byBrand[brand]) : null;
    if (hRate === null) throw new Error('No Hotel rate defined for your grade (' + (grade || 'unknown') + ') at ' + (brand || 'your company') + '. Contact Admin.');
    return { amount: nights * hRate, description: nights + ' night(s) @ ₹' + hRate + '/day' + (payload.city ? ' — ' + payload.city : '') };
  }

  if (category === 'Food DA') {
    var days = Number(payload.days);
    if (!(days > 0)) throw new Error('Enter the number of days.');
    var fRow = findPolicyRowByCategoryAndGrade_(policyRows, 'Food DA', grade);
    var fRate = fRow ? parsePolicyNumber_(fRow.byBrand[brand]) : null;
    if (fRate === null) throw new Error('No Food DA rate defined for your grade (' + (grade || 'unknown') + ') at ' + (brand || 'your company') + '. Contact Admin.');
    return { amount: days * fRate, description: days + ' day(s) @ ₹' + fRate + '/day' };
  }

  return null; // Flight / Outstation Travel — flat amount, no rate table
}

/**
 * Sums an employee's claims in a category for the current IST calendar month, so policy
 * caps are enforced cumulatively (per month) rather than per individual claim. Rejected
 * claims don't count as real spend. Pass excludeClaimId when recomputing after an edit to
 * a claim that's already counted in the sheet, so it isn't added twice.
 */
function getMonthlySpend_(employeeId, category, excludeClaimId, preloadedRows) {
  var monthKey = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM');
  var rows = preloadedRows || readClaimRows_();
  var total = 0;
  rows.forEach(function (r) {
    if (r['Employee ID'] !== employeeId) return;
    if (r['Category'] !== category) return;
    if (r['Status'] === 'Rejected') return;
    if (excludeClaimId && r['Claim ID'] === excludeClaimId) return;
    if (!r['Expense Date']) return;
    if (Utilities.formatDate(new Date(r['Expense Date']), 'Asia/Kolkata', 'yyyy-MM') !== monthKey) return;
    total += Number(r['Amount']) || 0;
  });
  return total;
}

/** Powers the employee-facing "Policy" reference tab — the caller's own applicable rates, read live from the Travel policy sheet (brand + grade specific), so this page can never go stale the way a hardcoded copy would. */
function handleGetMyTravelRates(token) {
  var employee = requireActiveEmployee_(token);
  var policyRows = readTravelPolicyRows_();
  var brand = employee['Brand Name'];
  var grade = employee['Grade'];
  var conv2 = findPolicyRowByCategory_(policyRows, 'Local Conv. 2-Wheeler');
  var conv4 = findPolicyRowByCategory_(policyRows, 'Local Conv. 4-Wheeler');
  var hotel = findPolicyRowByCategoryAndGrade_(policyRows, 'Hotel', grade);
  var foodDa = findPolicyRowByCategoryAndGrade_(policyRows, 'Food DA', grade);
  var flightRow = findPolicyRowByCategory_(policyRows, 'Flight Eligible');
  var billsRow = findPolicyRowByCategory_(policyRows, 'Bills Required (Hotel)');
  return {
    success: true,
    brand: brand || '', grade: grade || '',
    conveyance2Wheeler: conv2 ? (conv2.byBrand[brand] || null) : null,
    conveyance4Wheeler: conv4 ? (conv4.byBrand[brand] || null) : null,
    monthlyCap: getMonthlyCap_(policyRows, brand),
    hotelRate: hotel ? (hotel.byBrand[brand] || null) : null,
    foodDaRate: foodDa ? (foodDa.byBrand[brand] || null) : null,
    flightEligible: flightRow ? (flightRow.byBrand[brand] || null) : null,
    billsRequiredHotel: billsRow ? (billsRow.byBrand[brand] || null) : null
  };
}

/** Live preview for the New Claim form — computes the amount from km/nights/days as the employee types, without saving anything. */
function handleComputeTravelAmount(payload, token) {
  var employee = requireActiveEmployee_(token);
  try {
    var computed = computeTravelAmount_(employee, payload || {});
    if (!computed) return { success: true, amount: null, description: null };
    return { success: true, amount: computed.amount, description: computed.description };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/** Nudge shown on the New Claim form: this month's cap, spend-so-far, and remaining for a category. */
function handleGetCategorySpend(payload, token) {
  var employee = requireActiveEmployee_(token);
  var category = payload && payload.category;
  if (!category) throw new Error('Category is required.');
  var cap = getPolicyMax_(category, employee['Brand Name']);
  var used = getMonthlySpend_(employee['Employee ID'], category);
  return { success: true, category: category, cap: cap, used: used, remaining: cap === null ? null : Math.max(0, cap - used) };
}

/** Finds a same-named subfolder scoped to this parent only (fast), or creates it. */
function getOrCreateSubfolder_(parentFolder, name) {
  var it = parentFolder.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parentFolder.createFolder(name);
}

/**
 * Decodes a data: URL receipt image and saves it under
 * Receipts/<Employee Name>/<date filed>/. Files inherit the root receipts folder's
 * sharing (set once — see migrateSchemaV4), so no per-file setSharing() call is needed
 * here; that call alone used to add a noticeable Drive round-trip to every submission.
 */
function saveReceipt_(dataUrl, fileName, claimId, employeeName) {
  var match = /^data:([^;]+);base64,(.+)$/.exec(dataUrl || '');
  if (!match) throw new Error('Receipt photo is required and must be a valid image.');
  var contentType = match[1];
  var bytes = Utilities.base64Decode(match[2]);
  var blob = Utilities.newBlob(bytes, contentType, claimId + '_' + (fileName || 'receipt'));
  var folderId = PropertiesService.getScriptProperties().getProperty('RECEIPTS_FOLDER_ID');
  if (!folderId) throw new Error('RECEIPTS_FOLDER_ID not set. Run migrateSchemaV2() first.');
  var rootFolder = DriveApp.getFolderById(folderId);
  var employeeFolder = getOrCreateSubfolder_(rootFolder, employeeName || 'Unknown');
  var dateFolderName = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
  var dateFolder = getOrCreateSubfolder_(employeeFolder, dateFolderName);
  var file = dateFolder.createFile(blob);
  return file.getUrl();
}

/** Saves every attached receipt file for one claim, returning their URLs joined with "|" for sheet storage. */
function saveReceipts_(dataUrls, fileNames, claimId, employeeName) {
  var urls = dataUrls.map(function (dataUrl, i) {
    return saveReceipt_(dataUrl, (fileNames && fileNames[i]) || 'receipt', claimId, employeeName);
  });
  return urls.join('|');
}

/** Guards against duplicate submissions caused by client retries after a slow/timed-out request. */
function checkIdempotency_(key) {
  if (!key) return null;
  var cached = CacheService.getScriptCache().get('idem_' + key);
  return cached ? JSON.parse(cached) : null;
}

function saveIdempotency_(key, result) {
  if (!key) return;
  CacheService.getScriptCache().put('idem_' + key, JSON.stringify(result), 1800);
}

/**
 * For Local Conveyance/Hotel/Food DA, amount is auto-computed from the
 * employee's brand+grade and whatever quantity they entered (km/nights/days)
 * via computeTravelAmount_ — no payload.amount expected from the New Claim
 * form for those. Flight/Outstation Travel (and any caller that already
 * supplies a flat payload.amount directly, e.g. bulk upload) keeps the old
 * flat-amount-plus-receipt behavior.
 */
function validateClaimPayload_(payload, employee) {
  if (!payload || !payload.category || !payload.expenseDate) {
    throw new Error('Category and expense date are required.');
  }
  var computed = computeTravelAmount_(employee, payload);
  var amount, autoDescription;
  if (computed) {
    amount = computed.amount;
    autoDescription = computed.description;
  } else {
    if (!payload.amount) throw new Error('Amount is required.');
    amount = Number(payload.amount);
    if (!(amount > 0)) throw new Error('Amount must be a positive number.');
    autoDescription = null;
  }
  var billsRequired = true;
  if (payload.category === 'Hotel') {
    billsRequired = getBillsRequiredForHotel_(readTravelPolicyRows_(), employee['Brand Name']);
  }
  if (billsRequired && (!payload.receiptDataUrls || !payload.receiptDataUrls.length)) {
    throw new Error('At least one receipt file is required.');
  }
  if (autoDescription) {
    payload.description = payload.description ? (payload.description + ' — ' + autoDescription) : autoDescription;
  }
  return amount;
}

/**
 * Core append-a-claim-row logic shared by single submit, batch submit, and resubmit.
 * extraCommitted lets a batch submission track same-category amounts already appended
 * earlier in the same batch, so the cap check is cumulative across the whole batch too.
 * preloadedRows (optional) lets a multi-claim batch read the Claims sheet once up front
 * and reuse it for every claim's monthly-spend check, instead of a fresh full-sheet read
 * per claim (handleSubmitBatch passes it; single-claim callers omit it and this falls
 * back to its own read, unchanged).
 */
function appendClaimRow_(employee, payload, amount, batchId, resubmittedFrom, extraCommitted, preloadedRows) {
  var sheet = getClaimsSheet_();
  var claimId = nextClaimId_();
  var receiptUrl = saveReceipts_(payload.receiptDataUrls, payload.receiptFileNames, claimId, employee['Name']);
  var policyMax = getPolicyMax_(payload.category, employee['Brand Name']);
  var priorSpend = getMonthlySpend_(employee['Employee ID'], payload.category, null, preloadedRows);
  var projected = priorSpend + (extraCommitted || 0) + amount;
  var overPolicy = policyMax !== null && projected > policyMax ? 'Yes' : 'No';
  var now = new Date();
  sheet.appendRow([
    claimId, employee['Employee ID'], employee['Name'], employee._email, employee['Manager Email'] || '',
    payload.category, amount, payload.description || '', payload.expenseDate, 'Pending Manager',
    now, '', '', '', '', '', receiptUrl, overPolicy, batchId || '', resubmittedFrom || '',
    employee['Brand Name'] || '', employee['Department'] || ''
  ]);
  var claim = { 'Claim ID': claimId, 'Employee Name': employee['Name'], 'Category': payload.category, 'Amount': amount, 'Manager Email': employee['Manager Email'] };
  notifyManagerOnSubmit_(claim);
  logAction_(employee._email, 'submit', claimId, 'amount=' + amount + ' category=' + payload.category + (overPolicy === 'Yes' ? ' OVER_POLICY' : ''));
  return { claimId: claimId, overPolicy: overPolicy };
}

function handleSubmitClaim(payload, token) {
  var cached = checkIdempotency_(payload && payload.idempotencyKey);
  if (cached) return cached;
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Employee')) throw new Error('Only Employee-role accounts can submit claims.');
  var amount = validateClaimPayload_(payload, employee);
  var res = appendClaimRow_(employee, payload, amount, null, null);
  var result = { success: true, claimId: res.claimId, overPolicy: res.overPolicy };
  saveIdempotency_(payload.idempotencyKey, result);
  return result;
}

function handleSubmitBatch(payload, token) {
  var claims = (payload && payload.claims) || payload;
  var idempotencyKey = payload && payload.idempotencyKey;
  var cached = checkIdempotency_(idempotencyKey);
  if (cached) return cached;
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Employee')) throw new Error('Only Employee-role accounts can submit claims.');
  if (!claims || !claims.length) throw new Error('No claims in batch.');
  var batchId = claims.length > 1 ? 'BATCH-' + Utilities.getUuid().slice(0, 8).toUpperCase() : null;
  var committedByCategory = {};
  // Read the Claims sheet once for the whole batch instead of once per claim (getMonthlySpend_ inside
  // appendClaimRow_ would otherwise re-read the entire sheet for every claim in the batch).
  var preloadedRows = claims.length > 1 ? readClaimRows_() : null;
  var results = claims.map(function (p) {
    var amount = validateClaimPayload_(p, employee);
    var extra = committedByCategory[p.category] || 0;
    var res = appendClaimRow_(employee, p, amount, batchId, null, extra, preloadedRows);
    committedByCategory[p.category] = extra + amount;
    return res;
  });
  var result = { success: true, claimIds: results.map(function (r) { return r.claimId; }) };
  saveIdempotency_(idempotencyKey, result);
  return result;
}

function handleListMyClaims(token) {
  var employee = requireActiveEmployee_(token);
  var rows = readClaimRows_().filter(function (r) { return r['Employee Email'] === employee._email; });
  rows.sort(function (a, b) { return new Date(b['Submitted At']) - new Date(a['Submitted At']); });
  return { success: true, claims: rows };
}

function handleListTeamClaims(token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Manager')) throw new Error('Manager role required.');
  var rows = readClaimRows_().filter(function (r) { return r['Manager Email'] === employee._email; });
  rows.sort(function (a, b) { return new Date(b['Submitted At']) - new Date(a['Submitted At']); });
  return { success: true, claims: rows };
}

function handleListManagerReturned(token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Manager')) throw new Error('Manager role required.');
  var rows = readClaimRows_().filter(function (r) {
    return r['Manager Email'] === employee._email && r['Status'] === 'Returned to Manager';
  });
  rows.sort(function (a, b) { return new Date(b['Returned At']) - new Date(a['Returned At']); });
  return { success: true, claims: rows };
}

function approveClaimRow_(claim, remarks) {
  var sheet = getClaimsSheet_();
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.status).setValue('Manager Approved');
  // managerRemarks + managerActionAt are adjacent columns — one setValues() round-trip instead of two.
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.managerRemarks, 1, 2).setValues([[remarks || '', new Date()]]);
  notifyEmployeeOnDecision_(claim, true, remarks);
}

function handleApproveClaim(payload, token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Manager')) throw new Error('Manager role required.');
  var claim = findClaimRow_(payload.claimId);
  if (!claim) throw new Error('Claim not found.');
  if (claim['Manager Email'] !== employee._email) throw new Error('Not your team\'s claim.');
  if (claim['Status'] !== 'Pending Manager') throw new Error('Claim already actioned.');

  var amountChanged = false;
  var originalAmount = Number(claim['Amount']);
  if (payload.editedAmount !== undefined && payload.editedAmount !== null && String(payload.editedAmount).trim() !== '') {
    var newAmount = Number(payload.editedAmount);
    if (!(newAmount > 0)) throw new Error('Edited amount must be a positive number.');
    if (newAmount !== originalAmount) {
      var sheet = getClaimsSheet_();
      sheet.getRange(claim._rowIndex, CLAIM_COLS_.amount).setValue(newAmount);
      var policyMax = getPolicyMax_(claim['Category'], claim['Brand Name']);
      var priorSpend = getMonthlySpend_(claim['Employee ID'], claim['Category'], claim['Claim ID']);
      var overPolicy = policyMax !== null && (priorSpend + newAmount) > policyMax ? 'Yes' : 'No';
      sheet.getRange(claim._rowIndex, CLAIM_COLS_.overPolicy).setValue(overPolicy);
      claim['Amount'] = newAmount;
      claim['Over Policy'] = overPolicy;
      amountChanged = true;
    }
  }

  approveClaimRow_(claim, payload.remarks);
  logAction_(employee._email, 'manager_approve', claim['Claim ID'],
    (amountChanged ? 'amount_changed_from=' + originalAmount + '_to=' + claim['Amount'] + '; ' : '') + (payload.remarks || ''));
  return { success: true, amountChanged: amountChanged };
}

function handleRejectClaim(payload, token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Manager')) throw new Error('Manager role required.');
  var claim = findClaimRow_(payload.claimId);
  if (!claim) throw new Error('Claim not found.');
  if (claim['Manager Email'] !== employee._email) throw new Error('Not your team\'s claim.');
  if (claim['Status'] !== 'Pending Manager') throw new Error('Claim already actioned.');
  var sheet = getClaimsSheet_();
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.status).setValue('Rejected');
  // managerRemarks + managerActionAt are adjacent columns — one setValues() round-trip instead of two.
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.managerRemarks, 1, 2).setValues([[payload.remarks || '', new Date()]]);
  notifyEmployeeOnDecision_(claim, false, payload.remarks);
  logAction_(employee._email, 'manager_reject', claim['Claim ID'], payload.remarks || '');
  return { success: true };
}

function handleBulkApproveWithinPolicy(token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Manager')) throw new Error('Manager role required.');
  var rows = readClaimRows_().filter(function (r) {
    return r['Manager Email'] === employee._email && r['Status'] === 'Pending Manager';
  });
  var approved = [], skipped = [];
  rows.forEach(function (claim) {
    if (claim['Over Policy'] === 'Yes') { skipped.push(claim['Claim ID']); return; }
    approveClaimRow_(claim, 'Bulk-approved (within policy)');
    logAction_(employee._email, 'manager_bulk_approve', claim['Claim ID'], '');
    approved.push(claim['Claim ID']);
  });
  return { success: true, approvedCount: approved.length, approvedIds: approved, skippedOverPolicyIds: skipped };
}

function handleListAdminQueue(token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Admin')) throw new Error('Admin role required.');
  var rows = readClaimRows_().filter(function (r) { return r['Status'] === 'Manager Approved'; });
  rows.sort(function (a, b) { return new Date(a['Manager Action At']) - new Date(b['Manager Action At']); });
  return { success: true, claims: rows };
}

function handleListAdminReturned(token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Admin')) throw new Error('Admin role required.');
  var rows = readClaimRows_().filter(function (r) { return r['Status'] === 'Returned to Admin'; });
  rows.sort(function (a, b) { return new Date(b['Returned At']) - new Date(a['Returned At']); });
  return { success: true, claims: rows };
}

function handleListAllClaims(token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Admin')) throw new Error('Admin role required.');
  var rows = readClaimRows_();
  rows.sort(function (a, b) { return new Date(b['Submitted At']) - new Date(a['Submitted At']); });
  return { success: true, claims: rows };
}

function handleForwardToFinance(payload, token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Admin')) throw new Error('Admin role required.');
  var claim = findClaimRow_(payload.claimId);
  if (!claim) throw new Error('Claim not found.');
  if (['Manager Approved', 'Returned to Admin'].indexOf(claim['Status']) === -1) {
    throw new Error('Claim is not awaiting Admin action.');
  }
  var sheet = getClaimsSheet_();
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.status).setValue('Ready for Finance');
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.adminActionAt).setValue(new Date());
  notifyFinanceOnForward_(claim);
  logAction_(employee._email, 'admin_forward_to_finance', claim['Claim ID'], '');
  return { success: true };
}

function handleBulkForwardToFinance(payload, token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Admin')) throw new Error('Admin role required.');
  if (!payload || !payload.claimIds || !payload.claimIds.length) throw new Error('No claims selected.');
  var sheet = getClaimsSheet_();
  var rows = readClaimRows_();
  var forwarded = [];
  payload.claimIds.forEach(function (claimId) {
    var claim = findClaimInRows_(rows, claimId);
    if (!claim || ['Manager Approved', 'Returned to Admin'].indexOf(claim['Status']) === -1) return;
    sheet.getRange(claim._rowIndex, CLAIM_COLS_.status).setValue('Ready for Finance');
    sheet.getRange(claim._rowIndex, CLAIM_COLS_.adminActionAt).setValue(new Date());
    notifyFinanceOnForward_(claim);
    logAction_(employee._email, 'admin_bulk_forward_to_finance', claimId, '');
    forwarded.push(claimId);
  });
  return { success: true, forwardedCount: forwarded.length, forwardedIds: forwarded };
}

/** Shared "send it back a step" write: sets Status + Return Reason/By/At in one place. */
function returnClaimRow_(claim, newStatus, actorRole, reason) {
  var sheet = getClaimsSheet_();
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.status).setValue(newStatus);
  // returnReason/returnedBy/returnedAt are adjacent columns — one setValues() round-trip instead of three.
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.returnReason, 1, 3).setValues([[reason || '', actorRole, new Date()]]);
}

function handleManagerReturnToEmployee(payload, token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Manager')) throw new Error('Manager role required.');
  var claim = findClaimRow_(payload.claimId);
  if (!claim) throw new Error('Claim not found.');
  if (claim['Manager Email'] !== employee._email) throw new Error('Not your team\'s claim.');
  if (['Pending Manager', 'Returned to Manager'].indexOf(claim['Status']) === -1) {
    throw new Error('Claim is not awaiting your action.');
  }
  returnClaimRow_(claim, 'Returned to Employee', 'Manager', payload.remarks);
  notifyEmployeeOnReturn_(claim, payload.remarks);
  logAction_(employee._email, 'manager_return_to_employee', claim['Claim ID'], payload.remarks || '');
  return { success: true };
}

function handleAdminReturnToManager(payload, token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Admin')) throw new Error('Admin role required.');
  var claim = findClaimRow_(payload.claimId);
  if (!claim) throw new Error('Claim not found.');
  if (['Manager Approved', 'Returned to Admin'].indexOf(claim['Status']) === -1) {
    throw new Error('Claim is not awaiting Admin action.');
  }
  returnClaimRow_(claim, 'Returned to Manager', 'Admin', payload.remarks);
  notifyManagerOnReturn_(claim, payload.remarks);
  logAction_(employee._email, 'admin_return_to_manager', claim['Claim ID'], payload.remarks || '');
  return { success: true };
}

function handleFinanceHold(payload, token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Finance')) throw new Error('Finance role required.');
  var claim = findClaimRow_(payload.claimId);
  if (!claim) throw new Error('Claim not found.');
  assertFinanceScope_(employee, claim);
  if (claim['Status'] !== 'Ready for Finance') throw new Error('Claim is not in the Finance queue.');
  returnClaimRow_(claim, 'On Hold', 'Finance', payload.remarks);
  logAction_(employee._email, 'finance_hold', claim['Claim ID'], payload.remarks || '');
  return { success: true };
}

function handleFinanceReleaseHold(payload, token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Finance')) throw new Error('Finance role required.');
  var claim = findClaimRow_(payload.claimId);
  if (!claim) throw new Error('Claim not found.');
  assertFinanceScope_(employee, claim);
  if (claim['Status'] !== 'On Hold') throw new Error('Claim is not on hold.');
  var sheet = getClaimsSheet_();
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.status).setValue('Ready for Finance');
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.returnReason).setValue('');
  logAction_(employee._email, 'finance_release_hold', claim['Claim ID'], '');
  return { success: true };
}

function handleFinanceReturnToAdmin(payload, token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Finance')) throw new Error('Finance role required.');
  var claim = findClaimRow_(payload.claimId);
  if (!claim) throw new Error('Claim not found.');
  assertFinanceScope_(employee, claim);
  if (['Ready for Finance', 'On Hold'].indexOf(claim['Status']) === -1) {
    throw new Error('Claim is not in the Finance queue.');
  }
  returnClaimRow_(claim, 'Returned to Admin', 'Finance', payload.remarks);
  notifyAdminOnReturn_(claim, payload.remarks);
  logAction_(employee._email, 'finance_return_to_admin', claim['Claim ID'], payload.remarks || '');
  return { success: true };
}

/**
 * Which brand/company names this Finance employee is authorized to process,
 * from the Employees Data "Finance Company" column (comma-separated, e.g. two
 * Finance people both employed under "Curefoods" payroll but actually
 * responsible for different brands - "Sharief Bhai, Frozen Bottle" vs
 * "Curefoods, Cakezone"). Empty = unrestricted, so Finance accounts that
 * haven't been assigned a scope yet keep seeing everything (back-compat).
 */
function getFinanceScope_(employee) {
  return splitCommaList_(employee['Finance Company']);
}

function scopeFinanceRows_(rows, employee) {
  var scope = getFinanceScope_(employee);
  if (!scope.length) return rows;
  return rows.filter(function (r) { return scope.indexOf(r['Brand Name']) !== -1; });
}

/** Guards the single-claim Finance actions (Hold/Release/Return/Mark Paid) so a scoped Finance account can't act on a claim outside its assigned companies by claim ID, even though the queue listing already hides it. */
function assertFinanceScope_(employee, claim) {
  var scope = getFinanceScope_(employee);
  if (scope.length && scope.indexOf(claim['Brand Name']) === -1) {
    throw new Error('This claim belongs to a company outside your Finance scope.');
  }
}

function handleListFinanceHeld(token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Finance')) throw new Error('Finance role required.');
  var rows = scopeFinanceRows_(readClaimRows_().filter(function (r) { return r['Status'] === 'On Hold'; }), employee);
  rows.sort(function (a, b) { return new Date(a['Returned At']) - new Date(b['Returned At']); });
  attachManagerNames_(rows);
  return { success: true, claims: rows };
}

/** Enriches claim rows with the manager's display name (claims only store Manager Email). */
/** Looks up each row's Manager Name via the cached per-email employee lookup instead of a full Employees-sheet scan — rows usually span only a handful of distinct managers, so this is a few cache hits instead of a ~3600-row read every time a Finance tab loads. */
function attachManagerNames_(rows) {
  var nameByEmail = {};
  rows.forEach(function (r) {
    var mgrEmail = r['Manager Email'];
    var key = String(mgrEmail || '').toLowerCase();
    if (!nameByEmail.hasOwnProperty(key)) {
      var mgr = findEmployeeByEmail_(mgrEmail);
      nameByEmail[key] = (mgr && mgr['Name']) || mgrEmail;
    }
    r['Manager Name'] = nameByEmail[key];
  });
  return rows;
}

function handleListFinanceQueue(token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Finance')) throw new Error('Finance role required.');
  var rows = scopeFinanceRows_(readClaimRows_().filter(function (r) { return r['Status'] === 'Ready for Finance'; }), employee);
  rows.sort(function (a, b) { return new Date(a['Admin Action At']) - new Date(b['Admin Action At']); });
  attachManagerNames_(rows);
  return { success: true, claims: rows };
}

function handleListFinanceHistory(token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Finance')) throw new Error('Finance role required.');
  var rows = scopeFinanceRows_(readClaimRows_().filter(function (r) { return r['Status'] === 'Paid'; }), employee);
  rows.sort(function (a, b) { return new Date(b['Paid At']) - new Date(a['Paid At']); });
  attachManagerNames_(rows);
  return { success: true, claims: rows };
}

function markPaidRow_(claim, transactionRef) {
  var sheet = getClaimsSheet_();
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.status).setValue('Paid');
  // transactionRef + paidAt are adjacent columns — one setValues() round-trip instead of two.
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.transactionRef, 1, 2).setValues([[transactionRef, new Date()]]);
  notifyEmployeeOnPaid_(claim, transactionRef);
}

function handleMarkPaid(payload, token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Finance')) throw new Error('Finance role required.');
  var claim = findClaimRow_(payload.claimId);
  if (!claim) throw new Error('Claim not found.');
  assertFinanceScope_(employee, claim);
  if (['Ready for Finance', 'On Hold'].indexOf(claim['Status']) === -1) throw new Error('Claim is not in the Finance queue.');
  var ref = (payload.transactionRef || '').trim();
  markPaidRow_(claim, ref);
  logAction_(employee._email, 'finance_mark_paid', claim['Claim ID'], 'ref=' + (ref || '(none)'));
  return { success: true };
}

function handleBulkMarkPaid(payload, token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Finance')) throw new Error('Finance role required.');
  if (!payload || !payload.claimIds || !payload.claimIds.length) throw new Error('No claims selected.');
  var ref = (payload.transactionRef || '').trim();
  var scope = getFinanceScope_(employee);
  var rows = readClaimRows_();
  var paid = [];
  payload.claimIds.forEach(function (claimId) {
    var claim = findClaimInRows_(rows, claimId);
    if (!claim || claim['Status'] !== 'Ready for Finance') return;
    if (scope.length && scope.indexOf(claim['Brand Name']) === -1) return;
    markPaidRow_(claim, ref);
    logAction_(employee._email, 'finance_bulk_mark_paid', claimId, 'ref=' + (ref || '(none)'));
    paid.push(claimId);
  });
  return { success: true, paidCount: paid.length, paidIds: paid };
}

function csvEscape_(value) {
  var s = String(value === undefined || value === null ? '' : value);
  if (/[",\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
  return s;
}

function handleExportFinanceCsv(payload, token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Finance')) throw new Error('Finance role required.');
  var scope = (payload && payload.scope) || 'queue';
  var rows = scopeFinanceRows_(readClaimRows_().filter(function (r) {
    return scope === 'history' ? r['Status'] === 'Paid' : r['Status'] === 'Ready for Finance';
  }), employee);
  var header = ['Claim ID', 'Employee ID', 'Employee Name', 'Brand Name', 'Department', 'Employee Email', 'Category', 'Amount', 'Expense Date', 'Status', 'Transaction Ref'];
  var lines = [header.join(',')];
  rows.forEach(function (r) {
    lines.push([r['Claim ID'], r['Employee ID'], r['Employee Name'], r['Brand Name'], r['Department'], r['Employee Email'], r['Category'], r['Amount'], r['Expense Date'], r['Status'], r['Transaction Ref']].map(csvEscape_).join(','));
  });
  logAction_(employee._email, 'finance_export_csv', '', 'scope=' + scope + ' rows=' + rows.length);
  return { success: true, csv: lines.join('\n'), rowCount: rows.length };
}

function handleResubmitClaim(payload, token) {
  var cached = checkIdempotency_(payload && payload.idempotencyKey);
  if (cached) return cached;
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Employee')) throw new Error('Only Employee-role accounts can resubmit claims.');
  var original = findClaimRow_(payload.originalClaimId);
  if (!original) throw new Error('Original claim not found.');
  if (original['Employee Email'] !== employee._email) throw new Error('Not your claim.');
  if (['Rejected', 'Returned to Employee'].indexOf(original['Status']) === -1) {
    throw new Error('Only rejected or returned claims can be resubmitted.');
  }
  var amount = validateClaimPayload_(payload, employee);
  var res = appendClaimRow_(employee, payload, amount, null, original['Claim ID']);
  logAction_(employee._email, 'resubmit', res.claimId, 'from=' + original['Claim ID']);
  var result = { success: true, claimId: res.claimId, overPolicy: res.overPolicy };
  saveIdempotency_(payload.idempotencyKey, result);
  return result;
}

/** Inclusive date-range check against a claim's Expense Date; either bound may be omitted. */
function withinExpenseRange_(expenseDateVal, startDate, endDate) {
  if (!startDate && !endDate) return true;
  var d = new Date(expenseDateVal);
  if (startDate && d < new Date(startDate + 'T00:00:00')) return false;
  if (endDate && d > new Date(endDate + 'T23:59:59')) return false;
  return true;
}

function computeDashboardStats_(startDate, endDate, preloadedRows) {
  var rows = preloadedRows || readClaimRows_().filter(function (r) { return withinExpenseRange_(r['Expense Date'], startDate, endDate); });
  var now = new Date();
  var pendingStatuses = [
    'Pending Manager', 'Manager Approved', 'Ready for Finance',
    'Returned to Employee', 'Returned to Manager', 'Returned to Admin', 'On Hold'
  ];
  var totalPendingAmount = 0, byCategory = {}, teamTotals = {}, aging = [];

  rows.forEach(function (r) {
    if (pendingStatuses.indexOf(r['Status']) !== -1) {
      totalPendingAmount += Number(r['Amount']) || 0;
      var days = Math.floor((now - new Date(r['Submitted At'])) / (24 * 60 * 60 * 1000));
      aging.push({ claimId: r['Claim ID'], employeeName: r['Employee Name'], amount: r['Amount'], status: r['Status'], daysPending: days });
    }
    if (r['Status'] !== 'Rejected') {
      byCategory[r['Category']] = (byCategory[r['Category']] || 0) + (Number(r['Amount']) || 0);
      var mgr = r['Manager Email'] || '(no manager)';
      teamTotals[mgr] = (teamTotals[mgr] || 0) + (Number(r['Amount']) || 0);
    }
  });

  var totalPaid = rows
    .filter(function (r) { return r['Status'] === 'Paid'; })
    .reduce(function (sum, r) { return sum + (Number(r['Amount']) || 0); }, 0);

  aging.sort(function (a, b) { return b.daysPending - a.daysPending; });

  return {
    totalPendingAmount: totalPendingAmount,
    totalPaid: totalPaid,
    byCategory: byCategory,
    teamTotals: teamTotals,
    aging: aging.slice(0, 10)
  };
}

function handleGetDashboardStats(payload, token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Admin') && !hasRole_(employee._roles, 'Finance')) throw new Error('Admin or Finance role required.');
  var startDate = (payload && payload.startDate) || '';
  var endDate = (payload && payload.endDate) || '';
  var stats = computeDashboardStats_(startDate, endDate);
  stats.success = true;
  stats.startDate = startDate;
  stats.endDate = endDate;
  return stats;
}

/**
 * DriveApp's File.getAs() does not support converting a native Google Sheet to xlsx/pdf
 * (confirmed live: throws "Converting from application/vnd.google-apps.spreadsheet ...
 * is not supported"). Hitting the Sheets export endpoint directly with the script's own
 * OAuth token is the standard reliable way to get those formats.
 */
function exportSpreadsheetBlob_(ssId, format, extraParams) {
  var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=' + format + (extraParams || '');
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('Export to ' + format + ' failed: HTTP ' + response.getResponseCode());
  }
  return response.getBlob();
}

function formatDateOnly_(iso) {
  if (!iso) return '';
  return Utilities.formatDate(new Date(iso), 'Asia/Kolkata', 'yyyy-MM-dd');
}

function formatDateTime_(iso) {
  if (!iso) return '';
  return Utilities.formatDate(new Date(iso), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm');
}

/**
 * Builds a temporary Google Sheet with every claim entry (full transaction detail,
 * including transaction ref / paid date where applicable) on tab 1, and a category-level
 * rollup on tab 2 — the source for the report's Excel/PDF attachments. Caller is
 * responsible for trashing the returned Spreadsheet once its blobs are extracted.
 */
function buildReportWorkbook_(rows, rangeLabel) {
  var ss = SpreadsheetApp.create('Curefoods Reimbursement Report - ' + rangeLabel);
  var claimsSheet = ss.getSheets()[0];
  claimsSheet.setName('Claims');
  var headers = [
    'Claim ID', 'Employee ID', 'Employee Name', 'Brand Name', 'Department', 'Employee Email', 'Manager Email',
    'Category', 'Amount', 'Description', 'Expense Date', 'Status', 'Over Policy',
    'Submitted At', 'Manager Remarks', 'Transaction Ref', 'Paid At'
  ];
  claimsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  claimsSheet.setFrozenRows(1);
  if (rows.length) {
    var data = rows.map(function (r) {
      return [
        r['Claim ID'], r['Employee ID'], r['Employee Name'], r['Brand Name'], r['Department'], r['Employee Email'], r['Manager Email'],
        r['Category'], Number(r['Amount']) || 0, r['Description'], formatDateOnly_(r['Expense Date']),
        r['Status'], r['Over Policy'], formatDateTime_(r['Submitted At']), r['Manager Remarks'],
        r['Transaction Ref'], formatDateTime_(r['Paid At'])
      ];
    });
    claimsSheet.getRange(2, 1, data.length, headers.length).setValues(data);
  }
  claimsSheet.autoResizeColumns(1, headers.length);

  var byCategory = {};
  rows.forEach(function (r) {
    var cat = r['Category'] || '(none)';
    if (!byCategory[cat]) byCategory[cat] = { count: 0, total: 0, paid: 0, pending: 0 };
    var amt = Number(r['Amount']) || 0;
    byCategory[cat].count++;
    byCategory[cat].total += amt;
    if (r['Status'] === 'Paid') byCategory[cat].paid += amt;
    else if (r['Status'] !== 'Rejected') byCategory[cat].pending += amt;
  });
  var summarySheet = ss.insertSheet('Summary by Category');
  summarySheet.getRange(1, 1, 1, 5).setValues([['Category', 'Claim Count', 'Total Amount', 'Paid Amount', 'Pending Amount']]);
  summarySheet.setFrozenRows(1);
  var catNames = Object.keys(byCategory).sort(function (a, b) { return byCategory[b].total - byCategory[a].total; });
  if (catNames.length) {
    var summaryData = catNames.map(function (cat) {
      var c = byCategory[cat];
      return [cat, c.count, c.total, c.paid, c.pending];
    });
    summarySheet.getRange(2, 1, summaryData.length, 5).setValues(summaryData);
  }
  summarySheet.autoResizeColumns(1, 5);
  SpreadsheetApp.flush();
  return ss;
}

function handleEmailDashboardReport(payload, token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Admin') && !hasRole_(employee._roles, 'Finance')) throw new Error('Admin or Finance role required.');
  var startDate = (payload && payload.startDate) || '';
  var endDate = (payload && payload.endDate) || '';
  var recipient = (payload && payload.recipient) || employee._email;
  var rows = readClaimRows_().filter(function (r) { return withinExpenseRange_(r['Expense Date'], startDate, endDate); });
  var stats = computeDashboardStats_(startDate, endDate, rows);
  var rangeLabel = (startDate || endDate) ? (startDate || 'earliest') + ' to ' + (endDate || 'latest') : 'All time';

  var lines = [];
  lines.push('Curefoods Reimbursement Portal - Admin Report');
  lines.push('Range: ' + rangeLabel + ' (' + rows.length + ' claims)');
  lines.push('');
  lines.push('Total Pending: Rs.' + stats.totalPendingAmount.toLocaleString('en-IN'));
  lines.push('Total Paid: Rs.' + stats.totalPaid.toLocaleString('en-IN'));
  lines.push('');
  lines.push('Spend by Category:');
  var categories = Object.keys(stats.byCategory).sort(function (a, b) { return stats.byCategory[b] - stats.byCategory[a]; });
  if (!categories.length) lines.push('  None');
  categories.forEach(function (cat) { lines.push('  ' + cat + ': Rs.' + stats.byCategory[cat].toLocaleString('en-IN')); });
  lines.push('');
  lines.push('Spend by Team (Manager):');
  var teams = Object.keys(stats.teamTotals).sort(function (a, b) { return stats.teamTotals[b] - stats.teamTotals[a]; });
  if (!teams.length) lines.push('  None');
  teams.forEach(function (mgr) { lines.push('  ' + mgr + ': Rs.' + stats.teamTotals[mgr].toLocaleString('en-IN')); });
  lines.push('');
  lines.push('Oldest Pending (aging):');
  if (!stats.aging.length) lines.push('  None');
  stats.aging.forEach(function (a) { lines.push('  ' + a.claimId + ' - ' + a.employeeName + ' - Rs.' + a.amount + ' - ' + a.daysPending + 'd pending'); });
  lines.push('');
  lines.push('Full claim-by-claim detail and a category summary are attached (Excel + PDF).');

  var reportSs = buildReportWorkbook_(rows, rangeLabel);
  var reportFile = DriveApp.getFileById(reportSs.getId());
  var fileNameBase = 'Curefoods Reimbursement Report (' + rangeLabel + ')';
  var xlsxBlob = exportSpreadsheetBlob_(reportSs.getId(), 'xlsx', '').setName(fileNameBase + '.xlsx');
  var pdfBlob = exportSpreadsheetBlob_(reportSs.getId(), 'pdf', '&size=A4&portrait=false&fitw=true&gridlines=false').setName(fileNameBase + '.pdf');

  MailApp.sendEmail({
    to: recipient,
    subject: 'Curefoods Reimbursement Report (' + rangeLabel + ')',
    body: lines.join('\n'),
    attachments: [xlsxBlob, pdfBlob]
  });
  reportFile.setTrashed(true);

  logAction_(employee._email, 'admin_email_report', '', 'range=' + rangeLabel + ' to=' + recipient + ' claims=' + rows.length);
  return { success: true, recipient: recipient };
}

// ---------- Bulk upload (Employee) ----------

var BULK_TEMPLATE_HEADERS_ = ['Category', 'Amount', 'Expense Date', 'Description'];

/**
 * Builds the downloadable bulk-upload template as a real .xlsx: a fresh Google Sheet
 * with the exact expected headers, a Category dropdown (data validation) sourced from
 * the live Policy tab, and a sample row — then exported and trashed. Attachments are
 * never part of this file; those are still attached per-claim in the app after upload.
 *
 * Building + exporting a temp Sheet is a handful of sequential Drive/Sheets round trips
 * (confirmed live: ~5-6s per request), which felt slow for a file that's identical for
 * every employee until the Policy category list changes. Caching the finished base64
 * (keyed on the category list itself, so a Policy change invalidates it automatically)
 * makes every request after the first one instant.
 */
function handleGetBulkTemplate(token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Employee')) throw new Error('Employee role required.');

  var categories = getCategoryList_();
  var cache = CacheService.getScriptCache();
  var cacheKey = 'bulk_template_v1_' + Utilities.base64EncodeWebSafe(categories.join('|'));
  var cached = cache.get(cacheKey);
  if (cached) {
    return { success: true, base64: cached, fileName: 'Curefoods_Bulk_Claim_Template.xlsx' };
  }

  var ss = SpreadsheetApp.create('CFI Bulk Claim Template');
  var sheet = ss.getSheets()[0];
  sheet.setName('Claims');
  sheet.getRange(1, 1, 1, BULK_TEMPLATE_HEADERS_.length).setValues([BULK_TEMPLATE_HEADERS_]).setFontWeight('bold');
  sheet.getRange(2, 1, 1, 4).setValues([[
    categories[0] || 'Other', 1500, Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd'),
    'Example row - delete before uploading'
  ]]);
  if (categories.length) {
    var catRule = SpreadsheetApp.newDataValidation().requireValueInList(categories, true).setAllowInvalid(false).build();
    sheet.getRange(2, 1, 200, 1).setDataValidation(catRule);
  }
  sheet.getRange(2, 3, 200, 1).setNumberFormat('yyyy-mm-dd');
  SpreadsheetApp.flush();

  var xlsxBlob = exportSpreadsheetBlob_(ss.getId(), 'xlsx', '');
  var base64 = Utilities.base64Encode(xlsxBlob.getBytes());
  DriveApp.getFileById(ss.getId()).setTrashed(true);

  try {
    cache.put(cacheKey, base64, 21600); // 21600s = 6h, CacheService's max TTL
  } catch (e) {
    Logger.log('Bulk template cache put failed (serving uncached this time): ' + e.message);
  }
  return { success: true, base64: base64, fileName: 'Curefoods_Bulk_Claim_Template.xlsx' };
}

/** Converts uploaded .xlsx bytes into a temporary native Google Sheet via the Advanced Drive Service, so its cells can be read with SpreadsheetApp. Caller must trash the returned id. */
function convertXlsxBytesToSheet_(bytes, name) {
  var blob = Utilities.newBlob(bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', name || 'upload.xlsx');
  var resource = { title: name || 'CFI Bulk Upload Temp', mimeType: 'application/vnd.google-apps.spreadsheet' };
  var file = Drive.Files.insert(resource, blob, { convert: true });
  return file.id;
}

/**
 * Parses + strictly validates an uploaded bulk-claim file. Any deviation from the
 * template (wrong headers, unrecognized category, bad amount/date) rejects the whole
 * file with a specific reason rather than silently dropping or auto-fixing rows —
 * the user re-downloads the template and re-uploads.
 */
function handleParseBulkUpload(payload, token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Employee')) throw new Error('Employee role required.');

  var dataUrl = payload && payload.fileDataUrl;
  var match = /^data:([^;]+);base64,(.+)$/.exec(dataUrl || '');
  if (!match) throw new Error('Could not read the uploaded file. Please choose a valid .xlsx file.');
  var bytes = Utilities.base64Decode(match[2]);

  var sheetId;
  try {
    sheetId = convertXlsxBytesToSheet_(bytes, 'CFI Bulk Upload - ' + employee._email);
  } catch (e) {
    throw new Error('Could not read that file as an Excel spreadsheet. Please upload the unmodified .xlsx template.');
  }

  try {
    var sheet = SpreadsheetApp.openById(sheetId).getSheets()[0];
    var values = sheet.getDataRange().getValues();
    if (!values.length) throw new Error('That file is empty.');

    var headers = values[0].map(function (h) { return String(h || '').trim(); });
    var headersMatch = BULK_TEMPLATE_HEADERS_.every(function (h, i) { return headers[i] === h; });
    if (!headersMatch) {
      throw new Error('This file\'s columns don\'t match the template (expected: ' + BULK_TEMPLATE_HEADERS_.join(', ') +
        '). Please download the template again and re-upload without changing the headers.');
    }

    var categories = getCategoryList_();
    var rows = [];
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      if (r.every(function (v) { return v === '' || v === null; })) continue;
      var rowNum = i + 1;
      var category = String(r[0] || '').trim();
      if (categories.indexOf(category) === -1) {
        throw new Error('Row ' + rowNum + ': "' + category + '" is not a recognized category. Valid categories: ' +
          categories.join(', ') + '. Please fix and re-upload.');
      }
      var amount = Number(r[1]);
      if (!(amount > 0)) throw new Error('Row ' + rowNum + ': Amount must be a positive number.');
      var expenseDate;
      if (Object.prototype.toString.call(r[2]) === '[object Date]') {
        expenseDate = Utilities.formatDate(r[2], 'Asia/Kolkata', 'yyyy-MM-dd');
      } else if (/^\d{4}-\d{2}-\d{2}$/.test(String(r[2] || '').trim())) {
        expenseDate = String(r[2]).trim();
      } else {
        throw new Error('Row ' + rowNum + ': Expense Date must be a valid date (YYYY-MM-DD).');
      }
      rows.push({ category: category, amount: amount, expenseDate: expenseDate, description: String(r[3] || '').trim() });
    }
    if (!rows.length) throw new Error('No data rows found below the header. Fill in at least one row and re-upload.');
    if (rows.length > 100) throw new Error('Too many rows (' + rows.length + '). Please upload at most 100 claims at a time.');

    return { success: true, rows: rows };
  } finally {
    try { DriveApp.getFileById(sheetId).setTrashed(true); } catch (e2) { /* best-effort cleanup */ }
  }
}

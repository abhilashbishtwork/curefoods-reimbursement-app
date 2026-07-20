var CLAIM_COLS_ = {
  claimId: 1, employeeId: 2, employeeName: 3, employeeEmail: 4, managerEmail: 5,
  category: 6, amount: 7, description: 8, expenseDate: 9, status: 10,
  submittedAt: 11, managerRemarks: 12, managerActionAt: 13, adminActionAt: 14,
  transactionRef: 15, paidAt: 16, receiptUrl: 17, overPolicy: 18, batchId: 19,
  resubmittedFrom: 20, brandName: 21, department: 22
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

function findClaimRow_(claimId) {
  var rows = readClaimRows_();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i]['Claim ID'] === claimId) return rows[i];
  }
  return null;
}

function getPolicyMax_(category) {
  var sheet = getPolicySheet_();
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === category) return Number(values[i][1]);
  }
  return null;
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

function validateClaimPayload_(payload) {
  if (!payload || !payload.category || !payload.amount || !payload.expenseDate) {
    throw new Error('Category, amount and expense date are required.');
  }
  var amount = Number(payload.amount);
  if (!(amount > 0)) throw new Error('Amount must be a positive number.');
  if (!payload.receiptDataUrl) throw new Error('A receipt photo is required.');
  return amount;
}

/** Core append-a-claim-row logic shared by single submit, batch submit, and resubmit. */
function appendClaimRow_(employee, payload, amount, batchId, resubmittedFrom) {
  var sheet = getClaimsSheet_();
  var claimId = nextClaimId_();
  var receiptUrl = saveReceipt_(payload.receiptDataUrl, payload.receiptFileName, claimId, employee['Name']);
  var policyMax = getPolicyMax_(payload.category);
  var overPolicy = policyMax !== null && amount > policyMax ? 'Yes' : 'No';
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
  var amount = validateClaimPayload_(payload);
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
  var results = claims.map(function (p) {
    var amount = validateClaimPayload_(p);
    return appendClaimRow_(employee, p, amount, batchId, null);
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

function approveClaimRow_(claim, remarks) {
  var sheet = getClaimsSheet_();
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.status).setValue('Manager Approved');
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.managerRemarks).setValue(remarks || '');
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.managerActionAt).setValue(new Date());
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
      var policyMax = getPolicyMax_(claim['Category']);
      var overPolicy = policyMax !== null && newAmount > policyMax ? 'Yes' : 'No';
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
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.managerRemarks).setValue(payload.remarks || '');
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.managerActionAt).setValue(new Date());
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
  if (claim['Status'] !== 'Manager Approved') throw new Error('Claim is not awaiting Admin action.');
  var sheet = getClaimsSheet_();
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.status).setValue('Ready for Finance');
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.adminActionAt).setValue(new Date());
  notifyFinanceOnForward_(claim);
  logAction_(employee._email, 'admin_forward_to_finance', claim['Claim ID'], '');
  return { success: true };
}

function handleListFinanceQueue(token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Finance')) throw new Error('Finance role required.');
  var rows = readClaimRows_().filter(function (r) { return r['Status'] === 'Ready for Finance'; });
  rows.sort(function (a, b) { return new Date(a['Admin Action At']) - new Date(b['Admin Action At']); });
  return { success: true, claims: rows };
}

function handleListFinanceHistory(token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Finance')) throw new Error('Finance role required.');
  var rows = readClaimRows_().filter(function (r) { return r['Status'] === 'Paid'; });
  rows.sort(function (a, b) { return new Date(b['Paid At']) - new Date(a['Paid At']); });
  return { success: true, claims: rows };
}

function markPaidRow_(claim, transactionRef) {
  var sheet = getClaimsSheet_();
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.status).setValue('Paid');
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.transactionRef).setValue(transactionRef);
  sheet.getRange(claim._rowIndex, CLAIM_COLS_.paidAt).setValue(new Date());
  notifyEmployeeOnPaid_(claim, transactionRef);
}

function handleMarkPaid(payload, token) {
  var employee = requireActiveEmployee_(token);
  if (!hasRole_(employee._roles, 'Finance')) throw new Error('Finance role required.');
  var claim = findClaimRow_(payload.claimId);
  if (!claim) throw new Error('Claim not found.');
  if (claim['Status'] !== 'Ready for Finance') throw new Error('Claim is not in the Finance queue.');
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
  var paid = [];
  payload.claimIds.forEach(function (claimId) {
    var claim = findClaimRow_(claimId);
    if (!claim || claim['Status'] !== 'Ready for Finance') return;
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
  var rows = readClaimRows_().filter(function (r) {
    return scope === 'history' ? r['Status'] === 'Paid' : r['Status'] === 'Ready for Finance';
  });
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
  if (original['Status'] !== 'Rejected') throw new Error('Only rejected claims can be resubmitted.');
  var amount = validateClaimPayload_(payload);
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

function computeDashboardStats_(startDate, endDate) {
  var rows = readClaimRows_().filter(function (r) { return withinExpenseRange_(r['Expense Date'], startDate, endDate); });
  var now = new Date();
  var pendingStatuses = ['Pending Manager', 'Manager Approved', 'Ready for Finance'];
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
  if (!hasRole_(employee._roles, 'Admin')) throw new Error('Admin role required.');
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
  if (!hasRole_(employee._roles, 'Admin')) throw new Error('Admin role required.');
  var startDate = (payload && payload.startDate) || '';
  var endDate = (payload && payload.endDate) || '';
  var recipient = (payload && payload.recipient) || employee._email;
  var rows = readClaimRows_().filter(function (r) { return withinExpenseRange_(r['Expense Date'], startDate, endDate); });
  var stats = computeDashboardStats_(startDate, endDate);
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

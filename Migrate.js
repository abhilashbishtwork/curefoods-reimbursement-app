/**
 * One-off schema v2 migration: adds Policy + Logs tabs and extra Claims columns.
 * Safe to re-run — each step checks whether it already applied before acting.
 */
function migrateSchemaV2() {
  var ss = getMasterSheet_();

  var claims = ss.getSheetByName('Claims');
  var headerRange = claims.getRange(1, 1, 1, claims.getLastColumn());
  var headers = headerRange.getValues()[0];
  var newCols = ['Receipt URL', 'Over Policy', 'Batch ID', 'Resubmitted From'];
  newCols.forEach(function (name) {
    if (headers.indexOf(name) === -1) {
      var col = claims.getLastColumn() + 1;
      claims.getRange(1, col).setValue(name);
      headers.push(name);
    }
  });

  if (!ss.getSheetByName('Policy')) {
    var policy = ss.insertSheet('Policy');
    policy.getRange(1, 1, 1, 2).setValues([['Category', 'Max Amount']]);
    policy.setFrozenRows(1);
    policy.getRange(2, 1, 6, 2).setValues([
      ['Travel', 5000],
      ['Food', 1000],
      ['Accommodation', 8000],
      ['Communication', 1000],
      ['Office Supplies', 2000],
      ['Other', 1500]
    ]);
    Logger.log('Created Policy tab with placeholder limits - adjust these to real policy amounts.');
  }

  if (!ss.getSheetByName('Logs')) {
    var logs = ss.insertSheet('Logs');
    logs.getRange(1, 1, 1, 5).setValues([['Timestamp', 'Actor Email', 'Action', 'Claim ID', 'Details']]);
    logs.setFrozenRows(1);
  }

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('OTP_SECRET')) {
    props.setProperty('OTP_SECRET', Utilities.getUuid() + Utilities.getUuid());
  }
  if (!props.getProperty('RECEIPTS_FOLDER_ID')) {
    var folder = DriveApp.createFolder('Curefoods Reimbursement Portal - Receipts');
    props.setProperty('RECEIPTS_FOLDER_ID', folder.getId());
    Logger.log('Created receipts folder: ' + folder.getUrl());
  }

  Logger.log('Migration complete. Claims headers now: ' + headers.join(', '));
}

/**
 * Schema v3: switches Claim ID generation from sheet.getLastRow() (not concurrency-safe —
 * caused two real duplicate IDs in production) to an atomic Script Properties counter.
 * Initializes each year's counter to the current max ID found in the sheet, so numbering
 * continues from where it is instead of restarting. Safe to re-run — only sets a counter
 * if it isn't already set.
 */
function migrateSchemaV3() {
  var claims = getClaimsSheet_();
  var values = claims.getDataRange().getValues();
  var headers = values[0];
  var idCol = headers.indexOf('Claim ID');
  var maxByYear = {};
  for (var i = 1; i < values.length; i++) {
    var id = values[i][idCol];
    var m = /^CFI-(\d{4})-(\d+)$/.exec(id);
    if (m) {
      var yr = m[1], num = Number(m[2]);
      if (!maxByYear[yr] || num > maxByYear[yr]) maxByYear[yr] = num;
    }
  }
  var props = PropertiesService.getScriptProperties();
  Object.keys(maxByYear).forEach(function (yr) {
    var key = 'CLAIM_COUNTER_' + yr;
    if (!props.getProperty(key)) {
      props.setProperty(key, String(maxByYear[yr]));
      Logger.log('Initialized ' + key + ' = ' + maxByYear[yr]);
    } else {
      Logger.log(key + ' already set to ' + props.getProperty(key) + ', not touching it.');
    }
  });
}

/**
 * Schema v4: shares the root receipts folder once (domain-with-link, view) so every
 * file and subfolder created inside it — including the new per-employee/per-date
 * subfolder structure — inherits that access automatically. This removes the need for
 * a setSharing() call on every single uploaded receipt, which was a real Drive API
 * round-trip adding noticeable latency to every claim submission. Safe to re-run.
 */
function migrateSchemaV4() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty('RECEIPTS_FOLDER_ID');
  if (!folderId) throw new Error('RECEIPTS_FOLDER_ID not set. Run migrateSchemaV2() first.');
  var folder = DriveApp.getFolderById(folderId);
  folder.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
  Logger.log('Receipts root folder shared (domain, view). New subfolders/files will inherit this.');
}

/**
 * Schema v5: replaces the placeholder claim categories with the real approved list.
 * Clears existing Policy rows and reseeds — limits are still placeholders, adjust to
 * real policy amounts. Existing claims already submitted under the old category names
 * (Travel/Food/Communication/Other) are untouched — they just won't match any Policy
 * row going forward, so getPolicyMax_ treats them as "no limit" rather than erroring.
 */
function migrateSchemaV5() {
  var policy = getPolicySheet_();
  var lastRow = policy.getLastRow();
  if (lastRow > 1) policy.getRange(2, 1, lastRow - 1, 2).clearContent();
  var categories = [
    ['Fuel', 3000],
    ['Accommodation', 8000],
    ['Meals', 1000],
    ['Local Conveyance', 1000],
    ['Outstation Travel', 6000],
    ['Internet / Telephone / Mobile Bill', 1500],
    ['Office Supplies', 2000],
    ['Training', 10000],
    ['Medical', 5000],
    ['Petty Cash', 2000],
    ['Other Business Expenses', 1500]
  ];
  policy.getRange(2, 1, categories.length, 2).setValues(categories);
  Logger.log('Policy categories replaced with the new list (' + categories.length + ' rows) - adjust limits to real policy amounts.');
}

/**
 * Schema v6: adds Brand Name + Department to the Claims tab, captured from the
 * employee's record at submission time (same pattern as Employee Name/Email already
 * being denormalized onto each claim row) so they're available on claim cards and in
 * the exported report without a live lookup against Employees on every render.
 */
function migrateSchemaV6() {
  var claims = getClaimsSheet_();
  var headerRange = claims.getRange(1, 1, 1, claims.getLastColumn());
  var headers = headerRange.getValues()[0];
  var newCols = ['Brand Name', 'Department'];
  newCols.forEach(function (name) {
    if (headers.indexOf(name) === -1) {
      var col = claims.getLastColumn() + 1;
      claims.getRange(1, col).setValue(name);
      headers.push(name);
    }
  });
  Logger.log('Claims headers now: ' + headers.join(', '));
}

/**
 * Schema v7: adds Return Reason / Returned By / Returned At to the Claims tab, backing
 * the cross-role "send it back a step" workflow (Manager -> Employee, Admin -> Manager,
 * Finance -> Admin, plus Finance's Hold). Safe to re-run.
 */
/** Extracts a Drive file ID from a /file/d/<id>/, /document/d/<id>/, or /spreadsheets/d/<id>/ URL. */
function extractDriveFileId_(url) {
  var m = /\/d\/([a-zA-Z0-9_-]{20,})/.exec(url);
  return m ? m[1] : null;
}

/**
 * One-time reset for a fresh round of testing, per Abhilash's explicit ask
 * (2026-09-28): clears every Claims row and every Logs row (headers kept),
 * and moves every receipt file referenced in the Receipt URL column to
 * Drive Trash — reversible, NOT permanently deleted, so anything can be
 * restored from Trash if something looks wrong afterward. Also resets the
 * atomic claim-ID counters so the next real claim starts back at 00001.
 * Employees Data / Travel policy / Policy sheets are untouched.
 */
function wipeAllClaimsForTesting() {
  var claimsSheet = getClaimsSheet_();
  var lastRow = claimsSheet.getLastRow();
  var lastCol = claimsSheet.getLastColumn();
  var trashedCount = 0, unresolvedCount = 0;

  if (lastRow > 1) {
    var receiptCol = CLAIM_COLS_.receiptUrl;
    var urls = claimsSheet.getRange(2, receiptCol, lastRow - 1, 1).getValues();
    urls.forEach(function (row) {
      var cell = row[0];
      if (!cell) return;
      String(cell).split('|').forEach(function (url) {
        url = url.trim();
        if (!url) return;
        var id = extractDriveFileId_(url);
        if (!id) { unresolvedCount++; return; }
        try {
          DriveApp.getFileById(id).setTrashed(true);
          trashedCount++;
        } catch (e) {
          unresolvedCount++;
        }
      });
    });
    claimsSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }

  var logsSheet = getLogsSheet_();
  var logsLastRow = logsSheet.getLastRow();
  if (logsLastRow > 1) {
    logsSheet.getRange(2, 1, logsLastRow - 1, logsSheet.getLastColumn()).clearContent();
  }

  var props = PropertiesService.getScriptProperties();
  var allProps = props.getProperties();
  Object.keys(allProps).forEach(function (key) {
    if (key.indexOf('CLAIM_COUNTER_') === 0) props.deleteProperty(key);
  });

  Logger.log('Wipe complete. Receipt files moved to Trash: ' + trashedCount +
    '. Could not resolve/trash: ' + unresolvedCount +
    '. Claims + Logs rows cleared (headers kept). Claim ID counters reset.');
}

function migrateSchemaV7() {
  var claims = getClaimsSheet_();
  var headerRange = claims.getRange(1, 1, 1, claims.getLastColumn());
  var headers = headerRange.getValues()[0];
  var newCols = ['Return Reason', 'Returned By', 'Returned At'];
  newCols.forEach(function (name) {
    if (headers.indexOf(name) === -1) {
      var col = claims.getLastColumn() + 1;
      claims.getRange(1, col).setValue(name);
      headers.push(name);
    }
  });
  Logger.log('Claims headers now: ' + headers.join(', '));
}

/**
 * Schema v8: the Status column (J) had a stale dropdown data-validation rule
 * (Reject input) from before schema v7's return/hold workflow was added - it
 * only allowed the original 5 statuses (Pending Manager, Manager Approved,
 * Rejected, Ready for Finance, Paid). Root cause of a real bug (2026-09-28):
 * Manager "Return" silently reported success while the claim stayed visible
 * in Pending, because Range.setValue('Returned to Employee') on that cell was
 * rejected by validation with no exception at the call site - the eventual
 * error surfaced later, misleadingly caught by logAction_'s unrelated
 * try/catch on the next sheet write. Replaces the rule with the full list of
 * every status value the app actually writes. Safe to re-run.
 */
function migrateSchemaV8() {
  var claims = getClaimsSheet_();
  var allStatuses = [
    'Pending Manager', 'Manager Approved', 'Rejected', 'Ready for Finance', 'Paid',
    'Returned to Employee', 'Returned to Manager', 'Returned to Admin', 'On Hold'
  ];
  var lastRow = Math.max(claims.getMaxRows(), 2);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(allStatuses, true)
    .setAllowInvalid(false)
    .build();
  claims.getRange(2, CLAIM_COLS_.status, lastRow - 1, 1).setDataValidation(rule);
  Logger.log('Status column validation replaced. Allowed values: ' + allStatuses.join(', '));
}

/**
 * Schema v9: adds a "Finance Company" column to Employees Data, so Finance
 * accounts can be scoped to the brand(s) they actually process instead of
 * seeing every company's claims just because their own Role is "Finance"
 * (the "Company" column is their own employer/payroll entity, not what they
 * approve - e.g. Mayank Agarwal and Vivek Rai are both payrolled under
 * "Curefoods" but Mayank handles Sharief Bhai + Frozen Bottle finance while
 * Vivek handles Curefoods + Cakezone). Comma-separated, same convention as
 * the Role column (see splitCommaList_/getFinanceScope_). Blank = unrestricted,
 * so every other Finance account is unaffected until explicitly scoped.
 * Also seeds the two known cases from this migration's own ask (2026-10-08).
 * Safe to re-run.
 */
function migrateSchemaV9() {
  var sheet = getEmployeesSheet_();
  var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  var headers = headerRange.getValues()[0];
  if (headers.indexOf('Finance Company') === -1) {
    var col = sheet.getLastColumn() + 1;
    sheet.getRange(1, col).setValue('Finance Company');
    headers.push('Finance Company');
  }
  var financeCompanyCol = headers.indexOf('Finance Company') + 1;
  var empCodeCol = headers.indexOf('Emp Code') + 1;

  var seed = {
    'CFD11351': 'Sharief Bhai, Frozen Bottle', // Mayank Agarwal
    'CFD11233': 'Curefoods, Cakezone'          // Vivek Rai
  };
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var seeded = [];
  values.forEach(function (row, i) {
    var empCode = row[empCodeCol - 1];
    if (seed.hasOwnProperty(empCode)) {
      sheet.getRange(2 + i, financeCompanyCol).setValue(seed[empCode]);
      seeded.push(empCode);
    }
  });
  Logger.log('Employees Data headers now: ' + headers.join(', ') + '. Seeded Finance Company for: ' + seeded.join(', '));
}

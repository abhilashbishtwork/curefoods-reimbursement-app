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

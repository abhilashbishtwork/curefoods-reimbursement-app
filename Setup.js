/**
 * One-off setup: creates the Master Sheet (Employees + Claims tabs) and
 * stores its ID in Script Properties. Run once from the Apps Script editor.
 * Safe to re-run only if SHEET_ID script property is cleared first.
 */
function runFullSetup() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('SHEET_ID')) {
    throw new Error('SHEET_ID already set (' + props.getProperty('SHEET_ID') + '). Clear the property first if you really want to recreate the sheet.');
  }

  var ss = SpreadsheetApp.create('Curefoods Reimbursement Portal - Master Sheet');

  var employees = ss.getSheets()[0];
  employees.setName('Employees');
  employees.getRange(1, 1, 1, 6).setValues([
    ['Employee ID', 'Name', 'Email', 'Manager Email', 'Role', 'Status']
  ]);
  employees.setFrozenRows(1);
  var roleValidation = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=REGEXMATCH(TO_TEXT(E2), "^(Employee|Manager|Admin|Finance)(,(Employee|Manager|Admin|Finance))*$")')
    .setAllowInvalid(true)
    .build();
  employees.getRange(2, 5, 999, 1).setDataValidation(roleValidation);
  var statusValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive'], true)
    .setAllowInvalid(false)
    .build();
  employees.getRange(2, 6, 999, 1).setDataValidation(statusValidation);

  // Basic test data: manager-ID mapping to be refined later per Abhilash's note.
  employees.getRange(2, 1, 4, 6).setValues([
    ['CFI-1001', 'Abhilash Bisht', 'abhilash@curefoods.in', '', 'Employee,Manager,Admin,Finance', 'Active'],
    ['CFI-1002', 'Ram', 'ram@curefoods.in', 'abhilash@curefoods.in', 'Employee,Manager', 'Active'],
    ['CFI-1003', 'Test Employee One', 'test.employee1@curefoods.in', 'ram@curefoods.in', 'Employee', 'Active'],
    ['CFI-1004', 'Test Employee Two', 'test.employee2@curefoods.in', 'ram@curefoods.in', 'Employee', 'Active']
  ]);

  var claims = ss.insertSheet('Claims');
  claims.getRange(1, 1, 1, 16).setValues([[
    'Claim ID', 'Employee ID', 'Employee Name', 'Employee Email', 'Manager Email',
    'Category', 'Amount', 'Description', 'Expense Date', 'Status',
    'Submitted At', 'Manager Remarks', 'Manager Action At', 'Admin Action At',
    'Transaction Ref', 'Paid At'
  ]]);
  claims.setFrozenRows(1);
  var statusValidationClaims = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending Manager', 'Manager Approved', 'Rejected', 'Ready for Finance', 'Paid'], true)
    .setAllowInvalid(false)
    .build();
  claims.getRange(2, 10, 999, 1).setDataValidation(statusValidationClaims);

  props.setProperty('SHEET_ID', ss.getId());

  Logger.log('Created Master Sheet: ' + ss.getUrl());
  return ss.getUrl();
}

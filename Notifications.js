function claimSummary_(claim) {
  return claim['Claim ID'] + ' - ' + claim['Category'] + ' - Rs.' + claim['Amount'] + ' (' + claim['Employee Name'] + ')';
}

/**
 * Kill switch for every claim-lifecycle notification email (submit/approve/
 * reject/forward-to-finance/paid/return, at every notify*_ function below —
 * they all funnel through here). Flip to true to turn them back on.
 * Does NOT affect the OTP login email (handleRequestOtp in Employees.js,
 * sent directly via MailApp, required for login) or the Admin's manual
 * "Email Report" button (handleEmailDashboardReport in Claims.js, an
 * explicit one-off send, not an automatic notification) - ask if those
 * should be disabled too.
 */
var NOTIFICATIONS_ENABLED_ = false;

function notifySafe_(to, subject, body) {
  if (!NOTIFICATIONS_ENABLED_) return;
  if (!to) return;
  try {
    MailApp.sendEmail({ to: to, subject: subject, body: body });
  } catch (e) {
    Logger.log('Notification failed to ' + to + ': ' + e.message);
  }
}

function notifyManagerOnSubmit_(claim) {
  if (!claim['Manager Email']) return;
  notifySafe_(
    claim['Manager Email'],
    'New reimbursement claim awaiting your approval - ' + claim['Claim ID'],
    claimSummary_(claim) + ' has been submitted and is awaiting your approval.\n\nOpen the Reimbursement Portal to review it.'
  );
}

function notifyEmployeeOnDecision_(claim, approved, remarks) {
  notifySafe_(
    claim['Employee Email'],
    'Your claim ' + claim['Claim ID'] + ' was ' + (approved ? 'approved' : 'rejected'),
    claimSummary_(claim) + ' was ' + (approved ? 'approved by your manager' : 'rejected') + '.' +
    (remarks ? '\n\nRemarks: ' + remarks : '') +
    (approved ? '\n\nIt now moves to Admin for processing.' : '\n\nYou can edit and resubmit it from the Reimbursement Portal.')
  );
}

function notifyFinanceOnForward_(claim) {
  var financeEmails = getEmployeesWithRole_('Finance');
  financeEmails.forEach(function (email) {
    notifySafe_(
      email,
      'Claim ready for payment - ' + claim['Claim ID'],
      claimSummary_(claim) + ' has been approved and forwarded by Admin. It is ready for payment processing.'
    );
  });
}

function notifyEmployeeOnPaid_(claim, transactionRef) {
  notifySafe_(
    claim['Employee Email'],
    'Your claim ' + claim['Claim ID'] + ' has been paid',
    claimSummary_(claim) + ' has been paid.\n\nTransaction reference (UTR): ' + (transactionRef || 'Not provided')
  );
}

function notifyEmployeeOnReturn_(claim, reason) {
  notifySafe_(
    claim['Employee Email'],
    'Your claim ' + claim['Claim ID'] + ' was sent back for changes',
    claimSummary_(claim) + ' was sent back by your manager for changes.' +
    (reason ? '\n\nReason: ' + reason : '') +
    '\n\nEdit and resubmit it from the Reimbursement Portal (Returned tab).'
  );
}

function notifyManagerOnReturn_(claim, reason) {
  notifySafe_(
    claim['Manager Email'],
    'Claim ' + claim['Claim ID'] + ' was sent back to you by Admin',
    claimSummary_(claim) + ' was sent back by Admin.' +
    (reason ? '\n\nReason: ' + reason : '') +
    '\n\nReview it again from the Reimbursement Portal (Returned tab).'
  );
}

function notifyAdminOnReturn_(claim, reason) {
  var adminEmails = getEmployeesWithRole_('Admin');
  adminEmails.forEach(function (email) {
    notifySafe_(
      email,
      'Claim ' + claim['Claim ID'] + ' was sent back to Admin by Finance',
      claimSummary_(claim) + ' was sent back by Finance.' +
      (reason ? '\n\nReason: ' + reason : '') +
      '\n\nReview it again from the Reimbursement Portal (Returned tab).'
    );
  });
}

function logAction_(actorEmail, action, claimId, details) {
  try {
    getLogsSheet_().appendRow([new Date(), actorEmail, action, claimId || '', details || '']);
  } catch (e) {
    Logger.log('Failed to log action: ' + e.message);
  }
}

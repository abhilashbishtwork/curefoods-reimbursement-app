function claimSummary_(claim) {
  return claim['Claim ID'] + ' - ' + claim['Category'] + ' - Rs.' + claim['Amount'] + ' (' + claim['Employee Name'] + ')';
}

function notifySafe_(to, subject, body) {
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

function logAction_(actorEmail, action, claimId, details) {
  try {
    getLogsSheet_().appendRow([new Date(), actorEmail, action, claimId || '', details || '']);
  } catch (e) {
    Logger.log('Failed to log action: ' + e.message);
  }
}

function testMigration() {
  var out = [];
  function log(s) { out.push(s); Logger.log(s); }

  // 1. Employee lookup from the new "Employees Data" tab
  var emp = findEmployeeByEmail_('abhilash@curefoods.in');
  log('Employee lookup: ' + JSON.stringify({
    id: emp && emp['Employee ID'], name: emp && emp['Name'], brand: emp && emp['Brand Name'],
    grade: emp && emp['Grade'], dept: emp && emp['Department'], status: emp && emp['Status'],
    role: emp && emp['Role']
  }));

  // 2. Row count sanity check
  log('Total employee rows read: ' + readEmployeeRows_().length);

  // 3. Travel policy parsing
  var policyRows = readTravelPolicyRows_();
  log('Travel policy rows parsed: ' + policyRows.length);
  policyRows.forEach(function (r) {
    log('Row category=[' + r.category + '] grades=' + JSON.stringify(r.grades) + ' curefoods=[' + r.byBrand['Curefoods'] + ']');
  });

  function tryLog(label, fn) {
    try { log(label + ': ' + JSON.stringify(fn())); }
    catch (e) { log(label + ': THREW ' + e.message); }
  }

  // 4. Rate computation for a known case: Curefoods CF05A (Senior tier)
  var fakeEmployee = { 'Brand Name': 'Curefoods', 'Grade': 'CF05A' };
  tryLog('Local Conveyance 100km 2-Wheeler @ Curefoods (expect amount=550)', function () {
    return computeTravelAmount_(fakeEmployee, { category: 'Local Conveyance', vehicleType: '2-Wheeler', km: 100 });
  });
  tryLog('Hotel 2 nights @ Curefoods CF05A Senior (expect amount=5600)', function () {
    return computeTravelAmount_(fakeEmployee, { category: 'Hotel', nights: 2, city: 'Mumbai' });
  });
  tryLog('Food DA 3 days @ Curefoods CF05A Senior (expect amount=2400)', function () {
    return computeTravelAmount_(fakeEmployee, { category: 'Food DA', days: 3 });
  });

  var juniorEmployee = { 'Brand Name': 'Curefoods', 'Grade': 'CF02B' };
  tryLog('Hotel 1 night @ Curefoods CF02B Junior (expect amount=1200)', function () {
    return computeTravelAmount_(juniorEmployee, { category: 'Hotel', nights: 1 });
  });

  // 5. Monthly cap + bills-required helpers
  tryLog('Monthly cap Curefoods (expect 2500)', function () { return getMonthlyCap_(policyRows, 'Curefoods'); });
  tryLog('Monthly cap Sharief Bhai (expect null)', function () { return getMonthlyCap_(policyRows, 'Sharief Bhai'); });
  tryLog('Bills required Hotel Curefoods (expect false)', function () { return getBillsRequiredForHotel_(policyRows, 'Curefoods'); });
  tryLog('Bills required Hotel Sharief Bhai (expect true)', function () { return getBillsRequiredForHotel_(policyRows, 'Sharief Bhai'); });

  // 6. Full identity payload (as the frontend receives it)
  tryLog('buildIdentityPayload_', function () { return buildIdentityPayload_('abhilash@curefoods.in'); });
}

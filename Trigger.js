function createTrigger_DailyInstallmentChecker() {
  // Remove existing triggers for checkDailyInstallment to avoid duplicates
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkDailyInstallment')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('checkDailyInstallment')
    .timeBased().atHour(8).everyDays(1).create();
}

function checkDailyInstallment() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const today = new Date();
  const todayDay = today.getDate();

  // Build account → statement day map
  const acctRows = ss.getSheetByName('Accounts').getDataRange().getValues().slice(1);
  const stmtDayMap = {};
  acctRows.forEach(row => { if (row[0]) stmtDayMap[row[0]] = parseInt(row[7]) || null; }); // col H = index 7

  const instRows = ss.getSheetByName('Installments').getDataRange().getValues().slice(1);
  const txSheet = ss.getSheetByName('2026 Transactions');

  instRows.forEach(([name, account, , terms, termPayment, startDate]) => {
    const stmtDay = stmtDayMap[account];
    if (!stmtDay) return;

    // Trigger day = 1 day before statement day
    const triggerDay = stmtDay === 1 ? new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate() : stmtDay - 1;
    if (todayDay !== triggerDay) return;

    // Determine which term this is: billing month index from start
    const start = new Date(startDate);
    const startMonth = new Date(start.getFullYear(), start.getMonth(), 1);
    // The statement that falls after today belongs to next month's cycle
    const billingMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);
    const termIdx = (billingMonth.getFullYear() - startMonth.getFullYear()) * 12
      + billingMonth.getMonth() - startMonth.getMonth();

    if (termIdx < 0 || termIdx >= parseInt(terms)) return; // not active

    // Check not already inserted this cycle (avoid duplicates)
    const txRows = txSheet.getDataRange().getValues().slice(1);
    const note = `Auto: Installment - ${name}`;
    const alreadyExists = txRows.some(r =>
      r[1] === account && r[8] === note && r[3] === 'PHP' &&
      Math.abs(Number(r[2]) - Number(termPayment)) < 0.01 &&
      new Date(r[0]).getMonth() === today.getMonth() &&
      new Date(r[0]).getFullYear() === today.getFullYear()
    );
    if (alreadyExists) return;

    const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy');
    txSheet.appendRow([dateStr, account, Number(termPayment), 'PHP', '', 'Expense', 'Installment', '', note]);
  });

  // Re-sort by date
  const last = txSheet.getLastRow();
  if (last > 1) txSheet.getRange(2, 1, last - 1, txSheet.getLastColumn()).sort({ column: 1, ascending: true });
}

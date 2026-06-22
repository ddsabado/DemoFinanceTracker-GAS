const props = PropertiesService.getScriptProperties();
const SPREADSHEET_ID = props.getProperty('SPREADSHEET_ID');
const PRIVATE_DEPLOYMENT_ID = props.getProperty('PRIVATE_DEPLOYMENT_ID');
const PUBLIC_DEPLOYMENT_ID  = props.getProperty('PUBLIC_DEPLOYMENT_ID');

function getDeploymentId_() {
  const match = ScriptApp.getService().getUrl().match(/\/s\/([^/]+)\/(exec|dev)$/);
  return match ? match[1] : null;
}

function doGet(e) {
  const isPublic = getDeploymentId_() === PUBLIC_DEPLOYMENT_ID;
  if (isPublic && e.parameter.action !== 'shortcuts_data')
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  if (e.parameter.action === 'shortcuts_data') return getShortcutsData_();
  if (e.parameter.page === 'control') return serveHtml_('control', 'Finance Control Center');
  return serveHtml_('index', 'Finance Tracker');
}

function serveHtml_(file, title) {
  return HtmlService.createHtmlOutputFromFile(file)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getShortcutsData_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const types = ss.getSheetByName('Txn_Type').getDataRange().getValues().slice(1).map(([t]) => t).filter(Boolean);
  const currencies = ss.getSheetByName('Currency').getDataRange().getValues().slice(1).map(([c]) => c).filter(Boolean);
  const catRows = ss.getSheetByName('Txn_Categories').getDataRange().getValues().slice(1);
  const catList = [], transaction_categories = {};
  catRows.forEach(([cat]) => { if (cat) { catList.push(cat); transaction_categories[cat] = []; } });
  ss.getSheetByName('Txn_Subcategories').getDataRange().getValues().slice(1)
    .forEach(([sub, cat]) => { if (sub && cat && transaction_categories[cat] !== undefined) transaction_categories[cat].push(sub); });
  const accounts = {}, acct_meta = {};
  ss.getSheetByName('Accounts').getDataRange().getValues().slice(1)
    .forEach(([account, category, , , , , , , , , metadata]) => {
      if (!account || !category) return;
      if (!accounts[category]) accounts[category] = [];
      accounts[category].push(account);
      if (metadata) acct_meta[account] = metadata;
    });
  return ContentService.createTextOutput(JSON.stringify({ transaction_categories, accounts, acct_meta, transactionTypes: types, currencies }))
    .setMimeType(ContentService.MimeType.JSON);
}

function driveToDirectUrl(fileUrl) {
  const match = fileUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? `https://drive.google.com/thumbnail?id=${match[1]}&sz=w200` : '';
}

function testCache() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const rows = ss.getSheetByName('Accounts').getDataRange().getValues();
  const urls = rows.slice(1)
    .filter(r => r[1] === 'Credit Cards' && r[6])
    .map(r => r[6]);
  urls.forEach(url => driveToDirectUrl(url));
}

function getCCImages(imageUrls) {
  const result = imageUrls.map(driveToDirectUrl);
  return result;
}

function getChartData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const accountSheet = ss.getSheetByName('Accounts');
  const txSheet = ss.getSheetByName('2026 Transactions');
  const currencySheet = ss.getSheetByName('Currency');
  const rows = accountSheet.getDataRange().getValues();
  const txRows = txSheet.getDataRange().getValues().slice(1);

  // FX rates
  const fx = {};
  currencySheet.getDataRange().getValues().slice(1).forEach(([c, r]) => { if (c) fx[c] = Number(r) || 1; });

  // Account transactions map
  const acctTxs = {};
  txRows.forEach(([date, account, amount, currency, targetAccount, type, category, , note]) => {
    if (!account || !date) return;
    const d = new Date(date);
    if (isNaN(d)) return;
    const dateStr = d.toLocaleDateString('en-PH', { month: 'short', day: 'numeric', year: 'numeric' });
    const tx = { date: dateStr, amount: Number(amount), currency: currency || 'PHP', type: type || '', category: category || '', note: note || '' };
    if (!acctTxs[account]) acctTxs[account] = [];
    acctTxs[account].push({ ...tx, incoming: false });
    // Also add to target account as incoming transfer
    if (type === 'Transfer' && targetAccount) {
      if (!acctTxs[targetAccount]) acctTxs[targetAccount] = [];
      acctTxs[targetAccount].push({ ...tx, incoming: true });
    }
  });

  const nonCC = { labels: [], data: [], accounts: [] };
  const cc = { labels: [], limits: [], balances: [], images: [], statementDays: [], dueDays: [], cycles: [], allTxs: [] };

  for (let i = 1; i < rows.length; i++) {
    const [account, category, creditLimit, , currency, , imageLink, statementDay, dueDay, currentBalance] = rows[i];
    if (!account) continue;
    const balPHP = (Number(currentBalance) || 0) * (fx[currency] || 1);

    if (category === 'Credit Cards') {
      const today = new Date();
      const stDay = parseInt(statementDay) || 1;

      // Build last 6 cycles
      const cycles = [];
      for (let c = 0; c < 6; c++) {
        let cycleEnd = new Date(today.getFullYear(), today.getMonth() - c, stDay - 1);
        if (c === 0 && cycleEnd >= today) cycleEnd.setMonth(cycleEnd.getMonth() - 1);
        else if (c > 0) cycleEnd.setMonth(cycleEnd.getMonth() - (c === 0 ? 0 : 0)); // already offset by c
        // Recalculate properly
        let ce = new Date(today.getFullYear(), today.getMonth(), stDay - 1);
        if (ce >= today) ce.setMonth(ce.getMonth() - 1);
        ce.setMonth(ce.getMonth() - c);
        const cs = new Date(ce.getFullYear(), ce.getMonth() - 1, stDay);

        const txs = txRows.filter(([date, acct, , , , type]) => {
          if (acct !== account) return false;
          const d = new Date(date);
          return !isNaN(d) && d >= cs && d <= ce;
        }).map(([date, , amount, cur, , type, category, subcategory, note]) => ({
          date: new Date(date).toLocaleDateString('en-PH', { month: 'short', day: 'numeric' }),
          amount: Number(amount), currency: cur || 'PHP',
          type: type || '', category: category || '', subcategory: subcategory || '', note: note || ''
        })).sort((a, b) => new Date('1 ' + b.date) - new Date('1 ' + a.date));

        cycles.push({
          label: cs.toLocaleDateString('en-PH', { month: 'short', day: 'numeric' }) + ' – ' + ce.toLocaleDateString('en-PH', { month: 'short', day: 'numeric' }),
          cycleStart: cs.toLocaleDateString('en-PH', { month: 'short', day: 'numeric' }),
          cycleEnd: ce.toLocaleDateString('en-PH', { month: 'short', day: 'numeric' }),
          txs
        });
      }

      cc.labels.push(account);
      cc.limits.push(Number(String(creditLimit).replace(/,/g,'')) || 0);
      cc.balances.push(balPHP);
      cc.images.push(imageLink ? driveToDirectUrl(imageLink) : '');
      cc.statementDays.push(statementDay || '');
      cc.dueDays.push(dueDay || '');
      cc.cycles.push(cycles);
      // All transactions for this card regardless of cycle
      const allTxs = txRows.filter(([, acct]) => acct === account)
        .map(([date, , amount, cur, , type, category, subcategory, note]) => ({
          date: new Date(date).toLocaleDateString('en-PH', { month: 'short', day: 'numeric' }),
          amount: Number(amount), currency: cur || 'PHP',
          type: type || '', category: category || '', subcategory: subcategory || '', note: note || ''
        })).reverse(); // reverse since sheet is sorted ascending
      cc.allTxs.push(allTxs);
    } else {
      nonCC.labels.push(account);
      nonCC.data.push(balPHP);
      nonCC.accounts.push(acctTxs[account] || []);
    }
  }

  return { nonCC, cc };
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { date, account, amount, currency, target_account, type, category, subcategory, note } = body;
    const [month, day, year] = date.split('/');
    const dateObj = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('2026 Transactions');
    sheet.appendRow([
      Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'MM/dd/yyyy'),
      account || '',
      amount || '',
      currency || 'PHP',
      target_account || '',
      type || '',
      category || '',
      subcategory || '',
      note || ''
    ]);
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).sort({column: 1, ascending: true});

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: e.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getInstallments() {
  const rows = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Installments')
    .getDataRange().getValues().slice(1);

  const today = new Date();
  const currentMonth = new Date(today.getFullYear(), today.getMonth(), 1);

  // Build months from current to furthest end date
  const installments = rows.map(([name, account, total, terms, termPayment, startDate]) => {
    const start = new Date(startDate);
    const startMonth = new Date(start.getFullYear(), start.getMonth(), 1);
    return { name, account, total: Number(String(total).replace(/,/g,'')), terms: parseInt(terms), termPayment: Number(String(termPayment).replace(/,/g,'')), startMonth };
  });

  // Find all months from current to max end month
  const months = [];
  let maxMonths = 0;
  installments.forEach(inst => { if (inst.terms > maxMonths) maxMonths = inst.terms; });
  // Generate up to 24 future months
  for (let m = 0; m < 24; m++) {
    const monthDate = new Date(currentMonth.getFullYear(), currentMonth.getMonth() + m, 1);
    const label = monthDate.toLocaleString('en-US', { month: 'short', year: 'numeric' });
    const active = installments
      .map(inst => {
        const diffMonths = (monthDate.getFullYear() - inst.startMonth.getFullYear()) * 12
          + monthDate.getMonth() - inst.startMonth.getMonth();
        if (diffMonths < 0 || diffMonths >= inst.terms) return null;
        return { name: inst.name, account: inst.account, termPayment: inst.termPayment, iteration: diffMonths + 1, terms: inst.terms };
      })
      .filter(Boolean);
    if (active.length === 0 && m > 0) continue; // skip empty future months but include current
    months.push({ label, active, total: active.reduce((s, a) => s + a.termPayment, 0) });
    if (active.length === 0) break; // stop after first empty month
  }

  return months;
}

function getTransactionsByMonth() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const txSheet = ss.getSheetByName('2026 Transactions');
  const subSheet = ss.getSheetByName('Txn_Subcategories');
  const accountSheet = ss.getSheetByName('Accounts');
  const txRows = txSheet.getDataRange().getValues().slice(1);
  const subRows = subSheet.getDataRange().getValues().slice(1);
  const accountRows = accountSheet.getDataRange().getValues().slice(1);

  const subcatMap = {};
  subRows.forEach(([sub, cat]) => {
    if (!cat) return;
    if (!subcatMap[cat]) subcatMap[cat] = [];
    if (sub && !subcatMap[cat].includes(sub)) subcatMap[cat].push(sub);
  });

  // Build CC cycle map: { accountName: { cycleStart, cycleEnd } }
  const ccCycles = {};
  accountRows.forEach(([account, category, , , , , , statementDay]) => {
    if (category !== 'Credit Cards') return;
    const today = new Date();
    const stDay = parseInt(statementDay) || 1;
    let cycleEnd = new Date(today.getFullYear(), today.getMonth(), stDay - 1);
    if (cycleEnd >= today) cycleEnd.setMonth(cycleEnd.getMonth() - 1);
    const cycleStart = new Date(cycleEnd.getFullYear(), cycleEnd.getMonth() - 1, stDay);
    ccCycles[account] = { cycleStart, cycleEnd };
  });

  // CC transactions per account
  const ccTxs = {};
  Object.keys(ccCycles).forEach(acc => ccTxs[acc] = []);

  const result = {};
  txRows.forEach(([date, account, amount, currency, targetAccount, type, category, subcategory, note]) => {
    if (!date || !type || !category || (type !== 'Income' && type !== 'Expense')) return;
    const d = new Date(date);
    if (isNaN(d)) return;
    const amt = Number(amount);
    const dateStr = d.toLocaleDateString('en-PH', { month: 'short', day: 'numeric' });

    // CC cycle transactions
    if (ccCycles[account]) {
      const { cycleStart, cycleEnd } = ccCycles[account];
      if (d >= cycleStart && d <= cycleEnd) {
        ccTxs[account].push({ date: dateStr, amount: amt, currency: currency || 'PHP', category: category || '', subcategory: subcategory || '', note: note || '' });
      }
    }

    const monthKey = d.toLocaleString('en-US', { month: 'short', year: 'numeric' });
    if (!result[monthKey]) result[monthKey] = { Income: {}, Expense: {} };
    const m = result[monthKey][type];
    if (!m[category]) m[category] = { total: 0, subcats: {}, txs: [] };
    m[category].total += amt;
    const tx = { date: dateStr, account, amount: amt, currency: currency || 'PHP' };
    if (subcategory) {
      if (!m[category].subcats[subcategory]) m[category].subcats[subcategory] = { total: 0, txs: [] };
      m[category].subcats[subcategory].total += amt;
      m[category].subcats[subcategory].txs.push(tx);
    } else {
      m[category].txs.push(tx);
    }
  });

  return { months: result, subcatMap, ccTxs };
}

function getAccountList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const acctRows = ss.getSheetByName('Accounts').getDataRange().getValues().slice(1);
  const accounts = {};
  acctRows.forEach(([account, category]) => {
    if (!account || !category) return;
    if (!accounts[category]) accounts[category] = [];
    accounts[category].push(account);
  });
  const acctCats = ss.getSheetByName('Account Categories').getDataRange().getValues().slice(1)
    .map(r => r[0]).filter(Boolean);
  const currencies = ss.getSheetByName('Currency').getDataRange().getValues().slice(1)
    .map(r => r[0]).filter(Boolean);
  return { accounts, acctCats, currencies };
}

const CC_IMAGES_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('CC_IMAGES_FOLDER_ID');

function adjustBalance(body) {
  try {
    const { account, target_balance } = body;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const acctSheet = ss.getSheetByName('Accounts');
    const acctRow = acctSheet.getDataRange().getValues().find(r => r[0] === account);
    if (!acctRow) return { success: false, error: 'Account not found' };
    const delta = Number(target_balance) - (Number(acctRow[9]) || 0);
    if (delta === 0) return { success: false, error: 'Balance unchanged' };
    const today = new Date();
    ss.getSheetByName('2026 Transactions').appendRow([
      `${today.getMonth()+1}/${today.getDate()}/${today.getFullYear()}`,
      account, delta, acctRow[4] || 'PHP', '', 'Adjustment', 'Adjustment', '',
      `Manual adjustment: ${target_balance}`
    ]);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function addAccount(body) {
  try {
    const { name, category, init_balance, currency, metadata, credit_limit, image, statement_day, due_day } = body;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Accounts');
    const lastRow = sheet.getLastRow() + 1;
    const isCC = category === 'Credit Cards';

    let imageLink = '';
    if (image && image.base64) {
      const blob = Utilities.newBlob(Utilities.base64Decode(image.base64), image.mimeType, image.name);
      const file = DriveApp.getFolderById(CC_IMAGES_FOLDER_ID).createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      imageLink = file.getUrl();
    }

    const fxFormula = `=J${lastRow}*IFERROR(VLOOKUP(E${lastRow},Currency!$A:$B,2,0),1)`;
    const t = "'2026 Transactions'";
    const fx = `IFERROR(VLOOKUP(${t}!D2:D,Currency!$A:$B,2,0),1)/IFERROR(VLOOKUP(E${lastRow},Currency!$A:$B,2,0),1)`;
    const jFormula = isCC
      ? `=D${lastRow}+SUMPRODUCT((${t}!B2:B=A${lastRow})*(${t}!F2:F="Expense")*${t}!C2:C*${fx})-SUMPRODUCT((${t}!E2:E=A${lastRow})*(${t}!F2:F="Transfer")*${t}!C2:C*${fx})`
      : `=D${lastRow}+SUMPRODUCT((${t}!B2:B=A${lastRow})*(${t}!F2:F="Income")*${t}!C2:C*${fx})+SUMPRODUCT((${t}!B2:B=A${lastRow})*(${t}!F2:F="Adjustment")*${t}!C2:C*${fx})+SUMPRODUCT((${t}!E2:E=A${lastRow})*(${t}!F2:F="Transfer")*${t}!C2:C*${fx})-SUMPRODUCT((${t}!B2:B=A${lastRow})*(${t}!F2:F="Expense")*${t}!C2:C*${fx})-SUMPRODUCT((${t}!B2:B=A${lastRow})*(${t}!F2:F="Transfer")*${t}!C2:C*${fx})`;

    sheet.appendRow([
      name, category,
      isCC ? Number(credit_limit) || '' : '',
      Number(init_balance) || 0,
      currency || 'PHP',
      fxFormula,
      imageLink,
      isCC ? Number(statement_day) || '' : '',
      isCC ? Number(due_day) || '' : '',
      jFormula,
      metadata || ''
    ]);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function addCategory(body) {
  try {
    const { type, category, subcategory } = body;
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    ss.getSheetByName('Txn_Categories').appendRow([category, type]);
    if (subcategory) ss.getSheetByName('Txn_Subcategories').appendRow([subcategory, category]);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function addInstallment(body) {
  try {
    const { installment, account, total_amount, terms, start_date } = body;
    const term_payment = Number(total_amount) / Number(terms);
    SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Installments')
      .appendRow([installment, account, Number(total_amount), Number(terms), term_payment, start_date]);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

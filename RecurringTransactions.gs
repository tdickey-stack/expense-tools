function showAddRecurringTransactionDialog() {
  const html = HtmlService.createHtmlOutputFromFile('AddRecurringTransaction')
    .setWidth(1050)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'Manage Recurring Transactions'
  );
}

function getRecurringTransactionManagerData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const chartSheet = ss.getSheetByName('Chart of Accounts');
  const recurringSheet = ss.getSheetByName('Recurring Transactions');

  if (!chartSheet) {
    throw new Error('Could not find "Chart of Accounts" sheet.');
  }

  const lastRow = chartSheet.getLastRow();

  const ministryAccounts = chartSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String);
  const ministrySubaccounts = chartSheet.getRange(2, 3, lastRow - 1, 1).getValues().flat().filter(String);

  const operatingRows = chartSheet.getRange(2, 6, lastRow - 1, 3).getValues()
    .filter(row => row[0] && row[1]);

  const designatedAccounts = chartSheet.getRange(2, 10, lastRow - 1, 1).getValues().flat().filter(String);

  const operatingCategories = [...new Set(operatingRows.map(row => row[0]))];

  const accountOptions = [
    ...ministryAccounts,
    ...operatingCategories,
    'Designated'
  ].filter(String).sort();

  const operatingMap = {};

  operatingRows.forEach(row => {
    const category = row[0];
    const subaccount = row[1];

    if (!operatingMap[category]) {
      operatingMap[category] = [];
    }

    operatingMap[category].push(subaccount);
  });

  let rules = [];

  if (recurringSheet && recurringSheet.getLastRow() >= 2) {
    rules = recurringSheet
      .getRange(2, 1, recurringSheet.getLastRow() - 1, 4)
      .getValues()
      .filter(row => row[0] || row[1] || row[2] || row[3])
      .map(row => ({
        keyword: row[0],
        account: row[1],
        subaccount: row[2],
        reason: row[3]
      }));
  }

  return {
    accountOptions,
    ministrySubaccounts,
    operatingMap,
    designatedAccounts,
    rules
  };
}

function saveRecurringTransactions(rules) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Recurring Transactions');

  if (!sheet) {
    sheet = ss.insertSheet('Recurring Transactions');
  }

  sheet.clearContents();

  sheet.getRange(1, 1, 1, 4).setValues([
    ['Keyword', 'Account', 'Subaccount', 'Reason for Expense']
  ]);

  if (rules.length > 0) {
    const values = rules.map(rule => [
      rule.keyword,
      rule.account,
      rule.subaccount,
      rule.reason
    ]);

    sheet.getRange(2, 1, values.length, 4).setValues(values);
  }

  sheet.hideSheet();

  return `${rules.length} recurring transaction${rules.length === 1 ? '' : 's'} saved.`;
}

function applyRecurringTransactions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  applyRecurringTransactions_(sheet, false);
}

function applyRecurringTransactions_(sheet, silent = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const recurringSheet = ss.getSheetByName('Recurring Transactions');

  if (!recurringSheet) {
    if (!silent) {
      SpreadsheetApp.getUi().alert('No Recurring Transactions sheet found.');
    }
    return;
  }

  const firstDataRow = 6;
  const descriptionColumn = 3; // Column C
  const unassignedAmountColumn = 4; // Column D
  const reasonColumn = 19; // Column S
  const accountHeaderRow = 4;
  const subaccountHeaderRow = 5;

  const totalsRow = findTotalsRow_(sheet);
  const lastColumn = sheet.getLastColumn();

  const lastRuleRow = recurringSheet.getLastRow();

  if (lastRuleRow < 2) {
    if (!silent) {
      SpreadsheetApp.getUi().alert('No recurring transactions have been added yet.');
    }
    return;
  }

  const rules = recurringSheet
    .getRange(2, 1, lastRuleRow - 1, 4)
    .getValues()
    .filter(row => row[0] && row[1] && row[2]);

  if (rules.length === 0) {
    if (!silent) {
      SpreadsheetApp.getUi().alert('No valid recurring transactions found.');
    }
    return;
  }

  const accountHeaders = sheet
    .getRange(accountHeaderRow, 1, 1, lastColumn)
    .getValues()[0];

  const subaccountHeaders = sheet
    .getRange(subaccountHeaderRow, 1, 1, lastColumn)
    .getValues()[0];

  const numRows = totalsRow - firstDataRow;

  if (numRows <= 0) {
    if (!silent) {
      SpreadsheetApp.getUi().alert('No transaction rows found.');
    }
    return;
  }

  const descriptions = sheet
    .getRange(firstDataRow, descriptionColumn, numRows, 1)
    .getValues();

  const amounts = sheet
    .getRange(firstDataRow, unassignedAmountColumn, numRows, 1)
    .getValues();

  let matchedCount = 0;
  let skippedCount = 0;

  for (let i = 0; i < numRows; i++) {
    const description = String(descriptions[i][0]).toLowerCase().trim();
    const descriptionKey = normalizeRecurringVendorKey_(description);
    const amount = amounts[i][0];

    if (!description || amount === '') {
      skippedCount++;
      continue;
    }

    for (const rule of rules) {
      const keyword = String(rule[0]).toLowerCase().trim();
      const keywordKey = normalizeRecurringVendorKey_(keyword);
      const account = String(rule[1]).trim();
      const subaccount = String(rule[2]).trim();
      const reason = String(rule[3] || '').trim();

      if (
  description.includes(keyword) ||
  descriptionKey.includes(keywordKey) ||
  keywordKey.includes(descriptionKey)
) {
        const targetColumn = findRecurringTransactionColumn_(
          accountHeaders,
          subaccountHeaders,
          account,
          subaccount
        );

        if (targetColumn) {
          const row = firstDataRow + i;

          sheet.getRange(row, targetColumn).setValue(amount);
          sheet.getRange(row, unassignedAmountColumn).clearContent();

          if (reason) {
            sheet.getRange(row, reasonColumn).setValue(reason);
          }

          matchedCount++;
        }

        break;
      }
    }
  }

  if (!silent) {
    SpreadsheetApp.getUi().alert(
      `Recurring transactions applied.\n\nMatched: ${matchedCount}\nSkipped/Unassigned: ${skippedCount}`
    );
  }
}

function findRecurringTransactionColumn_(accountHeaders, subaccountHeaders, account, subaccount) {
  for (let i = 0; i < accountHeaders.length; i++) {
    const headerAccount = String(accountHeaders[i]).trim();
    const headerSubaccount = String(subaccountHeaders[i]).trim();

    if (headerAccount === account && headerSubaccount === subaccount) {
      return i + 1;
    }
  }

  return null;
}

/*
=== AUTO DETECT RECURRING TRANSACTIONS ===
*/
function detectRecurringTransactionsManual() {
  const ui = SpreadsheetApp.getUi();
  const candidates = findRecurringTransactionCandidates_();

  if (candidates.length === 0) {
    ui.alert('No recurring transaction patterns found.');
    return;
  }

  let addedCount = 0;
  let neverCount = 0;

  candidates.forEach(candidate => {
    const message =
      `Possible recurring transaction detected:\n\n` +
      `Vendor: ${candidate.vendorDisplay}\n` +
      `Typical Day: ${candidate.dayOfMonth}\n` +
      `Amount: ${formatCurrencyForRecurringDetection_(candidate.amount)}\n` +
      `Account: ${candidate.account}\n` +
      `Subaccount: ${candidate.subaccount}\n\n` +
      `Found in ${candidate.monthCount} different months.\n\n` +
      `Choose YES to add it.\n` +
      `Choose NO to skip for now.\n` +
      `Choose CANCEL to never ask again for this vendor.`;

    const response = ui.alert(
      'Recurring Transaction Detected',
      message,
      ui.ButtonSet.YES_NO_CANCEL
    );

    if (response === ui.Button.YES) {
      addRecurringTransactionFromCandidate_(candidate);
      addedCount++;
    }

    if (response === ui.Button.CANCEL) {
      addRecurringIgnoreVendor_(candidate.vendorKey, candidate.vendorDisplay);
      neverCount++;
    }
  });

  ui.alert(
    `Recurring transaction detection complete.\n\nAdded: ${addedCount}\nNever ask again: ${neverCount}`
  );
}

function findRecurringTransactionCandidates_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ignoredVendorKeys = getRecurringIgnoreVendorKeys_();
  const existingVendorKeys = getExistingRecurringVendorKeys_();

  const reportSheets = ss.getSheets().filter(sheet => isRecurringDetectionReportSheet_(sheet));

  const allTransactions = [];

  reportSheets.forEach(sheet => {
    allTransactions.push(...getRecurringDetectionTransactionsFromSheet_(sheet));
  });

  const groupedByVendor = {};

  allTransactions.forEach(tx => {
    if (!tx.vendorKey) return;
    if (ignoredVendorKeys.includes(tx.vendorKey)) return;
    if (existingVendorKeys.includes(tx.vendorKey)) return;

    if (!groupedByVendor[tx.vendorKey]) {
      groupedByVendor[tx.vendorKey] = [];
    }

    groupedByVendor[tx.vendorKey].push(tx);
  });

  const candidates = [];

  Object.keys(groupedByVendor).forEach(vendorKey => {
    const matches = groupedByVendor[vendorKey];

    const uniqueMonths = [...new Set(matches.map(tx => tx.monthKey))];

    // Change this to 3 later if 2 feels too noisy.
    if (uniqueMonths.length < 2) {
      return;
    }

    const days = matches.map(tx => tx.dayOfMonth);
    const averageDay = Math.round(
      days.reduce((sum, day) => sum + day, 0) / days.length
    );

    const closeDateMatches = matches.filter(tx =>
      Math.abs(tx.dayOfMonth - averageDay) <= 4
    );

    if (closeDateMatches.length < 2) {
      return;
    }

    const commonAccount = getMostCommonAccountPair_(matches);
    const sample = matches[0];

    candidates.push({
      vendorKey,
      vendorDisplay: sample.description,
      dayOfMonth: averageDay,
      amount: sample.amount,
      account: commonAccount.account || '',
      subaccount: commonAccount.subaccount || '',
      reason: sample.description,
      monthCount: uniqueMonths.length,
      matches
    });
  });

  return candidates;
}

function getRecurringDetectionTransactionsFromSheet_(sheet) {
  const firstDataRow = 6;
  const maxRowsToCheck = 300;

  const dateColumn = 2; // B
  const descriptionColumn = 3; // C
  const amountColumn = 4; // D

  const lastRow = sheet.getLastRow();

  if (lastRow < firstDataRow) {
    return [];
  }

  const rowsToCheck = Math.min(maxRowsToCheck, lastRow - firstDataRow + 1);

  const data = sheet
    .getRange(firstDataRow, dateColumn, rowsToCheck, 3) // B:D only
    .getDisplayValues();

  const account = String(sheet.getRange(4, amountColumn).getDisplayValue() || '').trim();
  const subaccount = String(sheet.getRange(5, amountColumn).getDisplayValue() || '').trim();

  const transactions = [];

  for (let i = 0; i < data.length; i++) {
    const dateText = data[i][0];
    const description = String(data[i][1] || '').trim();
    const amountText = data[i][2];

    if (!description || !amountText) {
      continue;
    }

    const parsedDate = parseRecurringDetectionDate_(dateText);

    if (!parsedDate) {
      continue;
    }

    const amount = parseRecurringDetectionCurrency_(amountText);

    if (!amount) {
      continue;
    }

    transactions.push({
      sheetName: sheet.getName(),
      rowNumber: firstDataRow + i,
      date: parsedDate,
      monthKey: Utilities.formatDate(
        parsedDate,
        Session.getScriptTimeZone(),
        'yyyy-MM'
      ),
      dayOfMonth: parsedDate.getDate(),
      description,
      vendorKey: normalizeRecurringVendorKey_(description),
      amount,
      account,
      subaccount
    });
  }

  return transactions;
}

function findAssignedRecurringDetectionAccount_(rowValues, accountHeaders, subaccountHeaders, unassignedAmountColumn) {
  for (let i = 0; i < rowValues.length; i++) {
    const columnNumber = i + 1;

    if (columnNumber === unassignedAmountColumn) {
      continue;
    }

    const value = rowValues[i];

    if (value === '' || value === null || value === 0) {
      continue;
    }

    const account = String(accountHeaders[i] || '').trim();
    const subaccount = String(subaccountHeaders[i] || '').trim();

    if (!account || !subaccount) {
      continue;
    }

    return {
      amount: Number(value),
      account,
      subaccount
    };
  }

  return null;
}

function normalizeRecurringVendorKey_(description) {
  return String(description || '')
    .toLowerCase()
    .replace(/\b\d{1,2}\/\d{1,2}(\/\d{2,4})?\b/g, '')
    .replace(/\b\d{4,}\b/g, '')
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\b(inc|llc|co|company|payment|purchase|pos|debit|card|online|web)\b/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .split(' ')
    .slice(0, 3)
    .join(' ');
}

function isRecurringDetectionReportSheet_(sheet) {
  const name = sheet.getName();

  const excludedNames = [
    'Chart of Accounts',
    'Recurring Transactions',
    'Recurring Ignore List',
    '_Helper'
  ];

  if (excludedNames.includes(name)) {
    return false;
  }

  if (name.startsWith('_')) {
    return false;
  }

  return sheet.getLastRow() >= 6;
}

function getExistingRecurringVendorKeys_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Recurring Transactions');

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  return sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 1)
    .getValues()
    .flat()
    .filter(String)
    .map(normalizeRecurringVendorKey_);
}

function getRecurringIgnoreVendorKeys_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Recurring Ignore List');

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  return sheet
    .getRange(2, 1, sheet.getLastRow() - 1, 1)
    .getValues()
    .flat()
    .filter(String);
}

function addRecurringIgnoreVendor_(vendorKey, vendorDisplay) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Recurring Ignore List');

  if (!sheet) {
    sheet = ss.insertSheet('Recurring Ignore List');
    sheet.getRange(1, 1, 1, 3).setValues([
      ['Vendor Key', 'Vendor Display', 'Date Added']
    ]);
    sheet.hideSheet();
  }

  sheet.appendRow([
    vendorKey,
    vendorDisplay,
    new Date()
  ]);
}

function addRecurringTransactionFromCandidate_(candidate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Recurring Transactions');

  if (!sheet) {
    sheet = ss.insertSheet('Recurring Transactions');
    sheet.getRange(1, 1, 1, 4).setValues([
      ['Keyword', 'Account', 'Subaccount', 'Reason for Expense']
    ]);
    sheet.hideSheet();
  }

  const keyword = candidate.vendorKey || candidate.vendorDisplay;

  if (!keyword) {
    throw new Error('No vendor keyword found.');
  }

  sheet.appendRow([
    keyword,
    candidate.account || '',
    candidate.subaccount || '',
    candidate.reason || ''
  ]);

  return true;
}

function formatCurrencyForRecurringDetection_(value) {
  return Number(value || 0).toLocaleString('en-US', {
    style: 'currency',
    currency: 'USD'
  });
}

function findDetectedTransactionAmountAndAccount_(rowValues, accountHeaders, subaccountHeaders) {
  for (let i = 0; i < rowValues.length; i++) {
    const value = rowValues[i];

    if (value === '' || value === null || value === 0) {
      continue;
    }

    const account = String(accountHeaders[i] || '').trim();
    const subaccount = String(subaccountHeaders[i] || '').trim();

    if (!account || !subaccount) {
      continue;
    }

    return {
      amount: Number(value),
      account,
      subaccount
    };
  }

  return null;
}

function getMostCommonAccountPair_(matches) {
  const counts = {};

  matches.forEach(tx => {
    if (!tx.account || !tx.subaccount) return;

    const key = `${tx.account}|||${tx.subaccount}`;

    if (!counts[key]) {
      counts[key] = {
        account: tx.account,
        subaccount: tx.subaccount,
        count: 0
      };
    }

    counts[key].count++;
  });

  const sorted = Object.values(counts).sort((a, b) => b.count - a.count);

  return sorted[0] || {
    account: '',
    subaccount: ''
  };
}

function parseRecurringDetectionDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value;
  }

  const text = String(value || '').trim();

  if (!text) {
    return null;
  }

  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);

  if (!match) {
    return null;
  }

  const month = Number(match[1]);
  const day = Number(match[2]);
  let year = Number(match[3]);

  if (year < 100) {
    year += 2000;
  }

  const date = new Date(year, month - 1, day);

  return isNaN(date.getTime()) ? null : date;
}

function parseRecurringDetectionCurrency_(value) {
  const text = String(value || '')
    .trim()
    .replace(/\$\s*\(([\d,]+\.\d{2})\)/, '-$1')
    .replace(/\(([\d,]+\.\d{2})\)/, '-$1')
    .replace(/[$,\s]/g, '');

  return Number(text) || 0;
}

/*
==Detection Dialog Box==
*/
function showRecurringDetectionDialog() {
  const html = HtmlService.createHtmlOutputFromFile('DetectedRecurringTransaction')
    .setWidth(650)
    .setHeight(800);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'Detected Recurring Transaction'
  );
}

function showRecurringDetectionDialogAfterImport() {
  Utilities.sleep(800);
  showRecurringDetectionDialog();
}
/*
===Old Detection Next Candidate Script===
function getNextRecurringDetectionCandidate(excludedVendorKeys) {
  try {
    excludedVendorKeys = excludedVendorKeys || [];

    const candidates = findRecurringTransactionCandidates_()
      .filter(candidate => !excludedVendorKeys.includes(candidate.vendorKey));

    if (!candidates || candidates.length === 0) {
      return {
        done: true,
        message: 'No more recurring transaction patterns found.'
      };
    }

    const managerData = getRecurringTransactionManagerData();
    const candidate = candidates[0];

    return {
      done: false,
      candidate: {
        vendorKey: candidate.vendorKey || '',
        vendorDisplay: candidate.vendorDisplay || '',
        dayOfMonth: candidate.dayOfMonth || '',
        amount: candidate.amount || 0,
        account: candidate.account || '',
        subaccount: candidate.subaccount || '',
        reason: '',
        monthCount: candidate.monthCount || 0
      },
      accountOptions: managerData.accountOptions || [],
      ministrySubaccounts: managerData.ministrySubaccounts || [],
      operatingMap: managerData.operatingMap || {},
      designatedAccounts: managerData.designatedAccounts || []
    };

  } catch (err) {
    return {
      done: true,
      message: 'Error loading recurring detection: ' + err.message
    };
  }
}
*/

function handleRecurringDetectionDecision(decision) {
  if (!decision) {
    throw new Error('No decision data received.');
  }

  const action = decision.action;
  const candidate = decision.candidate || {};

  if (action === 'add') {
    if (!decision.account || !decision.subaccount) {
      throw new Error('Account and subaccount are required.');
    }

    addRecurringTransactionFromCandidate_({
      vendorKey: candidate.vendorKey,
      vendorDisplay: candidate.vendorDisplay,
      account: decision.account,
      subaccount: decision.subaccount,
      reason: decision.reason || ''
    });

    return 'Added recurring transaction.';
  }

  if (action === 'never') {
    addRecurringIgnoreVendor_(candidate.vendorKey, candidate.vendorDisplay);
    return 'Vendor added to Never Ask Again.';
  }

  if (action === 'skip') {
    return 'Skipped.';
  }

  throw new Error('Unknown decision action: ' + action);
}


function getRecurringDetectionCandidatesForUi() {
  try {
    const candidates = findRecurringTransactionCandidates_();

    if (!candidates || candidates.length === 0) {
      return {
        done: true,
        message: 'No recurring transaction patterns found.'
      };
    }

    const managerData = getRecurringTransactionManagerData();

    const cleanCandidates = candidates.map(candidate => ({
      vendorKey: candidate.vendorKey || '',
      vendorDisplay: candidate.vendorDisplay || '',
      dayOfMonth: candidate.dayOfMonth || '',
      amount: candidate.amount || 0,
      account: candidate.account || '',
      subaccount: candidate.subaccount || '',
      reason: '',
      monthCount: candidate.monthCount || 0
    }));

    return {
      done: false,
      candidates: cleanCandidates,
      accountOptions: managerData.accountOptions || [],
      ministrySubaccounts: managerData.ministrySubaccounts || [],
      operatingMap: managerData.operatingMap || {},
      designatedAccounts: managerData.designatedAccounts || []
    };

  } catch (err) {
    return {
      done: true,
      message: 'Error loading recurring detection: ' + err.message
    };
  }
}

function triggerRecurringDetectionAfterImport_() {
  try {
    const candidates = findRecurringTransactionCandidates_();

    if (!candidates || candidates.length === 0) {
      return;
    }

    showRecurringDetectionDialog();

  } catch (err) {
    console.log('Recurring detection after import failed: ' + err.message);
  }
}

/*
Debugging
*/

function debugOpenAiFast() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let message = '';

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();

    if (
      name === 'Chart of Accounts' ||
      name === 'Recurring Transactions' ||
      name === 'Recurring Ignore List'
    ) {
      return;
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < 6) {
      return;
    }

    const numRows = lastRow - 5;

    const dates = sheet.getRange(6, 2, numRows, 1).getDisplayValues(); // B
    const descriptions = sheet.getRange(6, 3, numRows, 1).getDisplayValues(); // C
    const amounts = sheet.getRange(6, 4, numRows, 1).getDisplayValues(); // D

    for (let i = 0; i < numRows; i++) {
      const description = String(descriptions[i][0] || '');

      if (!description.toLowerCase().includes('openai')) {
        continue;
      }

      const dateText = dates[i][0];
      const amountText = amounts[i][0];

      message +=
        `Sheet: ${name}\n` +
        `Row: ${i + 6}\n` +
        `Date: ${dateText}\n` +
        `Description: ${description}\n` +
        `Vendor Key: ${normalizeRecurringVendorKey_(description)}\n` +
        `Amount: ${amountText}\n` +
        `Parsed Amount: ${parseRecurringDetectionCurrency_(amountText)}\n\n`;
    }
  });

  ui.alert(message || 'No OpenAI transactions found.');
}

function debugOpenAiAcrossSheetsCapped() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const excludedSheets = [
    'Chart of Accounts',
    'Recurring Transactions',
    'Recurring Ignore List',
    '_Helper'
  ];

  const startRow = 6;
  const maxRows = 145; // rows 6–150

  let message = 'Checking OpenAI across sheets, capped to rows 6–150\n\n';
  let foundCount = 0;

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();

    if (excludedSheets.includes(name)) {
      return;
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < startRow) {
      return;
    }

    const rowsToCheck = Math.min(maxRows, lastRow - startRow + 1);

    const dates = sheet.getRange(startRow, 2, rowsToCheck, 1).getDisplayValues(); // B
    const descriptions = sheet.getRange(startRow, 3, rowsToCheck, 1).getDisplayValues(); // C
    const amounts = sheet.getRange(startRow, 4, rowsToCheck, 1).getDisplayValues(); // D

    for (let i = 0; i < rowsToCheck; i++) {
      const description = String(descriptions[i][0] || '');

      if (!description.toLowerCase().includes('openai')) {
        continue;
      }

      foundCount++;

      const dateText = dates[i][0];
      const amountText = amounts[i][0];

      message +=
        `Sheet: ${name}\n` +
        `Row: ${startRow + i}\n` +
        `Date: ${dateText}\n` +
        `Description: ${description}\n` +
        `Vendor Key: ${normalizeRecurringVendorKey_(description)}\n` +
        `Amount: ${amountText}\n` +
        `Parsed Amount: ${parseRecurringDetectionCurrency_(amountText)}\n\n`;
    }
  });

  message += `Total OpenAI rows found: ${foundCount}`;

  ui.alert(message);
}

function debugOpenAiNoHelpers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const excludedSheets = [
    'Chart of Accounts',
    'Recurring Transactions',
    'Recurring Ignore List',
    '_Helper'
  ];

  const startRow = 6;
  const maxRows = 145;

  let output = [];
  let foundCount = 0;

  const sheets = ss.getSheets();

  for (let s = 0; s < sheets.length; s++) {
    const sheet = sheets[s];
    const name = sheet.getName();

    if (excludedSheets.includes(name)) {
      continue;
    }

    const lastRow = sheet.getLastRow();

    if (lastRow < startRow) {
      continue;
    }

    const rowsToCheck = Math.min(maxRows, lastRow - startRow + 1);

    const data = sheet
      .getRange(startRow, 2, rowsToCheck, 4) // B:D only
      .getDisplayValues();

    for (let i = 0; i < data.length; i++) {
      const dateText = data[i][0];
      const description = String(data[i][2] || '');
      const amountText = data[i][3];

      if (!description.toLowerCase().includes('openai')) {
        continue;
      }

      foundCount++;

      output.push(
        `Sheet: ${name}`,
        `Row: ${startRow + i}`,
        `Date: ${dateText}`,
        `Description: ${description}`,
        `Amount: ${amountText}`,
        ''
      );
    }
  }

  Logger.log(output.join('\n'));
  Logger.log(`Total OpenAI rows found: ${foundCount}`);
}

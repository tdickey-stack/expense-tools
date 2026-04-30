function showSplitTransactionDialog() {
  const html = HtmlService.createHtmlOutputFromFile('SplitTransaction_UI')
    .setWidth(1100)
    .setHeight(700);

  SpreadsheetApp.getUi().showModalDialog(html, 'Split Transaction');
}

function getSplitTxData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const chartSheet = ss.getSheetByName('Chart of Accounts');

  if (!chartSheet) {
    throw new Error('Could not find "Chart of Accounts" sheet.');
  }

  const firstDataRow = 6;
  const totalsRow = findTotalsRow_(sheet);
  const numRows = totalsRow - firstDataRow;

  const values = sheet
    .getRange(firstDataRow, 3, numRows, 2) // C:D
    .getDisplayValues();

  const transactions = values
    .map((row, index) => {
      const description = row[0];
      const amountText = row[1];

      return {
        rowNumber: firstDataRow + index,
        label: `${description} | ${amountText}`,
        description,
        amountText
      };
    })
    .filter(tx => tx.description && tx.amountText);

  const chartData = getSplitChartData_(chartSheet);

  return {
    transactions,
    accountOptions: chartData.accountOptions,
    ministrySubaccounts: chartData.ministrySubaccounts,
    operatingMap: chartData.operatingMap,
    designatedAccounts: chartData.designatedAccounts
  };
}

function getSplitTransactionList_(sheet) {
  const firstDataRow = 6;
  const descriptionColumn = 3; // C
  const amountColumn = 4; // D
  const totalsRow = findTotalsRow_(sheet);

  const transactions = [];
  const numRows = totalsRow - firstDataRow;

  if (numRows <= 0) {
    return transactions;
  }

  const descriptions = sheet
    .getRange(firstDataRow, descriptionColumn, numRows, 1)
    .getDisplayValues();

  const amounts = sheet
    .getRange(firstDataRow, amountColumn, numRows, 1)
    .getDisplayValues();

  for (let i = 0; i < numRows; i++) {
    const description = descriptions[i][0];
    const amountText = amounts[i][0];

    if (!description || !amountText) {
      continue;
    }

    const amount = parseCurrency_(amountText);

    transactions.push({
      rowNumber: firstDataRow + i,
      label: `${description} | ${amountText}`,
      description: description,
      amount: amount
    });
  }

  return transactions;
}

function getSplitChartData_(chartSheet) {
  const lastRow = chartSheet.getLastRow();

  const ministryAccounts = chartSheet
    .getRange(2, 1, lastRow - 1, 1) // A
    .getDisplayValues()
    .flat()
    .filter(String);

  const ministrySubaccounts = chartSheet
    .getRange(2, 3, lastRow - 1, 1) // C
    .getDisplayValues()
    .flat()
    .filter(String);

  const operatingRows = chartSheet
    .getRange(2, 6, lastRow - 1, 3) // F:H
    .getDisplayValues()
    .filter(row => row[0] && row[1]);

  const designatedAccounts = chartSheet
    .getRange(2, 10, lastRow - 1, 1) // J
    .getDisplayValues()
    .flat()
    .filter(String);

  const operatingCategories = [...new Set(operatingRows.map(row => row[0]))];

  const accountOptions = [
    ...ministryAccounts,
    ...operatingCategories,
    'Designated'
  ]
    .filter(String)
    .sort();

  const operatingMap = {};

  operatingRows.forEach(row => {
    const category = row[0];
    const subaccount = row[1];

    if (!operatingMap[category]) {
      operatingMap[category] = [];
    }

    operatingMap[category].push(subaccount);
  });

  return {
    accountOptions: accountOptions,
    ministrySubaccounts: ministrySubaccounts,
    operatingMap: operatingMap,
    designatedAccounts: designatedAccounts
  };
}

function applySplitTx(splitData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const rowNumber = Number(splitData.rowNumber);
  const mode = splitData.mode;
  const splits = splitData.splits || [];

  const amountColumn = 4; // D
  const reasonColumn = 19; // S
  const accountHeaderRow = 4;
  const subaccountHeaderRow = 5;

  if (!rowNumber) {
    throw new Error('No transaction selected.');
  }

  if (splits.length < 2) {
    throw new Error('A split transaction needs at least two accounts.');
  }

  const originalAmount = Number(sheet.getRange(rowNumber, amountColumn).getValue());

  if (!originalAmount) {
    throw new Error('Selected transaction does not have an amount in Column D.');
  }

  const sign = originalAmount < 0 ? -1 : 1;
  const originalAbs = Math.abs(originalAmount);

  const cleanedSplits = splits
    .filter(split => split.account && split.subaccount && split.value !== '')
    .map(split => ({
      account: String(split.account).trim(),
      subaccount: String(split.subaccount).trim(),
      value: Number(split.value),
      reason: String(split.reason || '').trim()
    }));

  if (cleanedSplits.length < 2) {
    throw new Error('Please complete at least two split lines.');
  }

  let calculatedSplits = [];

  if (mode === 'percentage') {
    const totalPercent = cleanedSplits.reduce((sum, split) => sum + split.value, 0);

    if (Math.abs(totalPercent - 100) > 0.01) {
      throw new Error(`Percentages must total 100%. Current total: ${totalPercent.toFixed(2)}%`);
    }

    calculatedSplits = cleanedSplits.map(split => ({
      ...split,
      amount: roundCurrency_(originalAmount * (split.value / 100))
    }));
  } else if (mode === 'amount') {
    const totalAmount = cleanedSplits.reduce((sum, split) => sum + Math.abs(split.value), 0);

    if (Math.abs(totalAmount - originalAbs) > 0.01) {
      throw new Error(
        `Split amounts must total ${originalAbs.toFixed(2)}. Current total: ${totalAmount.toFixed(2)}`
      );
    }

    calculatedSplits = cleanedSplits.map(split => ({
      ...split,
      amount: roundCurrency_(Math.abs(split.value) * sign)
    }));
  } else {
    throw new Error('Invalid split mode.');
  }

  const lastColumn = sheet.getLastColumn();

  const accountHeaders = sheet
    .getRange(accountHeaderRow, 1, 1, lastColumn)
    .getDisplayValues()[0];

  const subaccountHeaders = sheet
    .getRange(subaccountHeaderRow, 1, 1, lastColumn)
    .getDisplayValues()[0];

  // First determine all target columns
  const targetColumns = calculatedSplits.map(split => {
    const targetColumn = findSplitTransactionColumn_(
      accountHeaders,
      subaccountHeaders,
      split.account,
      split.subaccount
    );

    if (!targetColumn) {
      throw new Error(`Could not find column for ${split.account} / ${split.subaccount}`);
    }

    return targetColumn;
  });

  // 🔥 CRITICAL FIX:
  // Clear original transaction BEFORE applying splits
  sheet.getRange(rowNumber, amountColumn).clearContent();

  // Apply split values
  calculatedSplits.forEach((split, index) => {
    const targetColumn = targetColumns[index];

    const currentValue = Number(sheet.getRange(rowNumber, targetColumn).getValue()) || 0;
    sheet.getRange(rowNumber, targetColumn).setValue(
      roundCurrency_(currentValue + split.amount)
    );
  });

  const reasons = calculatedSplits
    .map(split => split.reason)
    .filter(String);

  if (reasons.length > 0) {
    sheet.getRange(rowNumber, reasonColumn).setValue(reasons.join(' | '));
  }

  if (typeof resizeVisaSheet === 'function') {
    resizeVisaSheet(true);
  }

  return 'Transaction split successfully.';
}

function findSplitTransactionColumn_(accountHeaders, subaccountHeaders, account, subaccount) {
  for (let i = 0; i < accountHeaders.length; i++) {
    const headerAccount = String(accountHeaders[i]).trim();
    const headerSubaccount = String(subaccountHeaders[i]).trim();

    if (headerAccount === account && headerSubaccount === subaccount) {
      return i + 1;
    }
  }

  return null;
}

function parseCurrency_(value) {
  const text = String(value)
    .replace(/[$,\s]/g, '')
    .replace(/[()]/g, '-');

  return Number(text);
}

function roundCurrency_(value) {
  return Math.round(value * 100) / 100;
}

function getSplitTxTestData() {
  return {
    transactions: [
      {
        rowNumber: 6,
        label: 'TEST TRANSACTION | -123.45',
        description: 'TEST TRANSACTION',
        amount: -123.45
      }
    ]
  };
}

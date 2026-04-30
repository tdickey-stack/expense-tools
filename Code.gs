function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const importMenu = ui.createMenu('Import')
    .addItem('Import Statement (CSV)', 'showCsvImportDialog')
    .addItem('Import Statement (PDF)', 'showPdfImportDialog');

  const recurringMenu = ui.createMenu('Recurring Transactions')
    .addItem('Manage Recurring Transactions', 'showAddRecurringTransactionDialog')
    .addItem('Apply Recurring Transactions', 'applyRecurringTransactions')
    .addItem('Detect Recurring Transactions', 'showRecurringDetectionDialog');

  ui.createMenu('⚡ Expense Tools')
    .addSubMenu(importMenu)
    .addSubMenu(recurringMenu)
    .addItem('Split Transaction', 'showSplitTransactionDialog')
    .addSeparator()
    .addItem('Clear Report', 'clearVisaSheet')
    .addItem('Resize Sheet', 'resizeVisaSheet')
    .addSeparator()
    .addItem('Check for Updates', 'checkForExpenseToolsUpdates')
    .addItem('The Update Works', ' ')
    .addToUi();
}

function showCsvImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('CsvImport')
    .setWidth(500)
    .setHeight(610);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Bank CSV');
}

function showPdfImportDialog() {
  const html = HtmlService.createHtmlOutputFromFile('PdfImport')
    .setWidth(500)
    .setHeight(610);
  SpreadsheetApp.getUi().showModalDialog(html, 'Import Bank PDF');
}

function importBankPdf(base64Data, fileName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64Data),
    MimeType.PDF,
    fileName
  );

  const resource = {
  name: fileName,
  mimeType: MimeType.GOOGLE_DOCS
};

const docFile = Drive.Files.create(resource, blob);
const doc = DocumentApp.openById(docFile.id);
const text = doc.getBody().getText();

DriveApp.getFileById(docFile.id).setTrashed(true);

  const output = parsePdfTransactions_(text);

  if (output.length === 0) {
    throw new Error('No transactions found in the PDF.');
  }

  insertTransactions_(sheet, output);

  return `${output.length} transactions imported from PDF.`;
}

function importBankCsv(csvText) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const rows = Utilities.parseCsv(csvText);
  if (rows.length < 2) {
    throw new Error('No transaction rows found in the CSV.');
  }

  const headers = rows[0];

  const dateIndex = headers.indexOf('Value Date of Payment');
  const amountIndex = headers.indexOf('Amount');
  const detailIndex = headers.indexOf('Transaction Detail');

  if (dateIndex === -1 || amountIndex === -1 || detailIndex === -1) {
    throw new Error(
      'Required columns were not found. Check for Value Date of Payment, Amount, and Transaction Detail.'
    );
  }

  const output = rows
    .slice(1)
    .filter(row => row[dateIndex] || row[amountIndex] || row[detailIndex])
    .map(row => [
      row[dateIndex],
      row[detailIndex],
      Number(row[amountIndex])
    ]);

  if (output.length === 0) {
    throw new Error('No valid transactions found.');
  }

  insertTransactions_(sheet, output);

  return `${output.length} transactions imported from CSV.`;
}

function parsePdfTransactions_(text) {
  const transactions = [];

  let normalizedText = text
    .replace(/\r?\n/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  // Remove page-break / statement boilerplate that can appear between transactions
  normalizedText = normalizedText
  .replace(/Transactions continued on next page.*?Trans Post Reference Number Description Amount/gi, ' ')
  .replace(/NOTICE: SEE REVERSE SIDE FOR IMPORTANT INFORMATION.*?Trans Post Reference Number Description Amount/gi, ' ')
  .replace(/THIS IS A MEMO.*?Trans Post Reference Number Description Amount/gi, ' ')
  .replace(/FIRST CITIZENS BANK.*?Trans Post Reference Number Description Amount/gi, ' ')
  .replace(/PO Box 2360 Omaha NE 68103-2360.*?Trans Post Reference Number Description Amount/gi, ' ')
  .replace(/Transactions Since Last Statement \(continued\)/gi, ' ')
  .replace(/Page \d+ of \d+/gi, ' ')
  .replace(/Account Number: XXXX XXXX XXXX \d+/gi, ' ')
  .replace(/Trans Post Reference Number Description Amount/gi, ' ')
  .replace(/\s+/g, ' ')
  .trim();

  const transactionRegex =
    /(\d{2}\/\d{2})\s+(\d{2}\/\d{2})\s+([A-Z0-9]{10,})\s+(.+?)\s+(-?\$?[\d,]+\.\d{2})(?=\s+\d{2}\/\d{2}\s+\d{2}\/\d{2}\s+[A-Z0-9]{10,}|\s+TOTAL PURCHASES|$)/g;

  let match;

  while ((match = transactionRegex.exec(normalizedText)) !== null) {
    const transDate = match[1];
    let description = match[4].trim();
    const amount = Number(match[5].replace(/[$,]/g, '')) * -1;

    if (
      description.includes('TOTAL PURCHASES') ||
      description.includes('PAYMENT DUE') ||
      description.includes('FIRST CITIZENS BANK') ||
      description.includes('DO NOT PAY')
    ) {
      continue;
    }

    transactions.push([
      transDate,
      description,
      amount
    ]);
  }

  return transactions;
}

function insertTransactions_(sheet, output) {
  const firstDataRow = 6;
  const startColumn = 2; // Column B
  const numColumns = 3; // B:D
  const formulaColumn = 18; // Column R

  const totalsRow = findTotalsRow_(sheet);
  const nextRow = getNextAvailableTransactionRow_(
    sheet,
    firstDataRow,
    totalsRow,
    startColumn
  );

  const availableRowsBeforeTotals = totalsRow - nextRow;

  if (output.length > availableRowsBeforeTotals) {
    const rowsNeeded = output.length - availableRowsBeforeTotals;
    sheet.insertRowsBefore(totalsRow, rowsNeeded);
  }

  sheet.getRange(firstDataRow, 1, 1, sheet.getLastColumn())
    .copyTo(
      sheet.getRange(nextRow, 1, output.length, sheet.getLastColumn()),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );

  sheet.getRange(nextRow, startColumn, output.length, numColumns)
    .setValues(output);

  copyFormulaFromRow6_(sheet, nextRow, output.length, formulaColumn);

  // Automatically apply existing recurring transaction rules first.
  applyRecurringTransactions_(sheet, true);

  sortTransactionsByDate_(sheet);

  resizeVisaSheet(true);

  // After import + existing recurring rules, check for NEW recurring patterns.
  //triggerRecurringDetectionAfterImport_();
}

function findTotalsRow_(sheet) {
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === 'Account Totals') {
      return i + 1;
    }
  }

  throw new Error('Could not find "Account Totals" in column A.');
}

function getNextAvailableTransactionRow_(sheet, firstDataRow, totalsRow, checkColumn) {
  const numRows = totalsRow - firstDataRow;

  if (numRows <= 0) return firstDataRow;

  const values = sheet
    .getRange(firstDataRow, checkColumn, numRows, 1)
    .getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== '') {
      return firstDataRow + i + 1;
    }
  }

  return firstDataRow;
}

function copyFormulaFromRow6_(sheet, startRow, numRows, formulaColumn) {
  const sourceFormula = sheet.getRange(6, formulaColumn).getFormulaR1C1();

  if (!sourceFormula) return;

  sheet
    .getRange(startRow, formulaColumn, numRows, 1)
    .setFormulaR1C1(sourceFormula);
}

function sortTransactionsByDate_(sheet) {
  const firstDataRow = 6;
  const dateColumn = 2;
  const totalsRow = findTotalsRow_(sheet);
  const numRows = totalsRow - firstDataRow;

  if (numRows <= 1) return;

  sheet
    .getRange(firstDataRow, 1, numRows, sheet.getLastColumn())
    .sort({
      column: dateColumn,
      ascending: true
    });
}

function resizeVisaSheet(silent = false) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  SpreadsheetApp.flush();
  Utilities.sleep(500);

  const lastColumn = sheet.getLastColumn();
  const totalsRow = findTotalsRow_(sheet);

  // Resize columns one by one
  for (let col = 1; col <= lastColumn; col++) {
    sheet.autoResizeColumn(col);
  }

  // Resize transaction rows from row 6 through the row before Account Totals
  if (totalsRow > 6) {
    sheet.autoResizeRows(6, totalsRow - 6);
  }

  if (!silent) {
    SpreadsheetApp.getUi().alert('Sheet resized.');
  }
}

function clearVisaSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const firstDataRow = 6;
  const formulaColumn = 18; // Column R
  const totalsRow = findTotalsRow_(sheet);

  // How many rows exist between row 6 and Account Totals
  const rowsBetween = totalsRow - firstDataRow;

  if (rowsBetween > 1) {
    // Delete everything except row 6
    sheet.deleteRows(firstDataRow + 1, rowsBetween - 1);
  }

  // Preserve the formula in Row 6 (Column R)
  const row6Formula = sheet
    .getRange(firstDataRow, formulaColumn)
    .getFormulaR1C1();

  // Clear row 6 contents EXCEPT formatting
  sheet.getRange(firstDataRow, 1, 1, sheet.getLastColumn()).clearContent();

  // Restore formula
  if (row6Formula) {
    sheet.getRange(firstDataRow, formulaColumn).setFormulaR1C1(row6Formula);
  }

  SpreadsheetApp.getUi().alert(
    'Sheet cleared. Ready for new import.'
  );
}

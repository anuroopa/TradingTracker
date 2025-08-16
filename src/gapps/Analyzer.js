function analyzeOptions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  const data = range.getValues();

  const startRow = range.getRow();
  const startCol = range.getColumn();

  // Step 1: Process ticker, price, expiry rows. Mark unnecessary rows for removal
  ({ instrument, ticker, lastPrice, expiryDateStr, expiryDateObj, openInterestRowIndex, rowsToDelete } =
    filterAndMarkRows(data, startRow));

  // Inject header info into first row
  const headerRow = [instrument, ticker, lastPrice, expiryDateStr, expiryDateObj];
  sheet.getRange(startRow, startCol, 1, headerRow.length).setValues([headerRow]);
  // Insert formula for days to expiry
  const daysToExpiry = getDaysToExpiry(expiryDateObj);
  sheet.getRange(startRow, startCol + 5).setValue(daysToExpiry);

  const tableStartRow = startRow + openInterestRowIndex;
  const tableNumRows = data.length - openInterestRowIndex;
  const numCols = data[openInterestRowIndex].length;

  // Step 3: Insert 2 cells to the left
  const leftInsertRange = sheet.getRange(tableStartRow, startCol, tableNumRows, 1);
  leftInsertRange.insertCells(SpreadsheetApp.Dimension.COLUMNS);
  leftInsertRange.insertCells(SpreadsheetApp.Dimension.COLUMNS);

  // Step 4: Insert 2 cells to the right
  const dataStartCol = startCol + 2; // after left insert
  const rightInsertStart = dataStartCol + numCols;
  const rightInsertRange = sheet.getRange(tableStartRow, rightInsertStart, tableNumRows, 1);
  rightInsertRange.insertCells(SpreadsheetApp.Dimension.COLUMNS);
  rightInsertRange.insertCells(SpreadsheetApp.Dimension.COLUMNS);

  for (let i = 0; i < tableNumRows; i++) {
    const rowIndex = tableStartRow + i;
    // Step 4: Write values
    const isHeader = i === 0;
    if (isHeader) {
      sheet.getRange(rowIndex, startCol, 1, 2).setValues([["ARR", "Sell C BE"]]);
      sheet.getRange(rowIndex, rightInsertStart, 1, 2).setValues([["Sell P BE", "ARR"]]);
      copyCellFormat(sheet, rowIndex, dataStartCol, rowIndex, startCol);
      copyCellFormat(sheet, rowIndex, dataStartCol, rowIndex, startCol + 1);
      copyCellFormat(sheet, rowIndex, dataStartCol, rowIndex, rightInsertStart);
      copyCellFormat(sheet, rowIndex, dataStartCol, rowIndex, rightInsertStart + 1);
    } else {
      const row = data[openInterestRowIndex + i];
      Logger.log(`Processing row ${rowIndex} ${row}`);
      const strike = row[7];
      const callBid = row[3];
      const putBid = row[9];
      // Logger.log(`Processing row ${rowIndex}: Strike = ${strike} Call Bid = ${callBid}, Put Bid = ${putBid}`);
      const sellCBE = isNumeric(strike) && isNumeric(callBid) ? strike - callBid : "";
      const sellPBE = isNumeric(strike) && isNumeric(putBid) ? strike - putBid : "";
      const minInvestment = Math.min(lastPrice, strike);
      setArrCell(sheet, rowIndex, startCol + 0, getAnnualizedReturn(minInvestment, callBid, daysToExpiry), dataStartCol + 3);
      setBeCell(sheet, rowIndex, startCol + 1, sellCBE, dataStartCol + 3);
      setBeCell(sheet, rowIndex, rightInsertStart + 0, sellPBE, dataStartCol + 11);
      setArrCell(sheet, rowIndex, rightInsertStart + 1, getAnnualizedReturn(strike, putBid, daysToExpiry), dataStartCol + 11);
    }
  }

  // Delete rows from bottom to top
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }

  SpreadsheetApp.flush();
}

// Helpers
function copyCellFormat(sheet, fromRow, fromCol, toRow, toCol) {
  const toCell = sheet.getRange(toRow, toCol);
  copyCellFormatToCell(sheet, fromRow, fromCol, toCell);
}

// copy format from one cell to another
function copyCellFormatToCell(sheet, fromRow, fromCol, toCell) {
  const fromCell = sheet.getRange(fromRow, fromCol);
  fromCell.copyTo(toCell, { formatOnly: true });
}

// Helper to set and format BE cell as 2 decimals
function setBeCell(sheet, row, col, value, copyCellFormat) {
  const cell = sheet.getRange(row, col);
  cell.setValue(value);
  copyCellFormatToCell(sheet, row, copyCellFormat, cell);
  cell.setNumberFormat('0.00');
  return cell;
}

// Helper to set and format ARR cell as percent
function setArrCell(sheet, row, col, value, copyCellFormat) {
  Logger.log(`Setting ARR cell at (${row}, ${col}) to value: ${value}`);
  const cell = sheet.getRange(row, col);
  cell.setValue(value);
  copyCellFormatToCell(sheet, row, copyCellFormat, cell);
  cell.setNumberFormat('0.00%');
  return cell;
}


function filterAndMarkRows(data, startRow) {
  let instrument = "";
  let ticker = "";
  let lastPrice = "";
  let expiryDateStr = "";
  let expiryDateObj = "";
  let openInterestRowIndex = -1;
  const rowsToDelete = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowText = row.join(" ").toLowerCase();
    const isInstrumentLine = i === 0;
    const isLastPriceRow = rowText.includes("last price");
    const isCallsHeader = rowText.startsWith("calls");
    const isOpenInterestHeader = rowText.startsWith("open interest");
    const hasNumeric = row.some(cell => typeof cell === 'number' && !isNaN(cell));

    if (isInstrumentLine) {
      instrument = row[0];
      ticker = extractTicker(instrument);
    }

    if (isLastPriceRow) {
      lastPrice = extractLastPrice(row);
      rowsToDelete.push(startRow + i);
      continue;
    }

    if (isCallsHeader) {
      expiryDateStr = extractExpiryDate(row);
      expiryDateObj = parseExpiryDate(expiryDateStr);
      rowsToDelete.push(startRow + i);
      continue;
    }

    if (isOpenInterestHeader) {
      openInterestRowIndex = i;
    }

    const shouldKeep = isInstrumentLine || isOpenInterestHeader || hasNumeric;

    if (!shouldKeep) {
      rowsToDelete.push(startRow + i);
    }
  }
  return { instrument, ticker, lastPrice, expiryDateStr, expiryDateObj, openInterestRowIndex, rowsToDelete };
}

function extractTicker(instrumentLine) {
  if (!instrumentLine || typeof instrumentLine !== 'string') return "";
  const colonSplit = instrumentLine.split(":")[0];
  const words = colonSplit.trim().split(" ");
  return words.length > 0 ? words[words.length - 1] : "";
}

function extractLastPrice(row) {
  for (let i = 0; i < row.length; i++) {
    const cell = row[i];
    if (typeof cell === 'string') {
      const match = cell.match(/(\d+(\.\d+)?)/);
      if (match) return parseFloat(match[0]);
    }
    if (typeof cell === 'number' && !isNaN(cell)) {
      return cell;
    }
  }
  return "";
}

function extractExpiryDate(row) {
  for (let i = 1; i < row.length; i++) {
    const cell = row[i];
    if (typeof cell === 'string' && cell.trim().length > 0) {
      return cell.trim();
    }
  }
  return "";
}

function parseExpiryDate(expiryStr) {
  if (!expiryStr) return "";
  let cleaned = expiryStr.trim();
  // Split, slice first 3 words, replace leading ' in last word with 20, join, then parse
  let parts = cleaned.split(/\s+/).slice(0, 3);
  if (parts.length === 3) {
    parts[2] = parts[2].replace(/^'/, '20');
    parsed = new Date(parts.join(' '));
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return "";
}

// Returns the number of days between expiryDate (Date object) and today (local time)
function getDaysToExpiry(expiryDate) {
  if (!(expiryDate instanceof Date) || isNaN(expiryDate.getTime())) return "";
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const exp = new Date(expiryDate.getFullYear(), expiryDate.getMonth(), expiryDate.getDate());
  const diffMs = exp - today;
  return Math.round(diffMs / (1000 * 60 * 60 * 24));
}

function getAnnualizedReturn(investment, gain, days) {
  if (!isNumeric(investment) || !isNumeric(gain) || !isNumeric(days) || investment === 0 || days <= 0) {
    return "";
  }
  const totalReturn = gain / investment;
  const annualized = Math.pow(1 + totalReturn, 365 / days) - 1;
  return annualized;
}

function isNumeric(val) {
  return typeof val === 'number' && !isNaN(val);
}
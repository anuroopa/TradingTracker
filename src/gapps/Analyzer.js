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

  const lastPriceCol = startCol + 2; // assuming lastPrice is in startcol + 2 in headerRow
  const lastPriceCell = sheet.getRange(startRow, lastPriceCol).getA1Notation();

  // Insert formula for today in startCol+5
  sheet.getRange(startRow, startCol + 5).setFormula('=TODAY()');
  const daysCol = startCol + 6;
  // Insert formula for days to expiry in startCol+6 (expiry date - today)
  sheet.getRange(startRow, daysCol).setFormula(`=${sheet.getRange(startRow, startCol + 4).getA1Notation()} - ${sheet.getRange(startRow, startCol + 5).getA1Notation()}`);
  const daysCell = sheet.getRange(startRow, daysCol).getA1Notation();

  // const daysToExpiry = getDaysToExpiry(expiryDateObj);
  // sheet.getRange(startRow, startCol + 5).setValue(daysToExpiry);

  // Step 2: Delete rows from bottom to top
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }

  const tableStartRow = startRow + openInterestRowIndex - rowsToDelete.length;
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
      // Calculate cell references for formulas
      const strikeCol = dataStartCol + 7;
      const callBidCol = dataStartCol + 3;
      const putBidCol = dataStartCol + 9;

      const strikeCell = sheet.getRange(rowIndex, strikeCol).getA1Notation();
      const callBidCell = sheet.getRange(rowIndex, callBidCol).getA1Notation();
      const putBidCell = sheet.getRange(rowIndex, putBidCol).getA1Notation();

      // ARR Call: use min(lastPrice, strike) as investment
      const arrCallFormula = `=ANNUALIZED_RETURN(MIN(${lastPriceCell},${strikeCell}),${callBidCell},${daysCell})`;
      setArrCell(sheet, rowIndex, startCol + 0, arrCallFormula, dataStartCol + 3);
      // Sell C BE
      const sellCBEFormula = `=BREAK_EVEN(${strikeCell},${callBidCell})`;
      setBeCell(sheet, rowIndex, startCol + 1, sellCBEFormula, dataStartCol + 3);
      // Sell P BE
      const sellPBEFormula = `=BREAK_EVEN(${strikeCell},${putBidCell})`;
      setBeCell(sheet, rowIndex, rightInsertStart + 0, sellPBEFormula, dataStartCol + 11);
      // ARR Put
      const arrPutFormula = `=ANNUALIZED_RETURN(${strikeCell},${putBidCell},${daysCell})`;
      setArrCell(sheet, rowIndex, rightInsertStart + 1, arrPutFormula, dataStartCol + 11);
    }
  }

  SpreadsheetApp.flush();
}

// Pure custom function for annualized return usable in Google Sheets formulas
function ANNUALIZED_RETURN(investment, gain, days) {
  if (typeof investment !== 'number' || typeof gain !== 'number' || typeof days !== 'number' || investment === 0 || days <= 0) {
    return '';
  }
  var totalReturn = gain / investment;
  return Math.pow(1 + totalReturn, 365 / days) - 1;
}

// Pure custom function for break-even (BE) usable in Google Sheets formulas
function BREAK_EVEN(strike, bid) {
  if (typeof strike !== 'number' || typeof bid !== 'number') {
    return '';
  }
  return strike - bid;
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
  if (typeof value === 'string' && value.startsWith('=')) {
    cell.setFormula(value);
  } else {
    cell.setValue(value);
  }
  copyCellFormatToCell(sheet, row, copyCellFormat, cell);
  cell.setNumberFormat('0.00');
  return cell;
}

// Helper to set and format ARR cell as percent
function setArrCell(sheet, row, col, value, copyCellFormat) {
  Logger.log(`Setting ARR cell at (${row}, ${col}) to value: ${value}`);
  const cell = sheet.getRange(row, col);
  if (typeof value === 'string' && value.startsWith('=')) {
    cell.setFormula(value);
  } else {
    cell.setValue(value);
  }
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
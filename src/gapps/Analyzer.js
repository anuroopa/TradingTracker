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
  const expiryDateCell = sheet.getRange(startRow, startCol + 4).getA1Notation();
  const formula = `=${expiryDateCell}-TODAY()`;
  sheet.getRange(startRow, startCol + 5).setFormula(formula);

  const tableStartRow = startRow + openInterestRowIndex;
  const tableNumRows = data.length - openInterestRowIndex;
  const numCols = data[0].length;

  for (let i = 0; i < tableNumRows; i++) {
    const rowIndex = tableStartRow + i;

    // Step 3: Insert 2 cells to the left
    const leftInsertRange = sheet.getRange(rowIndex, startCol, 1, 1);
    leftInsertRange.insertCells(SpreadsheetApp.Dimension.COLUMNS);
    sheet.getRange(rowIndex, startCol, 1, 1).insertCells(SpreadsheetApp.Dimension.COLUMNS);

    // Step 3: Insert 2 cells to the right
    const rightInsertStart = startCol + numCols + 2; // after left insert
    const rightInsertRange = sheet.getRange(rowIndex, rightInsertStart, 1, 1);
    rightInsertRange.insertCells(SpreadsheetApp.Dimension.COLUMNS);
    sheet.getRange(rowIndex, rightInsertStart, 1, 1).insertCells(SpreadsheetApp.Dimension.COLUMNS);

    // Step 4: Write values
    const isHeader = i === 0;
    if (isHeader) {
      sheet.getRange(rowIndex, startCol, 1, 2).setValues([["ARR", "Sell C BE"]]);
      sheet.getRange(rowIndex, rightInsertStart, 1, 2).setValues([["Sell P BE", "ARR"]]);
    } else {
      const row = sheet.getRange(rowIndex, startCol + 2, 1, numCols).getValues()[0];
      const strike = parseFloat(row[7]);     // 8th column (index 7)
      const callBid = parseFloat(row[3]);    // 4th column (index 3)
      const putBid = parseFloat(row[9]);     // 10th column (index 9)

      const sellCBE = isNumeric(strike) && isNumeric(callBid) ? strike - callBid : "";
      const sellPBE = isNumeric(strike) && isNumeric(putBid) ? strike - putBid : "";

      sheet.getRange(rowIndex, startCol + 0).setValue("");         // ARR left
      sheet.getRange(rowIndex, startCol + 1).setValue(sellCBE);    // Sell C BE
      sheet.getRange(rowIndex, rightInsertStart + 0).setValue(sellPBE); // Sell P BE
      sheet.getRange(rowIndex, rightInsertStart + 1).setValue("");     // ARR right
    }
  }

  // Delete rows from bottom to top
  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    sheet.deleteRow(rowsToDelete[i]);
  }

  SpreadsheetApp.flush();
}

// Helpers
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

function isNumeric(val) {
  return typeof val === 'number' && !isNaN(val);
}
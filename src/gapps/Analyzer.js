function analyzeOptions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  const data = range.getValues();

  const startRow = range.getRow();
  const startCol = range.getColumn();
  const numCols = data[0].length;

  let lastPrice = "";
  let expiryDateStr = "";
  let expiryDateObj = "";
  let openInterestRowIndex = -1;
  const rowsToDelete = [];

  // Step 1: Process ticker, price, expiry rows. Mark unnecessary rows for removal
  ({ lastPrice, expiryDateStr, expiryDateObj, openInterestRowIndex } = filterAndMarkRows(data, lastPrice, rowsToDelete, startRow, expiryDateStr, expiryDateObj, openInterestRowIndex));

  if (openInterestRowIndex === -1) {
    throw new Error("Open Interest row not found in selected range.");
  }

  const tableStartRow = startRow + openInterestRowIndex;
  const tableNumRows = data.length - openInterestRowIndex;

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
function filterAndMarkRows(data, lastPrice, rowsToDelete, startRow, expiryDateStr, expiryDateObj, openInterestRowIndex) {
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowText = row.join(" ").toLowerCase();
    const isInstrumentLine = i === 0;
    const isLastPriceRow = rowText.includes("last price");
    const isCallsHeader = rowText.startsWith("calls");
    const isOpenInterestHeader = rowText.startsWith("open interest");
    const hasNumeric = row.some(cell => typeof cell === 'number' && !isNaN(cell));

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
  return { lastPrice, expiryDateStr, expiryDateObj, openInterestRowIndex };
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
  const cleaned = expiryStr.replace(/'/g, "");
  // Parse as local time to avoid timezone issues
  const parsed = new Date(cleaned + 'T00:00:00');
  return isNaN(parsed.getTime()) ? "" : parsed;
}

function isNumeric(val) {
  return typeof val === 'number' && !isNaN(val);
}
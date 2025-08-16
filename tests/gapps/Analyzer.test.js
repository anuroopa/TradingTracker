const fs = require('fs');
const path = require('path');
const code = fs.readFileSync(path.join(__dirname, '../../src/gapps/Analyzer.js'), 'utf8');
eval(code); // Functions available globally

describe('extractLastPrice', () => {
  test.each([
    [['foo', 'bar 123.45', 'baz'], 123.45],
    [['foo', 99.99, 'baz'], 99.99],
    [['foo', 'bar', 'baz'], ""],
    [['foo', NaN, 'bar 42'], 42],
    [['foo', '100', 'baz'], 100],
  ])('extractLastPrice(%j) -> %j', (input, expected) => {
    expect(() => extractLastPrice(input)).not.toThrow();
    expect(extractLastPrice(input)).toBe(expected);
  });
});

describe('extractTicker', () => {
  test.each([
    ['AAPL', 'AAPL'],
    ['Stock Option MSFT', 'MSFT'],
    ['GOOG:US Equity', 'GOOG'],
    ['Option Chain TSLA:NASDAQ', 'TSLA'],
    ['', ''],
    [null, ''],
    [undefined, ''],
    [123, ''],
  ])('extractTicker(%j) -> %j', (input, expected) => {
    expect(extractTicker(input)).toBe(expected);
  });
});

describe('extractExpiryDate', () => {
  test.each([
    [['', '2025-08-15', 'foo'], '2025-08-15'],
    [['', '  2025-08-15  ', 'foo'], '2025-08-15'],
    [['', '', '2025-12-31'], '2025-12-31'],
    [['', '', ''], ''],
    [['foo', '', ''], ''],
  ])('extractExpiryDate(%j) -> %j', (input, expected) => {
    expect(extractExpiryDate(input)).toBe(expected);
  });
});

describe('parseExpiryDate', () => {
  test.each([
    ["Aug 15 '25", 2025, 8, 15],
    ["Sep 19 '25", 2025, 9, 19],
    ["Sep 26 '25 w (Weekly)", 2025, 9, 26],
    ['not-a-date', '', '', ''],
    ['', '', '', ''],
    [null, '', '', ''],
  ])('parseExpiryDate(%j) -> %j-%j-%j', (input, year, month, date) => {
    const result = parseExpiryDate(input);
    if (year !== '') {
      console.log('parseExpiryDate result type:', typeof result, 'value:', result);
      expect(result instanceof Date).toBe(true);
      expect(result.getFullYear()).toBe(year);
      expect(result.getMonth()).toBe(month-1); // 0-based
      expect(result.getDate()).toBe(date);
    } else {
      expect(result).toBe('');
    }
  });
});

describe('getDaysToExpiry', () => {
  const today = new Date();
  today.setHours(0,0,0,0);
  test.each([
    [new Date(today.getFullYear(), today.getMonth(), today.getDate() + 10), 10],
    [today, 0],
    [new Date(today.getFullYear(), today.getMonth(), today.getDate() - 5), -5],
    ['not-a-date', ""],
    [null, ""],
    [new Date('invalid'), ""]
  ])('getDaysToExpiry(%j) -> %j', (input, expected) => {
    expect(getDaysToExpiry(input)).toBe(expected);
  });
});

describe('filterAndMarkRows', () => {
  test('basic extraction and marking', () => {
    const data = [
      ["BITFARMS LTD COM BITF: NSDQ"],
      ["Real Time Equity Quote: August 15, 2025 01:15:49 PM ET"],
      ["Last Price1.28"],
      ["Today's Change+0.01 (+1.18%)"],
      ["Volume12,705,474"],
      ["Bid1.28(66,000)"],
      ["Ask1.29(134,600)"],
      ["Aug 15 '25"],
      ["CALLS", "Aug 15 '25"],
      ["Open Interest"],
      ["100", 1, 2],
      ["remove me"],
    ];
    const startRow = 5;
    const result = filterAndMarkRows(data, startRow);
    expect(result.instrument).toBe("BITFARMS LTD COM BITF: NSDQ");
    expect(result.ticker).toBe("BITF");
    expect(result.lastPrice).toBe(1.28);
    expect(result.expiryDateStr).toBe("Aug 15 '25");
    expect(result.expiryDateObj instanceof Date).toBe(true);
    expect(result.expiryDateObj.getFullYear()).toBe(2025);
    expect(result.expiryDateObj.getMonth()).toBe(7); // August
    expect(result.expiryDateObj.getDate()).toBe(15);
    expect(result.openInterestRowIndex).toBe(9);
    // rowsToDelete: Last Price, CALLS, and all non-header, non-numeric rows
    expect(result.rowsToDelete).toEqual([6, 7, 8, 9, 10, 11, 12, 13, 16]);
  });

  test('no last price or expiry', () => {
    const data = [
      ['AAPL'],
      ['Open Interest'],
      ['100', 1, 2],
    ];
    const startRow = 0;
    const result = filterAndMarkRows(data, startRow);
    expect(result.lastPrice).toBe("");
    expect(result.expiryDateStr).toBe("");
    expect(result.expiryDateObj).toBe("");
    expect(result.openInterestRowIndex).toBe(1);
    expect(result.rowsToDelete).toEqual([]);
  });

  test('removes non-numeric, non-header rows', () => {
    const data = [
      ['AAPL'],
      ['foo'],
      ['Open Interest'],
      ['bar'],
      ['baz', ''],
      ['100', 1, 2],
    ];
    const startRow = 10;
    const result = filterAndMarkRows(data, startRow);
    // Only rows 1, 3, 4 should be deleted (non-header, non-numeric)
    expect(result.rowsToDelete).toEqual([11, 13, 14]);
  });
});

describe('getAnnualizedReturn', () => {
  test.each([
    // investment, gain, days, expected (rounded to 6 decimals)
    [1000, 100, 365, 0.1],
    [1000, 200, 365, 0.2],
    [1000, 100, 30, Math.pow(1 + 0.1, 365 / 30) - 1],
    [500, 50, 180, Math.pow(1 + 0.1, 365 / 180) - 1],
    [1000, 0, 365, 0],
    [1000, -100, 365, -0.1],
    [1000, 100, 0, ""],
    [0, 100, 365, ""],
    [1000, 100, -10, ""],
    [null, 100, 365, ""],
    [1000, null, 365, ""],
    [1000, 100, null, ""],
    [NaN, 100, 365, ""],
    [1000, NaN, 365, ""],
    [1000, 100, NaN, ""],
    ["1000", 100, 365, ""],
    [1000, "100", 365, ""],
    [1000, 100, "365", ""],
  ])('annualizedReturn(%j, %j, %j) -> %j', (investment, gain, days, expected) => {
    const result = getAnnualizedReturn(investment, gain, days);
    if (expected === "") {
      expect(result).toBe("");
    } else {
      expect(typeof result).toBe("number");
      expect(result).toBeCloseTo(expected, 6);
    }
  });
});
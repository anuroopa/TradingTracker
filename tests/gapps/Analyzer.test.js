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
    ['2025-08-15', 2025, 8, 15],
    ["'2025-08-15'", 2025, 8, 15],
    ['not-a-date', '', '', ''],
    ['', '', '', ''],
    [null, '', '', ''],
  ])('parseExpiryDate(%j) -> %j-%j-%j', (input, year, month, date) => {
    const result = parseExpiryDate(input);
    if (year !== '') {
      expect(result instanceof Date).toBe(true);
      expect(result.getFullYear()).toBe(year);
      expect(result.getMonth()).toBe(month-1); // 0-based
      expect(result.getDate()).toBe(date);
    } else {
      expect(result).toBe('');
    }
  });
});

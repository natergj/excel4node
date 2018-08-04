export default {
  generateRId,
  getExcelAlpha,
  getExcelCellRef,
  getExcelRowCol,
  getExcelTS,
};

/**
 * Generates an OOXML style Resource ID
 * @function generateRId
 * @returns {String} Resource ID
 */
function generateRId(): string {
  const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let text = 'R';
  for (let i = 0; i < 16; i++) {
    text += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return text;
}

/**
 * Translates a column number into the Alpha equivalent used by Excel
 * @function getExcelAlpha
 * @param {Number} colNum Column number that is to be translated
 * @returns {String} The Excel alpha representation of the column number
 * @example
 * // returns B
 * getExcelAlpha(2);
 */
function getExcelAlpha(colNum: number): string {
  const aCharCode = 65;
  let remaining = colNum;
  let columnName = '';

  while (remaining > 0) {
    const mod = (remaining - 1) % 26;
    columnName = String.fromCharCode(aCharCode + mod) + columnName;
    remaining = (remaining - 1 - mod) / 26;
  }

  return columnName;
}

/**
 * Translates a column number into the Alpha equivalent used by Excel
 * @function getExcelAlpha
 * @param {Number} rowNum Row number that is to be transalated
 * @param {Number} colNum Column number that is to be transalated
 * @returns {String} The Excel alpha representation of the column number
 * @example
 * // returns B1
 * getExcelCellRef(1, 2);
 */
function getExcelCellRef(rowNum: number, colNum: number) {
  const aCharCode = 65;
  let remaining = colNum;
  let columnName = '';

  while (remaining > 0) {
    const mod = (remaining - 1) % 26;
    columnName = String.fromCharCode(aCharCode + mod) + columnName;
    remaining = (remaining - 1 - mod) / 26;
  }
  return columnName + rowNum;
}

interface IExcelRowColObject {
  row: number;
  col: number;
}
/**
 * Translates a Excel cell representation into row and column numerical equivalents
 * @function getExcelRowCol
 * @param {String} str Excel cell representation
 * @returns {Object} Object keyed with row and col
 * @example
 * // returns {row: 2, col: 3}
 * getExcelRowCol('C2')
 */
function getExcelRowCol(str: string): IExcelRowColObject {
  const numeric = str.split(/\D/).filter(el => el !== '')[0];
  const alpha = str.split(/\d/).filter(el => el !== '')[0];
  const row = parseInt(numeric, 10);
  const col = alpha
    .toUpperCase()
    .split('')
    .reduce((a, b, index, arr) => a + (b.charCodeAt(0) - 64) * Math.pow(26, arr.length - index - 1), 0);
  return {row, col};
}

/**
 * Translates a date into Excel timestamp
 * @function getExcelTS
 * @param {Date} date Date to translate
 * @returns {Number} Excel timestamp
 * @example
 * // returns 29810.958333333332
 * getExcelTS(new Date('08/13/1981'));
 */
function getExcelTS(date: Date | string): number {
  let thisDt;
  if (date instanceof Date) {
    thisDt = date;
  } else {
    thisDt = new Date(date);
  }
  thisDt.setDate(thisDt.getDate() + 1);

  const epoch = new Date('1900-01-01T00:00:00.0000Z');

  // Handle legacy leap year offset as described in  §18.17.4.1
  const legacyLeapDate = new Date('1900-02-28T23:59:59.999Z');
  if (thisDt.getTime() - legacyLeapDate.getTime() > 0) {
    thisDt.setDate(thisDt.getDate() + 1);
  }

  // Get milliseconds between date sent to function and epoch
  const diff2 = thisDt.getTime() - epoch.getTime();

  const ts = diff2 / (1000 * 60 * 60 * 24);

  return parseFloat(ts.toFixed(7));
}

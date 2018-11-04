/**
 * Generates an OOXML style Resource ID
 * @function generateRId
 * @returns {String} Resource ID
 */
export function generateRId(): string {
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
export function getExcelAlpha(colNum: number): string {
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
export function getExcelCellRef(rowNum: number, colNum: number) {
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
export function getExcelRowCol(str: string): IExcelRowColObject {
  const numeric = str.split(/\D/).filter(el => el !== '')[0];
  const alpha = str.split(/\d/).filter(el => el !== '')[0];
  const row = parseInt(numeric, 10);
  const col = alpha
    .toUpperCase()
    .split('')
    .reduce((a, b, index, arr) => a + (b.charCodeAt(0) - 64) * Math.pow(26, arr.length - index - 1), 0);
  return { row, col };
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
export function getExcelTS(date: Date | string): number {
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

export function boolToInt(bool: boolean | 1 | '1' | 0 | '0') {
  if (bool === true) {
    return 1;
  }
  if (bool === false) {
    return 0;
  }
  if (parseInt(String(bool), 10) === 1) {
    return 1;
  }
  if (parseInt(String(bool), 10) === 0) {
    return 0;
  }
  throw new TypeError('Value sent to boolToInt must be true, false, 1 or 0');
}

export function sortCellRefs(a: string, b: string) {
  const aAtt = getExcelRowCol(a);
  const bAtt = getExcelRowCol(b);
  if (aAtt.col === bAtt.col) {
    return aAtt.row - bAtt.row;
  } else {
    return aAtt.col - bAtt.col;
  }
}

function _bitXOR(a, b) {
  const maxLength = a.length > b.length ? a.length : b.length;

  let padString = '';
  for (let i = 0; i < maxLength; i++) {
    padString += '0';
  }

  a = String(padString + a).substr(-maxLength);
  b = String(padString + b).substr(-maxLength);

  let response = '';
  for (let i = 0; i < a.length; i++) {
    response += a[i] === b[i] ? 0 : 1;
  }
  return response;
}

function _rotateBinary(bin) {
  return bin.substr(1, bin.length - 1) + bin.substr(0, 1);
}

function _getHashForChar(char, hash) {
  hash = hash ? hash : '0000';
  const charCode = char.charCodeAt(0);
  let hashBin = parseInt(hash, 16).toString(2);
  let charBin = parseInt(charCode, 10).toString(2);
  hashBin = String('000000000000000' + hashBin).substr(-15);
  charBin = String('000000000000000' + charBin).substr(-15);
  let nextHash = _bitXOR(hashBin, charBin);
  nextHash = _rotateBinary(nextHash);
  nextHash = parseInt(nextHash, 2).toString(16);

  return nextHash;
}

//  http://www.openoffice.org/sc/excelfileformat.pdf section 4.18.4
export function getHashOfPassword(str) {
  let curHash = '0000';
  for (let i = str.length - 1; i >= 0; i--) {
    curHash = _getHashForChar(str[i], curHash);
  }
  const curHashBin = parseInt(curHash, 16).toString(2);
  const charCountBin = parseInt(str.length, 10).toString(2);
  const saltBin = parseInt('CE4B', 16).toString(2);

  const firstXOR = _bitXOR(curHashBin, charCountBin);
  const finalHashBin = _bitXOR(firstXOR, saltBin);
  const finalHash = String(
    '0000' +
      parseInt(finalHashBin, 2)
        .toString(16)
        .toUpperCase(),
  ).slice(-4);

  return finalHash;
}

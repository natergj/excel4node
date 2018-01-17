function bitXOR(a: string, b: string) {
  const maxLength = a.length > b.length ? a.length : b.length;

  let padString = '';
  for (let i = 0; i < maxLength; i += 1) {
    padString += '0';
  }

  const aStr = String(padString + a).substr(-maxLength);
  const bStr = String(padString + b).substr(-maxLength);

  let response = '';
  for (let i = 0; i < aStr.length; i += 1) {
    response += aStr[i] === bStr[i] ? 0 : 1;
  }
  return response;
}

function rotateBinary(bin: string) {
  return bin.substr(1, bin.length - 1) + bin.substr(0, 1);
}

function getHashForChar(char: string, hash: string = '0000') {
  const charCode = char.charCodeAt(0);
  let hashBin = parseInt(hash, 16).toString(2);
  let charBin = charCode.toString(2);
  hashBin = String('000000000000000' + hashBin).substr(-15);
  charBin = String('000000000000000' + charBin).substr(-15);
  let nextHash = bitXOR(hashBin, charBin);
  nextHash = rotateBinary(nextHash);
  nextHash = parseInt(nextHash, 2).toString(16);

  return nextHash;
}

export function generateRId() {
  let text = 'R';
  const possible =
    'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (let i = 0; i < 16; i += 1) {
    text += possible.charAt(Math.floor(Math.random() * possible.length));
  }
  return text;
}

//  http://www.openoffice.org/sc/excelfileformat.pdf section 4.18.4
export function getHashOfPassword(str: string) {
  let curHash = '0000';
  for (let i = str.length - 1; i >= 0; i -= 1) {
    curHash = getHashForChar(str[i], curHash);
  }
  const curHashBin = parseInt(curHash, 16).toString(2);
  const charCountBin = str.length.toString(2);
  const saltBin = parseInt('CE4B', 16).toString(2);

  const firstXOR = bitXOR(curHashBin, charCountBin);
  const finalHashBin = bitXOR(firstXOR, saltBin);
  const finalHash = String(
    '0000' +
      parseInt(finalHashBin, 2)
        .toString(16)
        .toUpperCase(),
  ).slice(-4);

  return finalHash;
}

/**
 * Translates a column number into the Alpha equivalent used by Excel
 * @function getExcelAlpha
 * @param {Number} colNum Column number that is to be transalated
 * @returns {String} The Excel alpha representation of the column number
 * @example
 * // returns B
 * getExcelAlpha(2);
 */
export function getExcelAlpha(colNum: number) {
  let remaining = colNum;
  const aCharCode = 65;
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
 * @param {Number} rowNum Row number that is to be translated
 * @param {Number} colNum Column number that is to be translated
 * @returns {String} The Excel alpha representation of the column number
 * @example
 * // returns B1
 * getExcelCellRef(1, 2);
 */
export function getExcelCellRef(rowNum: number, colNum: number) {
  let remaining = colNum;
  const aCharCode = 65;
  let columnName = '';
  while (remaining > 0) {
    const mod = (remaining - 1) % 26;
    columnName = String.fromCharCode(aCharCode + mod) + columnName;
    remaining = (remaining - 1 - mod) / 26;
  }
  return columnName + rowNum;
}

/**
 * Translates a Excel cell represenation into row and column numerical equivalents
 * @function getExcelRowCol
 * @param {String} str Excel cell representation
 * @returns {Object} Object keyed with row and col
 * @example
 * // returns {row: 2, col: 3}
 * getExcelRowCol('C2')
 */
export function getExcelRowCol(str: string) {
  const numeric = str.split(/\D/).filter((el) => {
    return el !== '';
  })[0];
  const alpha = str.split(/\d/).filter((el) => {
    return el !== '';
  })[0];
  const row = parseInt(numeric, 10);
  const alphaReducer = (a, b, index, arr) => {
    return a + (b.charCodeAt(0) - 64) * Math.pow(26, arr.length - index - 1);
  };
  const col = alpha
    .toUpperCase()
    .split('')
    .reduce(alphaReducer, 0);
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
export function getExcelTS(date: Date | string) {
  const thisDt = new Date(date.toString());
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

export function sortCellRefs(a: string, b: string) {
  const aAtt = getExcelRowCol(a);
  const bAtt = getExcelRowCol(b);
  if (aAtt.col === bAtt.col) {
    return aAtt.row - bAtt.row;
  }
  return aAtt.col - bAtt.col;
}

export function arrayIntersectSafe(a: any[], b: any[]) {
  if (a instanceof Array && b instanceof Array) {
    let ai = 0;
    let bi = 0;
    const result = [];

    while (ai < a.length && bi < b.length) {
      if (a[ai] < b[bi]) {
        ai += 1;
      } else if (a[ai] > b[bi]) {
        bi += 1;
      } else {
        result.push(a[ai]);
        ai += 1;
        bi += 1;
      }
    }
    return result;
  }
  throw new TypeError(
    'Both variables sent to arrayIntersectSafe must be arrays',
  );
}

export function getAllCellsInExcelRange(range: string) {
  const cells = range.split(':');
  const cell1props = getExcelRowCol(cells[0]);
  const cell2props = getExcelRowCol(cells[1]);
  return getAllCellsInNumericRange(
    cell1props.row,
    cell1props.col,
    cell2props.row,
    cell2props.col,
  );
}

export function getAllCellsInNumericRange(
  row1: number,
  col1: number,
  row2: number = row1,
  col2: number = col1,
) {
  const response = [];
  for (let i = row1; i <= row2; i += 1) {
    for (let j = col1; j <= col2; j += 1) {
      response.push(getExcelAlpha(j) + i);
    }
  }
  return response.sort(sortCellRefs);
}

export function boolToInt(bool: boolean | 0 | 1) {
  if (bool === true) {
    return 1;
  }
  if (bool === false) {
    return 0;
  }
  if (bool === 1) {
    return 1;
  }
  if (bool === 0) {
    return 0;
  }
  throw new TypeError('Value sent to boolToInt must be true, false, 1 or 0');
}

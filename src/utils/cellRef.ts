/**
 * Translates a column number into the Alpha equivalent used by Excel
 * @function getExcelAlpha
 * @param {number} colNum Column number that is to be translated
 * @returns {string} The Excel alpha representation of the column number
 * @example
 * // returns B
 * getExcelAlpha(2);
 */
export const getExcelAlpha = (colNum: number): string => {
  const aCharCode = 65;
  let remaining = colNum;
  let columnName = '';

  while (remaining > 0) {
    const mod = (remaining - 1) % 26;
    columnName = `${String.fromCharCode(aCharCode + mod)}${columnName}`;
    remaining = (remaining - 1 - mod) / 26;
  }

  return columnName;
};

/**
 * Translates a column number into the Alpha equivalent used by Excel
 * @function getExcelCellRef
 * @param {number} rowNum Row number that is to be translated
 * @param {number} colNum Column number that is to be translated
 * @returns {string} The Excel alpha representation of the column number
 * @example
 * // returns B1
 * getExcelCellRef(1, 2);
 */
export const getExcelCellRef = (rowNum: number, colNum: number): string => {
  const aCharCode = 65;
  let remaining = colNum;
  let columnName = '';

  while (remaining > 0) {
    const mod = (remaining - 1) % 26;
    columnName = `${String.fromCharCode(aCharCode + mod)}${columnName}`;
    remaining = (remaining - 1 - mod) / 26;
  }
  return `${columnName}${rowNum}`;
};

type ExcelRowColObject = {
  row: number;
  col: number;
};
/**
 * Translates a Excel cell representation into row and column numerical equivalents
 * @function getExcelRowCol
 * @param {string} str Excel cell representation
 * @returns {Object} Object keyed with row and col
 * @example
 * // returns {row: 2, col: 3}
 * getExcelRowCol('C2')
 */
export const getExcelRowCol = (str: string): ExcelRowColObject => {
  const numeric = str.split(/\D/).filter(el => el !== '')[0];
  const alpha = str.split(/\d/).filter(el => el !== '')[0];
  const row = parseInt(numeric, 10);
  const col = alpha
    .toUpperCase()
    .split('')
    .reduce((a, b, index, arr) => a + (b.charCodeAt(0) - 64) * Math.pow(26, arr.length - index - 1), 0);
  return { row, col };
};

/**
 * Sorter function for an array of Excel reference strings (i.e. ["A1", "BA2", "AAA2"])
 * @param {string} a first cell reference
 * @param {string} b second cell reference
 * @returns {number} the returned sort order
 * @example
 * // Return ["A1", "BA2", "AAA2"]
 * ["AAA2", "A1", "BA2"].sort(sortCellRefs);
 */
export const sortCellRefs = (a: string, b: string) => {
  const aAtt = getExcelRowCol(a);
  const bAtt = getExcelRowCol(b);
  if (aAtt.col === bAtt.col) {
    return aAtt.row - bAtt.row;
  } else {
    return aAtt.col - bAtt.col;
  }
};

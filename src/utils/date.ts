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

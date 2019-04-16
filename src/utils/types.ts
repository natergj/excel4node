/**
 * converts 1 or 0 as expected by the final Excel XML files based on input
 * @param {string | 1 | 0 | "1" | "0"} bool
 * @returns {1 | 0}
 */
export const boolToInt = (bool: boolean | 1 | '1' | 0 | '0') => {
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
};

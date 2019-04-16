const bitXOR = (a, b) => {
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
};

const rotateBinary = bin => {
  return bin.substr(1, bin.length - 1) + bin.substr(0, 1);
};

const getHashForChar = (char, hash) => {
  hash = hash ? hash : '0000';
  const charCode = char.charCodeAt(0);
  let hashBin = parseInt(hash, 16).toString(2);
  let charBin = parseInt(charCode, 10).toString(2);
  hashBin = String('000000000000000' + hashBin).substr(-15);
  charBin = String('000000000000000' + charBin).substr(-15);
  let nextHash = bitXOR(hashBin, charBin);
  nextHash = rotateBinary(nextHash);
  nextHash = parseInt(nextHash, 2).toString(16);

  return nextHash;
};

//  http://www.openoffice.org/sc/excelfileformat.pdf section 4.18.4
/**
 * Create an Excel password hash to be used for workbook editing protection
 * @param {string} str password string
 * @return {string} the hash of the password string
 */
export const getHashOfPassword = str => {
  let curHash = '0000';
  for (let i = str.length - 1; i >= 0; i--) {
    curHash = getHashForChar(str[i], curHash);
  }
  const curHashBin = parseInt(curHash, 16).toString(2);
  const charCountBin = parseInt(str.length, 10).toString(2);
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
};

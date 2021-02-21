let types = require('./types/index.js');

let _bitXOR = (a, b) => {
    let maxLength = a.length > b.length ? a.length : b.length;

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

let generateRId = () => {
    let text = 'R';
    let possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    for (let i = 0; i < 16; i++) {
        text += possible.charAt(Math.floor(Math.random() * possible.length));
    }
    return text;
};

let _rotateBinary = (bin) => {
    return bin.substr(1, bin.length - 1) + bin.substr(0, 1);
};

let _getHashForChar = (char, hash) => {    
    hash = hash ? hash : '0000';
    let charCode = char.charCodeAt(0);
    let hashBin = parseInt(hash, 16).toString(2);
    let charBin = parseInt(charCode, 10).toString(2);
    hashBin = String('000000000000000' + hashBin).substr(-15);
    charBin = String('000000000000000' + charBin).substr(-15);
    let nextHash = _bitXOR(hashBin, charBin);
    nextHash = _rotateBinary(nextHash);
    nextHash = parseInt(nextHash, 2).toString(16);

    return nextHash;
};

//  http://www.openoffice.org/sc/excelfileformat.pdf section 4.18.4
let getHashOfPassword = (str) => {
    let curHash = '0000';
    for (let i = str.length - 1; i >= 0; i--) {
        curHash = _getHashForChar(str[i], curHash);
    }
    let curHashBin = parseInt(curHash, 16).toString(2);
    let charCountBin = parseInt(str.length, 10).toString(2);
    let saltBin = parseInt('CE4B', 16).toString(2);

    let firstXOR = _bitXOR(curHashBin, charCountBin);
    let finalHashBin = _bitXOR(firstXOR, saltBin);
    let finalHash = String('0000' + parseInt(finalHashBin, 2).toString(16).toUpperCase()).slice(-4);

    return finalHash;
};

/**
 * Translates a column number into the Alpha equivalent used by Excel
 * @function getExcelAlpha
 * @param {Number} colNum Column number that is to be transalated
 * @returns {String} The Excel alpha representation of the column number
 * @example
 * // returns B
 * getExcelAlpha(2);
 */
let getExcelAlpha = (colNum) => {
    let remaining = colNum;
    let aCharCode = 65;
    let columnName = '';
    while (remaining > 0) {
        let mod = (remaining - 1) % 26;
        columnName = String.fromCharCode(aCharCode + mod) + columnName;
        remaining = (remaining - 1 - mod) / 26;
    } 
    return columnName;
};

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
let getExcelCellRef = (rowNum, colNum) => {
    let remaining = colNum;
    let aCharCode = 65;
    let columnName = '';
    while (remaining > 0) {
        let mod = (remaining - 1) % 26;
        columnName = String.fromCharCode(aCharCode + mod) + columnName;
        remaining = (remaining - 1 - mod) / 26;
    } 
    return columnName + rowNum;
};

/**
 * Translates a Excel cell represenation into row and column numerical equivalents 
 * @function getExcelRowCol
 * @param {String} str Excel cell representation
 * @returns {Object} Object keyed with row and col
 * @example
 * // returns {row: 2, col: 3}
 * getExcelRowCol('C2')
 */
let getExcelRowCol = (str) => {
    let numeric = str.split(/\D/).filter(function (el) {
        return el !== '';
    })[0];
    let alpha = str.split(/\d/).filter(function (el) {
        return el !== '';
    })[0];
    let row = parseInt(numeric, 10);
    let col = alpha.toUpperCase().split('').reduce(function (a, b, index, arr) {
        return a + (b.charCodeAt(0) - 64) * Math.pow(26, arr.length - index - 1);
    }, 0);
    return { row: row, col: col };
};

/**
 * Translates a date into Excel timestamp
 * @function getExcelTS
 * @param {Date} date Date to translate
 * @returns {Number} Excel timestamp
 * @example
 * // returns 29810.958333333332
 * getExcelTS(new Date('08/13/1981'));
 */
let getExcelTS = (date) => {

    let thisDt = new Date(date);
    thisDt = new Date(thisDt.getTime() + 24 * 60 * 60 * 1000);

    let epoch = new Date('1900-01-01T00:00:00.0000Z');

    // Handle legacy leap year offset as described in  §18.17.4.1
    const legacyLeapDate = new Date('1900-02-28T23:59:59.999Z');
    if (thisDt - legacyLeapDate > 0) {
        thisDt = new Date(thisDt.getTime() + 24 * 60 * 60 * 1000);
    } 

    // Get milliseconds between date sent to function and epoch 
    let diff2 = thisDt.getTime() - epoch.getTime();

    let ts = diff2 / (1000 * 60 * 60 * 24);

    return parseFloat(ts.toFixed(7));
};

let sortCellRefs = (a, b) => {
    let aAtt = getExcelRowCol(a);
    let bAtt = getExcelRowCol(b);
    if (aAtt.col === bAtt.col) {
        return aAtt.row - bAtt.row;
    } else {
        return aAtt.col - bAtt.col;
    }
};

let arrayIntersectSafe = (a, b) => {

    if (a instanceof Array && b instanceof Array) {
        var ai = 0, bi = 0;
        var result = new Array();

        while (ai < a.length && bi < b.length) {
            if (a[ai] < b[bi]) {
                ai++;
            } else if (a[ai] > b[bi]) {
                bi++;
            } else {
                result.push(a[ai]);
                ai++;
                bi++;
            }
        }
        return result;
    } else {
        throw new TypeError('Both variables sent to arrayIntersectSafe must be arrays');
    }
};

let getAllCellsInExcelRange = (range) => {
    var cells = range.split(':');
    var cell1props = getExcelRowCol(cells[0]);
    var cell2props = getExcelRowCol(cells[1]);
    return getAllCellsInNumericRange(cell1props.row, cell1props.col, cell2props.row, cell2props.col);
};

let getAllCellsInNumericRange = (row1, col1, row2, col2) => {
    var response = [];
    row2 = row2 ? row2 : row1;
    col2 = col2 ? col2 : col1;
    for (var i = row1; i <= row2; i++) {
        for (var j = col1; j <= col2; j++) {
            response.push(getExcelAlpha(j) + i);
        }
    }
    return response.sort(sortCellRefs);
};

let boolToInt = (bool) => {
    if (bool === true) {
        return 1;
    }
    if (bool === false) {
        return 0;
    }
    if (parseInt(bool) === 1) {
        return 1;
    }
    if (parseInt(bool) === 0) {
        return 0;
    }
    throw new TypeError('Value sent to boolToInt must be true, false, 1 or 0');
};

/*
 * Helper Functions
 */

module.exports = {
    generateRId,
    getHashOfPassword,
    getExcelAlpha,
    getExcelCellRef,
    getExcelRowCol,
    getExcelTS,
    sortCellRefs,
    arrayIntersectSafe,
    getAllCellsInExcelRange,
    getAllCellsInNumericRange,
    boolToInt
};
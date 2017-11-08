'use strict';

var types = require('./types/index.js');

var _bitXOR = function _bitXOR(a, b) {
    var maxLength = a.length > b.length ? a.length : b.length;

    var padString = '';
    for (var i = 0; i < maxLength; i++) {
        padString += '0';
    }

    a = String(padString + a).substr(-maxLength);
    b = String(padString + b).substr(-maxLength);

    var response = '';
    for (var _i = 0; _i < a.length; _i++) {
        response += a[_i] === b[_i] ? 0 : 1;
    }
    return response;
};

var generateRId = function generateRId() {
    var text = 'R';
    var possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    for (var i = 0; i < 16; i++) {
        text += possible.charAt(Math.floor(Math.random() * possible.length));
    }
    return text;
};

var _rotateBinary = function _rotateBinary(bin) {
    return bin.substr(1, bin.length - 1) + bin.substr(0, 1);
};

var _getHashForChar = function _getHashForChar(char, hash) {
    hash = hash ? hash : '0000';
    var charCode = char.charCodeAt(0);
    var hashBin = parseInt(hash, 16).toString(2);
    var charBin = parseInt(charCode, 10).toString(2);
    hashBin = String('000000000000000' + hashBin).substr(-15);
    charBin = String('000000000000000' + charBin).substr(-15);
    var nextHash = _bitXOR(hashBin, charBin);
    nextHash = _rotateBinary(nextHash);
    nextHash = parseInt(nextHash, 2).toString(16);

    return nextHash;
};

//  http://www.openoffice.org/sc/excelfileformat.pdf section 4.18.4
var getHashOfPassword = function getHashOfPassword(str) {
    var curHash = '0000';
    for (var i = str.length - 1; i >= 0; i--) {
        curHash = _getHashForChar(str[i], curHash);
    }
    var curHashBin = parseInt(curHash, 16).toString(2);
    var charCountBin = parseInt(str.length, 10).toString(2);
    var saltBin = parseInt('CE4B', 16).toString(2);

    var firstXOR = _bitXOR(curHashBin, charCountBin);
    var finalHashBin = _bitXOR(firstXOR, saltBin);
    var finalHash = String('0000' + parseInt(finalHashBin, 2).toString(16).toUpperCase()).slice(-4);

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
var getExcelAlpha = function getExcelAlpha(colNum) {
    var remaining = colNum;
    var aCharCode = 65;
    var columnName = '';
    while (remaining > 0) {
        var mod = (remaining - 1) % 26;
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
var getExcelCellRef = function getExcelCellRef(rowNum, colNum) {
    var remaining = colNum;
    var aCharCode = 65;
    var columnName = '';
    while (remaining > 0) {
        var mod = (remaining - 1) % 26;
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
var getExcelRowCol = function getExcelRowCol(str) {
    var numeric = str.split(/\D/).filter(function (el) {
        return el !== '';
    })[0];
    var alpha = str.split(/\d/).filter(function (el) {
        return el !== '';
    })[0];
    var row = parseInt(numeric, 10);
    var col = alpha.toUpperCase().split('').reduce(function (a, b, index, arr) {
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
var getExcelTS = function getExcelTS(date) {

    var thisDt = new Date(date);
    thisDt.setDate(thisDt.getDate() + 1);
    // Take timezone into account when calculating date
    thisDt.setMinutes(thisDt.getMinutes() - thisDt.getTimezoneOffset());

    var epoch = new Date(1899, 11, 31);
    // Take timezone into account when calculating epoch
    epoch.setMinutes(epoch.getMinutes() - epoch.getTimezoneOffset());

    // Get milliseconds between date sent to function and epoch 
    var diff2 = thisDt.getTime() - epoch.getTime();

    var ts = diff2 / (1000 * 60 * 60 * 24);
    return ts;
};

var sortCellRefs = function sortCellRefs(a, b) {
    var aAtt = getExcelRowCol(a);
    var bAtt = getExcelRowCol(b);
    if (aAtt.col === bAtt.col) {
        return aAtt.row - bAtt.row;
    } else {
        return aAtt.col - bAtt.col;
    }
};

var arrayIntersectSafe = function arrayIntersectSafe(a, b) {

    if (a instanceof Array && b instanceof Array) {
        var ai = 0,
            bi = 0;
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

var getAllCellsInExcelRange = function getAllCellsInExcelRange(range) {
    var cells = range.split(':');
    var cell1props = getExcelRowCol(cells[0]);
    var cell2props = getExcelRowCol(cells[1]);
    return getAllCellsInNumericRange(cell1props.row, cell1props.col, cell2props.row, cell2props.col);
};

var getAllCellsInNumericRange = function getAllCellsInNumericRange(row1, col1, row2, col2) {
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

var boolToInt = function boolToInt(bool) {
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
    generateRId: generateRId,
    getHashOfPassword: getHashOfPassword,
    getExcelAlpha: getExcelAlpha,
    getExcelCellRef: getExcelCellRef,
    getExcelRowCol: getExcelRowCol,
    getExcelTS: getExcelTS,
    sortCellRefs: sortCellRefs,
    arrayIntersectSafe: arrayIntersectSafe,
    getAllCellsInExcelRange: getAllCellsInExcelRange,
    getAllCellsInNumericRange: getAllCellsInNumericRange,
    boolToInt: boolToInt
};
//# sourceMappingURL=utils.js.map
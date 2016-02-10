/* REFERENCES
    http://www.ecma-international.org/news/TC45_current_work/OpenXML%20White%20Paper.pdf
    http://www.ecma-international.org/publications/standards/Ecma-376.htm
    http://www.openoffice.org/sc/excelfileformat.pdf 
    http://officeopenxml.com/anatomyofOOXML-xlsx.php
*/

/**
 * Translates a column number into the Alpha equivalent used by Excel
 * @function getExcelAlpha
 * @param {Number} colNum Column number that is to be transalated
 * @param {Boolean} capitalize States whether return string should be capital, defaults to true
 * @returns {String} The Excel alpha representation of the column number
 * @example
 * // returns B
 * getExcelAlpha(2);
 */
let getExcelAlpha = (colNum, capitalize) => {
    capitalize = capitalize ? capitalize : true;
    var remaining = colNum;
    var aCharCode = capitalize ? 65 : 97;
    var columnName = '';
    while (remaining > 0) {
        var mod = (remaining - 1) % 26;
        columnName = String.fromCharCode(aCharCode + mod) + columnName;
        remaining = (remaining - 1 - mod) / 26;
    } 
    return columnName;
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
let getExcelTS = (date) => {
    var epoch = new Date(1899, 11, 31);
    var dt = date.setDate(date.getDate() + 1);
    var ts = (dt-epoch) / (1000 * 60 * 60 * 24);
    return ts;
};

module.exports = {
	WorkBook : require('./lib/workbook/index.js'),
	getExcelRowCol : getExcelRowCol,
	getExcelAlpha : getExcelAlpha,
	getExcelTS : getExcelTS
};

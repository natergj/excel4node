const _ = require('lodash');
const Cell = require('./cell.js');
const Row = require('../row/row.js');
const Column = require('../column/column.js');
const Style = require('../style/style.js');
const utils = require('../utils.js');
const util = require('util');

function stringSetter(val) {
    let logger = this.ws.wb.logger;
    let chars, chr;
    chars = /[\u0000-\u0008\u000B-\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE-\uFFFF]/;
    chr = val.match(chars);
    if (chr) {
        logger.warn('Invalid Character for XML "' + chr + '" in string "' + val + '"');
        val = val.replace(chr, '');
    }

    if (typeof(val) !== 'string') {
        logger.warn('Value sent to String function of cells %s was not a string, it has type of %s', 
                    JSON.stringify(this.excelRefs), 
                    typeof(val));
        val = '';
    }

    val = val.toString();
    // Remove Control characters, they aren't understood by xmlbuilder
    val = val.replace(/[\u0000-\u0008\u000B-\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE-\uFFFF]/, '');

    if (!this.merged) {
        this.cells.forEach((c) => {
            c.string(this.ws.wb.getStringIndex(val));
        });
    } else {
        let c = this.cells[0];
        c.string(this.ws.wb.getStringIndex(val));
    }
    return this;
}

function complexStringSetter(val) {
    if (!this.merged) {
        this.cells.forEach((c) => {
            c.string(this.ws.wb.getStringIndex(val));
        });
    } else {
        let c = this.cells[0];
        c.string(this.ws.wb.getStringIndex(val));
    }
    return this;
}

function numberSetter(val) {
    if (val === undefined || parseFloat(val) !== val) {
        throw new TypeError(util.format('Value sent to Number function of cells %s was not a number, it has type of %s and value of %s',
            JSON.stringify(this.excelRefs),
            typeof(val),
            val
        ));
    }
    val = parseFloat(val);

    if (!this.merged) {
        this.cells.forEach((c, i) => {
            c.number(val);
        });
    } else {
        var c = this.cells[0];
        c.number(val);
    }
    return this;    
}

function booleanSetter(val) {
    if (val === undefined || typeof (val.toString().toLowerCase() === 'true' || ((val.toString().toLowerCase() === 'false') ? false : val)) !== 'boolean') {
        throw new TypeError(util.format('Value sent to Bool function of cells %s was not a bool, it has type of %s and value of %s',
            JSON.stringify(this.excelRefs),
            typeof(val),
            val
        ));
    }
    val = val.toString().toLowerCase() === 'true';

    if (!this.merged) {
        this.cells.forEach((c, i) => {
            c.bool(val.toString());
        });
    } else {
        var c = this.cells[0];
        c.bool(val.toString());
    }
    return this;
}

function formulaSetter(val) {
    if (typeof(val) !== 'string') {
        throw new TypeError(util.format('Value sent to Formula function of cells %s was not a string, it has type of %s', JSON.stringify(this.excelRefs), typeof(val)));
    }
    if (this.merged !== true) {
        this.cells.forEach((c, i) => {
            c.formula(val);
        });
    } else {
        var c = this.cells[0];
        c.formula(val);
    }

    return this;
}

function dateSetter(val) {
    let thisDate = new Date(val);
    if (isNaN(thisDate.getTime())) {
        throw new TypeError(util.format('Invalid date sent to date function of cells. %s could not be converted to a date.', val));
    }
    if (this.merged !== true) {
        this.cells.forEach((c, i) => {
            c.date(thisDate);
        });
    } else {
        var c = this.cells[0];
        c.date(thisDate);
    }
    return styleSetter.bind(this)({
        numberFormat: '[$-409]' + this.ws.wb.opts.dateFormat
    });
}

function styleSetter(val) {
    let thisStyle;
    if (val instanceof Style) {
        thisStyle = val.toObject();
    } else if (val instanceof Object) {
        thisStyle = val;
    } else {
        throw new TypeError(util.format('Parameter sent to Style function must be an instance of a Style or a style configuration object'));
    }

    let borderEdges = {};
    if (thisStyle.border && thisStyle.border.outline) {
        borderEdges.left = this.firstCol;
        borderEdges.right = this.lastCol;
        borderEdges.top = this.firstRow;
        borderEdges.bottom = this.lastRow;
    }

    this.cells.forEach((c) => {
        if (thisStyle.border && thisStyle.border.outline) {
            let thisCellsBorder = {};
            if (c.row === borderEdges.top && thisStyle.border.top) {
                thisCellsBorder.top = thisStyle.border.top;
            }
            if (c.row === borderEdges.bottom && thisStyle.border.bottom) {
                thisCellsBorder.bottom = thisStyle.border.bottom;
            }
            if (c.col === borderEdges.left && thisStyle.border.left) {
                thisCellsBorder.left = thisStyle.border.left;
            }
            if (c.col === borderEdges.right && thisStyle.border.right) {
                thisCellsBorder.right = thisStyle.border.right;
            }
            thisStyle.border = thisCellsBorder;
        }

        if (c.s === 0) {
            let thisCellStyle = this.ws.wb.createStyle(thisStyle);
            c.style(thisCellStyle.ids.cellXfs);
        } else {
            let curStyle = this.ws.wb.styles[c.s];
            let newStyleOpts = _.merge({}, curStyle.toObject(), thisStyle);
            let mergedStyle = this.ws.wb.createStyle(newStyleOpts);
            c.style(mergedStyle.ids.cellXfs);
        }
    });
    return this;
}

function hyperlinkSetter(url, displayStr, tooltip) {
    this.excelRefs.forEach((ref) => {
        displayStr = typeof displayStr === 'string' ? displayStr : url;
        this.ws.hyperlinkCollection.add({
            location: url,
            display: displayStr,
            tooltip: tooltip,
            ref: ref
        });
    });
    stringSetter.bind(this)(displayStr);
    return styleSetter.bind(this)({
        font: {
            color: 'Blue',
            underline: true
        }
    });
}

function mergeCells(cellBlock) {
    let excelRefs = cellBlock.excelRefs;
    if (excelRefs instanceof Array && excelRefs.length > 0) {
        excelRefs.sort(utils.sortCellRefs);

        let cellRange = excelRefs[0] + ':' + excelRefs[excelRefs.length - 1];
        let rangeCells = excelRefs;

        let okToMerge = true;
        cellBlock.ws.mergedCells.forEach((cr) => {
            // Check to see if currently merged cells contain cells in new merge request
            let curCells = utils.getAllCellsInExcelRange(cr);
            let intersection = utils.arrayIntersectSafe(rangeCells, curCells);
            if (intersection.length > 0) {
                okToMerge = false;
                cellBlock.ws.wb.logger.error(`Invalid Range for: ${cellRange}. Some cells in this range are already included in another merged cell range: ${cr}.`);
            }
        });
        if (okToMerge) {
            cellBlock.ws.mergedCells.push(cellRange);
        }
    } else {
        throw new TypeError(util.format('excelRefs variable sent to mergeCells function must be an array with length > 0'));
    }
}

/**
 * @class cellBlock
 */
class cellBlock {

    constructor() {
        this.ws;
        this.cells = [];
        this.excelRefs = [];
        this.merged = false;
    }

    get matrix() {
        let matrix = [];
        let tmpObj = {};
        this.cells.forEach((c) => {
            if (!tmpObj[c.row]) {
                tmpObj[c.row] = [];
            }
            tmpObj[c.row].push(c);
        });
        let rows = Object.keys(tmpObj);
        rows.forEach((r) => {
            tmpObj[r].sort((a, b) => {
                return a.col - b.col;
            });
            matrix.push(tmpObj[r]);
        });
        return matrix;
    }

    get firstRow() {
        let firstRow;
        this.cells.forEach((c) => {
            if (c.row < firstRow || firstRow === undefined) {
                firstRow = c.row;
            }
        });
        return firstRow;
    }

    get lastRow() {
        let lastRow;
        this.cells.forEach((c) => {
            if (c.row > lastRow || lastRow === undefined) {
                lastRow = c.row;
            }
        });
        return lastRow;
    }

    get firstCol() {
        let firstCol;
        this.cells.forEach((c) => {
            if (c.col < firstCol || firstCol === undefined) {
                firstCol = c.col;
            }
        });
        return firstCol;
    }

    get lastCol() {
        let lastCol;
        this.cells.forEach((c) => {
            if (c.col > lastCol || lastCol === undefined) {
                lastCol = c.col;
            }
        });
        return lastCol;
    }
}

/**
 * Module repesenting a Cell Accessor
 * @alias Worksheet.cell
 * @namespace
 * @func Worksheet.cell
 * @desc Access a range of cells in order to manipulate values
 * @param {Number} row1 Row of top left cell
 * @param {Number} col1 Column of top left cell
 * @param {Number} row2 Row of bottom right cell (optional)
 * @param {Number} col2 Column of bottom right cell (optional)
 * @param {Boolean} isMerged Merged the cell range into a single cell
 * @returns {cellBlock}
 */
function cellAccessor(row1, col1, row2, col2, isMerged) {
    let theseCells = new cellBlock();
    theseCells.ws = this;

    row2 = row2 ? row2 : row1;
    col2 = col2 ? col2 : col1;

    if (row2 > this.lastUsedRow) {
        this.lastUsedRow = row2;
    }

    if (col2 > this.lastUsedCol) {
        this.lastUsedCol = col2;
    }

    for (let r = row1; r <= row2; r++) {
        for (let c = col1; c <= col2; c++) {
            let ref = `${utils.getExcelAlpha(c)}${r}`;
            if (!this.cells[ref]) {
                this.cells[ref] = new Cell(r, c);
            }
            if (!this.rows[r]) {
                this.rows[r] = new Row(r, this);
            }
            if (this.rows[r].cellRefs.indexOf(ref) < 0) {
                this.rows[r].cellRefs.push(ref);
            }

            theseCells.cells.push(this.cells[ref]);
            theseCells.excelRefs.push(ref);
        }
    }
    if (isMerged) {
        theseCells.merged = true;
        mergeCells(theseCells);
    }

    return theseCells;
}

/**
 * @alias cellBlock.string
 * @func cellBlock.string
 * @param {String} val Value of String
 * @returns {cellBlock} Block of cells with attached methods
 */
cellBlock.prototype.string = function (val) {
    if (val instanceof Array) {
        return complexStringSetter.bind(this)(val);
    } else {
        return stringSetter.bind(this)(val);
    }
};

/**
 * @alias cellBlock.style
 * @func cellBlock.style
 * @param {Object} style One of a Style instance or an object with Style parameters
 * @returns {cellBlock} Block of cells with attached methods
 */
cellBlock.prototype.style = styleSetter;

/**
 * @alias cellBlock.number
 * @func cellBlock.number
 * @param {Number} val Value of Number
 * @returns {cellBlock} Block of cells with attached methods
 */
cellBlock.prototype.number = numberSetter;

/**
 * @alias cellBlock.bool
 * @func cellBlock.bool
 * @param {Boolean} val Value of Boolean
 * @returns {cellBlock} Block of cells with attached methods
 */
cellBlock.prototype.bool = booleanSetter;

/**
 * @alias cellBlock.formula
 * @func cellBlock.formula
 * @param {String} val Excel style formula as string
 * @returns {cellBlock} Block of cells with attached methods
 */
cellBlock.prototype.formula = formulaSetter;

/**
 * @alias cellBlock.date
 * @func cellBlock.date
 * @param {Date} val Value of Date
 * @returns {cellBlock} Block of cells with attached methods
 */
cellBlock.prototype.date = dateSetter;

/**
 * @alias cellBlock.link
 * @func cellBlock.link
 * @param {String} url Value of Hyperlink URL
 * @param {String} displayStr Value of String representation of URL
 * @param {String} tooltip Value of text to display as hover
 * @returns {cellBlock} Block of cells with attached methods
 */
cellBlock.prototype.link = hyperlinkSetter;

module.exports = cellAccessor;
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = require('lodash');
var Cell = require('./cell.js');
var Row = require('../row/row.js');
var Column = require('../column/column.js');
var Style = require('../style/style.js');
var utils = require('../utils.js');
var util = require('util');

function stringSetter(val) {
    var _this = this;

    var logger = this.ws.wb.logger;
    var chars = void 0,
        chr = void 0;
    chars = /[\u0000-\u0008\u000B-\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE-\uFFFF]/;
    chr = val.match(chars);
    if (chr) {
        logger.warn('Invalid Character for XML "' + chr + '" in string "' + val + '"');
        val = val.replace(chr, '');
    }

    if (typeof val !== 'string') {
        logger.warn('Value sent to String function of cells %s was not a string, it has type of %s', JSON.stringify(this.excelRefs), typeof val === 'undefined' ? 'undefined' : _typeof(val));
        val = '';
    }

    val = val.toString();
    // Remove Control characters, they aren't understood by xmlbuilder
    val = val.replace(/[\u0000-\u0008\u000B-\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE-\uFFFF]/, '');

    if (!this.merged) {
        this.cells.forEach(function (c) {
            c.string(_this.ws.wb.getStringIndex(val));
        });
    } else {
        var c = this.cells[0];
        c.string(this.ws.wb.getStringIndex(val));
    }
    return this;
}

function complexStringSetter(val) {
    var _this2 = this;

    if (!this.merged) {
        this.cells.forEach(function (c) {
            c.string(_this2.ws.wb.getStringIndex(val));
        });
    } else {
        var c = this.cells[0];
        c.string(this.ws.wb.getStringIndex(val));
    }
    return this;
}

function numberSetter(val) {
    if (val === undefined || parseFloat(val) !== val) {
        throw new TypeError(util.format('Value sent to Number function of cells %s was not a number, it has type of %s and value of %s', JSON.stringify(this.excelRefs), typeof val === 'undefined' ? 'undefined' : _typeof(val), val));
    }
    val = parseFloat(val);

    if (!this.merged) {
        this.cells.forEach(function (c, i) {
            c.number(val);
        });
    } else {
        var c = this.cells[0];
        c.number(val);
    }
    return this;
}

function booleanSetter(val) {
    if (val === undefined || typeof (val.toString().toLowerCase() === 'true' || (val.toString().toLowerCase() === 'false' ? false : val)) !== 'boolean') {
        throw new TypeError(util.format('Value sent to Bool function of cells %s was not a bool, it has type of %s and value of %s', JSON.stringify(this.excelRefs), typeof val === 'undefined' ? 'undefined' : _typeof(val), val));
    }
    val = val.toString().toLowerCase() === 'true';

    if (!this.merged) {
        this.cells.forEach(function (c, i) {
            c.bool(val.toString());
        });
    } else {
        var c = this.cells[0];
        c.bool(val.toString());
    }
    return this;
}

function formulaSetter(val) {
    if (typeof val !== 'string') {
        throw new TypeError(util.format('Value sent to Formula function of cells %s was not a string, it has type of %s', JSON.stringify(this.excelRefs), typeof val === 'undefined' ? 'undefined' : _typeof(val)));
    }
    if (this.merged !== true) {
        this.cells.forEach(function (c, i) {
            c.formula(val);
        });
    } else {
        var c = this.cells[0];
        c.formula(val);
    }

    return this;
}

function dateSetter(val) {
    var thisDate = new Date(val);
    if (isNaN(thisDate.getTime())) {
        throw new TypeError(util.format('Invalid date sent to date function of cells. %s could not be converted to a date.', val));
    }
    if (this.merged !== true) {
        this.cells.forEach(function (c, i) {
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
    var _this3 = this;

    var thisStyle = void 0;
    if (val instanceof Style) {
        thisStyle = val.toObject();
    } else if (val instanceof Object) {
        thisStyle = val;
    } else {
        throw new TypeError(util.format('Parameter sent to Style function must be an instance of a Style or a style configuration object'));
    }

    var borderEdges = {};
    if (thisStyle.border && thisStyle.border.outline) {
        borderEdges.left = this.firstCol;
        borderEdges.right = this.lastCol;
        borderEdges.top = this.firstRow;
        borderEdges.bottom = this.lastRow;
    }

    this.cells.forEach(function (c) {
        if (thisStyle.border && thisStyle.border.outline) {
            var thisCellsBorder = {};
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
            var thisCellStyle = _this3.ws.wb.createStyle(thisStyle);
            c.style(thisCellStyle.ids.cellXfs);
        } else {
            var curStyle = _this3.ws.wb.styles[c.s];
            var newStyleOpts = _.merge({}, curStyle.toObject(), thisStyle);
            var mergedStyle = _this3.ws.wb.createStyle(newStyleOpts);
            c.style(mergedStyle.ids.cellXfs);
        }
    });
    return this;
}

function hyperlinkSetter(url, displayStr, tooltip) {
    var _this4 = this;

    this.excelRefs.forEach(function (ref) {
        displayStr = typeof displayStr === 'string' ? displayStr : url;
        _this4.ws.hyperlinkCollection.add({
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
    var excelRefs = cellBlock.excelRefs;
    if (excelRefs instanceof Array && excelRefs.length > 0) {
        excelRefs.sort(utils.sortCellRefs);

        var cellRange = excelRefs[0] + ':' + excelRefs[excelRefs.length - 1];
        var rangeCells = excelRefs;

        var okToMerge = true;
        cellBlock.ws.mergedCells.forEach(function (cr) {
            // Check to see if currently merged cells contain cells in new merge request
            var curCells = utils.getAllCellsInExcelRange(cr);
            var intersection = utils.arrayIntersectSafe(rangeCells, curCells);
            if (intersection.length > 0) {
                okToMerge = false;
                cellBlock.ws.wb.logger.error('Invalid Range for: ' + cellRange + '. Some cells in this range are already included in another merged cell range: ' + cr + '.');
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

var cellBlock = function () {
    function cellBlock() {
        _classCallCheck(this, cellBlock);

        this.ws;
        this.cells = [];
        this.excelRefs = [];
        this.merged = false;
    }

    _createClass(cellBlock, [{
        key: 'matrix',
        get: function get() {
            var matrix = [];
            var tmpObj = {};
            this.cells.forEach(function (c) {
                if (!tmpObj[c.row]) {
                    tmpObj[c.row] = [];
                }
                tmpObj[c.row].push(c);
            });
            var rows = Object.keys(tmpObj);
            rows.forEach(function (r) {
                tmpObj[r].sort(function (a, b) {
                    return a.col - b.col;
                });
                matrix.push(tmpObj[r]);
            });
            return matrix;
        }
    }, {
        key: 'firstRow',
        get: function get() {
            var firstRow = void 0;
            this.cells.forEach(function (c) {
                if (c.row < firstRow || firstRow === undefined) {
                    firstRow = c.row;
                }
            });
            return firstRow;
        }
    }, {
        key: 'lastRow',
        get: function get() {
            var lastRow = void 0;
            this.cells.forEach(function (c) {
                if (c.row > lastRow || lastRow === undefined) {
                    lastRow = c.row;
                }
            });
            return lastRow;
        }
    }, {
        key: 'firstCol',
        get: function get() {
            var firstCol = void 0;
            this.cells.forEach(function (c) {
                if (c.col < firstCol || firstCol === undefined) {
                    firstCol = c.col;
                }
            });
            return firstCol;
        }
    }, {
        key: 'lastCol',
        get: function get() {
            var lastCol = void 0;
            this.cells.forEach(function (c) {
                if (c.col > lastCol || lastCol === undefined) {
                    lastCol = c.col;
                }
            });
            return lastCol;
        }
    }]);

    return cellBlock;
}();

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
    var theseCells = new cellBlock();
    theseCells.ws = this;

    row2 = row2 ? row2 : row1;
    col2 = col2 ? col2 : col1;

    if (row2 > this.lastUsedRow) {
        this.lastUsedRow = row2;
    }

    if (col2 > this.lastUsedCol) {
        this.lastUsedCol = col2;
    }

    for (var r = row1; r <= row2; r++) {
        for (var c = col1; c <= col2; c++) {
            var ref = '' + utils.getExcelAlpha(c) + r;
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
//# sourceMappingURL=index.js.map
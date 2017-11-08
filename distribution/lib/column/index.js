'use strict';

var _ = require('lodash');
var Cell = require('../cell/cell.js');
var Row = require('../row/row.js');
var Column = require('../column/column.js');
var utils = require('../utils.js');

/**
 * Module repesenting a Column Accessor
 * @alias Worksheet.column
 * @namespace
 * @func Worksheet.column
 * @desc Access a column in order to manipulate values
 * @param {Number} col Column of top left cell
 * @returns {Column}
 */
var colAccessor = function colAccessor(ws, col) {
    if (!(ws.cols[col] instanceof Column)) {
        ws.cols[col] = new Column(col, ws);
    }
    return ws.cols[col];
};

module.exports = colAccessor;
//# sourceMappingURL=index.js.map
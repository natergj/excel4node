'use strict';

var _ = require('lodash');
var Cell = require('../cell/cell.js');
var Row = require('../row/row.js');
var Column = require('../column/column.js');
var WorkSheet = require('../worksheet/worksheet.js');
var utils = require('../utils.js');

/**
 * Module repesenting a Row Accessor
 * @alias Worksheet.row
 * @namespace
 * @func Worksheet.row
 * @desc Access a row in order to manipulate values
 * @param {Number} row Row of top left cell
 * @returns {Row}
 */
var rowAccessor = function rowAccessor(ws, row) {

    if (typeof row !== 'number') {
        throw new TypeError('Row sent to row accessor was not a number.');
    }

    if (!(ws.rows[row] instanceof Row)) {
        ws.rows[row] = new Row(row, ws);
    }

    return ws.rows[row];
};

module.exports = rowAccessor;
//# sourceMappingURL=index.js.map
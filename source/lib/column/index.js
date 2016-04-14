const _ = require('lodash');
const Cell = require('../cell/cell.js');
const Row = require('../row/row.js');
const Column = require('../column/column.js');
const utils = require('../utils.js');

/**
 * Module repesenting a Column Accessor
 * @alias Worksheet.column
 * @namespace
 * @func Worksheet.column
 * @desc Access a column in order to manipulate values
 * @param {Number} col Column of top left cell
 * @returns {Column}
 */
let colAccessor = (ws, col) => {
    if (!(ws.cols[col] instanceof Column)) {
        ws.cols[col] = new Column(col, ws);
    }
    return ws.cols[col];
};

module.exports = colAccessor;
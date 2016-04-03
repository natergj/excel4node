const _ = require('lodash');
const Cell = require('../cell/cell.js');
const Row = require('../row/row.js');
const Column = require('../column/column.js');
const WorkSheet = require('../worksheet/worksheet.js');
const utils = require('../utils.js');

let rowAccessor = function (ws, row) {

    if (typeof row !== 'number') {
        throw new TypeError('Row sent to row accessor was not a number.');
    }

    if (!(ws.rows[row] instanceof Row)) {
        ws.rows[row] = new Row(row, ws);
    }

    return ws.rows[row];
};



module.exports = rowAccessor;
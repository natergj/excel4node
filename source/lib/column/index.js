const _ = require('lodash');
const Cell = require('../cell/cell.js');
const Row = require('../row/row.js');
const Column = require('../column/column.js');
const utils = require('../utils.js');

let colAccessor = (ws, col) => {
    if (!(ws.cols[col] instanceof Column)) {
        ws.cols[col] = new Column(col, ws);
    }
    return  ws.cols[col];
};

module.exports = colAccessor;
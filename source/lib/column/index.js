const _ = require('lodash');
const Cell = require('../cell/cell.js');
const Row = require('../row/row.js');
const Column = require('../column/column.js');
const utils = require('../utils.js');

let colAccessor = (ws, col) => {

    let returnObj = {};
    let thisCol = ws.cols[col] instanceof Column ? ws.cols[col] : new col(col);

    returnObj.ws = ws;
    returnObj.col = thisCol;

    return returnObj;
};

module.exports = colAccessor;
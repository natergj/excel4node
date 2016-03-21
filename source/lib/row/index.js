const _ = require('lodash');
const Cell = require('../cell/cell.js');
const Row = require('../row/row.js');
const Column = require('../column/column.js');
const WorkSheet = require('../worksheet/worksheet.js');
const utils = require('../utils.js');

let rowAccessor = function (ws, row) {
    let returnObj = {};

    if (typeof row !== 'number') {
        throw new TypeError('Row sent to row accessor was not a number.');
    }

    let thisRow = ws.rows[row] instanceof Row ? ws.rows[row] : new Row(row);

    /*-- 
        opts : {
            firstColumn: Number,
            lastColumn: Number,
            lastRow: Number
        }
    */
    returnObj.Filter = (opts, filters) => {
        let theseOpts = opts instanceof Object ? opts : {};
        let theseFilters = filters instanceof Array ? filters :
                              opts instanceof Array ? opts : [];

        ws.opts.autoFilter.startRow = row;
        if (typeof theseOpts.lastRow === 'number') {
            ws.opts.autoFilter.endRow = theseOpts.lastRow;
        }

        if (typeof theseOpts.firstColumn === 'number' && typeof theseOpts.lastColumn === 'number') {
            ws.opts.autoFilter.startCol = theseOpts.firstColumn;
            ws.opts.autoFilter.endCol = theseOpts.lastColumn;
        }

        ws.opts.autoFilter.filters = theseFilters;
    };

    return returnObj;
};



module.exports = rowAccessor;
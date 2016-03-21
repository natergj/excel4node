const _ = require('lodash');
const Cell = require('./cell.js');
const Row = require('../row/row.js');
const Column = require('../column/column.js');
const Style = require('../style/style.js');
const utils = require('../utils.js');

let stringSetter = (val, theseCells) => {
    let logger = theseCells.ws.wb.logger;
    let chars, chr;
    chars = /[\u0000-\u0008\u000B-\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE-\uFFFF]/;
    chr = val.match(chars);
    if (chr) {
        logger.warn('Invalid Character for XML "' + chr + '" in string "' + val + '"');
        val = val.replace(chr, '');
    }

    if (typeof(val) !== 'string') {
        logger.warn('Value sent to String function of cells %s was not a string, it has type of %s', 
                    JSON.stringify(theseCells.excelRefs), 
                    typeof(val));
        val = '';
    }

    val = val.toString();
    // Remove Control characters, they aren't understood by xmlbuilder
    val = val.replace(/[\u0000-\u0008\u000B-\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE-\uFFFF]/, '');

    if (!theseCells.merged) {
        theseCells.cells.forEach((c) => {
            c.String(theseCells.ws.wb.getStringIndex(val));
        });
    } else {
        let c = theseCells.cells[0];
        c.String(theseCells.ws.wb.getStringIndex(val));
    }
    return theseCells;
};

let numberSetter = (val, theseCells) => {
    if (val === undefined || parseFloat(val) !== val) {
        console.log('Value sent to Number function of cells %s was not a number, it has type of %s and value of %s',
            JSON.stringify(theseCells.excelRefs),
            typeof(val),
            val
        );
        val = '';
    }
    val = parseFloat(val);

    if (!theseCells.merged) {
        theseCells.cells.forEach(function (c, i) {
            c.Number(val);
        });
    } else {
        var c = theseCells.cells[0];
        c.Number(val);
    }
    return theseCells;    
};

let booleanSetter = (val, theseCells) => {
    if (val === undefined || typeof (val.toString().toLowerCase() === 'true' || ((val.toString().toLowerCase() === 'false') ? false : val)) !== 'boolean') {
        console.log('Value sent to Bool function of cells %s was not a bool, it has type of %s and value of %s',
            JSON.stringify(theseCells.excelRefs),
            typeof(val),
            val
        );
        val = '';
    }
    val = val.toString().toLowerCase() === 'true';

    if (!theseCells.merged) {
        theseCells.cells.forEach(function (c, i) {
            c.Bool(val.toString());
        });
    } else {
        var c = theseCells.cells[0];
        c.Bool(val.toString());
    }
    return theseCells;
};

let formulaSetter = (val, theseCells) => {
    if (typeof(val) !== 'string') {
        console.log('Value sent to Formula function of cells %s was not a string, it has type of %s', JSON.stringify(theseCells.excelRefs), typeof(val));
        val = '';
    }
    if (theseCells.merged !== true) {
        theseCells.cells.forEach(function (c, i) {
            c.Formula(val);
        });
    } else {
        var c = theseCells.cells[0];
        c.Formula(val);
    }

    return theseCells;
};

let styleSetter = (val, theseCells) => {
    let styleXFid;
    let thisStyle;
    if (val instanceof Style) {
        thisStyle = val;
    } else if (val instanceof Object) {
        thisStyle = theseCells.wb.Style(val);
    } else {
        throw new TypeError('Parameter sent to Style function must be an instance of a Style or a style configuration object');
    }

    styleXFid = thisStyle.ids.cellXfs;
    
    if (theseCells.merged !== true) {
        theseCells.cells.forEach(function (c, i) {
            c.Style(styleXFid);
        });
    } else {
        var c = theseCells.cells[0];
        c.Style(styleXFid);
    }   
};

let mergeCells = (ws, excelRefs) => {
    if (excelRefs instanceof Array && excelRefs.length > 0) {
        excelRefs.sort(utils.sortCellRefs);

        let cellRange = excelRefs[0] + ':' + excelRefs[excelRefs.length - 1];
        let rangeCells = excelRefs;

        let okToMerge = true;
        ws.mergedCells.forEach(function (cr) {
            // Check to see if currently merged cells contain cells in new merge request
            let curCells = utils.getAllCellsInExcelRange(cr);
            let intersection = utils.arrayIntersectSafe(rangeCells, curCells);
            if (intersection.length > 0) {
                okToMerge = false;
                ws.wb.logger.error(`Invalid Range for: ${cellRange}. Some cells in this range are already included in another merged cell range: ${cr}.`);
            }
        });
        if (okToMerge) {
            ws.mergedCells.push(cellRange);
        }
    } else {
        throw new TypeError('excelRefs variable sent to mergeCells function must be an array with length > 0');
    }
};

let cellAccessor = (ws, row1, col1, row2, col2, isMerged) => {
    
    let theseCells = {
        ws: ws,
        cells: [],
        excelRefs: [],
        merged: isMerged
    };

    row2 = row2 ? row2 : row1;
    col2 = col2 ? col2 : col1;

    if (row2 > ws.lastUsedRow) {
        ws.lastUsedRow = row2;
    }

    if (col2 > ws.lastUsedCol) {
        ws.lastUsedCol = col2;
    }

    for (let r = row1; r <= row2; r++) {
        for (let c = col1; c <= col2; c++) {
            let ref = `${utils.getExcelAlpha(c)}${r}`;
            if (!ws.cells[ref]) {
                ws.cells[ref] = new Cell(r, c);
            }
            if (!ws.rows[r]) {
                ws.rows[r] = new Row(r);
            }
            if (!ws.cols[c]) {
                ws.cols[c] = new Column(c);
            }
            if (ws.rows[r].cellRefs.indexOf(ref) < 0) {
                ws.rows[r].cellRefs.push(ref);
            }

            theseCells.cells.push(ws.cells[ref]);
            theseCells.excelRefs.push(ref);
        }
    }
    if (isMerged) {
        mergeCells(ws, theseCells.excelRefs);
    }

    theseCells.String = (val) => stringSetter(val, theseCells);
    theseCells.Number = (val) => numberSetter(val, theseCells);
    theseCells.Bool = (val) => booleanSetter(val, theseCells);
    theseCells.Formula = (val) => formulaSetter(val, theseCells);
    theseCells.Style = (val) => styleSetter(val, theseCells);

    return theseCells;
};

module.exports = cellAccessor;
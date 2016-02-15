let Cell = require('./cell.js');
let utils = require('../utils.js');
let logger = require('../logger.js');

let stringSetter = (val, theseCells) => {
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
}

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
}

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
}

let formulaSetter = (val, theseCells) => {
    if (typeof(val) !== 'string') {
        console.log('Value sent to Formula function of cells %s was not a string, it has type of %s', JSON.stringify(theseCells.excelRefs), typeof(val));
        val = '';
    }
    if (!theseCells.merged) {
        theseCells.cells.forEach(function (c, i) {
            c.Formula(val);
        });
    } else {
        var c = theseCells.cells[0];
        c.Formula(val);
    }

    return theseCells;
}

let cellAccessor = (ws, row1, col1, row2, col2, isMerged) => {

    let theseCells = {
        ws: ws,
        cells: [],
        excelRefs: [],
        merged: isMerged
    };

    row2 = row2 ? row2 : row1;
    col2 = col2 ? col2 : col1;

    for (let r = row1; r <= row2; r++) {
    	for (let c = col1; c <= col2; c++) {
    		let ref = `${utils.getExcelAlpha(c)}${r}`;
    		if(!ws.cells[ref]){
    			ws.cells[ref] = new Cell(r, c);
    		}
    		if(!ws.rows[r]){
    			ws.rows[r] = {
    				cellRefs : []
    			};
    		}
    		if(ws.rows[r].cellRefs.indexOf(ref) < 0){
    			ws.rows[r].cellRefs.push(ref);
    		}
    		theseCells.cells.push(ws.cells[ref]);
    		theseCells.excelRefs.push(ref);
    	}
    }

    theseCells.String = (val) => stringSetter(val, theseCells);
    theseCells.Number = (val) => numberSetter(val, theseCells);
    theseCells.Bool = (val) => booleanSetter(val, theseCells);
    theseCells.Formula = (val) => formulaSetter(val, theseCells);

    return theseCells;
};

module.exports = cellAccessor;
let Cell = require('./cell.js');
let utils = require('../utils.js');
let logger = require('../logger.js');

let cellAccessor = (ws, row1, col1, row2, col2, isMerged) => {

    let theseCells = {
        cells: [],
        excelRefs: []
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

    let string = (val) => {

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

        if (!isMerged) {
            theseCells.cells.forEach((c) => {
                c.String(ws.wb.getStringIndex(val));
            });
        } else {
            let c = theseCells.cells[0];
            c.String(ws.wb.getStringIndex(val));
        }
        return theseCells;
    };


    theseCells.String = string;

    return theseCells;
};

module.exports = cellAccessor;
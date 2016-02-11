let utils = require('../utils.js');

// §18.3.1.4 c (Cell)
class Cell {
	constructor(row, col) {
    	this.r = `${utils.getExcelAlpha(col)}${row}`; // 'r' attribute
    	this.s = 1; // 's' attribute refering to style index
    	this.t = null; // 't' attribute stating Cell data type - §18.18.11 ST_CellType (Cell Type)
    	this.f = null; // 'f' child element used for formulas
    	this.v = null; // 'v' child element for values
	}

	String(index) {
		this.v = index;
		this.t = 's';
	}
}

module.exports = Cell;


const utils = require('../utils.js');

// §18.3.1.4 c (Cell)
class Cell {
    constructor(row, col) {
        this.r = `${utils.getExcelAlpha(col)}${row}`; // 'r' attribute
        this.s = 0; // 's' attribute refering to style index
        this.t = null; // 't' attribute stating Cell data type - §18.18.11 ST_CellType (Cell Type)
        this.f = null; // 'f' child element used for formulas
        this.v = null; // 'v' child element for values
    }

    String(index) {
        this.t = 's';
        this.v = index;
        this.f = null;
    }

    Number(val) {
        this.t = 'n';
        this.v = val;
        this.f = null;
    }

    Formula(formula) {
        this.t = null;
        this.v = null;
        this.f = formula;
    }

    Bool(val) {
        this.t = 'b';
        this.v = val;
        this.f = null;
    }

    Date(dt) {
        this.t = 'd';
        this.v = utils.getExcelTS(dt);
        this.f = null;
    }

    Style(sId) {
        this.s = sId;
    }
}

module.exports = Cell;


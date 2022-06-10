const utils = require('../utils.js');
const Comment = require('../classes/comment');

// §18.3.1.4 c (Cell)
class Cell {
    /**
   * Create an Excel Cell
   * @private
   * @param {Number} row Row of cell.
   * @param {Number} col Column of cell
   */
    constructor(row, col) {
        if (row <= 0) throw 'Row parameter must not be zero or negative.';
        if (col <= 0) throw 'Col parameter must not be zero or negative.';
        this.r = `${utils.getExcelAlpha(col)}${row}`; // 'r' attribute
        this.s = 0; // 's' attribute refering to style index
        this.t = null; // 't' attribute stating Cell data type - §18.18.11 ST_CellType (Cell Type)
        this.f = null; // 'f' child element used for formulas
        this.v = null; // 'v' child element for values
        this.row = row; // used internally throughout code. Does not go into XML
        this.col = col; // used internally throughout code. Does not go into XML
    }

    get comment() {
        return this.comments[this.r];
    }

    string(index) {
        this.t = 's';
        this.v = index;
        this.f = null;
    }

    number(val) {
        this.t = 'n';
        this.v = val;
        this.f = null;
    }

    formula(formula) {
        this.t = null;
        this.v = null;
        this.f = formula;
    }

    bool(val) {
        this.t = 'b';
        this.v = val;
        this.f = null;
    }

    date(dt) {
        this.t = null;
        this.v = utils.getExcelTS(dt);
        this.f = null;
    }

    style(sId) {
        this.s = sId;
    }

    addToXMLele(ele) {
        if (this.v === null && this.is === null) {
            return;
        }

        let cEle = ele.ele('c').att('r', this.r).att('s', this.s);
        if (this.t !== null) {
            cEle.att('t', this.t);
        }
        if (this.f !== null) {
            cEle.ele('f').txt(this.f).up();
        }
        if (this.v !== null) {
            cEle.ele('v').txt(this.v).up();
        }
        cEle.up();
    }
}

module.exports = Cell;

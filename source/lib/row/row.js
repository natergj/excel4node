const utils = require('../utils.js');
const logger = require('../logger.js');
const _ = require('lodash');

class Row {
    constructor(row) {
        this.cellRefs = [];
        this.collapsed = null;
        this.customFormat = null;
        this.customHeight = null;
        this.hidden = null;
        this.ht = null;
        this.outlineLevel = null;
        this.r = row;
        this.s = null;
        this.thickBot = null;
        this.thicktop = null;
    }

    get spans() {
        if (this.cellRefs instanceof Array && this.cellRefs.length > 0) {
            return `${utils.getExcelRowCol(this.cellRefs[0]).row}:${utils.getExcelRowCol(this.cellRefs[this.cellRefs.length - 1]).row}`;
        } else {
            return `${this.r}:${this.r}`;
        }
    }

    get firstColumn() {
        if (this.cellRefs instanceof Array && this.cellRefs.length > 0) {
            return utils.getExcelRowCol(this.cellRefs[0]).col;
        } else {
            return 1;
        }
    }

    get firstColumnAlpha() {
        if (this.cellRefs instanceof Array && this.cellRefs.length > 0) {
            return utils.getExcelAlpha(utils.getExcelRowCol(this.cellRefs[0]).col);
        } else {
            return 'A';
        }  
    }

    get lastColumn() {
        if (this.cellRefs instanceof Array && this.cellRefs.length > 0) {
            return utils.getExcelRowCol(this.cellRefs[this.cellRefs.length - 1]).col;
        } else {
            return 1;
        }
    }

    get lastColumnAlpha() {
        if (this.cellRefs instanceof Array && this.cellRefs.length > 0) {
            return utils.getExcelAlpha(utils.getExcelRowCol(this.cellRefs[this.cellRefs.length - 1]).col);
        } else {
            return 'A';
        }  
    }
}

module.exports = Row;
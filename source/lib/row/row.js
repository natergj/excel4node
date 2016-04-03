const utils = require('../utils.js');
const _ = require('lodash');

class Row {
    constructor(row, ws) {
        this.ws = ws;
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
        this.thickTop = null;
    }

    set height(h) {
        if (typeof h === 'number') {
            this.ht = h;
            this.customHeight = true;
        } else {
            throw new TypeError('Row height must be a number');
        }
        return this.ht;
    }
    get height() {
        return this.ht;
    }

    setHeight(h) {
        if (typeof h === 'number') {
            this.ht = h;
            this.customHeight = true;
        } else {
            throw new TypeError('Row height must be a number');
        }
        return this;
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

    filter(opts) {

        let theseOpts = opts instanceof Object ? opts : {};
        let theseFilters = opts.filters instanceof Array ? opts.filters : [];

        let o = this.ws.opts.autoFilter;
        o.startRow = this.r;
        if (typeof theseOpts.lastRow === 'number') {
            o.endRow = theseOpts.lastRow;
        }

        if (typeof theseOpts.firstColumn === 'number' && typeof theseOpts.lastColumn === 'number') {
            o.startCol = theseOpts.firstColumn;
            o.endCol = theseOpts.lastColumn;
        }

        // Programmer Note: DefinedName class is added to workbook during workbook write process for filters

        this.ws.opts.autoFilter.filters = theseFilters;
    }

    hide() {
        this.hidden = true;
        return this;
    }

    group(level, collapsed) {
        if (parseInt(level) === level) {
            this.outlineLevel = level;
        } else {
            throw new TypeError('Row group level must be a positive integer');
        }

        if (collapsed === undefined) {
            return this;
        }

        if (typeof collapsed === 'boolean') {
            this.collapsed = collapsed;
            this.hidden = collapsed;
        } else {
            throw new TypeError('Row group collapse flag must be a boolean');
        }

        return this;
    }


    freeze(jumpTo) {
        let o = this.ws.opts.sheetView.pane;
        jumpTo = typeof jumpTo === 'number' && jumpTo > this.r ? jumpTo : this.r + 1;
        o.state = 'frozen';
        o.ySplit = this.r;
        o.activePane = 'bottomRight';
        o.xSplit === null ? 
            o.topLeftCell = utils.getExcelCellRef(jumpTo, 1) : 
            o.topLeftCell = utils.getExcelCellRef(jumpTo, utils.getExcelRowCol(o.topLeftCell).col);
        return this;
    }
}

module.exports = Row;
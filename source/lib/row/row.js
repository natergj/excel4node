const utils = require('../utils.js');

class Row {
    /**
     * Element representing an Excel Row
     * @param {Number} row Row of cell
     * @param {Worksheet} Worksheet that contains row
     * @property {Worksheet} ws Worksheet that contains the specified Row
     * @property {Array.String} cellRefs Array of excel cell references
     * @property {Boolean} collapsed States whether row is collapsed when grouped
     * @property {Boolean} customFormat States whether the row has a custom format
     * @property {Boolean} customHeight States whether the row's height is different than default
     * @property {Boolean} hidden States whether the row is hidden
     * @property {Number} ht Height of the row (internal property)
     * @property {Number} outlineLevel Grouping level of row
     * @property {Number} r Row index
     * @property {Number} s Style index
     * @property {Boolean} thickBot States whether row has a thick bottom border
     * @property {Boolean} thickTop States whether row has a thick top border
     * @property {Number} height Height of row
     * @property {String} spans String representation of excel cell range i.e. A1:A10
     * @property {Number} firstColumn Index of the first column of the row containg data
     * @property {String} firstColumnAlpha Alpha representation of the first column of the row containing data
     * @property {Number} lastColumn Index of the last column of the row cotaining data
     * @property {String} lastColumnAlpha Alpha representation of the last column of the row containing data
     */
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

    /**
     * @alias Row.setHeight
     * @desc Sets the height of a row
     * @func Row.setHeight
     * @param {Number} val New Height of row
     * @returns {Row} Excel Row with attached methods
     */
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
        if (this.cellRefs.length > 0) {
            const startCol = utils.getExcelRowCol(this.cellRefs[0]).col;
            const endCol = utils.getExcelRowCol(this.cellRefs[this.cellRefs.length - 1]).col;
            return `${startCol}:${endCol}`;
        } else {
            return null;
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

    /**
     * @alias Row.filter
     * @desc Add autofilter dropdowns to the items of the row
     * @func Row.filter
     * @param {Object} opts Object containing options for the fitler. 
     * @param {Number} opts.lastRow Last row in which the filter show effect filtered results (optional)
     * @param {Number} opts.startCol First column that a filter dropdown should be added (optional)
     * @param {Number} opts.lastCol Last column that a filter dropdown should be added (optional)
     * @param {Array.DefinedName} opts.filters Array of filter paramaters
     * @returns {Row} Excel Row with attached methods
     */
    filter(opts = {}) {

        let theseFilters = opts.filters instanceof Array ? opts.filters : [];

        let o = this.ws.opts.autoFilter;
        o.startRow = this.r;
        if (typeof opts.lastRow === 'number') {
            o.endRow = opts.lastRow;
        }

        if (typeof opts.firstColumn === 'number' && typeof opts.lastColumn === 'number') {
            o.startCol = opts.firstColumn;
            o.endCol = opts.lastColumn;
        }

        // Programmer Note: DefinedName class is added to workbook during workbook write process for filters

        this.ws.opts.autoFilter.filters = theseFilters;
    }

    /**
     * @alias Row.hide
     * @desc Hides the row
     * @func Row.hide
     * @returns {Row} Excel Row with attached methods
     */
    hide() {
        this.hidden = true;
        return this;
    }

    /**
     * @alias Row.group
     * @desc Hides the row
     * @func Row.group
     * @param {Number} level Group level of row
     * @param {Boolean} collapsed States whether group should be collapsed or expanded by default
     * @returns {Row} Excel Row with attached methods
     */
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

    /**
     * @alias Row.freeze
     * @desc Creates Worksheet panes and freezes the top pane
     * @func Row.freeze
     * @param {Number} jumpTo Row that the bottom pane should be scrolled to by default
     * @returns {Row} Excel Row with attached methods
     */
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
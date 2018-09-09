const utils = require('../utils.js');

class Column {
    /**
     * Element representing an Excel Column
     * @param {Number} col Column of cell
     * @param {Worksheet} Worksheet that contains column
     * @property {Worksheet} ws Worksheet that contains the specified Column
     * @property {Boolean} collapsed States whether the column is collapsed if part of a group
     * @property {Boolean} customWidth States whether or not the column as a width that is not default
     * @property {Boolean} hidden States whether or not the specified column is hiddent
     * @property {Number} max The greatest column if part of a range
     * @property {Number} min The least column if part of a range
     * @property {Number} outlineLevel The grouping leve of the Column
     * @property {Number} style ID of style
     * @property {Number} width Width of the Column
     */
    constructor(col, ws) {
        this.ws = ws;
        this.collapsed = null;
        this.customWidth = null;
        this.hidden = null;
        this.max = col;
        this.min = col;
        this.outlineLevel = null;
        this.style = null;
        this.colWidth = null;
    }

    get width() {
        return this.colWidth;
    }

    set width(w) {
        if (typeof w === 'number') {
            this.colWidth = w;
            this.customWidth = true;
        } else {
            throw new TypeError('Column width must be a number');
        }
        return this.colWidth;
    }

    /**
     * @alias Column.setWidth
     * @desc Sets teh width of a column
     * @func Column.setWidth
     * @param {Number} val New Width of column
     * @returns {Column} Excel Column with attached methods
     */
    setWidth(w) {
        this.width = w;
        return this;
    }

    /**
     * @alias Column.hide
     * @desc Sets a Column to be hidden
     * @func Column.hide
     * @returns {Column} Excel Column with attached methods
     */
    hide() {
        this.hidden = true;
        return this;
    }

    /**
     * @alias Column.group
     * @desc Adds column to the specified group
     * @func Column.group
     * @param {Number} level Level of excel grouping
     * @param {Boolean} collapsed States wheter column grouping level should be collapsed by default
     * @returns {Column} Excel Column with attached methods
     */
    group(level, collapsed) {
        if (parseInt(level) === level) {
            this.outlineLevel = level;
        } else {
            throw new TypeError('Column group level must be a positive integer');
        }

        if (collapsed === undefined) {
            return this;
        }

        if (typeof collapsed === 'boolean') {
            this.collapsed = collapsed;
            this.hidden = collapsed;
        } else {
            throw new TypeError('Column group collapse flag must be a boolean');
        }

        return this;
    }

    /**
     * @alias Column.freeze
     * @desc Creates an Excel pane at the specificed column and Freezes that column from scolling
     * @func Column.freeze
     * @param {Number} jumptTo Specifies the column that the active pane will be scrolled to by default
     * @returns {Column} Excel Column with attached methods
     */
    freeze(jumpTo) {
        let o = this.ws.opts.sheetView.pane;
        jumpTo = typeof jumpTo === 'number' && jumpTo > this.min ? jumpTo : this.min + 1;
        o.state = 'frozen';
        o.xSplit = this.min;
        o.activePane = 'bottomRight';
        o.ySplit === null ?
            o.topLeftCell = utils.getExcelCellRef(1, jumpTo) :
            o.topLeftCell = utils.getExcelCellRef(utils.getExcelRowCol(o.topLeftCell).row, jumpTo);
        return this;
    }
}

module.exports = Column;
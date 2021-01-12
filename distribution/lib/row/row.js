'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var utils = require('../utils.js');

var Row = function () {
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
    function Row(row, ws) {
        _classCallCheck(this, Row);

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

    _createClass(Row, [{
        key: 'setHeight',


        /**
         * @alias Row.setHeight
         * @desc Sets the height of a row
         * @func Row.setHeight
         * @param {Number} val New Height of row
         * @returns {Row} Excel Row with attached methods
         */
        value: function setHeight(h) {
            if (typeof h === 'number') {
                this.ht = h;
                this.customHeight = true;
            } else {
                throw new TypeError('Row height must be a number');
            }
            return this;
        }
    }, {
        key: 'filter',


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
        value: function filter() {
            var opts = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};


            var theseFilters = opts.filters instanceof Array ? opts.filters : [];

            var o = this.ws.opts.autoFilter;
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

    }, {
        key: 'hide',
        value: function hide() {
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

    }, {
        key: 'group',
        value: function group(level, collapsed) {
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

    }, {
        key: 'freeze',
        value: function freeze(jumpTo) {
            var o = this.ws.opts.sheetView.pane;
            jumpTo = typeof jumpTo === 'number' && jumpTo > this.r ? jumpTo : this.r + 1;
            o.state = 'frozen';
            o.ySplit = this.r;
            o.activePane = 'bottomRight';
            o.xSplit === null ? o.topLeftCell = utils.getExcelCellRef(jumpTo, 1) : o.topLeftCell = utils.getExcelCellRef(jumpTo, utils.getExcelRowCol(o.topLeftCell).col);
            return this;
        }
    }, {
        key: 'height',
        set: function set(h) {
            if (typeof h === 'number') {
                this.ht = h;
                this.customHeight = true;
            } else {
                throw new TypeError('Row height must be a number');
            }
            return this.ht;
        },
        get: function get() {
            return this.ht;
        }
    }, {
        key: 'spans',
        get: function get() {
            if (this.cellRefs.length > 0) {
                var startCol = utils.getExcelRowCol(this.cellRefs[0]).col;
                var endCol = utils.getExcelRowCol(this.cellRefs[this.cellRefs.length - 1]).col;
                return startCol + ':' + endCol;
            } else {
                return null;
            }
        }
    }, {
        key: 'firstColumn',
        get: function get() {
            if (this.cellRefs instanceof Array && this.cellRefs.length > 0) {
                return utils.getExcelRowCol(this.cellRefs[0]).col;
            } else {
                return 1;
            }
        }
    }, {
        key: 'firstColumnAlpha',
        get: function get() {
            if (this.cellRefs instanceof Array && this.cellRefs.length > 0) {
                return utils.getExcelAlpha(utils.getExcelRowCol(this.cellRefs[0]).col);
            } else {
                return 'A';
            }
        }
    }, {
        key: 'lastColumn',
        get: function get() {
            if (this.cellRefs instanceof Array && this.cellRefs.length > 0) {
                return utils.getExcelRowCol(this.cellRefs[this.cellRefs.length - 1]).col;
            } else {
                return 1;
            }
        }
    }, {
        key: 'lastColumnAlpha',
        get: function get() {
            if (this.cellRefs instanceof Array && this.cellRefs.length > 0) {
                return utils.getExcelAlpha(utils.getExcelRowCol(this.cellRefs[this.cellRefs.length - 1]).col);
            } else {
                return 'A';
            }
        }
    }]);

    return Row;
}();

module.exports = Row;
//# sourceMappingURL=row.js.map
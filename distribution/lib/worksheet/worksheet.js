'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = require('lodash');
var CfRulesCollection = require('./cf/cf_rules_collection');
var cellAccessor = require('../cell');
var rowAccessor = require('../row');
var colAccessor = require('../column');
var wsDefaultParams = require('./sheet_default_params.js');
var HyperlinkCollection = require('./classes/hyperlink.js').HyperlinkCollection;
var DataValidation = require('./classes/dataValidation.js');
var wsDrawing = require('../drawing/index.js');
var xmlBuilder = require('./builder.js');
var optsValidator = require('./optsValidator.js');

var Worksheet = function () {
    /**
     * Create a Worksheet.
     * @class Worksheet
     * @param {Workbook} wb Workbook that the Worksheet will belong to
     * @param {String} name Name of Worksheet
     * @param {Object} opts Worksheet settings
     * @param {Object} opts.margins
     * @param {Number} opts.margins.bottom Bottom margin in inches
     * @param {Number} opts.margins.footer Footer margin in inches
     * @param {Number} opts.margins.header Header margin in inches
     * @param {Number} opts.margins.left Left margin in inches
     * @param {Number} opts.margins.right Right margin in inches
     * @param {Number} opts.margins.top Top margin in inches
     * @param {Object} opts.printOptions Print Options object
     * @param {Boolean} opts.printOptions.centerHorizontal Should data be centered horizontally when printed
     * @param {Boolean} opts.printOptions.centerVertical Should data be centered vertically when printed
     * @param {Boolean} opts.printOptions.printGridLines Should gridlines by printed
     * @param {Boolean} opts.printOptions.printHeadings Should Heading be printed
     * @param {String} opts.headerFooter Set Header and Footer strings and options. 
     * @param {String} opts.headerFooter.evenFooter Even footer text
     * @param {String} opts.headerFooter.evenHeader Even header text
     * @param {String} opts.headerFooter.firstFooter First footer text
     * @param {String} opts.headerFooter.firstHeader First header text
     * @param {String} opts.headerFooter.oddFooter Odd footer text
     * @param {String} opts.headerFooter.oddHeader Odd header text
     * @param {Boolean} opts.headerFooter.alignWithMargins Should header/footer align with margins
     * @param {Boolean} opts.headerFooter.differentFirst Should header/footer show a different header/footer on first page
     * @param {Boolean} opts.headerFooter.differentOddEven Should header/footer show a different header/footer on odd and even pages
     * @param {Boolean} opts.headerFooter.scaleWithDoc Should header/footer scale when doc zoom is changed
     * @param {Object} opts.pageSetup
     * @param {Boolean} opts.pageSetup.blackAndWhite
     * @param {String} opts.pageSetup.cellComments one of 'none', 'asDisplayed', 'atEnd'
     * @param {Number} opts.pageSetup.copies How many copies to print
     * @param {Boolean} opts.pageSetup.draft Should quality be draft
     * @param {String} opts.pageSetup.errors One of 'displayed', 'blank', 'dash', 'NA'
     * @param {Number} opts.pageSetup.firstPageNumber Should the page number of the first page be printed
     * @param {Number} opts.pageSetup.fitToHeight Number of vertical pages to fit to
     * @param {Number} opts.pageSetup.fitToWidth Number of horizontal pages to fit to
     * @param {Number} opts.pageSetup.horizontalDpi 
     * @param {String} opts.pageSetup.orientation One of 'default', 'portrait', 'landscape'
     * @param {String} opts.pageSetup.pageOrder One of 'downThenOver', 'overThenDown'
     * @param {String} opts.pageSetup.paperHeight Value must a positive Float immediately followed by unit of measure from list mm, cm, in, pt, pc, pi. i.e. '10.5cm'
     * @param {String} opts.pageSetup.paperSize see lib/types/paperSize.js for all types and descriptions of types. setting paperSize overrides paperHeight and paperWidth settings
     * @param {String} opts.pageSetup.paperWidth Value must a positive Float immediately followed by unit of measure from list mm, cm, in, pt, pc, pi. i.e. '10.5cm'
     * @param {Number} opts.pageSetup.scale zoom of worksheet
     * @param {Boolean} opts.pageSetup.useFirstPageNumber
     * @param {Boolean} opts.pageSetup.usePrinterDefaults
     * @param {Number} opts.pageSetup.verticalDpi 
     * @param {Object} opts.sheetView 
     * @param {Object} opts.sheetView.pane 
     * @param {String} opts.sheetView.pane.activePane one of 'bottomLeft', 'bottomRight', 'topLeft', 'topRight'
     * @param {String} opts.sheetView.pane.state ne of 'split', 'frozen', 'frozenSplit'
     * @param {String} opts.sheetView.pane.topLeftCell Cell Reference i.e. 'A1'
     * @param {String} opts.sheetView.pane.xSplit Horizontal position of the split, in 1/20th of a point; 0 (zero) if none. If the pane is frozen, this value indicates the number of columns visible in the top pane.
     * @param {String} opts.sheetView.pane.ySplit Vertical position of the split, in 1/20th of a point; 0 (zero) if none. If the pane is frozen, this value indicates the number of rows visible in the left pane.
     * @param {Boolean} opts.sheetView.rightToLeft Flag indicating whether the sheet is in 'right to left' display mode. When in this mode, Column A is on the far right, Column B ;is one column left of Column A, and so on. Also, information in cells is displayed in the Right to Left format.
     * @param {Number} opts.sheetView.zoomScale  Defaults to 100
     * @param {Number} opts.sheetView.zoomScaleNormal Defaults to 100
     * @param {Number} opts.sheetView.zoomScalePageLayoutView Defaults to 100
     * @param {Object} opts.sheetFormat 
     * @param {Number} opts.sheetFormat.baseColWidth Defaults to 10. Specifies the number of characters of the maximum digit width of the normal style's font. This value does not include margin padding or extra padding for gridlines. It is only the number of characters.,
     * @param {Number} opts.sheetFormat.defaultColWidth
     * @param {Number} opts.sheetFormat.defaultRowHeight
     * @param {Boolean} opts.sheetFormat.thickBottom 'True' if rows have a thick bottom border by default.
     * @param {Boolean} opts.sheetFormat.thickTop 'True' if rows have a thick top border by default.
     * @param {Object} opts.sheetProtection same as "Protect Sheet" in Review tab of Excel 
     * @param {Boolean} opts.sheetProtection.autoFilter True means that that user will be unable to modify this setting
     * @param {Boolean} opts.sheetProtection.deleteColumns True means that that user will be unable to modify this setting
     * @param {Boolean} opts.sheetProtection.deleteRows True means that that user will be unable to modify this setting
     * @param {Boolean} opts.sheetProtection.formatCells True means that that user will be unable to modify this setting
     * @param {Boolean} opts.sheetProtection.formatColumns True means that that user will be unable to modify this setting
     * @param {Boolean} opts.sheetProtection.formatRows True means that that user will be unable to modify this setting
     * @param {Boolean} opts.sheetProtection.insertColumns True means that that user will be unable to modify this setting
     * @param {Boolean} opts.sheetProtection.insertHyperlinks True means that that user will be unable to modify this setting
     * @param {Boolean} opts.sheetProtection.insertRows True means that that user will be unable to modify this setting
     * @param {Boolean} opts.sheetProtection.objects True means that that user will be unable to modify this setting
     * @param {String} opts.sheetProtection.password Password used to protect sheet
     * @param {Boolean} opts.sheetProtection.pivotTables True means that that user will be unable to modify this setting
     * @param {Boolean} opts.sheetProtection.scenarios True means that that user will be unable to modify this setting
     * @param {Boolean} opts.sheetProtection.selectLockedCells True means that that user will be unable to modify this setting
     * @param {Boolean} opts.sheetProtection.selectUnlockedCells True means that that user will be unable to modify this setting
     * @param {Boolean} opts.sheetProtection.sheet True means that that user will be unable to modify this setting
     * @param {Boolean} opts.sheetProtection.sort True means that that user will be unable to modify this setting
     * @param {Object} opts.outline 
     * @param {Boolean} opts.outline.summaryBelow Flag indicating whether summary rows appear below detail in an outline, when applying an outline/grouping.
     * @param {Boolean} opts.outline.summaryRight Flag indicating whether summary columns appear to the right of detail in an outline, when applying an outline/grouping.
     * @returns {Worksheet}
     */
    function Worksheet(wb, name, opts) {
        _classCallCheck(this, Worksheet);

        this.wb = wb;
        this.sheetId = this.wb.sheets.length + 1;
        this.localSheetId = this.wb.sheets.length;
        this.opts = _.merge({}, _.cloneDeep(wsDefaultParams), opts);
        optsValidator(opts);

        this.opts.sheetView.tabSelected = this.sheetId === 1 ? 1 : 0;
        this.name = name ? name : 'Sheet ' + this.sheetId;
        this.hasGroupings = false;
        this.cols = {}; // Columns keyed by column, contains column properties
        this.rows = {}; // Rows keyed by row, contains row properties and array of cellRefs
        this.cells = {}; // Cells keyed by Excel ref
        this.mergedCells = [];
        this.lastUsedRow = 1;
        this.lastUsedCol = 1;

        // conditional formatting rules hashed by sqref
        this.cfRulesCollection = new CfRulesCollection();
        this.hyperlinkCollection = new HyperlinkCollection();
        this.dataValidationCollection = new DataValidation.DataValidationCollection();
        this.drawingCollection = new wsDrawing.DrawingCollection();
    }

    _createClass(Worksheet, [{
        key: 'addConditionalFormattingRule',


        /**
         * @func Worksheet.addConditionalFormattingRule
         * @param {String} sqref Text represetation of Cell range where the conditional formatting will take effect
         * @param {Object} options Options for conditional formatting
         * @param {String} options.type Type of conditional formatting
         * @param {String} options.priority Priority level for this rule
         * @param {String} options.formula Formula that returns nonzero or 0 value. If not 0 then rule will be applied
         * @param {Style} options.style Style that should be applied if rule passes
         * @returns {Worksheet}
         */
        value: function addConditionalFormattingRule(sqref, options) {
            var style = options.style || this.wb.Style();
            var dxf = this.wb.dxfCollection.add(style);
            delete options.style;
            options.dxfId = dxf.id;
            this.cfRulesCollection.add(sqref, options);
            return this;
        }
        /**
         * @func Worksheet.addDataValidation
         * @desc Add a data validation rule to the Worksheet
         * @param {Object} opts Options for Data Validation rule
         * @param {String} opts.sqref Required. Specifies range of cells to apply validate. i.e. "A1:A100"
         * @param {Boolean} opts.allowBlank Allows cells to be empty
         * @param {String} opts.errorStyle One of 'stop', 'warning', 'information'. You must specify an error string for this to take effect
         * @param {String} opts.error Message to show on error
         * @param {String} opts.errorTitle: String Title of message shown on error
         * @param {Boolean} opts.showErrorMessage Defaults to true if error or errorTitle is set
         * @param {String} opts.imeMode Restricts input to a specific set of characters. One of 'noControl', 'off', 'on', 'disabled', 'hiragana', 'fullKatakana', 'halfKatakana', 'fullAlpha', 'halfAlpha', 'fullHangul', 'halfHangul'
         * @param {String} opts.operator Must be one of 'between', 'notBetween', 'equal', 'notEqual', 'lessThan', 'lessThanOrEqual', 'greaterThan', 'greaterThanOrEqual'
         * @param {String} opts.prompt Message text of input prompt
         * @param {String} opts.promptTitle Title of input prompt
         * @param {Boolean} opts.showInputMessage Defaults to true if prompt or promptTitle is set
         * @param {Boolean} opts.showDropDown A boolean value indicating whether to display a dropdown combo box for a list type data validation.
         * @param {String} opts.type One of 'none', 'whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom'
         * @param {Array.String} opts.formulas Minimum count 1, maximum count 2. Rules for validation
         */

    }, {
        key: 'addDataValidation',
        value: function addDataValidation(opts) {
            var newValidation = this.dataValidationCollection.add(opts);
            return newValidation;
        }
        /**
         * @func Worksheet.generateRelsXML
         * @desc When Workbook is being built, generate the XML that will go into the Worksheet .rels file
         */

    }, {
        key: 'generateRelsXML',
        value: function generateRelsXML() {
            return xmlBuilder.relsXML(this);
        }
        /**
         * @func Worksheet.generateXML
         * @desc When Workbook is being built, generate the XML that will go into the Worksheet xml file 
         */

    }, {
        key: 'generateXML',
        value: function generateXML() {
            return xmlBuilder.sheetXML(this);
        }
    }, {
        key: 'row',
        value: function row(_row) {
            return rowAccessor(this, _row);
        }
    }, {
        key: 'column',
        value: function column(col) {
            return colAccessor(this, col);
        }
        /**
         * @func Worksheet.addImage
         * @param {Object} opts
         * @param {String} opts.path File system path of image
         * @param {String} opts.type Type of image. Currently only 'picture' is supported
         * @param {Object} opts.position Position object for image
         * @param {String} opts.position.type Type of positional anchor to use. One of 'absoluteAnchor', 'oneCellAnchor', 'twoCellAnchor'
         * @param {Object} opts.position.from Object containg position of top left corner of image.  Used with oneCellAnchor and twoCellAchor types
         * @param {Number} opts.position.from.col Left edge of image will align with left edge of this column
         * @param {String} opts.position.from.colOff Offset from left edge of column
         * @param {Number} opts.position.from.row Top edge of image will align with top edge of this row
         * @param {String} opts.position.from.rowOff Offset from top edge of row
         * @param {Object} opts.position.to Object containing position of bottom right corner of image
         * @param {Number} opts.position.to.col Right edge of image will align with Left edge of this column
         * @param {String} opts.position.to.colOff Offset of left edge of column
         * @param {Number} opts.position.to.row Bottom edge of image will align with Top edge of this row
         * @param {String} opts.position.to.rowOff Offset of top edge of row
         * @param {String} opts.position.x X position of top left corner of image. Used with absoluteAchor type
         * @param {String} opts.position.y Y position of top left corner of image
         */

    }, {
        key: 'addImage',
        value: function addImage(opts) {
            opts = opts ? opts : {};
            var mediaID = this.wb.mediaCollection.add(opts.path);
            var newImage = this.drawingCollection.add(opts);
            newImage.id = mediaID;

            return newImage;
        }
    }, {
        key: 'relationships',
        get: function get() {
            var rels = [];
            this.hyperlinkCollection.links.forEach(function (l) {
                rels.push(l);
            });
            if (!this.drawingCollection.isEmpty) {
                rels.push('drawing');
            }
            return rels;
        }
    }, {
        key: 'columnCount',
        get: function get() {
            return Math.max.apply(Math, Object.keys(this.cols));
        }
    }, {
        key: 'rowCount',
        get: function get() {
            return Math.max.apply(Math, Object.keys(this.rows));
        }
    }, {
        key: 'cell',
        get: function get() {
            return cellAccessor.bind(this);
        }
    }]);

    return Worksheet;
}();

module.exports = Worksheet;
//# sourceMappingURL=worksheet.js.map
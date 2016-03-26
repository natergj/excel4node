const _ = require('lodash');
const CfRulesCollection = require('./cf/cf_rules_collection');
const cellAccessor = require('../cell');
const rowAccessor = require('../row');
const colAccessor = require('../column');
const wsDefaultParams = require('./sheet_default_params.js');
const HyperlinkCollection = require('./classes/hyperlink.js');
const DataValidation = require('./classes/dataValidation.js');
const xmlBuilder = require('./builder.js');


/**
 * Class repesenting a WorkBook
 * @namespace WorkBook
 */
class WorkSheet {
    /**
     * Create a WorkSheet.
     * @param {Object} opts Workbook settings
     */
    constructor(wb, name, opts) {
        
        this.wb = wb;
        this.sheetId = this.wb.sheets.length + 1;
        this.opts = _.merge({}, wsDefaultParams, opts);
        this.opts.sheetView.tabSelected = this.sheetId === 1 ? 1 : 0;
        this.name = name ? name : `Sheet ${this.sheetId}`;
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

        this.wb.sheets.push(this);
    }

    addConditionalFormattingRule(sqref, options) {
        let style = options.style || this.wb.Style();
        let dxf = this.wb.dxfCollection.add(style);
        delete options.style;
        options.dxfId = dxf.id;
        this.cfRulesCollection.add(sqref, options);
        return this;
    }

    addDataValidation(opts) {
        let newValidation = this.dataValidationCollection.add(opts);
        return newValidation;
    }

    generateRelsXML() {
        this.wb.logger.debug('generateRelsXML called from instance');
        return xmlBuilder.relsXML(this);
    }

    generateXML() {
        this.wb.logger.debug('generateXML called from instance');
        return xmlBuilder.sheetXML(this);
    }

    Cell(row1, col1, row2, col2, isMerged) {
        return cellAccessor(this, row1, col1, row2, col2, isMerged);
    }

    Row(row) {
        return rowAccessor(this, row);
    }

    Column(col) {
        return colAccessor(this, col);
    }

}

module.exports = WorkSheet;
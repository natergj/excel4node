const _ = require('lodash');
const xml = require('xmlbuilder');
const CfRulesCollection = require('./cf/cf_rules_collection');
const logger = require('../logger.js');
const utils = require('../utils.js');
const cellAccessor = require('../cell');
const wsDefaultParams = require('./sheet_default_params.js');

// ------------------------------------------------------------------------------
// Private WorkSheet Functions
let _addSheetPr = (promiseObj) => {
    // §18.3.1.82 sheetPr (Sheet Properties)
    return new Promise((resolve, reject) => {
        let o = promiseObj.ws.opts.printOptions;

        // Check if any option that would require the sheetPr element to be added exists
        if (o.fitToHeight || o.fitToWidth || o.orientation || o.horizontalDpi || o.verticalDpi) {
            let ele = promiseObj.xml.ele('sheetPr');

            // §18.3.1.65 pageSetUpPr (Page Setup Properties)
            if (o.fitToHeight || o.fitToWidth) {
                ele.ele('pageSetUpPr').att('fitToPage', 1);
            }
        }
        resolve(promiseObj);
    });
};

let _addDimension = (promiseObj) => {
    // §18.3.1.35 dimension (Worksheet Dimensions)
    return new Promise((resolve, reject) => {
        let firstCell = 'A1';
        let lastCell = `${utils.getExcelAlpha(promiseObj.ws.lastUsedCol)}${promiseObj.ws.lastUsedRow}`;
        let ele = promiseObj.xml.ele('dimension');
        ele.att('ref', `${firstCell}:${lastCell}`);

        resolve(promiseObj);
    });
};

let _addSheetViews = (promiseObj) => {
    // §18.3.1.88 sheetViews (Sheet Views)
    return new Promise((resolve, reject) => {
        let o = promiseObj.ws.opts.sheetView;
        let ele = promiseObj.xml.ele('sheetViews');
        let tabSelected = promiseObj.ws.opts;
        let sv = ele.ele('sheetView')
        .att('tabSelected', o.tabSelected)
        .att('workbookViewId', o.workbookViewId)
        .att('rightToLeft', o.rightToLeft)
        .att('zoomScale', o.zoomScale)
        .att('zoomScaleNormal', o.zoomScaleNormal)
        .att('zoomScalePageLayoutView', o.zoomScalePageLayoutView);

        let modifiedPaneParams = [];
        Object.keys(o.pane).forEach((k) => {
            if (o.pane[k] !== null) {
                modifiedPaneParams.push(k);
            }
        });
        if (modifiedPaneParams.lenth > 0) {
            let pEle = ele.ele('pane');
            modifiedPaneParams.forEach((k) => {
                pEle.att(k, o.pane[k]);
            });
        }

        resolve(promiseObj);
    });
};

let _addSheetFormatPr = (promiseObj) => {
    // §18.3.1.81 sheetFormatPr (Sheet Format Properties)
    return new Promise((resolve, reject) => {
        let o = promiseObj.ws.opts.sheetFormat;
        let ele = promiseObj.xml.ele('sheetFormatPr');
        Object.keys(o).forEach((k) => {
            if (o[k] !== null) {
                ele.att(k, o[k]);
            } 
        });
        resolve(promiseObj);
    });
};

let _addCols = (promiseObj) => {
    // §18.3.1.17 cols (Column Information)
    return new Promise((resolve, reject) => {

        resolve(promiseObj);
    });
};

let _addSheetData = (promiseObj) => {
    // §18.3.1.80 sheetData (Sheet Data)
    return new Promise((resolve, reject) => {

        let ele = promiseObj.xml.ele('sheetData');
        let rows = Object.keys(promiseObj.ws.rows);
        
        let processNextRow = () => {
            let r = rows.shift();
            if (r) {
                let thisRow = promiseObj.ws.rows[r];
                thisRow.cellRefs.sort(utils.sortCellRefs);

                let firstCell = thisRow.cellRefs[0];
                let firstCol = utils.getExcelRowCol(firstCell).col;
                let lastCell = thisRow.cellRefs[thisRow.cellRefs.length - 1];
                let lastCol = utils.getExcelRowCol(lastCell).col;

                let rEle = ele.ele('row');
                // If defaultRowHeight !== 16, set customHeight attribute to 1 as stated in §18.3.1.81
                if (promiseObj.ws.opts.sheetFormat.defaultRowHeight !== 16) {
                    rEle.att('customHeight', '1');
                }
                
                rEle.att('r', r);
                rEle.att('spans', `${firstCol}:${lastCol}`);

                thisRow.cellRefs.forEach((c) => {
                    let thisCell = promiseObj.ws.cells[c];
                    let cEle = rEle.ele('c').att('r', thisCell.r).att('s', thisCell.s);
                    if (thisCell.t !== null) {
                        cEle.att('t', thisCell.t);
                    }
                    if (thisCell.f !== null) {
                        cEle.ele('f').txt(thisCell.f);
                    }
                    if (thisCell.v !== null) {
                        cEle.ele('v').txt(thisCell.v);
                    }
                });
                processNextRow();
            } else {
                resolve(promiseObj);
            }
        };
        processNextRow();

    });
};

let _addSheetProtection = (promiseObj) => {
    // §18.3.1.85 sheetProtection (Sheet Protection Options)
    return new Promise((resolve, reject) => {

        resolve(promiseObj);
    });
};

let _addAutoFilter = (promiseObj) => {
    // §18.3.1.2 autoFilter (AutoFilter Settings)
    return new Promise((resolve, reject) => {

        resolve(promiseObj);
    });
};

let _addMergeCells = (promiseObj) => {
    // §18.3.1.55 mergeCells (Merge Cells)
    return new Promise((resolve, reject) => {

        resolve(promiseObj);
    });
};

let _addConditionalFormatting = (promiseObj) => {
    // §18.3.1.18 conditionalFormatting (Conditional Formatting)
    return new Promise((resolve, reject) => {

        resolve(promiseObj);
    });
};

let _addHyperlinks = (promiseObj) => {
    // §18.3.1.48 hyperlinks (Hyperlinks)
    return new Promise((resolve, reject) => {

        resolve(promiseObj);
    });
};

let _addDataValidations = (promiseObj) => {
    // §18.3.1.33 dataValidations (Data Validations)
    return new Promise((resolve, reject) => {

        resolve(promiseObj);
    });
};

let _addPrintOptions = (promiseObj) => {
    // §18.3.1.70 printOptions (Print Options)
    return new Promise((resolve, reject) => {

        resolve(promiseObj);
    });
};

let _addPageMargins = (promiseObj) => {
    // §18.3.1.62 pageMargins (Page Margins)
    return new Promise((resolve, reject) => {

        resolve(promiseObj);
    });
};

let _addPageSetup = (promiseObj) => {
    // §18.3.1.63 pageSetup (Page Setup Settings)
    return new Promise((resolve, reject) => {

        resolve(promiseObj);
    });
};

let _addHeaderFooter = (promiseObj) => {
    // §18.3.1.46 headerFooter (Header Footer Settings)
    return new Promise((resolve, reject) => {

        resolve(promiseObj);
    });
};

let _addDrawing = (promiseObj) => {
    // §18.3.1.36 drawing (Drawing)
    return new Promise((resolve, reject) => {

        resolve(promiseObj);
    });
};


// ------------------------------------------------------------------------------


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
        this.cols = {};
        this.rows = {}; // Rows keyed by row, contains array of cellRefs
        this.cells = {}; // Cells keyed by Excel ref
        this.lastUsedRow = 1;
        this.lastUsedCol = 1;

        // conditional formatting rules hashed by sqref
        this.cfRulesCollection = new CfRulesCollection();

        this.wb.sheets.push(this);
    }

    generateXML() {
        return new Promise((resolve, reject) => {

            let wsXML = xml.create(
                'worksheet',
                {
                    'version': '1.0', 
                    'encoding': 'UTF-8', 
                    'standalone': true
                }
            )
            .att('mc:Ignorable', 'x14ac')
            .att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            .att('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
            .att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            .att('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');

            // Excel complains if specific elements on not in the correct order in the XML doc.
            // Elements must be added to the XML in this order
            //  - sheetPr
            //  - dimension
            //  - sheetViews
            //  - sheetFormatPr
            //  - cols
            //  - sheetData
            //  - sheetProtection
            //  - autoFilter
            //  - mergeCells
            //  - conditionalFormatting
            //  - hyperlinks
            //  - dataValidations
            //  - printOptions
            //  - pageMargins
            //  - pageSetup
            //  - headerFooter
            //  - drawing
            let promiseObj = { xml: wsXML, ws: this };
            _addSheetPr(promiseObj)
            .then(_addDimension)
            .then(_addSheetViews)
            .then(_addSheetFormatPr)
            .then(_addCols)
            .then(_addSheetData)
            .then(_addSheetProtection)
            .then(_addAutoFilter)
            .then(_addMergeCells)
            .then(_addConditionalFormatting)
            .then(_addHyperlinks)
            .then(_addDataValidations)
            .then(_addPrintOptions)
            .then(_addPageMargins)
            .then(_addPageSetup)
            .then(_addHeaderFooter)
            .then(_addDrawing)
            .then((promiseObj) => {
                resolve(promiseObj.xml.doc().end({ pretty: true, indent: '  ', newline: '\n' }));
            })
            .catch((e) => {
                console.error(e.stack);
            });


        });
    }

    Cell(row1, col1, row2, col2, isMerged) {
        return cellAccessor(this, row1, col1, row2, col2, isMerged);
    }

}

module.exports = WorkSheet;
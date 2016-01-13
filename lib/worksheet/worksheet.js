var xml = require('xmlbuilder');

var Image = require('../image');
var cellAccessor = require('../cell');
var columnAccessor = require('../column');
var rowAccessor = require('../row');
var utils = require('../utils');

var SheetStructureTemplate = require('./sheet_structure_template');
var CfRulesCollection = require('./cf/cf_rules_collection');

module.exports = WorkSheet;

// ------------------------------------------------------------------------------

// Excel complains if specific attributes on not in the correct order in the XML doc.
var EXCEL_XML_TAG_ORDER = [
    'sheetPr',
    'dimension',
    'sheetViews',
    'sheetFormatPr',
    'cols',
    'sheetData',
    'sheetProtection',
    'autoFilter',
    'mergeCells',
    'conditionalFormatting',
    'hyperlinks',
    'dataValidations',
    'printOptions',
    'pageMargins',
    'pageSetup',
    'drawing'
];

// ------------------------------------------------------------------------------

function WorkSheet(wb) {
    this.wb = wb;
    this.opts = {};
    this.name = '';
    this.hasGroupings = false;

    this.margins = {
        bottom: 1.0,
        footer: .5,
        header: .5,
        left: .75,
        right: .75,
        top: 1.0
    };

    this.printOptions = {
        centerHorizontal: false,
        centerVertical: true
    };

    this.sheetView = {
        workbookViewId: 0,
        rightToLeft: 0,
        zoomScale: 100,
        zoomScaleNormal: 100,
        zoomScalePageLayoutView: 100
    };

    // clone template
    this.sheet = JSON.parse(JSON.stringify(SheetStructureTemplate));

    // conditional formatting rules hashed by sqref
    this.cfRulesCollection = new CfRulesCollection();

    this.cols = {};
    this.rows = {};

    this.Settings = {
        Outline: {
            SummaryBelow: settings(this).outlineSummaryBelow
        }
    };
}

// ------------------------------------------------------------------------------

WorkSheet.prototype.Cell = cellAccessor;
WorkSheet.prototype.Row = rowAccessor;
WorkSheet.prototype.Column = columnAccessor;

// ------------------------------------------------------------------------------


WorkSheet.prototype.setName = function (name) {
    this.name = name;
};


WorkSheet.prototype.toXML = function () {
    var thisWS = this;
    var thisSheet = JSON.parse(JSON.stringify(thisWS.sheet));
    var sheetData = thisSheet.sheetData;
    var wsXML = xml.create('worksheet', { version: '1.0', encoding: 'UTF-8', standalone: true });

    // TODO can WorkSheet#toXML 'process' functions be stateless?
    processGroupings(thisWS, thisSheet);
    processColumns(thisWS, thisSheet);
    processRows(thisWS, thisSheet, sheetData);
    processCells(thisWS, thisSheet);
    processHyperlinks(thisWS, thisSheet);
    processDimension(thisWS, thisSheet);

    // XML tag order is important.
    // Start with sheet keys that start with '@'
    Object.keys(thisSheet).e4nForEach(function (k) {
        if (k.charAt(0) === '@') {
            var def = {};
            def[k] = thisSheet[k];
            wsXML.ele(def);
        }
    });

    // Now add the rest of the tags
    EXCEL_XML_TAG_ORDER.e4nForEach(function (k) {
        if (k === 'autoFilter' && thisSheet[k]) {
            thisSheet.sheetPr['@enableFormatConditionsCalculation'] = '1';
            thisSheet.sheetPr['@filterMode'] = '1';
        }
        if (k === 'conditionalFormatting') {
            thisWS.cfRulesCollection.getBuilderElements().forEach(function (builderElement) {
                wsXML.ele(builderElement);
            });
        }
        if (thisSheet[k]) {
            var def = {};
            def[k] = thisSheet[k];
            wsXML.ele(def);
        }
    });

    return wsXML.end({});
};


WorkSheet.prototype.addConditionalFormattingRule = function (sqref, options) {
    var style = options.style || this.wb.Style();
    var dxf = this.wb.dxfCollection.createFromStyle(style);
    delete options.style;
    options.dxfId = dxf.getId();
    this.cfRulesCollection.add(sqref, options);
    return this;
};

// -----------------------------------------------------------------------------

WorkSheet.prototype.Image = Image.Image;
WorkSheet.prototype.getCell = getCell;
WorkSheet.prototype.setWSOpts = setWSOpts;
WorkSheet.prototype.setValidation = setValidation;

var countValidation = 0;

function setValidation(data) {
    countValidation++;
    this.sheet.dataValidations = this.sheet.dataValidations || {};
    this.sheet.dataValidations['@count'] = countValidation;
    this.sheet.dataValidations['#list'] = this.sheet.dataValidations['#list'] || [];

    var dataValidation = {};

    Object.keys(data).e4nForEach(function (key) {
        if (key === 'formulas') {
            return;
        }
        dataValidation['@' + key] = data[key];
    });

    data.formulas.e4nForEach(function (f, i) {
        if (typeof f === 'number' || f.substr(0, 1) === '=') {
            f = f;
        } else {
            f = '"' + f + '"';
        }
        dataValidation['formula' + (i + 1)] = f;
    });
    this.sheet.dataValidations['#list'].push({
        dataValidation: dataValidation
    });
}

// TODO WorkSheet#getCell appears unused
function getCell(a, b) {
    var props = {};
    if (typeof(a) === 'string') {
        props = a.toExcelRowCol();
    } else if (typeof(a) === 'number' && typeof(b) === 'number') {
        props.row = a;
        props.col = b;
    } else {
        return undefined;
    }
    return thisWS.rows[props.row].cells[props.col];
}

function settings(thisWS) {
    var theseSettings = {};
    theseSettings.outlineSummaryBelow = function (val) {
        console.log('####################################################################################');
        console.log('# WorkSheet.Settings is deprecated and will be removed in version 1.0.0            #');
        console.log('# Create WorkBooks with opts paramater instead.                                    #');
        console.log('####################################################################################');
        val = val ? 1 : 0;
        thisWS.sheet.sheetPr.outlinePr['@summaryBelow'] = val;
    };
    return theseSettings;
}

function setWSOpts(opts) {
    var opts = opts ? opts : {};
    var thisWS = this;
    // Set Margins
    if (opts.margins) {
        this.margins.bottom = opts.margins.bottom ? opts.margins.bottom : 1.0;
        this.margins.footer = opts.margins.footer ? opts.margins.footer : .5;
        this.margins.header = opts.margins.header ? opts.margins.header : .5;
        this.margins.left = opts.margins.left ? opts.margins.left : .75;
        this.margins.right = opts.margins.right ? opts.margins.right : .75;
        this.margins.top = opts.margins.top ? opts.margins.top : 1.0;
    }
    Object.keys(this.margins).e4nForEach(function (k) {
        var margin = {};
        margin['@' + k] = thisWS.margins[k];
        thisWS.sheet.pageMargins.push(margin);
    });

    // Set Print Options
    if (opts.printOptions) {
        this.printOptions.centerHorizontal = opts.printOptions.centerHorizontal ? opts.printOptions.centerHorizontal : false;
        this.printOptions.centerVertical = opts.printOptions.centerVertical ? opts.printOptions.centerVertical : false;
    }
    this.sheet.printOptions.push({ '@horizontalCentered': this.printOptions.centerHorizontal ? 1 : 0 });
    this.sheet.printOptions.push({ '@verticalCentered': this.printOptions.centerVertical ? 1 : 0 });

    // Set Page View options
    var thisView = this.sheet.sheetViews[0].sheetView;
    if (opts.view) {
        if (parseInt(opts.view.zoom) !== opts.view.zoom) {
            console.log('invalid value for zoom. value must be an integer. value was %s', opts.view.zoom);
            opts.view.zoom = 100;
        }
        this.sheetView.zoomScale = opts.view.zoom ? opts.view.zoom : 100;
        this.sheetView.zoomScaleNormal = opts.view.zoom ? opts.view.zoom : 100;
        this.sheetView.zoomScalePageLayoutView = opts.view.zoom ? opts.view.zoom : 100;

        if (opts.view.rtl) {
            this.sheetView.rightToLeft = 1;
        }
    }

    // Set Outline Options
    if (opts.outline) {
        thisWS.sheet.sheetPr.outlinePr = {
            '@summaryBelow': opts.outline.summaryBelow === false ? 0 : 1
        };
    }

    // Set Page Setup
    if (opts.fitToPage) {
        this.sheet.sheetPr.pageSetUpPr = { '@fitToPage': 1 };
        this.sheet.pageSetup.push({ '@fitToHeight': opts.fitToPage.fitToHeight ? opts.fitToPage.fitToHeight : 1 });
        this.sheet.pageSetup.push({ '@fitToWidth': opts.fitToPage.fitToWidth ? opts.fitToPage.fitToWidth : 1 });
        this.sheet.pageSetup.push({ '@orientation': opts.fitToPage.orientation ? opts.fitToPage.orientation : 'portrait' });
        this.sheet.pageSetup.push({ '@horizontalDpi': opts.fitToPage.horizontalDpi ? opts.fitToPage.horizontalDpi : 4294967292 });
        this.sheet.pageSetup.push({ '@verticalDpi': opts.fitToPage.verticalDpi ? opts.fitToPage.verticalDpi : 4294967292 });
    }

    // Set WorkSheet protections
    if (opts.sheetProtection) {
        thisWS.sheet.sheetProtection = {
            '@autoFilter': (opts.sheetProtection.autoFilter ? opts.sheetProtection.autoFilter : false),
            '@deleteColumns': (opts.sheetProtection.deleteColumns ? opts.sheetProtection.deleteColumns : false),
            '@deleteRows': (opts.sheetProtection.deleteRows ? opts.sheetProtection.deleteRows : false),
            '@formatCells': (opts.sheetProtection.formatCells ? opts.sheetProtection.formatCells : false),
            '@formatColumns': (opts.sheetProtection.formatColumns ? opts.sheetProtection.formatColumns : false),
            '@formatRows': (opts.sheetProtection.formatRows ? opts.sheetProtection.formatRows : false),
            '@insertColumns': (opts.sheetProtection.insertColumns ? opts.sheetProtection.insertColumns : false),
            '@insertHyperlinks': (opts.sheetProtection.insertHyperlinks ? opts.sheetProtection.insertHyperlinks : false),
            '@insertRows': (opts.sheetProtection.insertRows ? opts.sheetProtection.insertRows : false),
            '@objects': (opts.sheetProtection.objects ? opts.sheetProtection.objects : false),
            '@pivotTables': (opts.sheetProtection.pivotTables ? opts.sheetProtection.pivotTables : false),
            '@scenarios': (opts.sheetProtection.scenarios ? opts.sheetProtection.scenarios : false),
            '@selectLockedCells': (opts.sheetProtection.selectLockedCells ? opts.sheetProtection.selectLockedCells : false),
            '@selectUnlockedCells': (opts.sheetProtection.selectUnlockedCells ? opts.sheetProtection.selectUnlockedCells : false),
            '@sheet': (opts.sheetProtection.sheet ? opts.sheetProtection.sheet : true),
            '@sort': (opts.sheetProtection.sort ? opts.sheetProtection.sort : false)
        };
        if (opts.sheetProtection.password) {
            thisWS.sheet.sheetProtection['@password'] = utils.getHashOfPassword(opts.sheetProtection.password);
        }
    }

    thisView.push({ '@workbookViewId': this.sheetView.workbookViewId ? this.sheetView.workbookViewId : 0 });
    thisView.push({ '@zoomScale': this.sheetView.zoomScale ? this.sheetView.zoomScale : 100 });
    thisView.push({ '@zoomScaleNormal': this.sheetView.zoomScaleNormal ? this.sheetView.zoomScaleNormal : 100 });
    thisView.push({ '@zoomScalePageLayoutView': this.sheetView.zoomScalePageLayoutView ? this.sheetView.zoomScalePageLayoutView : 100 });
    thisView.push({ '@rightToLeft': this.sheetView.rightToLeft ? 1 : 0 });
}


// Process groupings and add collapsed attributes to rows where applicable
// TODO break apart Worksheet processGroupings()
function processGroupings(thisWS, thisSheet) {
    if (!thisWS.hasGroupings) {
        return;
    }
    var lastRowNum = Object.keys(thisWS.rows).sort()[Object.keys(thisWS.rows).length - 1];
    var lastColNum = Object.keys(thisWS.cols).sort()[Object.keys(thisWS.cols).length - 1];

    var rOutlineLevels = {
        curHighestLevel: 0,
        0: {
            startRow: 1,
            endRow: 1,
            isHidden: 0
        }
    };

    var cOutlineLevels = {
        curHighestLevel: 0,
        0: {
            startCol: 1,
            endCol: 1,
            isHidden: 0
        }
    };

    var summaryBelow = parseInt(thisSheet.sheetPr.outlinePr['@summaryBelow']) === 0 ? false : true;
    if (summaryBelow && thisWS.rows[lastRowNum].attributes.hidden) {
        thisWS.Row(parseInt(lastRowNum) + 1);
    }
    if (summaryBelow && thisWS.cols[lastColNum].hidden) {
        thisWS.Column(parseInt(lastColNum) + 1);
    }

    Object.keys(thisWS.rows).e4nForEach(function (rNum, i) {
        var rID = parseInt(rNum);
        var curRow = thisWS.rows[rNum];
        var thisLevel = curRow.attributes.outlineLevel ? curRow.attributes.outlineLevel : 0;
        var isHidden = curRow.attributes.hidden === 1 ? curRow.attributes.hidden : 0;
        var rowNum = curRow.attributes.r;

        rOutlineLevels[0].endRow = rID;
        rOutlineLevels[0].isHidden = isHidden;

        if (typeof(rOutlineLevels[thisLevel]) === 'undefined') {
            rOutlineLevels[thisLevel] = {
                startRow: rID,
                endRow: rID,
                isHidden: isHidden
            };
        }

        if (thisLevel <= rOutlineLevels.curHighestLevel) {
            rOutlineLevels[thisLevel].endRow = rID;
            rOutlineLevels[thisLevel].isHidden = isHidden;
        }

        if (thisLevel !== rOutlineLevels.curHighestLevel || rID === lastRowNum) {
            if (summaryBelow && thisLevel !== rOutlineLevels.curHighestLevel) {
                if (rID === lastRowNum) {
                    thisLevel = 1;
                }
                var oLi;
                for (oLi = rOutlineLevels.curHighestLevel; oLi > thisLevel; oLi--) {
                    if (rOutlineLevels[oLi]) {
                        var rowToCollapse = rOutlineLevels[oLi].endRow + 1;
                        var lastRow = thisWS.Row(rowToCollapse);
                        lastRow.setAttribute('collapsed', rOutlineLevels[oLi].isHidden);
                        delete rOutlineLevels[oLi];
                    }
                }
            } else if (!summaryBelow && thisLevel !== rOutlineLevels.curHighestLevel) {
                if (rOutlineLevels[thisLevel]) {
                    if (thisLevel > rOutlineLevels.curHighestLevel) {
                        var rowToCollapse = rOutlineLevels[rOutlineLevels.curHighestLevel].startRow;
                    } else {
                        var rowToCollapse = rOutlineLevels[thisLevel].startRow;
                    }
                    var lastRow = thisWS.Row(rowToCollapse);
                    lastRow.setAttribute('collapsed', rOutlineLevels[thisLevel].isHidden);
                    rOutlineLevels[thisLevel].startRow = rowNum;
                }
            }
        }
        if (thisLevel !== rOutlineLevels.curHighestLevel) {
            rOutlineLevels.curHighestLevel = thisLevel;
        }
    });

    Object.keys(thisWS.cols).e4nForEach(function (cNum, i) {
        var cID = parseInt(cNum);
        var curCol = thisWS.cols[cNum];
        var thisLevel = curCol.outlineLevel ? curCol.outlineLevel : 0;
        var isHidden = curCol.hidden === 1 ? curCol.hidden : 0;
        var colNum = curCol.min;

        cOutlineLevels[0].endCol = cID;
        cOutlineLevels[0].isHidden = isHidden;

        if (typeof(cOutlineLevels[thisLevel]) === 'undefined') {
            cOutlineLevels[thisLevel] = {
                startCol: cID,
                endCol: cID,
                isHidden: isHidden
            };
        }

        if (thisLevel <= cOutlineLevels.curHighestLevel) {
            cOutlineLevels[thisLevel].endCol = cID;
            cOutlineLevels[thisLevel].isHidden = isHidden;
        }

        if (thisLevel !== cOutlineLevels.curHighestLevel || cID === lastColNum) {
            if (summaryBelow && thisLevel !== cOutlineLevels.curHighestLevel) {
                if (cID === lastColNum) {
                    thisLevel = 1;
                }
                var oLi;
                for (oLi = cOutlineLevels.curHighestLevel; oLi > thisLevel; oLi--) {
                    if (cOutlineLevels[oLi]) {
                        var colToCollapse = cOutlineLevels[oLi].endCol + 1;
                        var lastCol = thisWS.Column(colToCollapse);
                        lastCol.setAttribute('collapsed', cOutlineLevels[oLi].isHidden);
                        delete cOutlineLevels[oLi];
                    }
                }
            } else if (!summaryBelow && thisLevel !== cOutlineLevels.curHighestLevel) {
                if (cOutlineLevels[thisLevel]) {
                    if (thisLevel > cOutlineLevels.curHighestLevel) {
                        var colToCollapse = cOutlineLevels[cOutlineLevels.curHighestLevel].startCol;
                    } else {
                        var colToCollapse = cOutlineLevels[thisLevel].startCol;
                    }
                    var lastCol = thisWS.Column(colToCollapse);
                    lastCol.setAttribute('collapsed', cOutlineLevels[thisLevel].isHidden);
                    cOutlineLevels[thisLevel].startCol = colNum;
                }
            }
        }
        if (thisLevel !== cOutlineLevels.curHighestLevel) {
            cOutlineLevels.curHighestLevel = thisLevel;
        }
    });
}

function processColumns(thisWS, thisSheet) {
    if (Object.keys(thisWS.cols).length > 0) {
        if (thisSheet.cols instanceof Array === false) {
            thisSheet.cols = [];
        }

        Object.keys(thisWS.cols).e4nForEach(function (i) {
            var c = thisWS.cols[i];
            var thisCol = { col: [] };
            Object.keys(c).e4nForEach(function (k) {
                var tmpObj = {};
                if (typeof(c[k]) !== 'object') {
                    tmpObj['@' + k] = c[k];
                    thisCol.col.push(tmpObj);
                }
            });
            thisSheet.cols.push(thisCol);
        });
    }
}

function processRows(thisWS, thisSheet, sheetData) {
    Object.keys(thisWS.rows).e4nForEach(function (r, i) {
        var thisRow = { row: [] };
        Object.keys(thisWS.rows[r].attributes).e4nForEach(function (a, i) {
            var attr = '@' + a;
            var obj = {};
            obj[attr] = thisWS.rows[r].attributes[a];
            thisRow.row.push(obj);
        });
        Object.keys(thisWS.rows[r].cells).e4nForEach(function (c, i) {
            var thisCellIndex = thisRow.row.push({ 'c': {} });
            var thisCell = thisRow.row[thisCellIndex - 1]['c'];
            Object.keys(thisWS.rows[r].cells[c].attributes).e4nForEach(function (a, i) {
                thisCell['@' + a] = thisWS.rows[r].cells[c].attributes[a];
            });
            Object.keys(thisWS.rows[r].cells[c].children).e4nForEach(function (v, i) {
                thisCell[v] = thisWS.rows[r].cells[c].children[v];
            });
        });
        sheetData.push(thisRow);
    });
}

function processCells(thisWS, thisSheet) {
    if (thisWS.mergeCells && thisWS.mergeCells.length >= 0) {
        thisSheet.mergeCells = [];
        thisSheet.mergeCells.push({ '@count': thisWS.mergeCells.length });
        thisWS.mergeCells.e4nForEach(function (cr) {
            thisSheet.mergeCells.push({ 'mergeCell': { '@ref': cr } });
        });
    }
}

function processHyperlinks(thisWS, thisSheet) {
    if (thisWS.hyperlinks && thisWS.hyperlinks.length >= 0) {
        thisSheet.hyperlinks = [];
        thisWS.hyperlinks.e4nForEach(function (cr) {
            thisSheet.hyperlinks.push({ 'hyperlink': { '@ref': cr.ref, '@r:id': 'rId' + cr.id } });
        });
    }
}

function processDimension(thisWS, thisSheet) {
    var rowCount = Object.keys(thisWS.rows).length;
    var colCount = Object.keys(thisWS.cols).length;
    if (rowCount && colCount) {
        thisSheet['dimension'] = [{}];
        thisSheet['dimension'][0]['@ref'] = 'A1:' + colCount.toExcelAlpha() + rowCount;
    }
}


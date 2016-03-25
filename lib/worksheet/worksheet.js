var xml = require('xmlbuilder');
var image = require('../image');
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
	'headerFooter',
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
	var wsXML = xml.create('worksheet', {version: '1.0', encoding: 'UTF-8', standalone: true});
	// TODO can WorkSheet#toXML 'process' functions be stateless?
	processGroupings(thisWS, thisSheet);
	processColumns(thisWS, thisSheet);
	processRows(thisWS, thisSheet, sheetData);
	processCells(thisWS, thisSheet);
	processHyperlinks(thisWS, thisSheet);
	processDimension(thisWS, thisSheet);
	// XML tag order is important.
	// Start with sheet keys that start with '@'
	Object.keys(thisSheet).forEach(function (k) {
		if (k.charAt(0) === '@') {
			var def = {};
			def[k] = thisSheet[k];
			wsXML.ele(def);
		}
	});
	// Now add the rest of the tags
	EXCEL_XML_TAG_ORDER.forEach(function (k) {
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

WorkSheet.prototype.Image = image.Image;
WorkSheet.prototype.getCell = getCell;
WorkSheet.prototype.setWSOpts = setWSOpts;
WorkSheet.prototype.setValidation = setValidation;
WorkSheet.prototype.printArea = printArea;
WorkSheet.prototype.printTitles = printTitles;
WorkSheet.prototype.printScaling = printScaling;
WorkSheet.prototype.headerFooter = headerFooter;
var countValidation = 0;
function setValidation(data) {
	countValidation++;
	this.sheet.dataValidations = this.sheet.dataValidations || {};
	this.sheet.dataValidations['@count'] = countValidation;
	this.sheet.dataValidations['#list'] = this.sheet.dataValidations['#list'] || [];
	var dataValidation = {};
	Object.keys(data).forEach(function (key) {
		if (key === 'formulas') {
			return;
		}
		dataValidation['@' + key] = data[key];
	});
	data.formulas.forEach(function (f, i) {
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
	if (typeof (a) === 'string') {
		props = a.toExcelRowCol();
	} else if (typeof (a) === 'number' && typeof (b) === 'number') {
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

/*
 https://support.office.com/en-US/article/Set-a-specific-print-area-BEEBCEB7-0D43-4E07-8895-5AFE0AEDFB32

 {
 rows: {
 begin: 1,
 end: 2
 },
 columns: {
 begin: 1,
 end: 3
 }
 }
 */
function printArea(opts) {
	this.printOptions.printArea = opts;
}

/*
 https://support.office.com/en-us/article/Repeat-specific-rows-or-columns-on-every-printed-page-0d6dac43-7ee7-4f34-8b08-ffcc8b022409

 {
 rows: {
 begin: 1,
 end: 2
 },
 columns: {
 begin: 1,
 end: 3
 }
 }
 */
function printTitles(opts) {
	this.printOptions.printTitles = opts;
}

/*
 https://poi.apache.org/apidocs/org/apache/poi/xssf/usermodel/extensions/XSSFHeaderFooter.html
 Example: &L&A&C&BCompany Name. Confidential&B&RPage &P of &N
 (&L-from left, &A-print a sheet tab name, C&B -in center bold a company name, &R&P-right side add page number)


 sheet.headerFooter({
 oddHeader: '&LDavid Gofman&R&D',
 oddFooter: '&L&A&C&BCompany, Inc. Confidential&B&RPage &P of &N'
 });
 */
function headerFooter(opts) {
	var headerFooter = {};
	this.sheet.headerFooter = headerFooter;
	headerFooter['@differentOddEven'] = !!opts.differentOddEven;
	headerFooter['@differentFirst'] = !!opts.differentFirst;
	headerFooter['@scaleWithDoc'] = opts.scaleWithDoc === undefined ? true : opts.scaleWithDoc;
	headerFooter['@alignWithMargins'] = opts.alignWithMargins === undefined ? true : opts.alignWithMargins;
	headerFooter.oddHeader = opts.oddHeader;
	headerFooter.oddFooter = opts.oddFooter;
	headerFooter.evenHeader = opts.evenHeader;
	headerFooter.evenFooter = opts.evenFooter;
	headerFooter.firstHeader = opts.firstHeader;
	headerFooter.firstFooter = opts.firstFooter;
	return headerFooter;
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
	Object.keys(this.margins).forEach(function (k) {
		var margin = {};
		margin['@' + k] = thisWS.margins[k];
		thisWS.sheet.pageMargins.push(margin);
	});
	// Set Print Options
	if (opts.printOptions) {
		this.printOptions.centerHorizontal = opts.printOptions.centerHorizontal ? opts.printOptions.centerHorizontal : false;
		this.printOptions.centerVertical = opts.printOptions.centerVertical ? opts.printOptions.centerVertical : false;
	}
	this.sheet.printOptions.push({'@horizontalCentered': this.printOptions.centerHorizontal ? 1 : 0});
	this.sheet.printOptions.push({'@verticalCentered': this.printOptions.centerVertical ? 1 : 0});
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
		this.sheet.sheetPr.pageSetUpPr = {'@fitToPage': 1};
		this.sheet.pageSetup.push({'@fitToHeight': opts.fitToPage.fitToHeight ? opts.fitToPage.fitToHeight : 1});
		this.sheet.pageSetup.push({'@fitToWidth': opts.fitToPage.fitToWidth ? opts.fitToPage.fitToWidth : 1});
		this.sheet.pageSetup.push({'@orientation': opts.fitToPage.orientation ? opts.fitToPage.orientation : 'portrait'});
		this.sheet.pageSetup.push({'@horizontalDpi': opts.fitToPage.horizontalDpi ? opts.fitToPage.horizontalDpi : 4294967292});
		this.sheet.pageSetup.push({'@verticalDpi': opts.fitToPage.verticalDpi ? opts.fitToPage.verticalDpi : 4294967292});
		opts.orientation = opts.fitToPage.orientation;
	} else if (opts.orientation) {
		this.sheet.pageSetup.push({'@orientation': opts.orientation ? opts.orientation : 'portrait'});
	}

// Set paper width/height or paper size
// According to the OOXML specification, paper width/height overrule paper size
// See chapter 18.3.1.64 the ECMA 376 standard (http://www.ecma-international.org/publications/standards/Ecma-376.htm)
	if (opts.paperDimensions) {
		if (opts.paperDimensions.paperWidth && opts.paperDimensions.paperHeight) {
			this.sheet.pageSetup.push({'@paperWidth': opts.paperDimensions.paperWidth});
			this.sheet.pageSetup.push({'@paperHeight': opts.paperDimensions.paperHeight});
		} else if (opts.paperDimensions.paperSize) {
			if (1 <= opts.paperDimensions.paperSize && opts.paperDimensions.paperSize <= 68) {
				this.sheet.pageSetup.push({'@paperSize': opts.paperDimensions.paperSize});
			} else {
				console.log('Legal paper sizes are (see WorkSheet.PaperSize):\n' + WorkSheet.PaperSize);
			}
		}
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

	thisView.push({'@workbookViewId': this.sheetView.workbookViewId ? this.sheetView.workbookViewId : 0});
	thisView.push({'@zoomScale': this.sheetView.zoomScale ? this.sheetView.zoomScale : 100});
	thisView.push({'@zoomScaleNormal': this.sheetView.zoomScaleNormal ? this.sheetView.zoomScaleNormal : 100});
	thisView.push({'@zoomScalePageLayoutView': this.sheetView.zoomScalePageLayoutView ? this.sheetView.zoomScalePageLayoutView : 100});
	thisView.push({'@rightToLeft': this.sheetView.rightToLeft ? 1 : 0});
	this.WSOpts = opts;
}

/*
 opts - NO_SCALING | FIT_ONE_PAGE | FIT_COLUMNS | FIT_ALL_ROWS | CUSTOM_SCALING

 or

 opt - {
 scale: NO_SCALING | FIT_ONE_PAGE | FIT_COLUMNS | FIT_ALL_ROWS | CUSTOM_SCALING,
 horizontalDpi: 4294967292,
 verticalDpi: 4294967292
 }
 */
function printScaling(opts) {
	var print = this.wb.Print;
	this.sheet.sheetPr.pageSetUpPr = {};
	this.sheet.pageSetup = [];
	if (typeof opts === 'number') {
		opts = {scale: opts || print.NO_SCALING};
	}

	if (opts && opts.scale !== print.NO_SCALING) {
		switch (opts.scale) {
			case print.FIT_ONE_PAGE:
				delete opts.fitToHeight;
				delete opts.fitToWidth;
				break;
			case print.FIT_ALL_COLUMNS:
				opts.fitToHeight = 0;
				delete opts.fitToWidth;
				break;
			case print.FIT_ALL_ROWS:
				opts.fitToWidth = 0;
				delete opts.fitToHeight;
				break;
		}

		this.sheet.sheetPr.pageSetUpPr['@fitToPage'] = 1;
		if (opts.fitToHeight !== undefined) {
			this.sheet.pageSetup.push({'@fitToHeight': opts.fitToHeight});
		}
		if (opts.fitToWidth !== undefined) {
			this.sheet.pageSetup.push({'@fitToWidth': opts.fitToWidth});
		}
		this.sheet.pageSetup.push({'@horizontalDpi': opts.horizontalDpi ? opts.horizontalDpi : 4294967292});
		this.sheet.pageSetup.push({'@verticalDpi': opts.verticalDpi ? opts.verticalDpi : 4294967292});
	}

	if (this.WSOpts && this.WSOpts.orientation) {
		this.sheet.pageSetup.push({'@orientation': this.WSOpts.orientation ? this.WSOpts.orientation : 'portrait'});
	}

	return opts;
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

	Object.keys(thisWS.rows).forEach(function (rNum, i) {
		var rID = parseInt(rNum);
		var curRow = thisWS.rows[rNum];
		var thisLevel = curRow.attributes.outlineLevel ? curRow.attributes.outlineLevel : 0;
		var isHidden = curRow.attributes.hidden === 1 ? curRow.attributes.hidden : 0;
		var rowNum = curRow.attributes.r;
		rOutlineLevels[0].endRow = rID;
		rOutlineLevels[0].isHidden = isHidden;
		if (typeof (rOutlineLevels[thisLevel]) === 'undefined') {
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
	Object.keys(thisWS.cols).forEach(function (cNum, i) {
		var cID = parseInt(cNum);
		var curCol = thisWS.cols[cNum];
		var thisLevel = curCol.outlineLevel ? curCol.outlineLevel : 0;
		var isHidden = curCol.hidden === 1 ? curCol.hidden : 0;
		var colNum = curCol.min;
		cOutlineLevels[0].endCol = cID;
		cOutlineLevels[0].isHidden = isHidden;
		if (typeof (cOutlineLevels[thisLevel]) === 'undefined') {
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

		Object.keys(thisWS.cols).forEach(function (i) {
			var c = thisWS.cols[i];
			var thisCol = {col: []};
			Object.keys(c).forEach(function (k) {
				var tmpObj = {};
				if (typeof (c[k]) !== 'object') {
					tmpObj['@' + k] = c[k];
					thisCol.col.push(tmpObj);
				}
			});
			thisSheet.cols.push(thisCol);
		});
	}
}

function processRows(thisWS, thisSheet, sheetData) {
	Object.keys(thisWS.rows).forEach(function (r, i) {
		var thisRow = {row: []};
		Object.keys(thisWS.rows[r].attributes).forEach(function (a, i) {
			var attr = '@' + a;
			var obj = {};
			obj[attr] = thisWS.rows[r].attributes[a];
			thisRow.row.push(obj);
		});
		Object.keys(thisWS.rows[r].cells).forEach(function (c, i) {
			var thisCellIndex = thisRow.row.push({'c': {}});
			var thisCell = thisRow.row[thisCellIndex - 1]['c'];
			Object.keys(thisWS.rows[r].cells[c].attributes).forEach(function (a, i) {
				thisCell['@' + a] = thisWS.rows[r].cells[c].attributes[a];
			});
			Object.keys(thisWS.rows[r].cells[c].children).forEach(function (v, i) {
				thisCell[v] = thisWS.rows[r].cells[c].children[v];
			});
		});
		sheetData.push(thisRow);
	});
}

function processCells(thisWS, thisSheet) {
	if (thisWS.mergeCells && thisWS.mergeCells.length >= 0) {
		thisSheet.mergeCells = [];
		thisSheet.mergeCells.push({'@count': thisWS.mergeCells.length});
		thisWS.mergeCells.forEach(function (cr) {
			thisSheet.mergeCells.push({'mergeCell': {'@ref': cr}});
		});
	}
}

function processHyperlinks(thisWS, thisSheet) {
	if (thisWS.hyperlinks && thisWS.hyperlinks.length >= 0) {
		thisSheet.hyperlinks = [];
		thisWS.hyperlinks.forEach(function (cr) {
			thisSheet.hyperlinks.push({'hyperlink': {'@ref': cr.ref, '@r:id': 'rId' + cr.id}});
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

WorkSheet.prototype.PaperSize = {
	LETTER_PAPER: 1, // Letter paper (8.5 in. by 11 in.)
	LETTER_SMALL_PAPER: 2, // Letter small paper (8.5 in. by 11 in.)
	TABLOID_PAPER: 3, // Tabloid paper (11 in. by 17 in.)
	LEDGER_PAPER: 4, // Ledger paper (17 in. by 11 in.)
	LEGAL_PAPER: 5, // Legal paper (8.5 in. by 14 in.)
	STATEMENT_PAPER: 6, // Statement paper (5.5 in. by 8.5 in.)
	EXECUTIVE_PAPER: 7, // Executive paper (7.25 in. by 10.5 in.)
	A3_PAPER: 8, // A3 paper (297 mm by 420 mm)
	A4_PAPER: 9, // A4 paper (210 mm by 297 mm)
	A4_SMALL_PAPER: 10, // A4 small paper (210 mm by 297 mm)
	A5_PAPER: 11, // A5 paper (148 mm by 210 mm)
	B4_PAPER: 12, // B4 paper (250 mm by 353 mm)
	B5_PAPER: 13, // B5 paper (176 mm by 250 mm)
	FOLIO_PAPER: 14, // Folio paper (8.5 in. by 13 in.)
	QUARTO_PAPER: 15, // Quarto paper (215 mm by 275 mm)
	STANDARD_PAPER_10_BY_14_IN: 16, // Standard paper (10 in. by 14 in.)
	STANDARD_PAPER_11_BY_17_IN: 17, // Standard paper (11 in. by 17 in.)
	NOTE_PAPER: 18, // Note paper (8.5 in. by 11 in.)
	NUMBER_9_ENVELOPE: 19, // #9 envelope (3.875 in. by 8.875 in.)
	NUMBER_10_ENVELOPE: 20, // #10 envelope (4.125 in. by 9.5 in.)
	NUMBER_11_ENVELOPE: 21, // #11 envelope (4.5 in. by 10.375 in.)
	NUMBER_12_ENVELOPE: 22, // #12 envelope (4.75 in. by 11 in.)
	NUMBER_14_ENVELOPE: 23, // #14 envelope (5 in. by 11.5 in.)
	C_PAPER: 24, // C paper (17 in. by 22 in.)
	D_PAPER: 25, // D paper (22 in. by 34 in.)
	E_PAPER: 26, // E paper (34 in. by 44 in.)
	DL_PAPER: 27, // DL envelope (110 mm by 220 mm)
	C5_ENVELOPE: 28, // C5 envelope (162 mm by 229 mm)
	C3_ENVELOPE: 29, // C3 envelope (324 mm by 458 mm)
	C4_ENVELOPE: 30, // C4 envelope (229 mm by 324 mm)
	C6_ENVELOPE: 31, // C6 envelope (114 mm by 162 mm)
	C65_ENVELOPE: 32, // C65 envelope (114 mm by 229 mm)
	B4_ENVELOPE: 33, // B4 envelope (250 mm by 353 mm)
	B5_ENVELOPE: 34, // B5 envelope (176 mm by 250 mm)
	B6_ENVELOPE: 35, // B6 envelope (176 mm by 125 mm)
	ITALY_ENVELOPE: 36, // Italy envelope (110 mm by 230 mm)
	MONARCH_ENVELOPE: 37, // Monarch envelope (3.875 in. by 7.5 in.).
	SIX_THREE_QUARTERS_ENVELOPE: 38, // 6 3/4 envelope (3.625 in. by 6.5 in.)
	US_STANDARD_FANFOLD: 39, // US standard fanfold (14.875 in. by 11 in.)
	GERMAN_STANDARD_FANFOLD: 40, // German standard fanfold (8.5 in. by 12 in.)
	GERMAN_LEGAL_FANFOLD: 41, // German legal fanfold (8.5 in. by 13 in.)
	ISO_B4: 42, // ISO B4 (250 mm by 353 mm)
	JAPANESE_DOUBLE_POSTCARD: 43, // Japanese double postcard (200 mm by 148 mm)
	STANDARD_PAPER_9_BY_11_IN: 44, // Standard paper (9 in. by 11 in.)
	STANDARD_PAPER_10_BY_11_IN: 45, // Standard paper (10 in. by 11 in.)
	STANDARD_PAPER_15_BY_11_IN: 46, // Standard paper (15 in. by 11 in.)
	INVITE_ENVELOPE: 47, // Invite envelope (220 mm by 220 mm)
	LETTER_EXTRA_PAPER: 50, // Letter extra paper (9.275 in. by 12 in.)
	LEGAL_EXTRA_PAPER: 51, // Legal extra paper (9.275 in. by 15 in.)
	TABLOID_EXTRA_PAPER: 52, // Tabloid extra paper (11.69 in. by 18 in.)
	A4_EXTRA_PAPER: 53, // A4 extra paper (236 mm by 322 mm)
	LETTER_TRANSVERSE_PAPER: 54, // Letter transverse paper (8.275 in. by 11 in.)
	A4_TRANSVERSE_PAPER: 55, // A4 transverse paper (210 mm by 297 mm)
	LETTER_EXTRA_TRANSVERSE_PAPER: 56, // Letter extra transverse paper (9.275 in. by 12 in.)
	SUPER_A_SUPER_A_A4_PAPER: 57, // SuperA/SuperA/A4 paper (227 mm by 356 mm)
	SUPER_B_SUPER_B_A3_PAPER: 58, // SuperB/SuperB/A3 paper (305 mm by 487 mm)
	LETTER_PLUS_PAPER: 59, // Letter plus paper (8.5 in. by 12.69 in.)
	A4_PLUS_PAPER: 60, // A4 plus paper (210 mm by 330 mm)
	A5_TRANSVERSE_PAPER: 61, // A5 transverse paper (148 mm by 210 mm)
	JIS_B5_TRANSVERSE_PAPER: 62, // JIS B5 transverse paper (182 mm by 257 mm)
	A3_EXTRA_PAPER: 63, // A3 extra paper (322 mm by 445 mm)
	A5_EXTRA_PAPER: 64, // A5 extra paper (174 mm by 235 mm)
	ISO_B5_EXTRA_PAPER: 65, // ISO B5 extra paper (201 mm by 276 mm)
	A2_PAPER: 66, // A2 paper (420 mm by 594 mm)
	A3_TRANSVERSE_PAPER: 67, // A3 transverse paper (297 mm by 420 mm)
	A3_EXTRA_TRANSVERSE_PAPER: 68 // A3 extra transverse paper (322 mm by 445 mm)
};
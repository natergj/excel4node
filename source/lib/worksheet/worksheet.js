let _ 					= require('lodash');
let xml 				= require('xmlbuilder');
let CfRulesCollection 	= require('./cf/cf_rules_collection');
let logger 				= require('../logger.js');
let utils 				= require('../utils.js');
let cellAccessor 		= require('../cell');

// ------------------------------------------------------------------------------
// Default Options for New WorkSheets
let sheetOpts = {
	'margins' : {
		'bottom'					: 0.75,
		'footer'					: 0.3,
		'header'					: 0.3,
		'left'						: 0.7,
		'right'						: 0.7,
		'top'						: 0.75
	},
	'printOptions' : {
        'centerHorizontal'			: false,
        'centerVertical'			: true,
		'fitToHeight' 				: null, // (Optional) Max number of pages high
		'fitToWidth' 				: null, // (Optional) Max number of pages wide
		'orientation' 				: null, // (Optional) 'potrait' or 'landscape'
		'horizontalDpi' 			: null, // (Optional) standard is 4294967292
		'verticalDpi' 				: null  // (Optional) standard is 4294967292
	
    },
    'sheetView' : {
		'workbookViewId'			: 0,
		'rightToLeft'				: 0,
		'zoomScale'					: 100,
		'zoomScaleNormal'			: 100,
		'zoomScalePageLayoutView'	: 100
	},
	'outline' : {
        'summaryBelow'				: false
    }
};

// ------------------------------------------------------------------------------
// Private WorkSheet Functions
let _addSheetPr = (promiseObj) => {
	// §18.3.1.82 sheetPr (Sheet Properties)
	return new Promise((resolve, reject) => {
		let o = promiseObj.ws.opts.printOptions;

		// Check if any option that would require the sheetPr element to be added exists
		if(o.fitToHeight || o.fitToWidth || o.orientation || o.horizontalDpi || o.verticalDpi){
			let ele = promiseObj.xml.ele('sheetPr');

			// §18.3.1.65 pageSetUpPr (Page Setup Properties)
			if(o.fitToHeight || o.fitToWidth) {
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

		resolve(promiseObj);
	});
};

let _addSheetFormatPr = (promiseObj) => {
	// §18.3.1.81 sheetFormatPr (Sheet Format Properties)
	return new Promise((resolve, reject) => {

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
			if(r){
				let thisRow = promiseObj.ws.rows[r];
				thisRow.cellRefs.sort(utils.sortCellRefs);

				let firstCell = thisRow.cellRefs[0];
				let firstCol = utils.getExcelRowCol(firstCell).col;
				let lastCell = thisRow.cellRefs[thisRow.cellRefs.length - 1];
				let lastCol = utils.getExcelRowCol(lastCell).col;

				let rEle = ele.ele('row');
				rEle.att('r', r);
				rEle.att('spans', `${firstCol}:${lastCol}`);
				thisRow.cellRefs.forEach((c) => {
					let thisCell = promiseObj.ws.cells[c];
					let cEle = rEle.ele('c').att('r', thisCell.r).att('s', thisCell.s);
					if(thisCell.t !== null){
						cEle.att('t', thisCell.t);
					}
					if(thisCell.f !== null){
						cEle.ele('f').txt(thisCell.f);
					}
					if(thisCell.v !== null){
						cEle.ele('v').txt(thisCell.v);
					}
				});
				processNextRow();
			} else {
				resolve(promiseObj);
			}
		}
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
	constructor( wb, name, opts ) {
		this.wb = wb;
		this.opts = _.merge({}, sheetOpts, opts);
	    this.name = name ? name : `Sheet ${wb.sheets.length + 1}`;
	    this.hasGroupings = false;
	    this.cols = {};
	    this.rows = {}; // Rows keyed by row, contains array of cellRefs
	    this.cells = {}; // Cells keyed by Excel ref
	    this.lastUsedRow = 1;
	    this.lastUsedCol = 1;
	    this.paneData = {}; // For use with §18.3.1.88 sheetViews (Sheet Views)

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
			let promiseObj = {xml: wsXML, ws: this};
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
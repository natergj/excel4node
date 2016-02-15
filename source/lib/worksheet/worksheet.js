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
let _addSheetPrEleToXML = (promiseObj) => {
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

let _addDimensionEleToXML = (promiseObj) => {
	// §18.3.1.35 dimension (Worksheet Dimensions)
	return new Promise((resolve, reject) => {
		let firstCell = 'A1';
		let lastCell = `${utils.getExcelAlpha(promiseObj.ws.lastUsedCol)}${promiseObj.ws.lastUsedRow}`;
		let ele = promiseObj.xml.ele('dimension');
		ele.att('ref', `${firstCell}:${lastCell}`);

		resolve(promiseObj);
	});
};

let _addSheetViewsEleToXML = (promiseObj) => {
	// §18.3.1.88 sheetViews (Sheet Views)
	return new Promise((resolve, reject) => {

		resolve(promiseObj);
	});
};

let _addSheetFormatPrEleToXML = (promiseObj) => {
	// §18.3.1.81 sheetFormatPr (Sheet Format Properties)
	return new Promise((resolve, reject) => {

		resolve(promiseObj);
	});
};

let _addColsEleToXML = (promiseObj) => {
	// §18.3.1.17 cols (Column Information)
	return new Promise((resolve, reject) => {

		resolve(promiseObj);
	});
};

let _addSheetDataEleToXML = (promiseObj) => {
	// §18.3.1.80 sheetData (Sheet Data)
	return new Promise((resolve, reject) => {

		let ele = promiseObj.xml.ele('sheetData');
		let rows = Object.keys(promiseObj.ws.rows);
		
		resolve(promiseObj);
	});
};

let _addSheetProtectionEleToXML = (promiseObj) => {
	// §18.3.1.85 sheetProtection (Sheet Protection Options)
	return new Promise((resolve, reject) => {

		resolve(promiseObj);
	});
};

let _addAutoFilterEleToXML = (promiseObj) => {
	// §18.3.1.2 autoFilter (AutoFilter Settings)
	return new Promise((resolve, reject) => {

		resolve(promiseObj);
	});
};

let _addMergeCellsEleToXML = (promiseObj) => {
	// §18.3.1.55 mergeCells (Merge Cells)
	return new Promise((resolve, reject) => {

		resolve(promiseObj);
	});
};

let _addConditionalFormattingEleToXML = (promiseObj) => {
	// §18.3.1.18 conditionalFormatting (Conditional Formatting)
	return new Promise((resolve, reject) => {

		resolve(promiseObj);
	});
};

let _addHyperlinksEleToXML = (promiseObj) => {
	// §18.3.1.48 hyperlinks (Hyperlinks)
	return new Promise((resolve, reject) => {

		resolve(promiseObj);
	});
};

let _addDataValidationsEleToXML = (promiseObj) => {
	// §18.3.1.33 dataValidations (Data Validations)
	return new Promise((resolve, reject) => {

		resolve(promiseObj);
	});
};

let _addPrintOptionsEleToXML = (promiseObj) => {
	// §18.3.1.70 printOptions (Print Options)
	return new Promise((resolve, reject) => {

		resolve(promiseObj);
	});
};

let _addPageMarginsEleToXML = (promiseObj) => {
	// §18.3.1.62 pageMargins (Page Margins)
	return new Promise((resolve, reject) => {

		resolve(promiseObj);
	});
};

let _addPageSetupEleToXML = (promiseObj) => {
	// §18.3.1.63 pageSetup (Page Setup Settings)
	return new Promise((resolve, reject) => {

		resolve(promiseObj);
	});
};

let _addHeaderFooterEleToXML = (promiseObj) => {
	// §18.3.1.46 headerFooter (Header Footer Settings)
	return new Promise((resolve, reject) => {

		resolve(promiseObj);
	});
};

let _addDrawingEleToXML = (promiseObj) => {
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
			_addSheetPrEleToXML(promiseObj)
			.then(_addDimensionEleToXML)
			.then(_addSheetViewsEleToXML)
			.then(_addSheetFormatPrEleToXML)
			.then(_addColsEleToXML)
			.then(_addSheetDataEleToXML)
			.then(_addSheetProtectionEleToXML)
			.then(_addAutoFilterEleToXML)
			.then(_addMergeCellsEleToXML)
			.then(_addConditionalFormattingEleToXML)
			.then(_addHyperlinksEleToXML)
			.then(_addDataValidationsEleToXML)
			.then(_addPrintOptionsEleToXML)
			.then(_addPageMarginsEleToXML)
			.then(_addPageSetupEleToXML)
			.then(_addHeaderFooterEleToXML)
			.then(_addDrawingEleToXML)
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
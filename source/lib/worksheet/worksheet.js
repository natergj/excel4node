let _ 					= require('lodash');
let Promise 			= require('bluebird');
let xml 				= require('xmlbuilder');
let CfRulesCollection 	= require('./cf/cf_rules_collection');
let logger 				= require('../logger.js');

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
	'pageSetup' : {
		'fitToHeight' 				: null, // (Optional) Max number of pages high
		'fitToWidth' 				: null, // (Optional) Max number of pages wide
		'orientation' 				: null, // (Optional) 'potrait' or 'landscape'
		'horizontalDpi' 			: null, // (Optional) standard is 4294967292
		'verticalDpi' 				: null  // (Optional) standard is 4294967292
	},
	'printOptions' : {
        'centerHorizontal'			: false,
        'centerVertical'			: true
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
	return new Promise((resolve, reject) => {
		try{
			let o = promiseObj.ws.opts.pageSetup;
			// Check if any option that would require the sheetPr element to be added exists
			if(o.fitToHeight || o.fitToWidth || o.orientation || o.horizontalDpi || o.verticalDpi){
				let ele = promiseObj.xml.ele('sheetPr');

				if(o.fitToHeight || o.fitToWidth) {
					ele.att('fitToPage', 1);
				}
			}

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addDimensionEleToXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addSheetViewsEleToXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addSheetFormatPrEleToXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addColsEleToXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addSheetDataEleToXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addSheetProtectionEleToXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addAutoFilterEleToXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addMergeCellsEleToXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addConditionalFormattingEleToXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addHyperlinksEleToXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addDataValidationsEleToXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addPrintOptionsEleToXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addPageMarginsEleToXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addPageSetupEleToXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addHeaderFooterEleToXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addDrawingEleToXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(promiseObj);
		}
		catch(e){
			reject(e);
		}
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
		logger.debug('New WorkSheet created');
		this.wb = wb;
		this.opts = _.merge({}, sheetOpts, opts);
	    this.name = name ? name : `Sheet ${wb.sheets.length + 1}`;
	    this.hasGroupings = false;
	    this.cols = {};
	    this.rows = {};

	    // conditional formatting rules hashed by sqref
	    this.cfRulesCollection = new CfRulesCollection();

	    this.wb.sheets.push(this);
	}

	generateXML() {
		return new Promise((resolve, reject) => {
			try {
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

			} 
			catch(e) {
				logger.error(e.stack);
				reject(e);
			}
		});
	}
}

module.exports = WorkSheet;
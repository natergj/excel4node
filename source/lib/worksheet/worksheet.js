let _ 					= require('lodash');
let Promise 			= require('bluebird');
let xml 				= require('xmlbuilder');
let CfRulesCollection 	= require('./cf/cf_rules_collection');
let logger 				= require('../logger.js');

// ------------------------------------------------------------------------------
// Default Options for New WorkSheets
let sheetDefaultOpts = {
	'margins' : {
		'bottom'					: 1.0,
		'footer'					: 0.5,
		'header'					: 0.5,
		'left'						: 0.75,
		'right'						: 0.75,
		'top'						: 1.0
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

let _addSheetPrEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addDimensionEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addSheetViewsEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addSheetFormatPrEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addColsEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addSheetDataEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addSheetProtectionEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addAutoFilterEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addMergeCellsEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addConditionalFormattingEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addHyperlinksEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addDataValidationsEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addPrintOptionsEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addPageMarginsEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addPageSetupEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addHeaderFooterEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
		}
		catch(e){
			reject(e);
		}
	});
};

let _addDrawingEleToXML = (wbXML) => {
	return new Promise((resolve, reject) => {
		try{

			resolve(wbXML);
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
		this.opts = _.merge({}, sheetDefaultOpts, opts);
	    this.name = name ? name : `Sheet ${wb.sheets.length + 1}`;
	    this.hasGroupings = false;
	    this.cols = {};
	    this.rows = {};

	    // conditional formatting rules hashed by sqref
	    this.cfRulesCollection = new CfRulesCollection();

	    this.wb.sheets.push(this);
	}

	toXML() {

		logger.debug('Called WorkSheet.toXML');
		let wsXML = xml.create(
			'worksheet',
			{
				'version': '1.0', 
				'encoding': 'UTF-8', 
				'standalone': true
			},
			{
				'mc:Ignorable':'x14ac',
				'xmlns':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
				'xmlns:mc':'http://schemas.openxmlformats.org/markup-compatibility/2006',
				'xmlns:r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
				'xmlns:x14ac':'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'
			}
		);

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
		_addSheetPrEleToXML(wsXML)
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
		.then((wsXML) => {
			console.log('xml generated');
			return wsXML.doc().end();
		})
		.catch((e) => {
			console.error(e.stack);
		});

	}
}

module.exports = WorkSheet;
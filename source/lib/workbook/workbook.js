let _ 			= require('lodash');
let JSZip 		= require('jszip');
let fs 			= require('fs');
let Promise 	= require('bluebird');
let xml 		= require('xmlbuilder');
let WorkSheet 	= require('../worksheet');
let logger 		= require('../logger.js');

// ------------------------------------------------------------------------------
// Private WorkBook Methods Start

let _addWorkSheetsXML = (promiseObj) => {
	return new Promise((resolve, reject) => {
		try {
			let curSheet = 0;

			let processNextSheet = () => {
				let thisSheet = promiseObj.wb.sheets[curSheet];
				if(thisSheet){
					thisSheet
					.generateXML()
					.then((xml) => {
						// Add worksheet to zip
						logger.debug(xml);
						curSheet++;
						processNextSheet();
					});
				} else {
					resolve(promiseObj);
				}
			};
			processNextSheet();
		} catch(e) {
			reject(e);
		}
	});
};

/**
 * Generate XML for SharedStrings.xml file and add it to zip file. Called from _writeToBuffer()
 * @private
 * @memberof WorkBook
 * @param {Object} promiseObj object containing jszip instance, workbook intance and xmlvars
 * @return {Promise} Resolves with promiseObj
 */
let _addSharedStringsXMLToZip = (promiseObj) => {
	return new Promise( (resolve, reject) => {
		try{
			let stringObj = {
				'sst' : [{
                    '@count': 0,
                    '@uniqueCount': 0,
                    '@xmlns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
                }]
			};

		    promiseObj.wb.sharedStrings.forEach(function (s) {
		        if (s instanceof Array) {
		            var si = [],
		                lastrPr = {}; //Remeber last font name, size and color
		            s.forEach(function (i) {
		                if (typeof(i) === 'string') {
		                    si.push({ 'r': { 'rPr': JSON.parse(JSON.stringify(lastrPr)), 't': { '@xml:space': 'preserve', '#text': i } } });
		                } else {
		                    var rPr = {};
		                    if (i.size !== undefined) {
		                        lastrPr.sz = rPr.sz = { '@val': i.size };
		                    } else if (lastrPr.sz) {
		                        rPr.sz = lastrPr.sz;
		                    }
		                    if (i.font !== undefined) {
		                        lastrPr.rFont = rPr.rFont = { '@val': i.font };
		                    } else if (lastrPr.rFont) {
		                        rPr.rFont = lastrPr.rFont;
		                    }
		                    if (i.color !== undefined) {
		                        lastrPr.color = rPr.color = { '@rgb': i.color };
		                    } else  if (lastrPr.color) {
		                        rPr.color = lastrPr.color;
		                    }
		                    if (i.bold) {
		                        rPr.b = {};
		                    }
		                    if (i.underline) {
		                        rPr.u = {};
		                    }
		                    if (i.italic) {
		                        rPr.i = {};
		                    }
		                    if (i.value !== undefined) {
		                        si.push({ 'r': { 'rPr': rPr, 't': { '@xml:space': 'preserve', '#text': i.value } } });
		                    }
		                }
		            });
		            if (si.length) {
		                stringObj.sst.push({ 'si': si });
		            }
		        } else {
		            stringObj.sst.push({ 'si': { 't': s } });
		        }
		    });

		    stringObj.sst[0]['@uniqueCount'] = promiseObj.wb.sharedStrings.length;

			promiseObj.xlsx.folder('xl').file('sharedStrings.xml', xml.create(stringObj).end(promiseObj.xmlOutVars));

			resolve(promiseObj);
		} catch(e) {
			reject(e);
		}
	});
};

/**
 * Use JSZip to generate file to a node buffer
 * @private
 * @memberof WorkBook
 * @param {WorkBook} wb WorkBook instance
 * @return {Promise} resolves with Buffer 
 */
let _writeToBuffer = (wb) => {
	return new Promise((resolve, reject) => {
		try {
			let promiseObj = {
				wb : wb, 
				xlsx : new JSZip(),
				xmlOutVars : {}
			};
			if(promiseObj.wb.sheets.length === 0){
				promiseObj.wb.sheets.push(promiseObj.wb.WorkSheet('Sheet 1'));
			}

			_addWorkSheetsXML(promiseObj)
			.then(_addSharedStringsXMLToZip)
			.then(() => {
				let buffer = promiseObj.xlsx.generate({
			    	type: 'nodebuffer',
			    	compression: wb.opts.jszip.compression
				});	
				resolve(buffer);
			})
			.catch((e) => {
				console.error(e.stack);
			});

		} catch(e) {
			reject(e);
		}
	});
};

// Private WorkBook Methods End
// ------------------------------------------------------------------------------


// Default Options for WorkBook
let workBookDefaultOpts = {
	jszip : {
		compression : 'DEFLATE'
	}
};

/**
 * Class repesenting a WorkBook
 * @namespace WorkBook
 */
class WorkBook {

	/**
	 * Create a WorkBook.
	 * @param {Object} opts Workbook settings
	 */
	constructor(opts) {

		this.opts = _.merge({}, workBookDefaultOpts, opts);
		this.sheets = [];
		this.sharedStrings = [];
		this.styles = [];

	}

	/**
	 * Generate .xlsx file. 
	 * @param {String} fileName Name of Excel workbook with .xslx extension
	 * @param {http.response | callback} http response object or callback function (optional). 
	 * If http response object is given, file is written to http response. Useful for web applications.
	 * If callback is given, callback called with (err, fs.Stats) passed
	 */
	write(fileName, handler) {
		_writeToBuffer(this)
		.then((buffer) => {

		    switch (typeof handler) {
		    	// handler passed as http response object. 
		    	case 'object':
			        handler.writeHead(200, {
			            'Content-Length': buffer.length,
			            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			            'Content-Disposition': 'attachment; filename="' + fileName + '"'
			        });
			        handler.end(buffer);
		    	break;

		    	// handler passed as callback function
		    	case 'function':
			        fs.writeFile(fileName, buffer, function (err) {
			            handler(err);
			        });
		    	break;

		    	// no handler passed, write file to FS.
		    	default:
			        fs.writeFile(fileName, buffer, function (err) {
			            if (err) { throw err; }
			        });
		    	break;
		    }
		})
		.catch((e) => {
			console.error(e.stack);
		});
	}

	WorkSheet(name, opts) {
		return new WorkSheet(this, name, opts);
	}
}

module.exports = WorkBook;
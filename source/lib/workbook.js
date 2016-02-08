let _ 			= require('lodash');
let JSZip 		= require('jszip');
let fs 			= require('fs');
let BbPromise 	= require('bluebird');
let xml 		= require('xmlbuilder');

/**
 * Generate XML for SharedStrings.xml file and add it to zip file
 * @private
 * @memberof WorkBook
 * @param {Object} promiseObj object containing jszip instance, workbook intance and xmlvars
 * @return {Promise} Resolves with promiseObj
 */
let addSharedStringsXMLToZip = (promiseObj) => {
	return new BbPromise( (resolve, reject) => {
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
 * @param {WorkBook} Workbook instance
 * @return {Promise} resolves with Buffer 
 */
let writeToBuffer = (WorkBook) => {
	return new BbPromise((resolve, reject) => {
		try {
			let promiseObj = {
				wb : WorkBook, 
				xlsx : new JSZip(),
				xmlOutVars : {}
			};

			addSharedStringsXMLToZip(promiseObj)
			.then(() => {
				let buffer = promiseObj.xlsx.generate({
			    	type: 'nodebuffer',
			    	compression: WorkBook.opts.jszip.compression
				});	
				resolve(buffer);
			})
			.catch((e) => {
				console.error(e);
			});

		} catch(e) {
			reject(e);
		}
	});
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
		let defaultOpts = {
			jszip : {
				compression : 'DEFLATE'
			}
		};

		this.opts = _.merge({}, defaultOpts, opts);
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
	write(fileName, httpResponse) {
		console.log('write file');
		writeToBuffer(this)
		.then((buffer) => {
		    // If `httpResponse` is an object (a node httpResponse object)
		    if (typeof httpResponse === 'object') {
		        httpResponse.writeHead(200, {
		            'Content-Length': buffer.length,
		            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		            'Content-Disposition': 'attachment; filename="' + fileName + '"'
		        });
		        httpResponse.end(buffer);
		    // Else if `httpResponse` is a function, use it as a callback
		    } else if (typeof httpResponse === 'function') {
		        fs.writeFile(fileName, buffer, function (err) {
		            httpResponse(err);
		        });
		    // Else httpResponse wasn't specified
		    } else {
		        fs.writeFile(fileName, buffer, function (err) {
		            if (err) { throw err; }
		        });
		    }
		})
		.catch((e) => {
			console.error(e);
		});
	}



	/**
	 * Generate JSON representation of WorkBook. Used in debugging.
	 * @return {String} WorkBook instance in JSON format
	 */
	toString() {
		console.log(JSON.stringify(this, null, '\t'));
	}
}

module.exports = WorkBook;
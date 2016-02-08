let _ 		= require('lodash');
let JSZip 	= require('jszip');
let fs 		= require('fs');

/**
 * Class repesenting a WorkBook
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
		let buffer = this.writeToBuffer();

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
	}

	/**
	 * Use JSZip to generate file to a node buffer
	 * @return {NodeBuffer} JSZip generated node buffer 
	 */
	writeToBuffer() {
		let xlsx = new JSZip();

		return xlsx.generate({
        	type: 'nodebuffer',
        	compression: this.opts.jszip.compression
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
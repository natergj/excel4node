let _ = require('lodash');

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
	 * @return {Object} Data about output file.
	 */
	write() {
		console.log('write file');
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
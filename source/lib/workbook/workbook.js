let _ 				= require('lodash');
let fs 				= require('fs');
let JSZip 			= require('jszip');
let logger 			= require('../logger.js');
let utils 			= require('../utils.js');
let WorkSheet 		= require('../worksheet');
let xmlbuilder 		= require('xmlbuilder');

// ------------------------------------------------------------------------------
// Private WorkBook Methods Start


let _addRootContentTypesXML = (promiseObj) => {
	// Required as stated in §12.2
	return new Promise((resolve, reject) => {
		let xml = xmlbuilder.create(
			'Types',
			{
				'version': '1.0', 
				'encoding': 'UTF-8', 
				'standalone': true
			}
		)
		.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');

		xml.ele('Default').att('ContentType', 'application/xml').att('Extension', 'xml');
		xml.ele('Default').att('ContentType', 'application/vnd.openxmlformats-package.relationships+xml').att('Extension', 'rels');
		xml.ele('Override').att('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml').att('PartName', '/xl/workbook.xml');
		promiseObj.wb.sheets.forEach((s, i) => {
			xml.ele('Override')
			.att('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
			.att('PartName', `/xl/worksheets/sheet${i+1}.xml`);
		});
		xml.ele('Override').att('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml').att('PartName', '/xl/styles.xml');
		xml.ele('Override').att('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml').att('PartName', '/xl/sharedStrings.xml');

		let xmlString = xml.doc().end(promiseObj.xmlOutVars);
		promiseObj.xlsx.file('[Content_Types].xml', xmlString);
		resolve(promiseObj);
	});
};

let _addRootRelsXML = (promiseObj) => {
	// Required as stated in §12.2
	return new Promise((resolve, reject) => {
		let xml = xmlbuilder.create(
			'Relationships',
			{
				'version': '1.0', 
				'encoding': 'UTF-8', 
				'standalone': true
			}
		)
		.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

		xml
		.ele('Relationship')
		.att('Id', 'rId1')
		.att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument')
		.att('Target', 'xl/workbook.xml');

		let xmlString = xml.doc().end(promiseObj.xmlOutVars);
		promiseObj.xlsx.folder('_rels').file('.rels', xmlString);
		resolve(promiseObj);

	});
};

let _addWorkBookXML = (promiseObj) => {
	// Required as stated in §12.2
	return new Promise((resolve, reject) => {

		let xml = xmlbuilder.create(
			'workbook',
			{
				'version': '1.0', 
				'encoding': 'UTF-8', 
				'standalone': true
			}
		)
		xml.att('mc:Ignorable', 'x15');
		xml.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
		xml.att('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
		xml.att('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
		xml.att('xmlns:x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main');

		let sheetsEle = xml.ele('sheets');
		promiseObj.wb.sheets.forEach((s, i) => {
			sheetsEle.ele('sheet')
			.att('name', s.name)
			.att('sheetId', i+1)
			.att('r:id', `rId${i+1}`);
		});

		let xmlString = xml.doc().end(promiseObj.xmlOutVars);
		promiseObj.xlsx.folder('xl').file('workbook.xml', xmlString);
		resolve(promiseObj);

	});
};

let _addWorkBookRelsXML = (promiseObj) => {
	// Required as stated in §12.2
	return new Promise((resolve, reject) => {

		let xml = xmlbuilder.create(
			'Relationships',
			{
				'version': '1.0', 
				'encoding': 'UTF-8', 
				'standalone': true
			}
		)
		.att('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');

		xml
		.ele('Relationship')
        .att('Id', `rId${promiseObj.wb.sheets.length+1}`)
        .att('Target', 'sharedStrings.xml')
        .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings');

		xml
		.ele('Relationship')
        .att('Id', `rId${promiseObj.wb.sheets.length+2}`)
        .att('Target', 'styles.xml')
        .att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles');

        promiseObj.wb.sheets.forEach((s, i) => {
        	xml
        	.ele('Relationship')
        	.att('Id', `rId${i+1}`)
        	.att('Target', `worksheets/sheet${i+1}.xml`)
        	.att('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
        });

        let xmlString = xml.doc().end(promiseObj.xmlOutVars);
		promiseObj.xlsx.folder('xl').folder('_rels').file('workbook.xml.rels', xmlString);
		resolve(promiseObj);

	});
};

let _addWorkSheetsXML = (promiseObj) => {
	// Required as stated in §12.2
	return new Promise((resolve, reject) => {

		let curSheet = 0;
		
		let processNextSheet = () => {
			let thisSheet = promiseObj.wb.sheets[curSheet];
			if(thisSheet){
				thisSheet
				.generateXML()
				.then((xml) => {
					// Add worksheet to zip
					curSheet++;
					promiseObj.xlsx.folder('xl').folder('worksheets').file(`sheet${curSheet}.xml`, xml);
					processNextSheet();
				});
			} else {
				resolve(promiseObj);
			}
		};
		processNextSheet();

	});
};

/**
 * Generate XML for SharedStrings.xml file and add it to zip file. Called from _writeToBuffer()
 * @private
 * @memberof WorkBook
 * @param {Object} promiseObj object containing jszip instance, workbook intance and xmlvars
 * @return {Promise} Resolves with promiseObj
 */
let _addSharedStringsXML = (promiseObj) => {
	// §12.3.15 Shared String Table Part
	return new Promise( (resolve, reject) => {

		let xml = xmlbuilder.create(
			'sst',
			{
				'version': '1.0', 
				'encoding': 'UTF-8', 
				'standalone': true
			}
		)
		.att('count', promiseObj.wb.sharedStrings.length)
		.att('uniqueCount', promiseObj.wb.sharedStrings.length)
		.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');

	    promiseObj.wb.sharedStrings.forEach((s) => {
	    	xml.ele('si').ele('t').txt(s);
	    });

	    let xmlString = xml.doc().end(promiseObj.xmlOutVars);
		promiseObj.xlsx.folder('xl').file('sharedStrings.xml', xmlString);

		resolve(promiseObj);

	});
};

let _addStylesXML = (promiseObj) => {
	// §12.3.20 Styles Part
	return new Promise((resolve, reject) => {
		let xml = xmlbuilder.create(
			'styleSheet',
			{
				'version': '1.0', 
				'encoding': 'UTF-8', 
				'standalone': true
			}
		)
		.att('mc:Ignorable', 'x14ac')
		.att('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
		.att('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
		.att('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');

		if(promiseObj.wb.styleData.numFmts.length > 0){
			let nfXML = xml
			.ele('numFmts')
			.att('count', promiseObj.wb.styleData.numFmts.length);
			promiseObj.wb.styleData.numFmts.forEach((nf) => {
				nfXML
				.ele('numFmt')
				.att('formatCode', nf.format)
				.att('numFmtId', nf.id);
			});
		}

		let fontXML = xml
		.ele('fonts')
		.att('count', promiseObj.wb.styleData.fonts.length);
		promiseObj.wb.styleData.fonts.forEach((f) => {
			let fEle = fontXML.ele('font');
			fEle.ele('sz').att('val', f.sz);
			fEle.ele('color').att('rgb', f.color);
			fEle.ele('name').att('val', f.name);
			fEle.ele('family').att('val', f.family);
			fEle.ele('scheme').att('val', f.scheme);
		});

		let fillXML = xml 
		.ele('fills')
		.att('count', promiseObj.wb.styleData.fills.length);
		promiseObj.wb.styleData.fills.forEach((f) => {
			let fXML = fillXML.ele('fill');
			let pFill = fXML.ele('patternFill').att('patternType', f.patternFill.patternType);
			if(f.patternFill.fgColor){
				pFill.ele('fgColor').att('rgb', f.patternFill.fgColor);
			}
			if(f.patternFill.bgColor){
				pFill.ele('bgColor').att('rgb', f.patternFill.bgColor);
			}

		});

		let borderXML = xml 
		.ele('borders')
		.att('count', promiseObj.wb.styleData.borders.length);
		promiseObj.wb.styleData.borders.forEach((b) => {
			let bXML = borderXML.ele('border');

			['left','right','top','bottom','diagonal'].forEach((o) => {
				let thisOEle = bXML.ele(o);
				if(b[o]){
					if(b[o].style){
						thisOEle.att('style', b[o].style);
					}
					if(b[o].color){
						thisOEle.ele('color').att('rgb', b[o].color);
					}
				}
			});
		});


		let cellXfsXML = xml 
		.ele('cellXfs')
		.att('count', promiseObj.wb.styleData.cellXfs.length);
		promiseObj.wb.styleData.cellXfs.forEach((c) => {
			let thisEle = cellXfsXML.ele('xf');
			Object.keys(c).forEach((a) => {
				thisEle.att(a, c[a]);
			});
		});

	    let xmlString = xml.doc().end(promiseObj.xmlOutVars);
		promiseObj.xlsx.folder('xl').file('styles.xml', xmlString);

		resolve(promiseObj);
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

		let promiseObj = {
			wb : wb, 
			xlsx : new JSZip(),
			xmlOutVars : { pretty: true, indent: '  ', newline: '\n' }
			//xmlOutVars : {}
		};


		if(promiseObj.wb.sheets.length === 0){
			promiseObj.wb.WorkSheet();
		}

		_addRootContentTypesXML(promiseObj)
		.then(_addRootRelsXML)
		.then(_addWorkBookXML)
		.then(_addWorkBookRelsXML)
		.then(_addWorkSheetsXML)
		.then(_addSharedStringsXML)
		.then(_addStylesXML)
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

	});
};

// Private WorkBook Methods End
// ------------------------------------------------------------------------------

/* Available options for WorkBook
{
	jszip : {
		compression : 'DEFLATE'
	},
	defaultFont : {
		size : 12,
		family : 'Calibri',
		color : 'FFFFFFFF'
	}
}
*/
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
		this.styleData = {
			'numFmts': [],
			'fonts': [],
			'fills': [
				{
					'patternFill': {
						'patternType' : 'none'
					}
				},
				{
					'patternFill': {
						'patternType' : 'none'
					}
				}
			],
			'borders': [
				{
					'left': null,
					'right': null,
					'top': null,
					'bottom': null,
					'diagonal': null
				}
			],
			'cellXfs': [
				{
					'borderId': '0',
					'fillId': '0',
					'fontId': '0',
					'numFmtId': '0'
				}
			]
		}

		// Set Default Font
		let defaultFont = {
			'sz': '12',
			'color': utils.cleanColor('Black'),
			'name': 'Calibri',
			'family': '2',
			'scheme': 'minor'
		};
		if(this.opts.defaultFont){
			defaultFont.sz = this.opts.defaultFont.size;
			defaultFont.color = utils.cleanColor(this.opts.defaultFont.color);
			defaultFont.name = this.opts.defaultFont.family;
		}
		this.styleData.fonts.push(defaultFont)

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

	getStringIndex(val) {
		if(this.sharedStrings.indexOf(val) < 0){
			this.sharedStrings.push(val);
		}
		return this.sharedStrings.indexOf(val);
	}
}

module.exports = WorkBook;
const _ = require('lodash');
const fs = require('fs');
const utils = require('../utils.js');
const Worksheet = require('../worksheet');
const Style = require('../style');
const Border = require('../style/classes/border.js');
const Fill = require('../style/classes/fill.js');
const DXFCollection = require('./dxfCollection.js');
const MediaCollection = require('./mediaCollection.js');
const DefinedNameCollection = require('../classes/definedNameCollection.js');
const SlothLogger = require('sloth-logger');
const types = require('../types/index.js');
const builder = require('./builder.js');
const http = require('http');


/* Available options for Workbook
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
// Default Options for Workbook
let workbookDefaultOpts = {
    jszip: {
        compression: 'DEFLATE'
    },
    defaultFont: {
        'color': 'FF000000',
        'name': 'Calibri',
        'size': 12,
        'family': 'roman'
    },
    dateFormat: 'm/d/yy'
};


class Workbook {

    /**
     * @class Workbook
     * @param {Object} opts Workbook settings
     * @param {Object} opts.jszip
     * @param {String} opts.jszip.compression JSZip compression type. defaults to 'DEFLATE'
     * @param {Object} opts.defaultFont
     * @param {String} opts.defaultFont.color HEX value of default font color. defaults to #000000
     * @param {String} opts.defaultFont.name Font name. defaults to Calibri
     * @param {Number} opts.defaultFont.size Font size. defaults to 12
     * @param {String} opts.defaultFont.family Font family. defaults to roman
     * @param {String} opts.dataFormat Specifies the format for dates in the Workbook. defaults to 'm/d/yy'
     * @returns {Workbook}
     */
    constructor(opts) {
        opts = opts ? opts : {};
        
        this.logger = new SlothLogger.Logger({
            logLevel: Number.isNaN(parseInt(opts.logLevel)) ? 0 : parseInt(opts.logLevel)
        });

        this.opts = _.merge({}, workbookDefaultOpts, opts);

        this.sheets = [];
        this.sharedStrings = [];
        this.styles = [];
        this.dxfCollection = new DXFCollection(this);
        this.mediaCollection = new MediaCollection();
        this.definedNameCollection = new DefinedNameCollection();
        this.styleData = {
            'numFmts': [],
            'fonts': [],
            'fills': [new Fill({ type: 'pattern', patternType: 'none' }), new Fill({ type: 'pattern', patternType: 'gray125' })],
            'borders': [new Border()],
            'cellXfs': [
                {
                    'borderId': null,
                    'fillId': null,
                    'fontId': 0,
                    'numFmtId': null
                }
            ]
        };

        // Set Default Font and Style
        this.createStyle({ font: this.opts.defaultFont });

    }

    /**
     * setSelectedTab
     * @param {Number} tab number of sheet that should be displayed when workbook opens. tabs are indexed starting with 1
     **/
    setSelectedTab(id) {
        this.sheets.forEach((s) => {
            if (s.sheetId === id) {
                s.opts.sheetView.tabSelected = 1;
            } else {
                s.opts.sheetView.tabSelected = 0;
            }
        });
    }

    /**
     * Generate .xlsx file. 
     * @param {String} fileName Name of Excel workbook with .xslx extension
     * @param {http.response | callback} http response object or callback function (optional). 
     * If http response object is given, file is written to http response. Useful for web applications.
     * If callback is given, callback called with (err, fs.Stats) passed
     */
    write(fileName, handler) {

        builder.writeToBuffer(this)
        .then((buffer) => {
            switch (typeof handler) {
                // handler passed as http response object. 

            case 'object':
                if (handler instanceof http.ServerResponse) {
                    handler.writeHead(200, {
                        'Content-Length': buffer.length,
                        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        'Content-Disposition': 'attachment; filename="' + fileName + '"'
                    });
                    handler.end(buffer);
                } else {
                    throw new TypeError('Unknown object sent to write function.');
                }
                break;

            // handler passed as callback function
            case 'function':
                fs.writeFile(fileName, buffer, function (err) {
                    if (err) {
                        handler(err);
                    } else {
                        fs.stat(fileName, handler);
                    }
                });
                break;

            // no handler passed, write file to FS.
            default:
                
                fs.writeFile(fileName, buffer, function (err) {
                    if (err) { 
                        throw err; 
                    }
                });
                break;
            }
        })
        .catch((e) => {
            console.error(e.stack);
        });
    }

    /**
     * Add a worksheet to the Workbook 
     * @param {String} name Name of the Worksheet
     * @param {Object} opts Options for Worksheet. See Worksheet class definition
     * @returns {Worksheet}
     */
    addWorksheet(name, opts) {
        let newLength = this.sheets.push(new Worksheet(this, name, opts));
        return this.sheets[newLength - 1];
    }

    /** 
     * Add a Style to the Workbook
     * @param {Object} opts Options for the style. See Style class definition
     * @returns {Style}
     */
    createStyle(opts) {
        let thisStyle;
        let checkCount = 0;
        while (thisStyle === undefined && checkCount < this.styles.length) {
            if (_.isEqual(this.styles[checkCount].toObject(), opts)) {
                thisStyle = this.styles[checkCount];
            }
            checkCount++;
        }
        if (thisStyle === undefined) {
            thisStyle = new Style(this, opts);
            let count = this.styles.push(thisStyle);
            this.styles[count - 1].ids.cellXfs = count - 1;
            return this.styles[count - 1];
        } else {
            return thisStyle;
        }
    }

    /**
     * Gets the index of a string from the shared string array if exists and adds the string if it does not and returns the new index
     * @param {String} val Text of string
     * @returns {Number} index of the string in the shared strings array
     */
    getStringIndex(val) {
        if (this.sharedStrings.indexOf(val) < 0) {
            this.sharedStrings.push(val);
        }
        return this.sharedStrings.indexOf(val);
    }
}

module.exports = Workbook;
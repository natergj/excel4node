const _ = require('lodash');
const fs = require('fs');
const utils = require('../utils.js');
const WorkSheet = require('../worksheet');
const Style = require('../style');
const Border = require('../style/classes/border.js');
const Fill = require('../style/classes/fill.js');
const DXFCollection = require('./dxfCollection.js');
const MediaCollection = require('./mediaCollection.js');
const DefinedNameCollection = require('../classes/definedNameCollection.js');
const SlothLogger = require('sloth-logger');
const types = require('../types/index.js');
const builder = require('./builder.js');


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
    jszip: {
        compression: 'DEFLATE'
    },
    defaultFont: {
        'color': 'FF000000',
        'name': 'Calibri',
        'size': 12,
        'family': 'roman'
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
        opts = opts ? opts : {};
        
        this.logger = new SlothLogger.Logger({
            logLevel: Number.isNaN(parseInt(opts.logLevel)) ? 0 : parseInt(opts.logLevel)
        });

        this.opts = _.merge({}, workBookDefaultOpts, opts);

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
        this.logger.debug('write called');

        builder.writeToBuffer(this)
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

    addWorksheet(name, opts) {
        let newLength = this.sheets.push(new WorkSheet(this, name, opts));
        return this.sheets[newLength - 1];
    }

    createStyle(opts) {
        let thisStyle = new Style(this, opts);
        let count = this.styles.push(thisStyle);
        this.styles[count - 1].ids.cellXfs = count - 1;
        return this.styles[count - 1];
    }

    getStringIndex(val) {
        if (this.sharedStrings.indexOf(val) < 0) {
            this.sharedStrings.push(val);
        }
        return this.sharedStrings.indexOf(val);
    }
}

module.exports = WorkBook;
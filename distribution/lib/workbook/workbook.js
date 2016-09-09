'use strict';

var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol ? "symbol" : typeof obj; };

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = require('lodash');
var fs = require('fs');
var utils = require('../utils.js');
var Worksheet = require('../worksheet');
var Style = require('../style');
var Border = require('../style/classes/border.js');
var Fill = require('../style/classes/fill.js');
var DXFCollection = require('./dxfCollection.js');
var MediaCollection = require('./mediaCollection.js');
var DefinedNameCollection = require('../classes/definedNameCollection.js');
var SlothLogger = require('sloth-logger');
var types = require('../types/index.js');
var builder = require('./builder.js');
var http = require('http');

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
var workbookDefaultOpts = {
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

var Workbook = function () {

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
    function Workbook(opts) {
        _classCallCheck(this, Workbook);

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
            'cellXfs': [{
                'borderId': null,
                'fillId': null,
                'fontId': 0,
                'numFmtId': null
            }]
        };

        // Set Default Font and Style
        this.createStyle({ font: this.opts.defaultFont });
    }

    /**
     * setSelectedTab
     * @param {Number} tab number of sheet that should be displayed when workbook opens. tabs are indexed starting with 1
     **/


    _createClass(Workbook, [{
        key: 'setSelectedTab',
        value: function setSelectedTab(id) {
            this.sheets.forEach(function (s) {
                if (s.sheetId === id) {
                    s.opts.sheetView.tabSelected = 1;
                } else {
                    s.opts.sheetView.tabSelected = 0;
                }
            });
        }

        /**
         * writeToBuffer
         * Writes Excel data to a node Buffer.
         */

    }, {
        key: 'writeToBuffer',
        value: function writeToBuffer() {
            return builder.writeToBuffer(this);
        }

        /**
         * Generate .xlsx file.
         * @param {String} fileName Name of Excel workbook with .xslx extension
         * @param {http.response | callback} http response object or callback function (optional).
         * If http response object is given, file is written to http response. Useful for web applications.
         * If callback is given, callback called with (err, fs.Stats) passed
         */

    }, {
        key: 'write',
        value: function write(fileName, handler) {

            builder.writeToBuffer(this).then(function (buffer) {
                switch (typeof handler === 'undefined' ? 'undefined' : _typeof(handler)) {
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
            }).catch(function (e) {
                throw new Error(e.stack);
            });
        }

        /**
         * Add a worksheet to the Workbook
         * @param {String} name Name of the Worksheet
         * @param {Object} opts Options for Worksheet. See Worksheet class definition
         * @returns {Worksheet}
         */

    }, {
        key: 'addWorksheet',
        value: function addWorksheet(name, opts) {
            var newLength = this.sheets.push(new Worksheet(this, name, opts));
            return this.sheets[newLength - 1];
        }

        /**
         * Add a Style to the Workbook
         * @param {Object} opts Options for the style. See Style class definition
         * @returns {Style}
         */

    }, {
        key: 'createStyle',
        value: function createStyle(opts) {
            var thisStyle = void 0;
            var checkCount = 0;
            while (thisStyle === undefined && checkCount < this.styles.length) {
                if (_.isEqual(this.styles[checkCount].toObject(), opts)) {
                    thisStyle = this.styles[checkCount];
                }
                checkCount++;
            }
            if (thisStyle === undefined) {
                thisStyle = new Style(this, opts);
                var count = this.styles.push(thisStyle);
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

    }, {
        key: 'getStringIndex',
        value: function getStringIndex(val) {
            if (this.sharedStrings.indexOf(val) < 0) {
                this.sharedStrings.push(val);
            }
            return this.sharedStrings.indexOf(val);
        }
    }]);

    return Workbook;
}();

module.exports = Workbook;
//# sourceMappingURL=workbook.js.map
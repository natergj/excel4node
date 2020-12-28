'use strict';

var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _isUndefined = require('lodash.isundefined');
var deepmerge = require('deepmerge');
var fs = require('fs');
var utils = require('../utils.js');
var Worksheet = require('../worksheet');
var Style = require('../style');
var Border = require('../style/classes/border.js');
var Fill = require('../style/classes/fill.js');
var Font = require('../style/classes/font');
var DXFCollection = require('./dxfCollection.js');
var MediaCollection = require('./mediaCollection.js');
var DefinedNameCollection = require('../classes/definedNameCollection.js');
var types = require('../types/index.js');
var builder = require('./builder.js');
var http = require('http');
var SimpleLogger = require('../logger');

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
     * @param {Number} opts.workbookView.activeTab Specifies an unsignedInt that contains the index to the active sheet in this book view.
     * @param {Boolean} opts.workbookView.autoFilterDateGrouping Specifies a boolean value that indicates whether to group dates when presenting the user with filtering options in the user interface.
     * @param {Number} opts.workbookView.firstSheet Specifies the index to the first sheet in this book view.
     * @param {Boolean} opts.workbookView.minimized Specifies a boolean value that indicates whether the workbook window is minimized.
     * @param {Boolean} opts.workbookView.showHorizontalScroll Specifies a boolean value that indicates whether to display the horizontal scroll bar in the user interface.
     * @param {Boolean} opts.workbookView.showSheetTabs Specifies a boolean value that indicates whether to display the sheet tabs in the user interface.
     * @param {Boolean} opts.workbookView.showVerticalScroll Specifies a boolean value that indicates whether to display the vertical scroll bar.
     * @param {Number} opts.workbookView.tabRatio Specifies ratio between the workbook tabs bar and the horizontal scroll bar.
     * @param {String} opts.workbookView.visibility Specifies visible state of the workbook window. ('hidden', 'veryHidden', 'visible') (§18.18.89)
     * @param {Number} opts.workbookView.windowHeight Specifies the height of the workbook window. The unit of measurement for this value is twips.
     * @param {Number} opts.workbookView.windowWidth Specifies the width of the workbook window. The unit of measurement for this value is twips..
     * @param {Number} opts.workbookView.xWindow Specifies the X coordinate for the upper left corner of the workbook window. The unit of measurement for this value is twips.
     * @param {Number} opts.workbookView.yWindow Specifies the Y coordinate for the upper left corner of the workbook window. The unit of measurement for this value is twips.
     * @param {Boolean} opts.workbookView
     * @param {Object} opts.logger Logger that supports warn and error method, defaults to console
     * @param {String} opts.author Name displayed as document's author
     * @returns {Workbook}
     */
    function Workbook() {
        var opts = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};

        _classCallCheck(this, Workbook);

        var hasCustomLogger = opts.logger !== undefined;
        var hasValidCustomLogger = hasCustomLogger && typeof opts.logger.warn === 'function' && typeof opts.logger.error === 'function';

        this.logger = hasValidCustomLogger ? opts.logger : new SimpleLogger({
            logLevel: Number.isNaN(parseInt(opts.logLevel)) ? 0 : parseInt(opts.logLevel)
        });
        if (hasCustomLogger && !hasValidCustomLogger) {
            this.logger.log('opts.logger is not a valid logger');
        }

        this.opts = deepmerge(workbookDefaultOpts, opts);
        this.author = this.opts.author || 'Microsoft Office User';

        this.sheets = [];
        this.sharedStrings = [];
        this.sharedStringLookup = new Map();
        this.styles = [];
        this.stylesLookup = new Map();
        this.dxfCollection = new DXFCollection(this);
        this.mediaCollection = new MediaCollection();
        this.definedNameCollection = new DefinedNameCollection();
        this.styleData = {
            'numFmts': [],
            'fonts': [],
            'fills': [new Fill({
                type: 'pattern',
                patternType: 'none'
            }), new Fill({
                type: 'pattern',
                patternType: 'gray125'
            })],
            'borders': [new Border()],
            'cellXfs': [{
                'borderId': null,
                'fillId': null,
                'fontId': 0,
                'numFmtId': null
            }]
        };

        // Lookups for style components to quickly find existing entries
        // - Lookup keys are stringified JSON of a style's toObject result
        // - Lookup values are the indexes for the actual entry in the styleData arrays
        this.styleDataLookup = {
            'fonts': {},
            'fills': this.styleData.fills.reduce(function (ret, fill, index) {
                ret[JSON.stringify(fill.toObject())] = index;
                return ret;
            }, {}),
            'borders': this.styleData.borders.reduce(function (ret, border, index) {
                ret[JSON.stringify(border.toObject())] = index;
                return ret;
            }, {})
        };

        // Set Default Font and Style
        this.createStyle({
            font: this.opts.defaultFont
        });
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
            var _this = this;

            builder.writeToBuffer(this).then(function (buffer) {
                switch (typeof handler === 'undefined' ? 'undefined' : _typeof(handler)) {
                    // handler passed as http response object.

                    case 'object':
                        if (handler instanceof http.ServerResponse) {
                            handler.writeHead(200, {
                                'Content-Length': buffer.length,
                                'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                                'Content-Disposition': 'attachment; filename="' + encodeURIComponent(fileName) + '"; filename*=utf-8\'\'' + encodeURIComponent(fileName) + ';'
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
                if (handler instanceof http.ServerResponse) {
                    _this.logger.error(e.stack);
                    handler.status = 500;
                    handler.setHeader('Content-Type', 'text/plain');
                    handler.end('500 Server Error');
                } else if (typeof handler === 'function') {
                    handler(e.stack);
                } else {
                    _this.logger.error(e.stack);
                }
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
            var thisStyle = new Style(this, opts);
            var lookupKey = JSON.stringify(thisStyle.toObject());

            // Use existing style if one exists
            if (this.stylesLookup.get(lookupKey)) {
                return this.stylesLookup.get(lookupKey);
            }

            this.stylesLookup.set(lookupKey, thisStyle);
            var index = this.styles.push(thisStyle) - 1;
            this.styles[index].ids.cellXfs = index;
            return this.styles[index];
        }

        /**
         * Gets the index of a string from the shared string array if exists and adds the string if it does not and returns the new index
         * @param {String} val Text of string
         * @returns {Number} index of the string in the shared strings array
         */

    }, {
        key: 'getStringIndex',
        value: function getStringIndex(val) {
            var lookupKey = typeof val === "string" ? val : JSON.stringify(val);
            var target = this.sharedStringLookup.get(lookupKey);
            if (_isUndefined(target)) {
                var index = this.sharedStrings.push(val) - 1;
                this.sharedStringLookup.set(lookupKey, index);
                return index;
            } else {
                return target;
            }
        }

        /**
         * @func Workbook._generateXML
         * @desc used for testing the Workbook XML generated by the builder
         * @return {Promise} resolves with Workbook XML 
         */

    }, {
        key: '_generateXML',
        value: function _generateXML() {
            return builder.workbookXML(this);
        }
    }]);

    return Workbook;
}();

module.exports = Workbook;
//# sourceMappingURL=workbook.js.map
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var utils = require('../utils.js');
var deepmerge = require('deepmerge');

var Alignment = require('./classes/alignment.js');
var Border = require('./classes/border.js');
var Fill = require('./classes/fill.js');
var Font = require('./classes/font.js');
var NumberFormat = require('./classes/numberFormat.js');

var _getFontId = function _getFontId(wb) {
    var font = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : {};


    // Create the Font and lookup key
    font = deepmerge(wb.opts.defaultFont, font);
    var thisFont = new Font(font);
    var lookupKey = JSON.stringify(thisFont.toObject());

    // Find an existing entry, creating a new one if it does not exist
    var id = wb.styleDataLookup.fonts[lookupKey];
    if (id === undefined) {
        id = wb.styleData.fonts.push(thisFont) - 1;
        wb.styleDataLookup.fonts[lookupKey] = id;
    }

    return id;
};

var _getFillId = function _getFillId(wb, fill) {
    if (fill === undefined) {
        return null;
    }

    // Create the Fill and lookup key
    var thisFill = new Fill(fill);
    var lookupKey = JSON.stringify(thisFill.toObject());

    // Find an existing entry, creating a new one if it does not exist
    var id = wb.styleDataLookup.fills[lookupKey];
    if (id === undefined) {
        id = wb.styleData.fills.push(thisFill) - 1;
        wb.styleDataLookup.fills[lookupKey] = id;
    }

    return id;
};

var _getBorderId = function _getBorderId(wb, border) {
    if (border === undefined) {
        return null;
    }

    // Create the Border and lookup key
    var thisBorder = new Border(border);
    var lookupKey = JSON.stringify(thisBorder.toObject());

    // Find an existing entry, creating a new one if it does not exist
    var id = wb.styleDataLookup.borders[lookupKey];
    if (id === undefined) {
        id = wb.styleData.borders.push(thisBorder) - 1;
        wb.styleDataLookup.borders[lookupKey] = id;
    }

    return id;
};

var _getNumFmt = function _getNumFmt(wb, val) {
    var fmt = void 0;
    wb.styleData.numFmts.forEach(function (f) {
        if (f.formatCode === val) {
            fmt = f;
        }
    });

    if (fmt === undefined) {
        var fmtId = wb.styleData.numFmts.length + 164;
        fmt = new NumberFormat(val);
        fmt.numFmtId = fmtId;
        wb.styleData.numFmts.push(fmt);
    }

    return fmt;
};

/*
    Style Opts
    {
        alignment: { // §18.8.1
            horizontal: ['center', 'centerContinuous', 'distributed', 'fill', 'general', 'justify', 'left', 'right'],
            indent: integer, // Number of spaces to indent = indent value * 3
            justifyLastLine: boolean,
            readingOrder: ['contextDependent', 'leftToRight', 'rightToLeft'], 
            relativeIndent: integer, // number of additional spaces to indent
            shrinkToFit: boolean,
            textRotation: integer, // number of degrees to rotate text counter-clockwise
            vertical: ['bottom', 'center', 'distributed', 'justify', 'top'],
            wrapText: boolean
        },
        font: { // §18.8.22
            bold: boolean,
            charset: integer,
            color: string,
            condense: boolean,
            extend: boolean,
            family: string,
            italics: boolean,
            name: string,
            outline: boolean,
            scheme: string, // §18.18.33 ST_FontScheme (Font scheme Styles)
            shadow: boolean,
            strike: boolean,
            size: integer,
            underline: boolean,
            vertAlign: string // §22.9.2.17 ST_VerticalAlignRun (Vertical Positioning Location)
        },
        border: { // §18.8.4 border (Border)
            left: {
                style: string,
                color: string
            },
            right: {
                style: string,
                color: string
            },
            top: {
                style: string,
                color: string
            },
            bottom: {
                style: string,
                color: string
            },
            diagonal: {
                style: string,
                color: string
            },
            diagonalDown: boolean,
            diagonalUp: boolean,
            outline: boolean
        },
        fill: { // §18.8.20 fill (Fill)
            type: 'pattern',
            patternType: 'solid',
            color: 'Yellow'
        },
        numberFormat: integer or string // §18.8.30 numFmt (Number Format)
    }
*/

var Style = function () {
    function Style(wb, opts) {
        _classCallCheck(this, Style);

        /**
         * Excel Style object
         * @class Style
         * @desc Style object for formatting Excel Cells
         * @param {Workbook} wb Excel Workbook object
         * @param {Object} opts Options for style
         * @param {Object} opts.alignment Options for creating an Alignment instance
         * @param {Object} opts.font Options for creating a Font instance
         * @param {Object} opts.border Options for creating a Border instance
         * @param {Object} opts.fill Options for creating a Fill instance
         * @param {String} opts.numberFormat
         * @property {Alignment} alignment Alignment instance associated with Style
         * @property {Border} border Border instance associated with Style
         * @property {Number} borderId ID of Border instance in the Workbook
         * @property {Fill} fill Fill instance associated with Style
         * @property {Number} fillId ID of Fill instance in the Workbook
         * @property {Font} font Font instance associated with Style
         * @property {Number} fontId ID of Font instance in the Workbook
         * @property {String} numberFormat String represenation of the way a number should be formatted
         * @property {Number} xf XF id of the Style in the Workbook
         * @returns {Style} 
         */
        opts = opts ? opts : {};
        opts = deepmerge(wb.styles[0] ? wb.styles[0] : {}, opts);

        if (opts.alignment !== undefined) {
            this.alignment = new Alignment(opts.alignment);
        }

        if (opts.border !== undefined) {
            this.borderId = _getBorderId(wb, opts.border); // attribute 0 based index
            this.border = wb.styleData.borders[this.borderId];
        }
        if (opts.fill !== undefined) {
            this.fillId = _getFillId(wb, opts.fill); // attribute 0 based index
            this.fill = wb.styleData.fills[this.fillId];
        }

        if (opts.font !== undefined) {
            this.fontId = _getFontId(wb, opts.font); // attribute 0 based index
            this.font = wb.styleData.fonts[this.fontId];
        }

        if (opts.numberFormat !== undefined) {
            if (typeof opts.numberFormat === 'number' && opts.numberFormat <= 164) {
                this.numFmtId = opts.numberFormat;
            } else if (typeof opts.numberFormat === 'string') {
                this.numFmt = _getNumFmt(wb, opts.numberFormat);
            }
        }

        if (opts.pivotButton !== undefined) {
            this.pivotButton = null; // attribute boolean
        }

        if (opts.quotePrefix !== undefined) {
            this.quotePrefix = null; // attribute boolean
        }

        this.ids = {};
    }

    _createClass(Style, [{
        key: 'toObject',


        /** 
         * @func Style.toObject
         * @desc Converts the Style instance to a javascript object
         * @returns {Object}
         */
        value: function toObject() {
            var obj = {};

            if (typeof this.fontId === 'number') {
                obj.font = this.font.toObject();
            }

            if (typeof this.fillId === 'number') {
                obj.fill = this.fill.toObject();
            }

            if (typeof this.borderId === 'number') {
                obj.border = this.border.toObject();
            }

            if (typeof this.numFmtId === 'number' && this.numFmtId < 164) {
                obj.numberFormat = this.numFmtId;
            } else if (this.numFmt !== undefined && this.numFmt !== null) {
                obj.numberFormat = this.numFmt.formatCode;
            }

            if (this.alignment instanceof Alignment) {
                obj.alignment = this.alignment.toObject();
            }

            if (this.pivotButton !== undefined) {
                obj.pivotButton = this.pivotButton;
            }

            if (this.quotePrefix !== undefined) {
                obj.quotePrefix = this.quotePrefix;
            }

            return obj;
        }

        /**
         * @alias Style.addToXMLele
         * @desc When generating Workbook output, attaches style to the styles xml file
         * @func Style.addToXMLele
         * @param {xmlbuilder.Element} ele Element object of the xmlbuilder module
         */

    }, {
        key: 'addXFtoXMLele',
        value: function addXFtoXMLele(ele) {
            var thisEle = ele.ele('xf');
            var thisXF = this.xf;
            Object.keys(thisXF).forEach(function (a) {
                if (a === 'alignment') {
                    thisXF[a].addToXMLele(thisEle);
                } else {
                    thisEle.att(a, thisXF[a]);
                }
            });
        }

        /**
         * @alias Style.addDXFtoXMLele
         * @desc When generating Workbook output, attaches style to the styles xml file as a dxf for use with conditional formatting rules
         * @func Style.addDXFtoXMLele
         * @param {xmlbuilder.Element} ele Element object of the xmlbuilder module
         */

    }, {
        key: 'addDXFtoXMLele',
        value: function addDXFtoXMLele(ele) {
            var thisEle = ele.ele('dxf');

            if (this.font instanceof Font) {
                this.font.addToXMLele(thisEle);
            }

            if (this.numFmt instanceof NumberFormat) {
                this.numFmt.addToXMLele(thisEle);
            }

            if (this.fill instanceof Fill) {
                this.fill.addToXMLele(thisEle.ele('fill'));
            }

            if (this.alignment instanceof Alignment) {
                this.alignment.addToXMLele(thisEle);
            }

            if (this.border instanceof Border) {
                this.border.addToXMLele(thisEle);
            }
        }
    }, {
        key: 'xf',
        get: function get() {
            var thisXF = {};

            if (typeof this.fontId === 'number') {
                thisXF.applyFont = 1;
                thisXF.fontId = this.fontId;
            }

            if (typeof this.fillId === 'number') {
                thisXF.applyFill = 1;
                thisXF.fillId = this.fillId;
            }

            if (typeof this.borderId === 'number') {
                thisXF.applyBorder = 1;
                thisXF.borderId = this.borderId;
            }

            if (typeof this.numFmtId === 'number') {
                thisXF.applyNumberFormat = 1;
                thisXF.numFmtId = this.numFmtId;
            } else if (this.numFmt !== undefined && this.numFmt !== null) {
                thisXF.applyNumberFormat = 1;
                thisXF.numFmtId = this.numFmt.numFmtId;
            }

            if (this.alignment instanceof Alignment) {
                thisXF.applyAlignment = 1;
                thisXF.alignment = this.alignment;
            }

            return thisXF;
        }
    }]);

    return Style;
}();

module.exports = Style;
//# sourceMappingURL=style.js.map
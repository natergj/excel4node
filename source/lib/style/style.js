const utils = require('../utils.js');
const _ = require('lodash');

const Alignment = require('./classes/alignment.js');
const Border = require('./classes/border.js');
const Fill = require('./classes/fill.js');
const Font = require('./classes/font.js');

let _getFontId = (wb, font) => {

    let thisFont = new Font(font);

    let fontId;
    wb.styleData.fonts.forEach((f, i) => {
        if (_.isEqual(thisFont.toObject(), f.toObject())) {
            fontId = i;
        }
    });
    if (fontId === undefined) {
        let count = wb.styleData.fonts.push(thisFont);
        fontId = count - 1;
    }

    return fontId;
};

let _getFillId = (wb, fill) => {
    if (fill === undefined) {
        return null;
    }

    let thisFill = new Fill(fill);

    let fillId;
    wb.styleData.fills.forEach((f, i) => {
        if (_.isEqual(thisFill.toObject(), f.toObject())) {
            fillId = i;
        }
    });
    if (fillId === undefined) {
        let count = wb.styleData.fills.push(thisFill);
        fillId = count - 1;
    }

    return fillId;
};

let _getBorderId = (wb, border) => {
    if (border === undefined) {
        return null;
    }

    let thisBorder = new Border(border);
    let borderId;
    wb.styleData.borders.forEach((b, i) => {
        if (_.isEqual(b.toObject(), thisBorder.toObject())) {
            borderId = i;
        }
    });

    if (borderId === undefined) {
        let count = wb.styleData.borders.push(thisBorder);
        borderId = count - 1;
    }

    return borderId;
};

let _getNumFmt = (wb, val) => {
    let fmt;
    wb.styleData.numFmts.forEach((f, i) => {
        if (_.isEqual(f.formatCode, val)) {
            fmt = f;
        }
    });

    if (fmt === undefined) {
        let fmtId = wb.styleData.numFmts.length + 164;
        fmt = {
            formatCode: val,
            numFmtId: fmtId
        };
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
module.exports = class Style {
    constructor(wb, opts) {
        opts = opts ? opts : {};

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
            if (typeof opts.numberFormat === 'number' && opts.numberFormat <= 164){
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

    get xf() {
        let thisXF = {};

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

    toObject() {
        let obj = {};

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
        } else if (this.numFmt !== undefined && this.numFmt !== null){
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

    addXFtoXMLele(ele) {
        let thisEle = ele.ele('xf');
        let thisXF = this.xf;
        Object.keys(thisXF).forEach((a) => {
            if(a === 'alignment') {
                thisXF[a].addToXMLele(thisEle);
            } else {
                thisEle.att(a, thisXF[a]);
            }
        });        
    }
};

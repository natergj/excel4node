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

let _getNumFmtId = (wb, fmt) => {
    if (fmt === undefined) {
        return null;
    }

    if (typeof fmt === 'number') {
        if (fmt <= 166) {
            return fmt;
        } else {
            return 0;
        }
    }

    if (typeof fmt === 'string') {

        let fmtId;
        wb.styleData.numFmts.forEach((f, i) => {
            if (_.isEqual(f.formatCode, fmt)) {
                fmtId = f.numFmtId;
            }
        });
        if (fmtId === undefined) {
            fmtId = wb.styleData.numFmts.length + 166;
            wb.styleData.numFmts.push({
                formatCode: fmt,
                numFmtId: fmtId
            });
        }

        return fmtId;
    }

    return null;
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
            pattern: string,
            color: string,
            gradient: object
        },
        numberFormat: integer or string // §18.8.30 numFmt (Number Format)
    }
*/
module.exports = class Style {
    constructor(wb, opts) {
        opts = opts ? opts : {};
        opts.alignment ? this.alignment = new Alignment(opts.alignment) : null; // Child Element
        opts.protection ? this.protection = opts.protection : null; // Child Element
        this.applyAlignment = null; // attribute boolean
        this.applyBorder = null; // attribute boolean
        this.applyFill = null; // attribute boolean
        this.applyFont = null; // attribute boolean
        this.applyNumberFormat = null; // attribute boolean
        this.applyProtection = null; // attribute boolean
        this.borderId = _getBorderId(wb, opts.border); // attribute 0 based index
        this.fillId = _getFillId(wb, opts.fill); // attribute 0 based index
        this.fontId = _getFontId(wb, opts.font); // attribute 0 based index
        this.numFmtId = _getNumFmtId(wb, opts.numberFormat); // attribute 0 based index
        this.pivotButton = null; // attribute boolean
        this.quotePrefix = null; // attribute boolean
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
            obj.applyFont = 1;
            obj.fontId = this.fontId;
        }

        if (typeof this.fillId === 'number') {
            obj.applyFill = 1;
            obj.fillId = this.fillId;
        }

        if (typeof this.borderId === 'number') {
            obj.applyBorder = 1;
            obj.borderId = this.borderId;
        }

        if (typeof this.numFmtId === 'number') {
            obj.applyNumberFormat = 1;
            obj.numFmtId = this.numFmtId;
        }

        if (this.alignment instanceof Alignment) {
            obj.applyAlignment = 1;
            obj.alignment = this.alignment.toObject();
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

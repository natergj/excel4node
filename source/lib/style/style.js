const utils = require('../utils.js');
const logger = require('../logger.js');
const _ = require('lodash');

let _getFontId = (wb, font) => {
    if (font === undefined) {
        return null;
    }

    let thisFont = {};

    typeof font.charset === 'number' ? thisFont.charset = font.charset : null;
    typeof font.color === 'string' ? thisFont.color = utils.cleanColor(font.color) : thisFont.color = wb.styleData.fonts[0].color;
    typeof font.family === 'number' ? thisFont.family = font.family : thisFont.family = wb.styleData.fonts[0].family;
    typeof font.name === 'string' ? thisFont.name = font.name : thisFont.name = wb.styleData.fonts[0].name;
    typeof font.scheme === 'string' ? thisFont.scheme = font.scheme : thisFont.scheme = wb.styleData.fonts[0].scheme;
    typeof font.size === 'number' ? thisFont.sz = font.size : thisFont.sz = wb.styleData.fonts[0].sz;

    font.condense === true ? thisFont.condense = true : null;
    font.extend === true ? thisFont.extend = true : null;
    font.bold === true ? thisFont.b = true : null;
    font.italics === true ? thisFont.i = true : null;
    font.outline === true ? thisFont.outline = true : null;
    font.shadow === true ? thisFont.shadow = true : null;
    font.strike === true ? thisFont.strike = true : null;
    font.underline === true ? thisFont.u = true : null;
    font.alignVertical === true ? thisFont.vertAlign = true : null;

    let fontId;
    wb.styleData.fonts.forEach((f, i) => {
        if (_.isEqual(f, thisFont)) {
            fontId = i;
        }
    });
    if (!fontId) {
        let count = wb.styleData.fonts.push(thisFont);
        fontId = count - 1;
    }

    return fontId;
};

let _getFontId = (wb, fill) => {
    if (fill === undefined) {
        return null;
    }

    let thisFill = {};

    let fillId;
    wb.styleData.fills.forEach((f, i) => {
        if (_.isEqual(f, thisFill)) {
            fillId = i;
        }
    });
    if (!fillId) {
        let count = wb.styleData.fills.push(thisFill);
        fillId = count - 1;
    }

    return fillId;
}

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
        if(!fmtId){
            fmtId = wb.styleData.numFmts.length + 166
            wb.styleData.numFmts.push({
                formatCode: fmt,
                numFmtId: fmtId
            });
        }

        return fmtId;
    }

    return null;
}


/*
    Style Opts
    {
        alignment: { // §18.8.1
            horizontal: ['center', 'centerContinuous', 'distributed', 'fill', 'general', 'justify', 'left', 'right'],
            indent: integer, // Number of spaces to indent = indent value * 3
            justifyLastLine: boolean,
            readingOrder: [0, 1, 2], // 0 = context dependent, 1 = left-to-right, 2 = right-to-left
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
            alignVertical: boolean
        },
        fill: { // §18.8.20 fill (Fill)
            pattern: string,
            color: string,
            gradient: object
        },
        number: integer or string // §18.8.30 numFmt (Number Format)
    }
*/
module.exports = class Style {
    constructor(wb, opts) {
        opts = opts ? opts : {};
        opts.alignment ? this.alignment = opts.alignment : null; // Child Element
        opts.protection ? this.protection = opts.protection : null; // Child Element
        this.applyAlignment = null; // attribute boolean
        this.applyBorder = null; // attribute boolean
        this.applyFill = null // attribute boolean
        this.applyFont = null // attribute boolean
        this.applyNumberFormat = null // attribute boolean
        this.applyProtection = null // attribute boolean
        this.borderId = null // attribute 0 based index
        this.fillId = _getFillId(wb, opts.fill); // attribute 0 based index
        this.fontId = _getFontId(wb, opts.font); // attribute 0 based index
        this.numFmtId = _getNumFmtId(wb, opts.number); // attribute 0 based index
        this.pivotButton = null // attribute boolean
        this.quotePrefix = null // attribute boolean
    }

    get xf() {
        /*
        <xsd:sequence>
           <xsd:element name="alignment" type="CT_CellAlignment" minOccurs="0" maxOccurs="1"/>
           <xsd:element name="protection" type="CT_CellProtection" minOccurs="0" maxOccurs="1"/>
           <xsd:element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
        </xsd:sequence>
        <xsd:attribute name="numFmtId" type="ST_NumFmtId" use="optional"/>
        <xsd:attribute name="fontId" type="ST_FontId" use="optional"/>
        <xsd:attribute name="fillId" type="ST_FillId" use="optional"/>
        <xsd:attribute name="borderId" type="ST_BorderId" use="optional"/>
        <xsd:attribute name="xfId" type="ST_CellStyleXfId" use="optional"/>
        <xsd:attribute name="quotePrefix" type="xsd:boolean" use="optional" default="false"/>
        <xsd:attribute name="pivotButton" type="xsd:boolean" use="optional" default="false"/>
        <xsd:attribute name="applyNumberFormat" type="xsd:boolean" use="optional"/>
        <xsd:attribute name="applyFont" type="xsd:boolean" use="optional"/>
        <xsd:attribute name="applyFill" type="xsd:boolean" use="optional"/>
        <xsd:attribute name="applyBorder" type="xsd:boolean" use="optional"/>
        <xsd:attribute name="applyAlignment" type="xsd:boolean" use="optional"/>
        <xsd:attribute name="applyProtection" type="xsd:boolean" use="optional"/>
        */

        let thisXF = {};

        if (typeof this.fontId === 'number') {
            thisXF.applyFont = 1;
            thisXF.fontId = this.fontId;
        }

        return thisXF;
    }
};

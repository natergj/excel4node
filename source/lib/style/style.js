const utils = require('../utils.js');
const logger = require('../logger.js');
const _ = require('lodash');

let _getFontId = (wb, font) => {
    if (font === undefined) {
        return null;
    }

    let thisFont = {};
    if (font.bold === true) {
        thisFont.b = true;
    }

    if (typeof font.charset === 'number') {
        thisFont.charset = font.charset;
    }

    if (typeof font.color === 'string') {
        thisFont.color = utils.cleanColor(font.color);
    } else {
        thisFont.color = wb.styleData.fonts[0].color;
    }

    if (font.condense === true) {
        thisFont.condense = true;
    }

    if (font.exend === true) {
        thisFont.extend = true;
    }

    if (typeof font.family === 'number') {
        thisFont.family = font.family;
    } else {
        thisFont.family = wb.styleData.fonts[0].family;
    }

    if (font.italics === true) {
        thisFont.i = true;
    }

    if (typeof font.name === 'string') {
        thisFont.name = font.name;
    } else {
        thisFont.name = wb.styleData.fonts[0].name;
    }

    if (font.outline === true) {
        thisFont.outline = true;
    }

    if (typeof font.scheme === 'string') {
        thisFont.scheme = font.scheme;
    } else {
        thisFont.scheme = wb.styleData.fonts[0].scheme;
    }

    if (font.shadow === true) {
        thisFont.shadow = true;
    }

    if (font.strike === true) {
        thisFont.strike = true;
    }

    if (typeof font.size === 'number') {
        thisFont.sz = font.size;
    } else {
        thisFont.sz = wb.styleData.fonts[0].sz;
    }

    if (font.underline === true) {
        thisFont.u = true;
    }

    if (font.alignVertical === true) {
        thisFont.vertAlign = true;
    }

    let fontId;
    wb.styleData.fonts.forEach((f, i) => {
        if (_.isEqual(f, thisFont)) {
            fontId = i;
        }
    });
    if (!fontId) {
        let count = wb.styleData.fonts.push(thisFont);
        let fontId = count - 1;
    }
    return fontId;
};

module.exports = class Style {
    constructor(wb, opts) {
        this.wb = wb;
        this.fontId = _getFontId(wb, opts.font);
    }

    get xfId() {
    }

    get dxfId() {
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

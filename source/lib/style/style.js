const utils = require('../utils.js');
const logger = require('../logger.js');
const _ = require('lodash');

let _getFont = (wb, font) => {
    if (font === undefined) {
        return null;
    }
};

module.exports = class Style {
    constructor(wb, opts) {
        this.wb = wb;
        this.font = _getFont(wb, opts.font);
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

        if (this.font !== null) {
            thisXF.applyFont = 1;
        }

        return thisXF;
    }
};

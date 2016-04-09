const utils = require('../../utils.js');
const _ = require('lodash');
const xmlbuilder = require('xmlbuilder');
const types = require('../../types/index.js');

class Font {
    constructor(opts) {
        opts = opts ? opts : {};
        console.log(opts);
        console.log(typeof(opts.underline) === 'boolean');

        typeof opts.color === 'string' ? this.color = types.excelColor.getColor(opts.color) : null;
        typeof opts.name === 'string' ? this.name = opts.name : null;
        typeof opts.scheme === 'string' ? this.scheme = opts.scheme : null;
        typeof opts.size === 'number' ? this.size = opts.size : null;
        typeof opts.family === 'string' && types.fontFamily.validate(opts.family) === true ? this.family = opts.family : null;

        typeof opts.vertAlign === 'string' ? this.vertAlign = opts.vertAlign : null;
        typeof opts.charset === 'number' ? this.charset = opts.charset : null;

        typeof opts.condense === 'boolean' ? this.condense = opts.condense : null;
        typeof opts.extend === 'boolean' ? this.extend = opts.extend : null;
        typeof opts.bold === 'boolean' ? this.bold = opts.bold : null;
        typeof opts.italics === 'boolean' ? this.italics = opts.italics : null;
        typeof opts.outline === 'boolean' ? this.outline = opts.outline : null;
        typeof opts.shadow === 'boolean' ? this.shadow = opts.shadow : null;
        typeof opts.strike === 'boolean' ? this.strike = opts.strike : null;
        typeof opts.underline === 'boolean' ? this.underline = opts.underline : null;
    }

    toObject() {
        let obj = {};

        typeof this.charset === 'number' ? obj.charset = this.charset : null;
        typeof this.color === 'string' ? obj.color = this.color : null;
        typeof this.family === 'string' ? obj.family = this.family : null;
        typeof this.name === 'string' ? obj.name = this.name : null;
        typeof this.scheme === 'string' ? obj.scheme = this.scheme : null;
        typeof this.size === 'number' ? obj.size = this.size : null;
        typeof this.vertAlign === 'string' ? obj.vertAlign = this.vertAlign : null;

        typeof this.condense === 'boolean' ? obj.condense = this.condense : null;
        typeof this.extend === 'boolean' ? obj.extend = this.extend : null;
        typeof this.bold === 'boolean' ? obj.bold = this.bold : null;
        typeof this.italics === 'boolean' ? obj.italics = this.italics : null;
        typeof this.outline === 'boolean' ? obj.outline = this.outline : null;
        typeof this.shadow === 'boolean' ? obj.shadow = this.shadow : null;
        typeof this.strike === 'boolean' ? obj.strike = this.strike : null;
        typeof this.underline === 'boolean' ? obj.underline = this.underline : null;

        return obj;
    }

    addToXMLele(fontXML) {
        let fEle = fontXML.ele('font');
        fEle.ele('sz').att('val', this.size !== undefined ? this.size : 12);
        fEle.ele('color').att('rgb', this.color !== undefined ? this.color : 'FF000000');
        fEle.ele('name').att('val', this.name !== undefined ? this.name : 'Calibri');
        if (this.family !== undefined) {
            fEle.ele('family').att('val', types.fontFamily[this.family.toLowerCase()]);
        }
        if (this.scheme !== undefined) {
            fEle.ele('scheme').att('val', this.scheme);
        }

        this.condense === true ? fEle.ele('condense') : null;
        this.extend === true ? fEle.ele('extend') : null;
        this.bold === true ? fEle.ele('b') : null;
        this.italics === true ? fEle.ele('i') : null;
        this.outline === true ? fEle.ele('outline') : null;
        this.shadow === true ? fEle.ele('shadow') : null;
        this.strike === true ? fEle.ele('strike') : null;
        this.underline === true ? fEle.ele('u') : null;
        this.vertAlign === true ? fEle.ele('vertAlign') : null;

        return true;
    }


}

module.exports = Font;
const utils = require('../../utils.js');
const _ = require('lodash');
const xmlbuilder = require('xmlbuilder');
const constants = require('../../constants.js');

class Font {
    constructor(opts) {
        opts = opts ? opts : {};

        typeof opts.color === 'string' ? this.color = utils.cleanColor(opts.color) : null;
        typeof opts.name === 'string' ? this.name = opts.name : null;
        typeof opts.scheme === 'string' ? this.scheme = opts.scheme : null;
        typeof opts.size === 'number' ? this.size = opts.size : null;
        
        if (opts.family !== undefined) {
            if (constants.fontFamilies.indexOf(opts.family.toLowerCase()) >= 0) {
                this.family = opts.family;
            } else {
                throw new TypeError('Font family must be one of ' + constants.fontFamilies.join(', '));
            }
        }

        typeof opts.vertAlign === 'string' ? this.vertAlign = opts.vertAlign : null;
        typeof opts.charset === 'number' ? this.charset = opts.charset : null;

        opts.condense === true ? this.condense = true : null;
        opts.extend === true ? this.extend = true : null;
        opts.bold === true ? this.bold = true : null;
        opts.italics === true ? this.italics = true : null;
        opts.outline === true ? this.outline = true : null;
        opts.shadow === true ? this.shadow = true : null;
        opts.strike === true ? this.strike = true : null;
        opts.underline === true ? this.underline = true : null;
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

        this.condense === true ? obj.condense = true : null;
        this.extend === true ? obj.extend = true : null;
        this.bold === true ? obj.bold = true : null;
        this.italics === true ? obj.italics = true : null;
        this.outline === true ? obj.outline = true : null;
        this.shadow === true ? obj.shadow = true : null;
        this.strike === true ? obj.strike = true : null;
        this.underline === true ? obj.underline = true : null;

        return obj;
    }

    addToXMLele(fontXML) {
        let fEle = fontXML.ele('font');
        fEle.ele('sz').att('val', this.size !== undefined ? this.size : constants.defaultFont.size);
        fEle.ele('color').att('rgb', this.color !== undefined ? this.color : constants.defaultFont.color);
        fEle.ele('name').att('val', this.name !== undefined ? this.name : constants.defaultFont.name);
        if (this.family !== undefined) {
            fEle.ele('family').att('val', constants.fontFamilies.indexOf(this.family.toLowerCase()));
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
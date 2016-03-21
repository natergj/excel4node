const utils = require('../../utils.js');
const _ = require('lodash');
const xmlbuilder = require('xmlbuilder');


const defaultFont = {
    'color': utils.cleanColor('Black'),
    'family': '2',
    'name': 'Calibri',
    'scheme': 'minor',
    'sz': '12'
};

class Font {
    constructor(opts) {
        opts = opts ? opts : {};

        typeof opts.color === 'string' ? this.color = utils.cleanColor(opts.color) : this.color = defaultFont.color;
        typeof opts.family === 'number' ? this.family = opts.family : this.family = defaultFont.family;
        typeof opts.name === 'string' ? this.name = opts.name : this.name = defaultFont.name;
        typeof opts.scheme === 'string' ? this.scheme = opts.scheme : this.scheme = defaultFont.scheme;
        typeof opts.size === 'number' ? this.sz = opts.size : this.sz = defaultFont.sz;
        
        typeof opts.vertAlign === 'string' ? this.vertAlign = opts.vertAlign : null;
        typeof opts.charset === 'number' ? this.charset = opts.charset : null;

        opts.condense === true ? this.condense = true : null;
        opts.extend === true ? this.extend = true : null;
        opts.bold === true ? this.b = true : null;
        opts.italics === true ? this.i = true : null;
        opts.outline === true ? this.outline = true : null;
        opts.shadow === true ? this.shadow = true : null;
        opts.strike === true ? this.strike = true : null;
        opts.underline === true ? this.u = true : null;
    }

    toObject() {
        let obj = {};

        typeof this.charset === 'number' ? obj.charset = this.charset : null;
        obj.color = this.color;
        obj.family = this.family;
        obj.name = this.name;
        obj.scheme = this.scheme;
        obj.sz = this.sz;

        this.condense === true ? obj.condense = true : null;
        this.extend === true ? obj.extend = true : null;
        this.bold === true ? obj.b = true : null;
        this.italics === true ? obj.i = true : null;
        this.outline === true ? obj.outline = true : null;
        this.shadow === true ? obj.shadow = true : null;
        this.strike === true ? obj.strike = true : null;
        this.underline === true ? obj.u = true : null;
        this.vertAlign === true ? obj.vertAlign = true : null;

        return obj;
    }

    addToXMLele(fontXML) {
        let fEle = fontXML.ele('font');
        fEle.ele('sz').att('val', this.sz);
        fEle.ele('color').att('rgb', this.color);
        fEle.ele('name').att('val', this.name);
        fEle.ele('family').att('val', this.family);
        fEle.ele('scheme').att('val', this.scheme);

        this.condense === true ? fEle.ele('condense') : null;
        this.extend === true ? fEle.ele('extend') : null;
        this.b === true ? fEle.ele('b') : null;
        this.i === true ? fEle.ele('i') : null;
        this.outline === true ? fEle.ele('outline') : null;
        this.shadow === true ? fEle.ele('shadow') : null;
        this.strike === true ? fEle.ele('strike') : null;
        this.u === true ? fEle.ele('u') : null;
        this.vertAlign === true ? fEle.ele('vertAlign') : null;

        return true;
    }
}

module.exports = Font;
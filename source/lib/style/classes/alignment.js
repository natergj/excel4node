const utils = require('../../utils.js');
const types = require('../../types/index.js');
const _ = require('lodash');
const xmlbuilder = require('xmlbuilder');

class Alignment { // §18.8.1 alignment (Alignment)
    constructor(opts) {

        if (opts.horizontal !== undefined) {
            this.horizontal = types.alignment.horizontal.validate(opts.horizontal) === true ? opts.horizontal : null;
        }
        
        if (opts.vertical !== undefined) {
            this.vertical = types.alignment.vertical.validate(opts.vertical) === true ? opts.vertical : null;
        }
        
        if (opts.readingOrder !== undefined) {
            this.readingOrder = types.alignment.readingOrder.validate(opts.readingOrder) === true ? opts.readingOrder : null;
        }
        
        if (opts.indent !== undefined) {
            if (typeof opts.indent === 'number' && parseInt(opts.indent) === opts.indent && opts.indent > 0) {
                this.indent = opts.indent;
            } else {
                throw new TypeError('alignment indent must be a positive integer.');
            }
        }
        
        if (opts.justifyLastLine !== undefined) {
            if (opts.justifyLastLine === true) {
                this.justifyLastLine = opts.justifyLastLine;
            } else if (typeof opts.justifyLastLine !== 'boolean') {
                throw new TypeError('justifyLastLine alignment option must be of type boolean');
            }
        }
        
        if (opts.relativeIndent !== undefined) {
            if (typeof opts.relativeIndent === 'number' && parseInt(opts.relativeIndent) === opts.relativeIndent && opts.relativeIndent > 0) {
                this.relativeIndent = opts.relativeIndent;
            } else {
                throw new TypeError('alignment indent must be a positive integer.');
            }
        }
        
        if (opts.shrinkToFit !== undefined) {
            if (opts.shrinkToFit === true) {
                this.shrinkToFit = opts.shrinkToFit;
            } else if (typeof opts.shrinkToFit !== 'boolean') {
                throw new TypeError('justifyLastLine alignment option must be of type boolean');
            }
        }
        
        if (opts.textRotation !== undefined) {
            if (typeof opts.textRotation === 'number' && parseInt(opts.textRotation) === opts.textRotation) {
                this.textRotation = opts.textRotation;
            } else if (opts.textRotation !== undefined) {
                throw new TypeError('alignment indent must be an integer.');
            }
        }
        
        if (opts.wrapText !== undefined) {
            if (opts.wrapText === true) {
                this.wrapText = opts.wrapText;
            } else if (typeof opts.wrapText !== 'boolean') {
                throw new TypeError('justifyLastLine alignment option must be of type boolean');
            }
        }
    }

    toObject() {
        let obj = {};

        this.horizontal !== undefined ? obj.horizontal = this.horizontal : null;
        this.indent !== undefined ? obj.indent = this.indent : null;
        this.justifyLastLine !== undefined ? obj.justifyLastLine = this.justifyLastLine : null;
        this.readingOrder !== undefined ? obj.readingOrder = this.readingOrder : null;
        this.relativeIndent !== undefined ? obj.relativeIndent = this.relativeIndent : null;
        this.shrinkToFit !== undefined ? obj.shrinkToFit = this.shrinkToFit : null;
        this.textRotation !== undefined ? obj.textRotation = this.textRotation : null;
        this.vertical !== undefined ? obj.vertical = this.vertical : null;
        this.wrapText !== undefined ? obj.wrapText = this.wrapText : null;

        return obj;
    }    

    addToXMLele(ele) {
        let thisEle = ele.ele('alignment');
        this.horizontal !== undefined ? thisEle.att('horizontal', this.horizontal) : null;
        this.indent !== undefined ? thisEle.att('indent', this.indent) : null;
        this.justifyLastLine === true ? thisEle.att('justifyLastLine', 1) : null;
        this.readingOrder !== undefined ? thisEle.att('readingOrder', this.readingOrder) : null;
        this.relativeIndent !== undefined ? thisEle.att('relativeIndent', this.relativeIndent) : null;
        this.shrinkToFit === true ? thisEle.att('shrinkToFit', 1) : null;
        this.textRotation !== undefined ? thisEle.att('textRotation', this.textRotation) : null;
        this.vertical !== undefined ? thisEle.att('vertical', this.vertical) : null;
        this.wrapText === true ? thisEle.att('wrapText', 1) : null;
    }
}

module.exports = Alignment;
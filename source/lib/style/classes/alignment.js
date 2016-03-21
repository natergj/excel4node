const utils = require('../../utils.js');
const constants = require('../../constants.js');
const _ = require('lodash');
const xmlbuilder = require('xmlbuilder');

class Alignment { // §18.8.1 alignment (Alignment)
    constructor(opts) {

        if(opts.horizontal !== undefined){
            if (constants.alignmentTypes.horizontal.indexOf(opts.horizontal) >= 0) {
                this.horizontal = opts.horizontal;
            } else if (opts.horizontal !== undefined) {
                throw new TypeError(`Horizontal alignment must be one of ${constants.alignmentTypes.horizontal.join(', ')}`);
            }
        }
        
        if(opts.indent !== undefined){
            if (typeof opts.indent === 'number' && parseInt(opts.indent) === opts.indent && opts.indent > 0) {
                this.indent = opts.indent;
            } else {
                throw new TypeError('alignment indent must be a positive integer.');
            }
        }
        
        if(opts.justifyLastLine !== undefined){
            if (opts.justifyLastLine === true) {
                this.justifyLastLine = opts.justifyLastLine;
            } else if(typeof opts.justifyLastLine !== 'boolean') {
                throw new TypeError('justifyLastLine alignment option must be of type boolean');
            }
        }
        
        if(opts.readingOrder !== undefined){
            if (constants.readingOrders.indexOf(opts.readingOrder) >= 0) {
                this.readingOrder = opts.readingOrder;
            } else {
                throw new TypeError('alignment readingOrder must be one of ' + constants.readingOrders.join(', ') );
            }
        }
        
        if(opts.relativeIndent !== undefined){
            if (typeof opts.relativeIndent === 'number' && parseInt(opts.relativeIndent) === opts.relativeIndent && opts.relativeIndent > 0) {
                this.relativeIndent = opts.relativeIndent;
            } else {
                throw new TypeError('alignment indent must be a positive integer.')
            }
        }
        
        if(opts.shrinkToFit !== undefined){
            if (opts.shrinkToFit === true) {
                this.shrinkToFit = opts.shrinkToFit;
            } else if(typeof opts.shrinkToFit !== 'boolean') {
                throw new TypeError('justifyLastLine alignment option must be of type boolean');
            }
        }
        
        if(opts.textRotation !== undefined){
            if (typeof opts.textRotation === 'number' && parseInt(opts.textRotation) === opts.textRotation) {
                this.textRotation = opts.textRotation;
            } else if(opts.textRotation !== undefined) {
                throw new TypeError('alignment indent must be an integer.');
            }
        }
        
        if(opts.vertical !== undefined){
            if (constants.alignmentTypes.vertical.indexOf(opts.vertical) >= 0) {
                this.vertical = opts.vertical;
            } else if(opts.vertical !== undefined) {
                throw new TypeError(`Vertical alignment must be one of ${constants.alignmentTypes.vertical.join(', ')}`);
            }
        }
        
        if(opts.wrapText !== undefined){
            if (opts.wrapText === true) {
                this.wrapText = opts.wrapText;
            } else if(typeof opts.wrapText !== 'boolean') {
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
        this.justifyLastLine !== undefined ? thisEle.att('justifyLastLine', this.justifyLastLine) : null;
        this.readingOrder !== undefined ? thisEle.att('readingOrder', this.readingOrder) : null;
        this.relativeIndent !== undefined ? thisEle.att('relativeIndent', this.relativeIndent) : null;
        this.shrinkToFit !== undefined ? thisEle.att('shrinkToFit', this.shrinkToFit) : null;
        this.textRotation !== undefined ? thisEle.att('textRotation', this.textRotation) : null;
        this.vertical !== undefined ? thisEle.att('vertical', this.vertical) : null;
        this.wrapText !== undefined ? thisEle.att('wrapText', this.wrapText) : null;
    }
}

module.exports = Alignment;
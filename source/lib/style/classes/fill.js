const utils = require('../../utils.js');
const types = require('../../types/index.js');
const _ = require('lodash');
const xmlbuilder = require('xmlbuilder');
const CTColor = require('./ctColor.js');

class Stop { //§18.8.38
    /** 
     * @class Stop
     * @desc Stops for Gradient fills
     * @param {Object} opts Options for Stop
     * @param {String} opts.color Color of Stop
     * @param {Number} opts.position Order of Stop with first stop being 0
     * @returns {Stop}
     */
    constructor(opts, position) {
        this.color = new CTColor(opts.color);
        this.position = position;
    }

    /** 
     * @func Stop.toObject
     * @desc Converts the Stop instance to a javascript object
     * @returns {Object}
     */
    toObject() {
        let obj = {};
        this.color !== undefined ? obj.color = this.color.toObject() : null;
        this.position !== undefined ? obj.position = this.position : null;
        return obj;
    }
}

class Fill { //§18.8.20 fill (Fill)

    /** 
     * @class Fill
     * @desc Excel Fill 
     * @param {Object} opts
     * @param {String} opts.type Type of Excel fill (gradient or pattern)
     * @param {Number} opts.bottom If Gradient fill, the position of the bottom edge of the inner rectange as a percentage in decimal form. (must be between 0 and 1)
     * @param {Number} opts.top If Gradient fill, the position of the top edge of the inner rectange as a percentage in decimal form. (must be between 0 and 1)
     * @param {Number} opts.left If Gradient fill, the position of the left edge of the inner rectange as a percentage in decimal form. (must be between 0 and 1)
     * @param {Number} opts.right If Gradient fill, the position of the right edge of the inner rectange as a percentage in decimal form. (must be between 0 and 1)
     * @param {Number} opts.degree Angle of the Gradient
     * @param {Array.Stop} opts.stops Array of position stops for gradient
     * @returns {Fill}
     */
    constructor(opts) {

        if (['gradient', 'pattern', 'none'].indexOf(opts.type) >= 0) {
            this.type = opts.type;
        } else {
            throw new TypeError('Fill type must be one of gradient, pattern or none.');
        }

        switch (this.type) {
        case 'gradient': //§18.8.24
            if (opts.bottom !== undefined) {
                if (opts.bottom < 0 || opts.bottom > 1) {
                    throw new TypeError('Values for gradient fill bottom attribute must be a decimal between 0 and 1');
                } else {
                    this.bottom = opts.bottom;
                }
            }

            if (opts.degree !== undefined) {
                if (typeof opts.degree === 'number') {
                    this.degree = opts.degree;
                } else {
                    throw new TypeError('Values of gradient fill degree must be of type number.');
                }
            }


            if (opts.left !== undefined) {
                if (opts.left < 0 || opts.left > 1) {
                    throw new TypeError('Values for gradient fill left attribute must be a decimal between 0 and 1');
                } else {
                    this.left = opts.left;
                }
            }

            if (opts.right !== undefined) {
                if (opts.right < 0 || opts.right > 1) {
                    throw new TypeError('Values for gradient fill right attribute must be a decimal between 0 and 1');
                } else {
                    this.right = opts.right;
                }
            }

            if (opts.top !== undefined) {
                if (opts.top < 0 || opts.top > 1) {
                    throw new TypeError('Values for gradient fill top attribute must be a decimal between 0 and 1');
                } else {
                    this.top = opts.top;
                }
            }

            if (opts.stops !== undefined) {
                if (opts.stops instanceof Array) {
                    opts.stops.forEach((s, i) => {
                        this.stops.push(new Stop(s, i));
                    });
                } else {
                    throw new TypeError('Stops for gradient fills must be sent as an Array');
                }
            }

            break;

        case 'pattern': //§18.8.32
            if (opts.bgColor !== undefined) {
                this.bgColor = new CTColor(opts.bgColor);
            }

            if (opts.fgColor !== undefined) {
                this.fgColor = new CTColor(opts.fgColor);
            }

            if (opts.patternType !== undefined) {
                types.fillPattern.validate(opts.patternType) === true ? this.patternType = opts.patternType : null; 
            }
            break;

        case 'none':
            this.patternType = 'none';
            break;
        }
    }

    /** 
     * @func Fill.toObject
     * @desc Converts the Fill instance to a javascript object
     * @returns {Object}
     */
    toObject() {
        let obj = {};

        this.type !== undefined ? obj.type = this.type : null;
        this.bottom !== undefined ? obj.bottom = this.bottom : null;
        this.degree !== undefined ? obj.degree = this.degree : null;
        this.left !== undefined ? obj.left = this.left : null;
        this.right !== undefined ? obj.right = this.right : null;
        this.top !== undefined ? obj.top = this.top : null;
        this.bgColor !== undefined ? obj.bgColor = this.bgColor.toObject() : null;
        this.fgColor !== undefined ? obj.fgColor = this.fgColor.toObject() : null;
        this.patternType !== undefined ? obj.patternType = this.patternType : null;

        if (this.stops !== undefined) {
            obj.stop = [];
            this.stops.forEach((s) => {
                obj.stops.push(s.toObject());
            });
        }

        return obj;
    }

    /**
     * @alias Fill.addToXMLele
     * @desc When generating Workbook output, attaches style to the styles xml file
     * @func Fill.addToXMLele
     * @param {xmlbuilder.Element} ele Element object of the xmlbuilder module
     */
    addToXMLele(fXML) {
        let pFill = fXML.ele('patternFill').att('patternType', this.patternType);

        if (this.fgColor instanceof CTColor) {
            pFill.ele('fgColor').att(this.fgColor.type, this.fgColor[this.fgColor.type]);
        }

        if (this.bgColor instanceof CTColor) {
            pFill.ele('bgColor').att(this.bgColor.type, this.bgColor[this.bgColor.type]);
        }
    }
}

module.exports = Fill;
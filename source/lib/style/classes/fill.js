const utils = require('../../utils.js');
const types = require('../../types/index.js');
const _ = require('lodash');
const xmlbuilder = require('xmlbuilder');
const CTColor = require('./ctColor.js');

class Stop { //§18.8.38
    constructor(opts, position) {
        this.color = new CTColor(opts.color);
        this.position = position;
    }

    toObject() {
        let obj = {};
        this.color !== undefined ? obj.color = this.color.toObject() : null;
        this.position !== undefined ? obj.position = this.position : null;
        return obj;
    }
}

class Fill { //§18.8.20 fill (Fill)

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
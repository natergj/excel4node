const utils = require('../../utils.js');
const types = require('../../types/index.js');
const _ = require('lodash');
const xmlbuilder = require('xmlbuilder');
const CTColor = require('./ctColor.js');

class BorderOrdinal {
    constructor(opts) {
        opts = opts ? opts : {};
        if (opts.color !== undefined) {
            this.color = new CTColor(opts.color);
        }
        if (opts.style !== undefined) {
            if (types.borderStyles.indexOf(opts.style) < 0) {
                throw new TypeError(`Border style must be one of the following: ${types.borderStyles.join(', ')}`);
            } else {
                this.style = opts.style;
            }
        }
    }

    toObject() {
        let obj = {};
        if (this.color !== undefined) {
            obj.color = this.color.toObject();
        }
        if (this.style !== undefined) {
            obj.style = this.style;
        }
        return obj;
    }
}

class Border {
    constructor(opts) {
        opts = opts ? opts : {};
        this.left;
        this.right;
        this.top;
        this.bottom;
        this.diagonal;
        this.outline;
        this.diagonalDown;
        this.diagonalUp;

        Object.keys(opts).forEach((opt) => {
            if (['outline', 'diagonalDown', 'diagonalUp'].indexOf(opt) >= 0) {
                if (typeof opts[opt] === 'boolean') {
                    this[opt] = opts[opt];
                } else {
                    throw new TypeError('Border outline option must be of type Boolean');
                }
            } else if (['left', 'right', 'top', 'bottom', 'diagonal'].indexOf(opt) < 0) {  //TODO: move logic to types folder
                throw new TypeError(`Invalid key for border declaration ${opt}. Must be one of left, right, top, bottom, diagonal`);
            } else {
                this[opt] = new BorderOrdinal(opts[opt]);
            }
        });
    }

    toObject() {
        let obj = {};
        obj.left;
        obj.right;
        obj.top;
        obj.bottom;
        obj.diagonal;

        if (this.left !== undefined) {
            obj.left = this.left.toObject();
        }
        if (this.right !== undefined) {
            obj.right = this.right.toObject();
        }
        if (this.top !== undefined) {
            obj.top = this.top.toObject();
        }
        if (this.bottom !== undefined) {
            obj.bottom = this.bottom.toObject();
        }
        if (this.diagonal !== undefined) {
            obj.diagonal = this.diagonal.toObject();
        }
        this.outline !== undefined ? obj.outline = this.outline : null;
        this.diagonalDown !== undefined ? obj.diagonalDown = this.diagonalDown : null;
        this.diagonalUp !== undefined ? obj.diagonalUp = this.diagonalUp : null;

        return obj;
    }

    addToXMLele(borderXML) {
        let bXML = borderXML.ele('border');
        if (this.outline !== undefined) {
            bXML.att('outline', '1');
        }
        if (this.diagonalUp !== undefined) {
            bXML.att('diagonalUp', '1');
        }
        if (this.diagonalDown !== undefined) {
            bXML.att('diagonalDown', '1');
        }

        ['left', 'right', 'top', 'bottom', 'diagonal'].forEach((ord) => {
            let thisOEle = bXML.ele(ord);
            if (this[ord] !== undefined) {
                if (this[ord].style !== undefined) {
                    thisOEle.att('style', this[ord].style);
                }
                if (this[ord].color instanceof CTColor) {
                    this[ord].color.addToXMLele(thisOEle);
                }
            }
        });
    }
}

module.exports = Border;
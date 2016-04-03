const utils = require('../../utils.js');
const types = require('../../types/index.js');
const _ = require('lodash');
const xmlbuilder = require('xmlbuilder');

class CTColor { //§18.8.3 && §18.8.19
    constructor(color) {
        this.type;
        this.rgb;
        this.theme; //§20.1.6.2 clrScheme (Color Scheme) : types.colorSchemes

        if (typeof color === 'string') {
            if (types.colorScheme[color.toLowerCase()] !== undefined) {
                this.theme = color;
                this.type = 'theme';
            } else {
                try {
                    this.rgb = types.excelColor.getColor(color);
                    this.type = 'rgb';
                } catch (e) {
                    throw new TypeError(`Fill color must be an RGB value, Excel color (${types.excelColor.opts.join(', ')}) or Excel theme (${types.colorScheme.opts.join(', ')})`);
                }
            }
        }
    }

    toObject() {
        return this[this.type];
    }

    addToXMLele(ele) {
        let colorEle = ele.ele('color');
        colorEle.att(this.type, this[this.type]);
    }
}

module.exports = CTColor;
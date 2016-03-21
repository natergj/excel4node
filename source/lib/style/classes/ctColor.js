const utils = require('../../utils.js');
const constants = require('../../constants.js');
const _ = require('lodash');
const xmlbuilder = require('xmlbuilder');

class CTColor { //§18.8.3 && §18.8.19
    constructor(color) {
        this.type;
        this.rgb;
        this.theme; //§20.1.6.2 clrScheme (Color Scheme) : constants.colorSchemes

        if (typeof color === 'string') {
            if (constants.colorSchemes.indexOf(color.toLowerCase()) >= 0) {
                this.theme = constants.colorSchemes.indexOf(color);
                this.type = 'theme';
            } else {
                try {
                    this.rgb = utils.cleanColor(color);
                    this.type = 'rgb';
                } catch (e) {
                    throw new TypeError(`Fill color must be an RGB value, Excel color (${Object.keys(constants.excelColors).join(', ')}) or Excel theme (${constants.colorSchemes.join(', ')})`);
                }
            }
        }
    }

    toObject() {
        let obj = {};
        this.rgb !== undefined ? obj.rgb = this.rgb : null;
        this.theme !== undefined ? obj.theme = this.theme : null;
        this.type !== undefined ? obj.type = this.type : null;
        return obj;
    }

    addToXMLele(ele) {
        let colorEle = ele.ele('color');
        colorEle.att(this.type, this[this.type]);
    }
}

module.exports = CTColor;
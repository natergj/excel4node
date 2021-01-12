'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var types = require('../../types/index.js');
var xmlbuilder = require('xmlbuilder');

var CTColor = function () {
    //§18.8.3 && §18.8.19
    /** 
     * @class CTColor
     * @desc Excel color representation
     * @param {String} color Excel Color scheme or Excel Color name or HEX value of Color
     * @properties {String} type Type of color object. defaults to rgb
     * @properties {String} rgb ARGB representation of Color
     * @properties {String} theme Excel Color Scheme
     * @returns {CTColor}
     */
    function CTColor(color) {
        _classCallCheck(this, CTColor);

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
                    throw new TypeError('Fill color must be an RGB value, Excel color (' + types.excelColor.opts.join(', ') + ') or Excel theme (' + types.colorScheme.opts.join(', ') + ')');
                }
            }
        }
    }

    /** 
     * @func CTColor.toObject
     * @desc Converts the CTColor instance to a javascript object
     * @returns {Object}
     */


    _createClass(CTColor, [{
        key: 'toObject',
        value: function toObject() {
            return this[this.type];
        }

        /**
         * @alias CTColor.addToXMLele
         * @desc When generating Workbook output, attaches style to the styles xml file
         * @func CTColor.addToXMLele
         * @param {xmlbuilder.Element} ele Element object of the xmlbuilder module
         */

    }, {
        key: 'addToXMLele',
        value: function addToXMLele(ele) {
            var colorEle = ele.ele('color');
            colorEle.att(this.type, this[this.type]);
        }
    }]);

    return CTColor;
}();

module.exports = CTColor;
//# sourceMappingURL=ctColor.js.map
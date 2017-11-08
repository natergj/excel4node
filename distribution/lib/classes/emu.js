'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var EMU = function () {

    /** 
     * The EMU was created in order to be able to evenly divide in both English and Metric units
     * @class EMU
     * @param {String} Number of EMUs or string representation of length in mm, cm or in. i.e. '10.5mm'
     * @property {Number} value Number of EMUs
     * @returns {EMU} Number of EMUs 
     */
    function EMU(val) {
        _classCallCheck(this, EMU);

        this._value;
        this.value = val;
    }

    _createClass(EMU, [{
        key: 'toInt',


        /**
         * @alias EMU.toInt
         * @desc Returns the number of EMUs as integer
         * @func EMU.toInt
         * @returns {Number} Number of EMUs
         */
        value: function toInt() {
            return this._value;
        }

        /**
         * @alias EMU.toInch
         * @desc Returns the number of Inches for the EMUs
         * @func EMU.toInch
         * @returns {Number} Number of Inches for the EMUs
         */

    }, {
        key: 'toInch',
        value: function toInch() {
            return this._value / 914400;
        }

        /**
         * @alias EMU.toCM
         * @desc Returns the number of Centimeters for the EMUs
         * @func EMU.toCM
         * @returns {Number} Number of Centimeters for the EMUs
         */

    }, {
        key: 'toCM',
        value: function toCM() {
            return this._value / 360000;
        }
    }, {
        key: 'value',
        get: function get() {
            return this._value;
        },
        set: function set(val) {
            if (val === undefined) {
                this._value = 0;
            } else if (typeof val === 'number') {
                this._value = val ? parseInt(val) : 0;
            } else if (typeof val === 'string') {
                var re = new RegExp('[0-9]+(\.[0-9]+)?(mm|cm|in)');
                if (re.test(val) === true) {
                    var measure = parseFloat(/[0-9]+(\.[0-9]+)?/.exec(val)[0]);
                    var unit = /(mm|cm|in)/.exec(val)[0];

                    switch (unit) {
                        case 'mm':
                            this._value = parseInt(measure * 36000);
                            break;

                        case 'cm':
                            this._value = parseInt(measure * 360000);
                            break;

                        case 'in':
                            this._value = parseInt(measure * 914400);
                            break;
                    }
                } else {
                    throw new TypeError('EMUs must be specified as whole integer EMUs or Floats immediately followed by unit of measure in cm, mm, or in. i.e. "1.5in"');
                }
            }
        }
    }]);

    return EMU;
}();

module.exports = EMU;

/*
M.4.1.1 EMU Unit of Measurement

1 emu  = 1/914400 in = 1/360000 cm

Throughout ECMA-376, the EMU is used as a unit of measurement for length. An EMU is defined as follows:
The EMU was created in order to be able to evenly divide in both English and Metric units, in order to 
avoid rounding errors during the calculation. The usage of EMUs also facilitates a more seamless system 
switch and interoperability between different locales utilizing different units of measurement. 
EMUs define an integer based, high precision coordinate system.
*/
//# sourceMappingURL=emu.js.map
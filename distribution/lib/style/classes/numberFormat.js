'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var NumberFormat = function () {
    /**
    * @class NumberFormat
    * @param {String} fmt Format of the Number
    * @returns {NumberFormat}
    */
    function NumberFormat(fmt) {
        _classCallCheck(this, NumberFormat);

        this.formatCode = fmt;
        this.id;
    }

    _createClass(NumberFormat, [{
        key: 'addToXMLele',


        /**
         * @alias NumberFormat.addToXMLele
         * @desc When generating Workbook output, attaches style to the styles xml file
         * @func NumberFormat.addToXMLele
         * @param {xmlbuilder.Element} ele Element object of the xmlbuilder module
         */
        value: function addToXMLele(ele) {
            if (this.formatCode !== undefined) {
                ele.ele('numFmt').att('formatCode', this.formatCode).att('numFmtId', this.numFmtId);
            }
        }
    }, {
        key: 'numFmtId',
        get: function get() {
            return this.id;
        },
        set: function set(id) {
            this.id = id;
        }
    }]);

    return NumberFormat;
}();

module.exports = NumberFormat;
//# sourceMappingURL=numberFormat.js.map
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var types = require('../../types/index.js');
var xmlbuilder = require('xmlbuilder');
var CTColor = require('./ctColor.js');

var BorderOrdinal = function () {
    function BorderOrdinal(opts) {
        _classCallCheck(this, BorderOrdinal);

        opts = opts ? opts : {};
        if (opts.color !== undefined) {
            this.color = new CTColor(opts.color);
        }
        if (opts.style !== undefined) {
            this.style = types.borderStyle.validate(opts.style) === true ? opts.style : null;
        }
    }

    _createClass(BorderOrdinal, [{
        key: 'toObject',
        value: function toObject() {
            var obj = {};
            if (this.color !== undefined) {
                obj.color = this.color.toObject();
            }
            if (this.style !== undefined) {
                obj.style = this.style;
            }
            return obj;
        }
    }]);

    return BorderOrdinal;
}();

var Border = function () {
    /** 
     * @class Border
     * @desc Border object for Style
     * @param {Object} opts Options for Border object
     * @param {Object} opts.left Options for left side of Border
     * @param {String} opts.left.color HEX represenation of color
     * @param {String} opts.left.style Border style
     * @param {Object} opts.right Options for right side of Border
     * @param {String} opts.right.color HEX represenation of color
     * @param {String} opts.right.style Border style
     * @param {Object} opts.top Options for top side of Border
     * @param {String} opts.top.color HEX represenation of color
     * @param {String} opts.top.style Border style
     * @param {Object} opts.bottom Options for bottom side of Border
     * @param {String} opts.bottom.color HEX represenation of color
     * @param {String} opts.bottom.style Border style
     * @param {Object} opts.diagonal Options for diagonal side of Border
     * @param {String} opts.diagonal.color HEX represenation of color
     * @param {String} opts.diagonal.style Border style
     * @param {Boolean} opts.outline States whether borders should be applied only to the outside borders of a cell range
     * @param {Boolean} opts.diagonalDown States whether diagonal border should go from top left to bottom right
     * @param {Boolean} opts.diagonalUp States whether diagonal border should go from bottom left to top right
     * @returns {Border}
     */
    function Border(opts) {
        var _this = this;

        _classCallCheck(this, Border);

        opts = opts ? opts : {};
        this.left;
        this.right;
        this.top;
        this.bottom;
        this.diagonal;
        this.outline;
        this.diagonalDown;
        this.diagonalUp;

        Object.keys(opts).forEach(function (opt) {
            if (['outline', 'diagonalDown', 'diagonalUp'].indexOf(opt) >= 0) {
                if (typeof opts[opt] === 'boolean') {
                    _this[opt] = opts[opt];
                } else {
                    throw new TypeError('Border outline option must be of type Boolean');
                }
            } else if (['left', 'right', 'top', 'bottom', 'diagonal'].indexOf(opt) < 0) {
                //TODO: move logic to types folder
                throw new TypeError('Invalid key for border declaration ' + opt + '. Must be one of left, right, top, bottom, diagonal');
            } else {
                _this[opt] = new BorderOrdinal(opts[opt]);
            }
        });
    }

    /** 
     * @func Border.toObject
     * @desc Converts the Border instance to a javascript object
     * @returns {Object}
     */


    _createClass(Border, [{
        key: 'toObject',
        value: function toObject() {
            var obj = {};
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
            typeof this.outline === 'boolean' ? obj.outline = this.outline : null;
            typeof this.diagonalDown === 'boolean' ? obj.diagonalDown = this.diagonalDown : null;
            typeof this.diagonalUp === 'boolean' ? obj.diagonalUp = this.diagonalUp : null;

            return obj;
        }

        /**
         * @alias Border.addToXMLele
         * @desc When generating Workbook output, attaches style to the styles xml file
         * @func Border.addToXMLele
         * @param {xmlbuilder.Element} ele Element object of the xmlbuilder module
         */

    }, {
        key: 'addToXMLele',
        value: function addToXMLele(borderXML) {
            var _this2 = this;

            var bXML = borderXML.ele('border');
            if (this.outline === true) {
                bXML.att('outline', '1');
            }
            if (this.diagonalUp === true) {
                bXML.att('diagonalUp', '1');
            }
            if (this.diagonalDown === true) {
                bXML.att('diagonalDown', '1');
            }

            ['left', 'right', 'top', 'bottom', 'diagonal'].forEach(function (ord) {
                var thisOEle = bXML.ele(ord);
                if (_this2[ord] !== undefined) {
                    if (_this2[ord].style !== undefined) {
                        thisOEle.att('style', _this2[ord].style);
                    }
                    if (_this2[ord].color instanceof CTColor) {
                        _this2[ord].color.addToXMLele(thisOEle);
                    }
                }
            });
        }
    }]);

    return Border;
}();

module.exports = Border;
//# sourceMappingURL=border.js.map
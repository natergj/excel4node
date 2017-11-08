'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var Drawing = require('./drawing.js');
var Picture = require('./picture.js');

var DrawingCollection = function () {
    function DrawingCollection() {
        _classCallCheck(this, DrawingCollection);

        this.drawings = [];
    }

    _createClass(DrawingCollection, [{
        key: 'add',
        value: function add(opts) {
            switch (opts.type) {
                case 'picture':
                    var newPic = new Picture(opts);
                    this.drawings.push(newPic);
                    return newPic;

                default:
                    throw new TypeError('this option is not yet supported');
            }
        }
    }, {
        key: 'length',
        get: function get() {
            return this.drawings.length;
        }
    }, {
        key: 'isEmpty',
        get: function get() {
            if (this.drawings.length === 0) {
                return true;
            } else {
                return false;
            }
        }
    }]);

    return DrawingCollection;
}();

module.exports = { DrawingCollection: DrawingCollection, Drawing: Drawing, Picture: Picture };
//# sourceMappingURL=index.js.map
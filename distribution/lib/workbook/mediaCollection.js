'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var fs = require('fs');

var MediaCollection = function () {
    function MediaCollection() {
        _classCallCheck(this, MediaCollection);

        this.items = [];
    }

    _createClass(MediaCollection, [{
        key: 'add',
        value: function add(item) {
            if (typeof item === 'string') {
                fs.accessSync(item, fs.R_OK);
            }

            this.items.push(item);
            return this.items.length;
        }
    }, {
        key: 'isEmpty',
        get: function get() {
            if (this.items.length === 0) {
                return true;
            } else {
                return false;
            }
        }
    }]);

    return MediaCollection;
}();

module.exports = MediaCollection;
//# sourceMappingURL=mediaCollection.js.map
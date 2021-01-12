'use strict';

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var uuid = require('uuid/v4');
var utils = require('../utils');

// §18.7.3 Comment

var Comment = function Comment(ref, comment) {
    var options = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : {};

    _classCallCheck(this, Comment);

    this.ref = ref;
    this.comment = comment;
    this.uuid = '{' + uuid().toUpperCase() + '}';
    this.row = utils.getExcelRowCol(ref).row;
    this.col = utils.getExcelRowCol(ref).col;
    this.marginLeft = options.marginLeft || this.col * 88 + 8 + 'pt';
    this.marginTop = options.marginTop || (this.row - 1) * 16 + 8 + 'pt';
    this.width = options.width || '104pt';
    this.height = options.height || '69pt';
    this.position = options.position || 'absolute';
    this.zIndex = options.zIndex || '1';
    this.fillColor = options.fillColor || '#ffffe1';
    this.visibility = options.visibility || 'hidden';
};

module.exports = Comment;
//# sourceMappingURL=comment.js.map
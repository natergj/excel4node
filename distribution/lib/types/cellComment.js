'use strict';

//§18.18.5 ST_CellComments (Cell Comments)

function items() {
    var _this = this;

    this.opts = ['none', 'asDisplayed', 'atEnd'];
    this.opts.forEach(function (o, i) {
        _this[o] = i + 1;
    });
}

items.prototype.validate = function (val) {
    if (this[val] === undefined) {
        var opts = [];
        for (var name in this) {
            if (this.hasOwnProperty(name)) {
                opts.push(name);
            }
        }
        throw new TypeError('Invalid value for ST_CellComments; Value must be one of ' + this.opts.join(', '));
    } else {
        return true;
    }
};

module.exports = new items();
//# sourceMappingURL=cellComment.js.map
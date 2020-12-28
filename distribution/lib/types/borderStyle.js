'use strict';

function items() {
    var _this = this;

    this.opts = [//§18.18.3 ST_BorderStyle (Border Line Styles)
    'none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot'];
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
        throw new TypeError('Invalid value for ST_BorderStyle; Value must be one of ' + this.opts.join(', '));
    } else {
        return true;
    }
};

module.exports = new items();
//# sourceMappingURL=borderStyle.js.map
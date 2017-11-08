'use strict';

function items() {
    var _this = this;

    this.opts = [//§18.8.18 family (Font Family)
    'n/a', 'roman', 'swiss', 'modern', 'script', 'decorative'];
    this.opts.forEach(function (o, i) {
        _this[o] = i;
    });
}

items.prototype.validate = function (val) {
    if (typeof val !== 'string') {
        throw new TypeError('Invalid value for Font Family ' + val + '; Value must be one of ' + this.opts.join(', '));
    }

    if (this[val.toLowerCase()] === undefined) {
        var opts = [];
        for (var name in this) {
            if (this.hasOwnProperty(name)) {
                opts.push(name);
            }
        }
        throw new TypeError('Invalid value for Font Family ' + val + '; Value must be one of ' + this.opts.join(', '));
    } else {
        return true;
    }
};

module.exports = new items();
//# sourceMappingURL=fontFamily.js.map
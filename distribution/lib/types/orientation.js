'use strict';

//§18.18.50 ST_Orientation (Orientation)

function items() {
    var _this = this;

    var opts = ['default', 'portrait', 'landscape'];
    opts.forEach(function (o, i) {
        _this[o] = i + 1;
    });
}

items.prototype.validate = function (val) {
    if (this[val.toLowerCase()] === undefined) {
        var opts = [];
        for (var name in this) {
            if (this.hasOwnProperty(name)) {
                opts.push(name);
            }
        }
        throw new TypeError('Invalid value for pageSetup.orientation; Value must be one of ' + opts.join(', '));
    } else {
        return true;
    }
};

module.exports = new items();
//# sourceMappingURL=orientation.js.map
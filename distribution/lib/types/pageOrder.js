'use strict';

//§18.18.51 ST_PageOrder (Page Order)

function items() {
    var _this = this;

    var opts = ['downThenOver', 'overThenDown'];
    opts.forEach(function (o, i) {
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
        throw new TypeError('Invalid value for pageSetup.pageOrder; Value must be one of ' + opts.join(', '));
    } else {
        return true;
    }
};

module.exports = new items();
//# sourceMappingURL=pageOrder.js.map
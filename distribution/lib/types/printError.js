'use strict';

//§18.18.60 ST_PrintError (Print Errors)
function items() {
    var _this = this;

    var opts = ['displayed', 'blank', 'dash', 'NA'];
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
        throw new TypeError('Invalid value for pageSetup.errors; Value must be one of ' + opts.join(', '));
    } else {
        return true;
    }
};

module.exports = new items();
//# sourceMappingURL=printError.js.map
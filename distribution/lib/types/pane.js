'use strict';

//§18.18.52 ST_Pane (Pane Types)

function items() {
    var _this = this;

    var opts = ['bottomLeft', 'bottomRight', 'topLeft', 'topRight'];
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
        throw new TypeError('Invalid value for sheetview.pane.activePane; Value must be one of ' + opts.join(', '));
    } else {
        return true;
    }
};

module.exports = new items();
//# sourceMappingURL=pane.js.map
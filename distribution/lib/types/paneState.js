'use strict';

//§ST_PaneState (Pane State)

function items() {
    var _this = this;

    var opts = ['split', 'frozen', 'frozenSplit'];
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
        throw new TypeError('Invalid value for sheetView.pane.state; Value must be one of ' + opts.join(', '));
    } else {
        return true;
    }
};

module.exports = new items();
//# sourceMappingURL=paneState.js.map
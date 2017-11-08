'use strict';

function horizontalAlignments() {
    var _this = this;

    this.opts = [// §18.18.40 ST_HorizontalAlignment (Horizontal Alignment Type)
    'center', 'centerContinuous', 'distributed', 'fill', 'general', 'justify', 'left', 'right'];
    this.opts.forEach(function (o, i) {
        _this[o] = i + 1;
    });
}

function verticalAlignments() {
    var _this2 = this;

    this.opts = [//§18.18.88 ST_VerticalAlignment (Vertical Alignment Types)
    'bottom', 'center', 'distributed', 'justify', 'top'];
    this.opts.forEach(function (o, i) {
        _this2[o] = i + 1;
    });
}

function readingOrders() {
    this['contextDependent'] = 0;
    this['leftToRight'] = 1;
    this['rightToLeft'] = 2;
    this.opts = ['contextDependent', 'leftToRight', 'rightToLeft'];
}

horizontalAlignments.prototype.validate = function (val) {
    if (this[val] === undefined) {
        var opts = [];
        for (var name in this) {
            if (this.hasOwnProperty(name)) {
                opts.push(name);
            }
        }
        throw new TypeError('Invalid value for alignment.horizontal ' + val + '; Value must be one of ' + this.opts.join(', '));
    } else {
        return true;
    }
};

verticalAlignments.prototype.validate = function (val) {
    if (this[val] === undefined) {
        var opts = [];
        for (var name in this) {
            if (this.hasOwnProperty(name)) {
                opts.push(name);
            }
        }
        throw new TypeError('Invalid value for alignment.vertical ' + val + '; Value must be one of ' + this.opts.join(', '));
    } else {
        return true;
    }
};

readingOrders.prototype.validate = function (val) {
    if (this[val] === undefined) {
        var opts = [];
        for (var name in this) {
            if (this.hasOwnProperty(name)) {
                opts.push(name);
            }
        }
        throw new TypeError('Invalid value for alignment.readingOrder ' + val + '; Value must be one of ' + this.opts.join(', '));
    } else {
        return true;
    }
};

module.exports.vertical = new verticalAlignments();
module.exports.horizontal = new horizontalAlignments();
module.exports.readingOrder = new readingOrders();
//# sourceMappingURL=alignment.js.map
function horizontalAlignments() {
    this.opts = [ // §18.18.40 ST_HorizontalAlignment (Horizontal Alignment Type)
        'center', 
        'centerContinuous', 
        'distributed', 
        'fill', 
        'general', 
        'justify', 
        'left', 
        'right'
    ];
    this.opts.forEach((o, i) => {
        this[o] = i + 1;
    });
}

function verticalAlignments() {
    this.opts = [ //§18.18.88 ST_VerticalAlignment (Vertical Alignment Types)
        'bottom', 
        'center', 
        'distributed', 
        'justify', 
        'top'
    ];
    this.opts.forEach((o, i) => {
        this[o] = i + 1;
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
        let opts = [];
        for (let name in this) {
            if (this.hasOwnProperty(name)) {
                opts.push(name);
            }
        }
        throw new TypeError(`Invalid value for alignment.horizontal ${val}; Value must be one of ${this.opts.join(', ')}`);
    } else {
        return true;
    }
};

verticalAlignments.prototype.validate = function (val) {
    if (this[val] === undefined) {
        let opts = [];
        for (let name in this) {
            if (this.hasOwnProperty(name)) {
                opts.push(name);
            }
        }
        throw new TypeError(`Invalid value for alignment.vertical ${val}; Value must be one of ${this.opts.join(', ')}`);
    } else {
        return true;
    }
};

readingOrders.prototype.validate = function (val) {
    if (this[val] === undefined) {
        let opts = [];
        for (let name in this) {
            if (this.hasOwnProperty(name)) {
                opts.push(name);
            }
        }
        throw new TypeError(`Invalid value for alignment.readingOrder ${val}; Value must be one of ${this.opts.join(', ')}`);
    } else {
        return true;
    }
};

module.exports.vertical = new verticalAlignments();
module.exports.horizontal = new horizontalAlignments();
module.exports.readingOrder = new readingOrders();
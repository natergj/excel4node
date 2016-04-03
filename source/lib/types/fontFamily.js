function items() {
    this.opts = [//§18.8.18 family (Font Family)
        'n/a', 
        'roman', 
        'swiss', 
        'modern', 
        'script', 
        'decorative'
    ];
    this.opts.forEach((o, i) => {
        this[o] = i;
    });
}


items.prototype.validate = function (val) {
    if (typeof val !== 'string') {
        throw new TypeError(`Invalid value for Font Family ${val}; Value must be one of ${this.opts.join(', ')}`);
    }

    if (this[val.toLowerCase()] === undefined) {
        let opts = [];
        for (let name in this) {
            if (this.hasOwnProperty(name)) {
                opts.push(name);
            }
        }
        throw new TypeError(`Invalid value for Font Family ${val}; Value must be one of ${this.opts.join(', ')}`);
    } else {
        return true;
    }
};

module.exports = new items();
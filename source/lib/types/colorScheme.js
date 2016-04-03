function items() {
    this.opts = [//§20.1.6.2 clrScheme (Color Scheme)
        'dark 1', 
        'light 1', 
        'dark 2', 
        'light 2', 
        'accent 1', 
        'accent 2', 
        'accent 3', 
        'accent 4', 
        'accent 5', 
        'accent 6', 
        'hyperlink', 
        'followed hyperlink'
    ];
    this.opts.forEach((o, i) => {
        this[o] = i + 1;
    });
}


items.prototype.validate = function (val) {
    if (this[val.toLowerCase()] === undefined) {
        let opts = [];
        for (let name in this) {
            if (this.hasOwnProperty(name)) {
                opts.push(name);
            }
        }
        throw new TypeError('Invalid value for clrScheme; Value must be one of ' + this.opts.join(', '));
    } else {
        return true;
    }
};

module.exports = new items();
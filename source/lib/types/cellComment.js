//§18.18.5 ST_CellComments (Cell Comments)

function items() {
    this.opts = ['none', 'asDisplayed', 'atEnd'];
    this.opts.forEach((o, i) => {
        this[o] = i + 1;
    });
}


items.prototype.validate = function (val) {
    if (this[val] === undefined) {
        let opts = [];
        for (let name in this) {
            if (this.hasOwnProperty(name)) {
                opts.push(name);
            }
        }
        throw new TypeError('Invalid value for ST_CellComments; Value must be one of ' + this.opts.join(', '));
    } else {
        return true;
    }
};

module.exports = new items();
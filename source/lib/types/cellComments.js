//§18.18.5 ST_CellComments (Cell Comments)

function items() {
    let opts = ['none', 'asDisplayed', 'atEnd'];
    opts.forEach((i) => {
        this[i] = i;
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
        throw new TypeError('Invalid value for pageSetup.cellComments; Value must be one of ' + opts.join(', '));
    } else {
        return true;
    }
};

module.exports = new items();
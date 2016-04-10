//§18.18.51 ST_PageOrder (Page Order)

function items() {
    let opts = ['downThenOver', 'overThenDown'];
    opts.forEach((o, i) => {
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
        throw new TypeError('Invalid value for pageSetup.pageOrder; Value must be one of ' + opts.join(', '));
    } else {
        return true;
    }
};

module.exports = new items();